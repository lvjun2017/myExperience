Option Explicit
' Model sheet中的列信息
'主题域(Pachage)
CONST CELL_A="A" 
'表注释
CONST CELL_B="B" 
'表英文名称
CONST CELL_C="C" 
'表中文名称
CONST CELL_D="D" 
'列名
CONST CELL_E="E" 
'列中文名称
CONST CELL_F="F" 
'列注释
CONST CELL_G="G" 
'数据类型
CONST CELL_H="H" 
'是否主键
CONST CELL_I="I" 
'是否可空
CONST CELL_J="J" 
'默认值
CONST CELL_K="K" 
'标识符
CONST str_iskey="Y"
'表的所属者
CONST str_username="srv"
'是否先删除表的所有列，如果是false则不会删除excel中没有的列，如果是true，则会重新创建相应表的所有列
CONST isclear_columns = true  

'定义当前的模型
DIM mdl 
'通过全局参数获得当前的模型
SET mdl = ActiveModel 

IF (mdl IS NOTHING) THEN
   MsgBox "没有选择模型，请选择一个模型并打开"
ELSEIF NOT mdl.IsKindOf(PdPDM.cls_Model) THEN
   MsgBox "当前选择的不是一个物理模型（PDM）."
ELSE

'选择需要导入的Excel文件
' 打开Excel
'定义Excel对象
DIM xlApp   
SET xlApp  = CreateObject("Excel.Application")
xlApp.DisplayAlerts = FALSE
'定义Excel Sheet
DIM xlBook  
SET xlBook = xlApp.WorkBooks.Open("D:\model\model_import.xlsx")
xlApp.Visible = TRUE

output "开始从Excel创建模型"
Create_From_Excel(xlBook)
output "模型创建完成，开始关闭Excel"

SET xlBook=NOTHING
xlApp.Quit
SET xlApp=NOTHING

END IF

PRIVATE SUB Create_From_Excel(xlBook)
  DIM xlsheet
  DIM rowcount
  dim pkg

  FOR EACH xlsheet IN xlBook.WORKSHEETS
	rowcount = xlsheet.UsedRange.Cells.Rows.Count
	output "本Excel["+xlsheet.name+"]共有行数为:"+CSTR(rowcount)
	IF rowcount>1 THEN
	  SET pkg = CreateOrReplacePackageByName( xlsheet.name , mdl)
	  Create_Model_From_Excel xlsheet,pkg 
	  SET xlsheet=NOTHING
	END IF
  NEXT
END SUB

'--------------------------------------------------------------------------------
'功能函数
'--------------------------------------------------------------------------------
PRIVATE SUB Create_Model_From_Excel(xlsheet,package)
	'定义数据表对象
	DIM Tab 
	DIM col
	DIM tabcode
	DIM tabcode1
	DIM i
	DIM col_code

	FOR i=2 TO xlsheet.UsedRange.Cells.Rows.Count
		'判断是否需要创建新表对象
		tabcode1 = xlsheet.Range(CELL_C+CSTR(i)).Value
		IF tabcode1<>"" and tabcode<>tabcode1 THEN
			SET Tab=NOTHING 
			tabcode=tabcode1
			IF tabcode<>"" THEN
			    '判断表是否存在，如果不存在则创建，存在则直接返回表对象
				SET tab = CreateOrReplaceTableByCode(tabcode,package)
				'将表的所有列删除,如果需要重新创建表的列
				IF isclear_columns THEN
					DeleteTableColumns(tab)
				END IF
				'更新表的属性
				'更新表的英文名称
				Tab.code=xlsheet.Range(CELL_C+CSTR(i)).Value
				'更新表的中文名称
				Tab.name=xlsheet.Range(CELL_D+CSTR(i)).Value
				'更新表的注释
				Tab.comment=xlsheet.Range(CELL_B+CSTR(i)).Value
				
				output "创建表模型OK:"+Tab.code+"——"+Tab.name
			END IF
		END IF

		'创建表的列
		IF NOT(Tab IS NOTHING) THEN 
			'列代码   	
			col_code=xlsheet.Range(CELL_E+CSTR(i)).Value 
			'判断是否已经存在列,不存在则创建
			SET col = CreateOrReplaceColumnByCode(col_code,Tab)
			'设置列属性
			'列英文名称
			col.code=xlsheet.Range(CELL_E+CSTR(i)).Value 
			'列中文名称
			col.name=xlsheet.Range(CELL_F+CSTR(i)).Value 
			'列注释
			col.comment=xlsheet.Range(CELL_G+CSTR(i)).Value 
			'列数据类型
			col.DataType=xlsheet.Range(CELL_H+CSTR(i)).Value 
			'列是否主键，如果是主键，则输出 Y
			IF CSTR(xlsheet.Range(CELL_I+CSTR(i)).Value)=str_iskey THEN
				col.primary= TRUE
			END IF
			'列是否非空，如果是非空，则输出 Y
			IF CSTR(xlsheet.Range(CELL_J+CSTR(i)).Value)=str_iskey THEN
				col.mandatory= TRUE
			END IF
			'列默认值
			col.DefaultValue=xlsheet.Range(CELL_K+CSTR(i)).Value 
			output "更新表模型的列OK:"+Tab.code+"——"+col.code+"--"+col.name
		END IF
	NEXT

END SUB

'--------------------------------------------------------------------------------
'功能函数
'--------------------------------------------------------------------------------
PRIVATE FUNCTION CreateOrReplacePackageByName(name,model)
	'Table 对象
	DIM pkg 
	SET pkg = FindPackageByName(name,model)
	IF pkg IS NOTHING THEN
	  SET pkg = model.Packages.CreateNew()
	  pkg.SetNameAndCode name, name
	  pkg.PhysicalDiagrams.Item(0).SetNameAndCode name, name
	END IF
	SET CreateOrReplacePackageByName = pkg
END FUNCTION

PRIVATE FUNCTION CreateOrReplaceTableByCode(code,package)
	'Table 对象
	DIM tab 
	SET tab = FindTableByCode(code,package)
	IF tab IS NOTHING THEN
	  SET tab = package.Tables.CreateNew()
	  tab.SetNameAndCode code, code
	END IF
	SET CreateOrReplaceTableByCode = tab
END FUNCTION

PRIVATE FUNCTION CreateOrReplaceColumnByCode(code,table)
	'Table 对象
	DIM col 
	SET col =FindColumnByCode(code,table) 
	IF col IS NOTHING THEN
	  SET col =table.Columns.CreateNew
	  col.SetNameAndCode code , code
	END IF
	SET CreateOrReplaceColumnByCode = col
END FUNCTION

PRIVATE FUNCTION FindPackageByName(name,model)
	'Table 对象
	DIM pkg 
	SET FindPackageByName = NOTHING
	FOR EACH pkg IN model.Packages
		IF NOT pkg.isShortcut THEN
			IF pkg.name =name THEN
				SET FindPackageByName=pkg
				Exit FOR
			END IF
		END IF
	NEXT
	
END FUNCTION

PRIVATE FUNCTION FindTableByName(name,package)
	'Table 对象
	DIM Tab1 
	SET FindTableByName = NOTHING
	FOR EACH Tab1 IN package.Tables
		IF NOT Tab1.isShortcut THEN
			IF Tab1.name =name THEN
				SET FindTableByName=Tab1
				Exit FOR
			END IF
		END IF
	NEXT
END FUNCTION

PRIVATE FUNCTION FindTableByCode(code,package)
	'Table 对象
	DIM Tab1 
	SET FindTableByCode = NOTHING
	FOR EACH Tab1 IN package.Tables
		IF NOT Tab1.isShortcut THEN
			'OUTPUT "循环表:"+Tab1.name
			IF Tab1.code =code THEN
				SET FindTableByCode=Tab1
				Exit FOR
			END IF
		END IF
	NEXT
END FUNCTION

PRIVATE FUNCTION FindColumnByCode(code,tabobj)
	'Column 对象
	DIM col1 
	'OUTPUT "code:"+code
	SET FindColumnByCode = NOTHING
	FOR EACH col1 IN tabobj.Columns
		'OUTPUT "code2:"+col1.code
		IF col1.code =code THEN
			SET FindColumnByCode=col1
			EXIT FOR
		END IF
	NEXT
END FUNCTION

PRIVATE FUNCTION FindColumnByName(name,tabobj)
	'Column 对象
	DIM col1 
	'OUTPUT "codename:"+name
	SET FindColumnByName = NOTHING
	FOR EACH col1 IN tabobj.Columns
		IF col1.name =name THEN
			SET FindColumnByName=col1
			EXIT FOR
		END IF
	NEXT
END FUNCTION

PRIVATE FUNCTION FindDomainByName(dmname,mdl)
	'Domain 对象
	DIM dm1 
	SET FindDomainByName = NOTHING

	FOR EACH dm1 IN mdl.domains
		IF NOT dm1.isShortcut THEN
			IF dm1.name =dmname THEN
				SET FindDomainByName =dm1
				EXIT FOR
			END IF
		END IF
	NEXT

END FUNCTION

PRIVATE FUNCTION FindUserByName(username)
	DIM user1
	SET FindUserByName = NOTHING
	FOR EACH user1 IN mdl.users
		IF user1.name=username THEN
			SET FindUserByName=user1
			EXIT FOR
		END IF
	NEXT

END FUNCTION

' 删除表的所有列
PRIVATE SUB DeleteTableColumns(table)
  IF NOT table.isShortcut THEN  
   DIM col
   FOR EACH col IN table.columns  
  	'output "Column deleted :"+table.name
  	col.Delete
  	SET col = NOTHING
   NEXT
  END IF
END SUB