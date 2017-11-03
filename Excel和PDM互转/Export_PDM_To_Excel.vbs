Option Explicit
'Model sheet中的列信息
'主题域(Pachage)
CONST CELL_A="A" 
'表注释
CONST CELL_B="B" 
'表英文名称
CONST CELL_C="C" 
'表中文名称
CONST CELL_D="D"
'列英文名称 
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
DIM nb
'the current model
DIM mdl 
SET mdl = ActiveModel
IF (mdl IS NOTHING) THEN
   MsgBox "没有选择一个Model"
END IF

DIM fldr
SET Fldr = ActiveDiagram.Parent
'是否需要合并表名称单元格
DIM isMerage 
'是否不同的Package不同的sheet
DIM isMulite 
DIM RQ

RQ = MsgBox ("是否不同的Package不同的sheet?", vbYesNo + vbInformation,"确认")
IF RQ= VbYes THEN
 isMulite= TRUE
ELSE
 isMulite= FALSE
END IF

'创建新的Excel
DIM x1  '
SET x1 = CreateObject("Excel.Application")
x1.Workbooks.Add
x1.Visible = TRUE

ExportModelToExcel( fldr)

MsgBox "成功将 Models 导出到Excel中！"

'--------------------------------------------------------------------------------
'功能函数:将模型导出到Sheet页【 MODEL 】
'--------------------------------------------------------------------------------
PRIVATE FUNCTION ExportModelToExcel(folder)
  '如果是每个package导出到不同的sheet页面，则采用folder的名称作为sheet页名称，否则使用"MODEL"作为sheet页名称
  IF isMulite THEN
    IF folder.Tables.count>0 THEN
	  AddExcelSheet(folder.name)
    END IF
  ELSE
    AddExcelSheet("MODEL")		
  END IF
  '写sheet页的第一行表头
  WriteExcelModelHead

  DIM nStart
  DIM nEnd
  '定义数据表对象
  DIM tabobj 
  
  nb=2
  '是否需要合并单元格
  isMerage=FALSE
  '开始循环处理所有的folder
  FOR EACH tabobj IN folder.Tables
    '快捷方式不处理
    IF NOT tabobj.isShortcut THEN 
      '合并表的单元格A、B、C
      IF isMerage THEN  
        '合并起始行
        nStart=nb 
        '合并结束行
        nEnd=nb+tabobj.Columns.count-1 
		IF nStart<>nEnd THEN
          '合并单元格
          '合并主题域
          x1.Range(CELL_A+CSTR(nStart)+":"+CELL_A+CSTR(nEnd)).SELECT
          x1.Selection.Merge
          '合并表注释
          x1.Range(CELL_B+CSTR(nStart)+":"+CELL_B+CSTR(nEnd)).SELECT
          x1.Selection.Merge
          '合并表英文名称
          x1.Range(CELL_C+CSTR(nStart)+":"+CELL_C+CSTR(nEnd)).SELECT
          x1.Selection.Merge
          '合并表中文名称
          x1.Range(CELL_D+CSTR(nStart)+":"+CELL_D+CSTR(nEnd)).SELECT
          x1.Selection.Merge
		END IF
        '将主题域、表名称、表注释填写到合并后单元格中
        '主题域
        x1.Range(CELL_A+CSTR(nb)).Value = folder.name 
        '表注释  
        x1.Range(CELL_B+CSTR(nb)).Value = tabobj.comment   
        '表英文名称
        x1.Range(CELL_C+CSTR(nb)).Value = tabobj.code   
        '表中文名称 
        x1.Range(CELL_D+CSTR(nb)).Value = tabobj.name   
      END IF
      '开始循环列出输出信息
      '定义列对象
      DIM colobj 
      FOR EACH colobj IN tabobj.Columns
		    '写表的信息
        '主题域
        x1.Range(CELL_A+CSTR(nb)).Value = folder.name 
        '表注释
        x1.Range(CELL_B+CSTR(nb)).Value = tabobj.comment
        '表英文名称
        x1.Range(CELL_C+CSTR(nb)).Value = tabobj.code   
        '表中文名称 
        x1.Range(CELL_D+CSTR(nb)).Value = tabobj.name    
		
        '写列的信息
        '列英文名称
        x1.Range(CELL_E+CSTR(nb)).Value = colobj.code    
        '列中文名称
        x1.Range(CELL_F+CSTR(nb)).Value = colobj.name    
        '列注释
	     	x1.Range(CELL_G+CSTR(nb)).Value = colobj.comment
        '数据类型
        x1.Range(CELL_H+CSTR(nb)).Value = colobj.DataType    
        '列是否主键，如果是主键，则输出 Y
        IF colobj.primary THEN
          x1.Range(CELL_I+CSTR(nb)).Value = "Y"
        END IF
        '列是否为非空
         IF colobj.mandatory THEN
          x1.Range(CELL_J+CSTR(nb)).Value = "Y"
        END IF
        '列的默认值
        x1.Range(CELL_K+CSTR(nb)).Value = colobj.DefaultValue
        '行号加1
        nb = nb+1  
      NEXT
    END IF
  NEXT

  '对子包进行递归，如果不使用递归只能取到第一个模型图内的表
  DIM subfolder
  FOR EACH subfolder IN folder.Packages
    ExportModelToExcel(subfolder) 
  NEXT

END FUNCTION

'--------------------------------------------------------------------------------
'功能函数:添加一个Sheet页
'--------------------------------------------------------------------------------
PRIVATE SUB AddExcelSheet(sheetname)
  x1.Sheets.Add
  x1.ActiveSheet.Name=sheetname
END SUB

'--------------------------------------------------------------------------------
'功能函数:写Excel的第一行信息
'--------------------------------------------------------------------------------
PRIVATE SUB WriteExcelModelHead()
   x1.Range(CELL_A+"1").Value = "主题域"
   x1.Range(CELL_B+"1").Value = "表注释"
   x1.Range(CELL_C+"1").Value = "表英文名称"
   x1.Range(CELL_D+"1").Value = "表中文名称"
   x1.Range(CELL_E+"1").Value = "列英文名称"
   x1.Range(CELL_F+"1").Value = "列中文名称"
   x1.Range(CELL_G+"1").Value = "列注释"
   x1.Range(CELL_H+"1").Value = "数据类型"
   x1.Range(CELL_I+"1").Value = "是否为主键"
   x1.Range(CELL_J+"1").Value = "是否为非空"
   x1.Range(CELL_K+"1").Value = "默认值"
   
   '设置字体
   x1.Columns(CELL_A+":"+CELL_K).SELECT
   WITH x1.Selection.Font
        .Name = "宋体"
        .Size = 10
   END WITH

   '设置首行可过滤,背景颜色为灰色,字体粗体
   x1.Range(CELL_A+"1:"+CELL_K+"1").SELECT
   x1.Selection.AutoFilter
   x1.Selection.Interior.ColorIndex = 15
   x1.Selection.Font.Bold = TRUE
   '设定首行固定
   x1.Range(CELL_A+"2").SELECT
   x1.ActiveWindow.FreezePanes = TRUE

END SUB