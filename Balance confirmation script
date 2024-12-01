Sub Saldenbestatigung()
    Dim SapGuiAuto As Object
    Dim SAPApp As Object
    Dim SAPCon As Object
    Dim session As Object
    Dim SAPGuiSession As Object

    ' Initialize the SAP GUI Scripting objects
    Set SapGuiAuto = GetObject("SAPGUI")
    Set SAPApp = SapGuiAuto.GetScriptingEngine
    Set SAPCon = SAPApp.Children(0)
    Set session = SAPCon.Children(0)
    

  
Dim KeyDate As String
Dim BP As String
Dim CompanyCode As String
     Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("SB")

KeyDate = ActiveWorkbook.ActiveSheet.Range("D2").Value
BP = ActiveWorkbook.ActiveSheet.Range("B2").Value
CompanyCode = ActiveWorkbook.ActiveSheet.Range("C2").Value
Sheets("Data").Visible = False
Application.Calculation = xlManual
    
    ws.Range("H1") = "PROCESSING."
 ws.Range("A4:H3500").clearcontents
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/okcd").Text = "FPO4"
session.findById("wnd[0]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[1]/btn[17]").press
session.findById("wnd[1]/usr/txtV-LOW").Text = "RS"
session.findById("wnd[1]/usr/ctxtENVIR-LOW").Text = ""
session.findById("wnd[1]/usr/txtENAME-LOW").Text = ""
session.findById("wnd[1]/usr/txtAENAME-LOW").Text = ""
session.findById("wnd[1]/usr/txtMLANGU-LOW").Text = ""
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/tabsTABSTRIP_TAB_B1/tabpTB11/ssub%_SUBSCREEN_TAB_B1:RFKKOP04:1001/ctxtP_KEYDT").Text = KeyDate
session.findById("wnd[0]/usr/tabsTABSTRIP_TAB_B1/tabpTB11/ssub%_SUBSCREEN_TAB_B1:RFKKOP04:1001/ctxtS_GPART-LOW").Text = BP
session.findById("wnd[0]/usr/tabsTABSTRIP_TAB_B1/tabpTB11/ssub%_SUBSCREEN_TAB_B1:RFKKOP04:1001/ctxtS_BUKRS-LOW").Text = CompanyCode
session.findById("wnd[0]/usr/tabsTABSTRIP_TAB_B1/tabpTB11/ssub%_SUBSCREEN_TAB_B1:RFKKOP04:1001/radP_OPEN").SetFocus

session.findById("wnd[0]/usr/tabsTABSTRIP_TAB_B1/tabpTB12").Select
session.findById("wnd[0]/usr/tabsTABSTRIP_TAB_B1/tabpTB12/ssub%_SUBSCREEN_TAB_B1:RFKKOP04:1002/chkP_ALV").SetFocus
session.findById("wnd[0]/usr/tabsTABSTRIP_TAB_B1/tabpTB12/ssub%_SUBSCREEN_TAB_B1:RFKKOP04:1002/chkP_ALV").Selected = True

ws.Range("H1") = "PROCESSING.."
session.findById("wnd[0]/usr/tabsTABSTRIP_TAB_B1/tabpTB12/ssub%_SUBSCREEN_TAB_B1:RFKKOP04:1002/ctxtP_ALVVAR").Text = "/MACRO R"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[8]").press

On Error GoTo Here
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = "ZZFCA0_NAME_ORG1_FKKOPRU"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").contextMenu
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectContextMenuItem "&XXL"
session.findById("wnd[1]/tbar[0]/btn[20]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
On Error GoTo 0

'activate last open workbook
    initialWorkbookCount = Application.Workbooks.Count
    maxWaitTime = 100
    startTime = Timer
    Do While Application.Workbooks.Count <= initialWorkbookCount
        DoEvents
        If Timer > startTime + maxWaitTime Then
            MsgBox "New workbook did not open within the expected time.", vbExclamation
            Exit Sub
        End If
    Loop
    Set NewWorkbook = Application.Workbooks(Application.Workbooks.Count)
    NewWorkbook.Sheets("Data").Range("A2:H2000").Select
Selection.Copy

    ws.Activate
    ws.Range("A4").Select
     DoEvents
    ws.Paste
    NewWorkbook.Activate
    Application.CutCopyMode = False
    NewWorkbook.Close SaveChanges:=False
    'reactivate the clipboard for future use
    Application.CutCopyMode = True
ws.Activate
    ws.Columns("I:Y").Select
    Selection.clearcontents
    'Value Formating
    
    With ThisWorkbook.Sheets("SB")
        .Range("C4:C4000").Value = .Range("C4:C4000").Value
    End With
session.findById("wnd[0]/tbar[0]/btn[15]").press
session.findById("wnd[0]/tbar[0]/btn[15]").press
ws.Range("H1") = "PROCESSING..."
    Dim rng As Range
    Dim cell As Range
    Dim cellValue As String
    Dim lastChar As String
    
        Range("A4:J5000").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    
    
           Range("A2").Copy
    Range("D2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
Range("D2").Value = Format(Range("D2").Value, "dd.mm.yyyy")
    
    ' Set the worksheet to "SB"
    Set ws = ThisWorkbook.Sheets("SB")
    
    ' Set the range F4:F200
    Set rng = ws.Range("F4:F3000")
    
    ' Loop through each cell in the range
    For Each cell In rng
        ' Substitute "." with ""
        cellValue = Replace(cell.Value, ".", "")
        
        ' Get the rightmost character
        lastChar = Right(cellValue, 1)
        
        ' Check if the rightmost character is "-"
        If lastChar = "-" Then
            ' Remove the "-" and make the number negative
            cellValue = Left(cellValue, Len(cellValue) - 1)
            cellValue = "-" & cellValue
        End If
        ' Update the cell value
        cell.Value = cellValue
        
        ' Convert the resulting text to a value
        On Error Resume Next ' To handle any errors during conversion
        On Error GoTo 0 ' Resume normal error handling
        ' Apply the VALUE() function to convert the string to a number
        cell.Formula = "=VALUE(""" & cellValue & """)"
        
        ' Replace the formula with its result to keep only the value
        cell.Value = cell.Value
        Calculate
    
         If ws.Cells(cell.Row, "E").Value = "" Then
            ' Clear the cell in column F
            cell.clearcontents
        End If
      ws.Range("H1") = "PROCESSING"
       Calculate
       
            Dim sourceRange As Range
    Dim fillRange As Range
   Next cell
   Calculate
   Application.Calculation = xlAutomatic
   ThisWorkbook.Sheets("SB").Range("I4").Select
    ActiveCell.FormulaR1C1 = "=SUMIFS(C[-3],C[-1],RC[-1])"
    Range("I4").Select
    Selection.AutoFill Destination:=Range("I4:I3000")
    Calculate
    ws.Columns("I:I").Select
    Selection.Copy
     DoEvents
    ws.Columns("F:F").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ws.Range("F3").Select
    Application.CutCopyMode = False
    ThisWorkbook.Sheets("SB").Range("F3").FormulaR1C1 = "Amount"
    ws.Columns("I:I").clearcontents
    ThisWorkbook.Sheets("SB").Range("A3:H9999").RemoveDuplicates Columns:=3, Header:=xlYes
        
ws.Range("H1") = "PROCESSING."
        Calculate
    ' Find the first empty cell in the range
    For Each cell In rng
        If IsEmpty(cell) Then
            Set firstEmptyCell = cell
            Exit For
        End If
    Next cell
    Calculate
    Application.Calculation = xlManual
    ' Place the sum in the first empty cell found
    If Not firstEmptyCell Is Nothing Then
        firstEmptyCell.Select
ActiveCell.FormulaR1C1 = "=SUMIFS(R4C6:R3000C6, R4C7:R3000C7, ""<>"")"
       firstEmptyCell.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    End If
    Set ws = ThisWorkbook.Sheets("SB")
    Calculate

    Dim lastRow As Long
    
    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("SB") ' Adjust to your sheet name
    
    ' Find the last used row in column E
    lastRow = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row
    
    ' Define the range to check
    Set sourceRange = ws.Range("E1:E" & lastRow)
    
    ' Loop through each cell in the range
    For Each cell In sourceRange
        ' Check if the cell contains a valid date
        If IsDate(cell.Value) Then
            ' Format the date as dd.mm.yyyy
            cell.Value = Format(CDate(cell.Value), "dd.mm.yyyy")
             cell.Value = cell.Value
              cell.Value = cell.Value
        End If
    Next cell
    
    Calculate
    With ws.Range("F4:F3000").NumberFormat = "_-* #,##0.00 _K_c_-;-* #,##0.00 _K_c_-;_-* ""-""?? _K_c_-;_-@_-"
    With ws.Range("J4:J3000").NumberFormat = "_-* #,##0.00 _K_c_-;-* #,##0.00 _K_c_-;_-* ""-""?? _K_c_-;_-@_-"
   Range("C4:C3000").NumberFormat = "0"
    Range("H4:H3000").NumberFormat = "0"
    ws.Range("A1").Select
    ws.Range("H1") = "PROCESSING.."
     Range("F3").Select
    Selection.End(xlDown).Select
    Application.CutCopyMode = False
    Selection.Copy
    ActiveCell.Offset(-1, 0).Range("A1").Select

    ActiveCell.Offset(1, 0).Range("A1").Select
    Application.CutCopyMode = False
    Calculate
    End With
    End With
     GoTo Boop
     
     
Here:
     MsgBox "No open entries found", vbInformation, "Message"
GoTo Ends

     
Boop:

      Columns("F:F").Select
    Selection.Style = "Comma"
    Calculate
ThisWorkbook.Sheets("SB").Range("C4:C2000").Select
ThisWorkbook.Sheets("SB").Range("C4:C2000").Copy
ws.Range("H1") = "PROCESSING..."
    session.findById("wnd[0]/tbar[0]/okcd").Text = "ZZCI1_REP_INVHEADER"
session.findById("wnd[0]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[1]/btn[17]").press
session.findById("wnd[1]/usr/txtV-LOW").Text = "RS"
session.findById("wnd[1]/usr/ctxtENVIR-LOW").Text = ""
session.findById("wnd[1]/usr/txtENAME-LOW").Text = ""
session.findById("wnd[1]/usr/txtAENAME-LOW").Text = ""
session.findById("wnd[1]/usr/txtMLANGU-LOW").Text = ""
session.findById("wnd[1]/usr/txtV-LOW").SetFocus
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/ctxtS_BUKRS-LOW").Text = CompanyCode
session.findById("wnd[0]/usr/txtS_EXBEL-LOW").SetFocus
session.findById("wnd[0]/usr/txtS_EXBEL-LOW").caretPosition = 0
session.findById("wnd[0]/usr/btn%_S_EXBEL_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[24]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[8]").press
Calculate
ws.Range("H1") = "PROCESSING"
    Sheets("Data").Range("A2:D20000").clearcontents
    
session.findById("wnd[0]/tbar[1]/btn[33]").press
session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").contextMenu
session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").selectContextMenuItem "&FILTER"
session.findById("wnd[2]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").Text = "/MACRO R"
session.findById("wnd[2]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").selectedRows = "0"
session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").clickCurrentCell



Calculate

session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItem "&XXL"
session.findById("wnd[1]/tbar[0]/btn[20]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press

'to open and set the new new workbook
initialWorkbookCount = Application.Workbooks.Count
    maxWaitTime = 100
    startTime = Timer
    Do While Application.Workbooks.Count <= initialWorkbookCount
        DoEvents
        If Timer > startTime + maxWaitTime Then
            MsgBox "New workbook did not open within the expected time.", vbExclamation
            Exit Sub
        End If
    Loop
    Set NewnewWorkbook = Application.Workbooks(Application.Workbooks.Count)
    NewnewWorkbook.Sheets("Data").Range("A2:D3000").Select
Selection.Copy



ws.Range("H1") = "PROCESSING."
 DoEvents
 ws.Activate
    Sheets("Data").Visible = True
    Sheets("Data").Select
    Range("A2").Select
    ActiveSheet.Paste
    
        Sheets("SB").Select
         Sheets("Data").Visible = False
    
        NewnewWorkbook.Activate
    Application.CutCopyMode = False
    NewnewWorkbook.Close SaveChanges:=False
    
     Application.CutCopyMode = True
Resume2:
    Calculate
    Sheets("SB").Select
    Range("I3").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "Bill to BP"
    Range("I4").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-6]="""","""",XLOOKUP(RC[-6],Value(Data!C[-8]),Data!C[-7],""""))"
    Range("I4").Select
    Selection.AutoFill Destination:=Range("I4:I2000")

    
    Application.Calculation = xlAutomatic
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press

  lastRow = ws.Cells(ws.Rows.Count, "F").End(xlUp).Row
    If lastRow >= 4 And firstEmptyCell.Row <= 3000 Then
          With ws.Range("A" & firstEmptyCell.Row - 1 & ":I" & firstEmptyCell.Row - 1).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThick
        End With
    End If

Ends:

'BP part
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/okcd").Text = "ZFCA0_1352"
session.findById("wnd[0]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/chkP_BP_AD").Selected = True
session.findById("wnd[0]/usr/ctxtS_STDBK-LOW").Text = CompanyCode
session.findById("wnd[0]/usr/ctxtS_GPART-LOW").Text = BP
session.findById("wnd[0]/usr/chkP_BP_AD").SetFocus
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[33]").press
session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").currentCellRow = 1
session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").selectColumn "VARIANT"
session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").selectedRows = ""
session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").contextMenu
session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").selectContextMenuItem "&FILTER"
session.findById("wnd[2]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").Text = "/MACRO ADDR"
session.findById("wnd[2]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").selectedRows = "0"
session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").clickCurrentCell
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectedRows = "0"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItemBytext "Copy Text"
 DoEvents
    
    
    
   Sheets("Data").Visible = True
    Sheets("Data").Select
    Range("N3").Select
    On Error GoTo Error1
    
    ActiveSheet.Paste
    On Error GoTo 0
Resumerino:
    
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
GoTo Skipp
Error1:
 DoEvents
  ActiveSheet.Paste
  GoTo Resumerino
Skipp:
  Sheets("SB").Select
  ws.Range("H1") = ""
  Sheets("Data").Visible = False
  
End Sub

Sub Macro16()
'
' Macro16 Macro
'
Sheets("Form").Range("C18:H18").AutoFill Destination:=Sheets("Form").Range("C18:H2000"), Type:=xlFillDefault

Dim KeyDate As String
Dim Supplier As String
Supplier = ActiveWorkbook.Sheets("Form").Range("C3").Value
KeyDate = ActiveWorkbook.ActiveSheet.Range("A2").Value
    Dim firstEmptyCell As Range
    Dim cell As Range
     Dim targetRange As Range

Sheets("Data").Visible = True
    Sheets("Data").Select
    Sheets("Data").Range("N18:S19").Select
    Selection.Copy

    For Each cell In Sheets("Form").Range("C17:C1999")
        ' Check if the cell appears empty, even if it contains a formula returning ""
        If cell.Value = "" And Not IsEmpty(cell) Then
            Set firstEmptyCell = cell
            Exit For
        End If
    Next cell
    
    ' Check if a suitable cell was found
    If Not firstEmptyCell Is Nothing Then
        ' Paste into the first found cell
    firstEmptyCell.PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False
    End If
    Sheets("Form").Select
            Set targetRange = firstEmptyCell.Resize(1, 6)
        
        ' Select the range
        targetRange.Select
    With Selection.Font
        .Name = "Calibri"
        .Size = 36
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
        Selection.Borders(xlEdgeTop).LineStyle = xlThick
        Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    
    End With
    Range("C44:C45").Select
    Selection.Font.Bold = True
    Range("H49").Select
    
    'print the doc
    
    
    Dim selectedRange As Range
    Dim pdfFileName As String
    Dim ws As Worksheet
    
    ' Set the worksheet and selected range
    Set ws = ThisWorkbook.Sheets("Form") ' Change "Sheet1" to your sheet name
    Set selectedRange = ws.Range("A1", firstEmptyCell.Resize(5, 7)) ' Change to your desired range
    
    
        Dim userResponse As VbMsgBoxResult
    
    ' Prompt the user to decide whether to fit to 1 page
    userResponse = MsgBox("Fit to 1 Page?", vbYesNo + vbQuestion, "Page Fit Setting")
    
    
       With ws.PageSetup
        .PaperSize = xlPaperA4
        .Orientation = xlPortrait ' Set to xlLandscape if needed
        .FitToPagesWide = 1 ' Fit content to 1 page wide
        .Zoom = False ' Disable manual zoom to allow scaling
        
        ' Set FitToPagesTall based on user response
        If userResponse = vbYes Then
            .FitToPagesTall = 1 ' Fit to 1 page tall
        Else
            .FitToPagesTall = False ' Allow splitting across multiple pages vertically
        End If
    End With
    
    ' Define the PDF file name and path
    pdfFileName = Supplier & " Saldenbestätigung " & KeyDate
    
    ' Print the selected range to a PDF
    selectedRange.ExportAsFixedFormat Type:=xlTypePDF, _
                                      Filename:=pdfFileName, _
                                      Quality:=xlQualityStandard, _
                                      IncludeDocProperties:=True, _
                                      IgnorePrintAreas:=False, _
                                      OpenAfterPublish:=True
    Sheets("SB").Select
    Sheets("Data").Visible = False
End Sub


Sub KK()
With ThisWorkbook.Sheets("Data")
    .Visible = xlSheetVisible
    .Activate
    .Range("A2").PasteSpecial Format:="Unicode Text", Link:=False, _
        DisplayAsIcon:=False, NoHTMLFormatting:=True
    .Visible = xlSheetHidden ' Or xlSheetVeryHidden
    End With
    
End Sub


Sub Macro4()
'
' Macro4 Macro
'

'
    Sheets("Form").Select
    Sheets("Data").Visible = True
    Sheets("Data").Select
    Range("A2").Select
    ActiveSheet.PasteSpecial Format:="Unicode Text", Link:=False, _
        DisplayAsIcon:=False, NoHTMLFormatting:=True
    Range("A1").Select
End Sub


Sub clearcontents()
'
' Macro6 Macro
    Range("A4:I20000").Select
    Selection.clearcontents
        Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
     End With
End Sub
Sub Macro7()
'
' Macro7 Macro
'

'
    Range("A2:A30000").Select
    Range("A239").Activate
    Selection.NumberFormat = "0"
End Sub
