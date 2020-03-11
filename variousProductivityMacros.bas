Attribute VB_Name = "Module1"
Sub unresourced1()
Attribute unresourced1.VB_Description = "unresourced"
Attribute unresourced1.VB_ProcData.VB_Invoke_Func = " \n14"
' unresourced1 Macro

    Range("B:B,C:C,D:D,F:F").Select
    Range("F1").Activate
    Selection.Delete Shift:=xlToLeft
    Columns("H:H").Select
    Columns("H:AR").Select
    Selection.Delete Shift:=xlToLeft
    Range("E2").Select
    ActiveSheet.Range("$A$1:$G$1144").RemoveDuplicates Columns:=5, Header:= _
        xlYes
End Sub
Sub mastercs()
Attribute mastercs.VB_ProcData.VB_Invoke_Func = " \n14"
' mastercs Macro

    Rows("8:8").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$8:$AD$1056").AutoFilter Field:=13, Criteria1:="RS"
    ActiveSheet.Range("$A$8:$AD$1056").AutoFilter Field:=16, Criteria1:="RS"
    Range("J23").Select
    ActiveSheet.Range("$A$8:$AD$1056").RemoveDuplicates Columns:=4, Header:= _
        xlYes
End Sub
Sub unresourced2()
Attribute unresourced2.VB_ProcData.VB_Invoke_Func = " \n14"
' unresourced2 Macro

    Range("E2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.FormatConditions.AddUniqueValues
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).DupeUnique = xlDuplicate
    With Selection.FormatConditions(1).Font
        .Color = -16383844
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    ActiveSheet.Range("$A$1:$G$793").AutoFilter Field:=5, Criteria1:=RGB(255, _
        199, 206), Operator:=xlFilterCellColor
End Sub

Sub import_sheets()
    
    Dim directory As String, fileName As String, sheet As Worksheet, total As Integer
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    'directory = "c:\test\"
    'fileName = Dir(directory & "*.xl??")
    Workbooks.Open (directory & fileName)
    
    For Each sheet In Workbooks(fileName).Worksheets
        total = Workbooks("import-sheets.xls").Worksheets.Count
        Workbooks(fileName).Worksheets(sheet.Name).Copy _
        after:=Workbooks("import-sheets.xls").Worksheets(total)
    Next sheet
    
    Workbooks(fileName).Close
    fileName = Dir()
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
End Sub

Option Explicit

Sub PortalLogon()
'This will load a webpage in IE
    Dim i As Long
    Dim URL As String
    Dim IE As Object
    Dim objElement As Object
    Dim objCollection As Object
 
    'Create InternetExplorer Object
    Set IE = CreateObject("InternetExplorer.Application")
 
    'Set IE.Visible = True to make IE visible, or False for IE to run in the background
    IE.Visible = True
 
    'Define URL(s)
    URL = "http://inspect.portal.manheim.co.uk/Login.aspx"
 
    'Navigate to URL
    IE.Navigate URL
 
    ' Wait while IE loading...
    'IE ReadyState = 4 signifies the webpage has loaded (the first loop is set to avoid inadvertently skipping over the second loop)
    Do While IE.ReadyState = 4: DoEvents: Loop   'Do While
    Do Until IE.ReadyState = 4: DoEvents: Loop   'Do Until
 
    'Enter Login Credentials
    Set htmldoc = IE.document

    For Each htmlel In htmldoc.getElementsByTagName("input")
    'MsgBox htmlel.Name
        If htmlel.ID = "ctl00_ContentPlaceHolder1_Login1_UserName" Then
            htmlel.Value = "<USERNAME>"
        End If
    Next

    For Each htmlel In htmldoc.getElementsByTagName("input")
        If htmlel.ID = "ctl00_ContentPlaceHolder1_Login1_Password" Then
            htmlel.Value = "<PASSWORD>"
        End If
    Next

    'Find and Click Login Button
    For Each htmlel In htmldoc.getElementsByTagName("input")
        If htmlel.ID = "ctl00_ContentPlaceHolder1_Login1_LoginButton" Then
            htmlel.Click
        End If
    Next
   
    ' Wait while IE loading...
    'IE ReadyState = 4 signifies the webpage has loaded (the first loop is set to avoid inadvertently skipping over the second loop)
    Do While IE.ReadyState = 4: DoEvents: Loop   'Do While
    Do Until IE.ReadyState = 4: DoEvents: Loop   'Do Until
   

    
    'Unload IE
    Set IE = Nothing
    Set objElement = Nothing
    Set objCollection = Nothing
    
End Sub

