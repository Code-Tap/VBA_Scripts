Attribute VB_Name = "emailFunctions"
Public Function ResolveDisplayName(sFromName) As Boolean

    Dim OLApp As Object 'Outlook.Application
    Dim oRecip As Object 'Outlook.Recipient
    Dim oEU As Object 'Outlook.ExchangeUser
    Dim oEDL As Object 'Outlook.ExchangeDistributionList

    Set OLApp = CreateObject("Outlook.Application")
    Set oRecip = OLApp.Session.CreateRecipient(sFromName)
    oRecip.Resolve
    If oRecip.Resolved Then
        ResolveDisplayName = True
    Else
        ResolveDisplayName = False
    End If

End Function

Public Function ResolveDisplayNameToSMTP(sFromName) As String
    Dim OLApp As Object 'Outlook.Application
    Dim oRecip As Object 'Outlook.Recipient
    Dim oEU As Object 'Outlook.ExchangeUser
    Dim oEDL As Object 'Outlook.ExchangeDistributionList

    Set OLApp = CreateObject("Outlook.Application")
    Set oRecip = OLApp.Session.CreateRecipient(sFromName)
    oRecip.Resolve
    If oRecip.Resolved Then
        Select Case oRecip.AddressEntry.AddressEntryUserType
            Case 0, 5 'olExchangeUserAddressEntry & olExchangeRemoteUserAddressEntry
                Set oEU = oRecip.AddressEntry.GetExchangeUser
                If Not (oEU Is Nothing) Then
                    ResolveDisplayNameToSMTP = oEU.PrimarySmtpAddress
                End If
            Case 10, 30 'olOutlookContactAddressEntry & 'olSmtpAddressEntry
                    ResolveDisplayNameToSMTP = oRecip.AddressEntry.Address
        End Select
    End If
End Function

Sub Test()

    MsgBox ResolveDisplayName("John Doe")
    MsgBox ResolveDisplayNameToSMTP("John Doe")

End Sub

Sub Validate()
    Dim cell As Range
    
    InsRange = "A2"
    TlrRange = "D2"
    RmngRange = "G2"
    LeaseRange = "J2"
    
    Set Insp = Range(InsRange, Range(InsRange).End(xlDown))
    Set Tldr = Range(TlrRange, Range(TlrRange).End(xlDown))
    Set Rmng = Range(RmngRange, Range(RmngRange).End(xlDown))
    Set Lesco = Range(LeaseRange, Range(LeaseRange).End(xlDown))

    For Each cell In Insp.Cells
        cell.Offset(0, 1).Value = ResolveDisplayName(cell)
        Next cell
       
    For Each cell In Tldr.Cells
        cell.Offset(0, 1).Value = ResolveDisplayName(cell)
        Next cell
        
    For Each cell In Rmng.Cells
        cell.Offset(0, 1).Value = ResolveDisplayName(cell)
        Next cell
        
    For Each cell In Lesco.Cells
        cell.Offset(0, 1).Value = ResolveDisplayName(cell)
        Next cell
    
End Sub

Sub ValidateLive()
    Dim cell As Range
    
    InsRange = "D2"
    
    Set Insp = Range(InsRange, Range(InsRange).End(xlDown))

    For Each cell In Insp.Cells
        cell.Offset(0, 2).Value = ResolveDisplayName(cell)
        Next cell
           
End Sub



Option Explicit
Sub Mail_RM_TL_AMX()
     
    ActiveWorkbook.RefreshAll
    Dim OutApp As Object
    Dim OutMail As Object
    Dim cell As Range
    Dim sTo As String
    Dim ccTo As String
    Dim pt1 As PivotTable
    Dim pt2 As PivotTable
    Dim pt3 As PivotTable
    Dim rng As Range
    
    With Application
        .EnableEvents = False
        .ScreenUpdating = False
    End With
    
    Set rng = Nothing
    Set rng = Sheets(1).ListObjects("Table1").Range '' Alternate to select everything on the page = Sheets(1).UsedRange
    Set pt1 = Sheet8.PivotTables("teamLeader") 'TeamLeaders
    Set pt2 = Sheet8.PivotTables("regMngr") 'Regional Managers
    Set pt3 = Sheet8.PivotTables("inspector") 'Inspectors
    
    For Each cell In pt1.DataBodyRange.Cells
        sTo = sTo & ";" & cell.Value
    Next

    For Each cell In pt2.DataBodyRange.Cells
        ccTo = ccTo & ";" & cell.Value
    Next

    sTo = Mid(sTo, 2)

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    On Error Resume Next
    With OutMail
        .To = sTo
        .CC = ccTo
        .BCC = ""
        .Subject = "AMX Overdue report "
        .HTMLBody = RangetoHTML(rng)
         '.Attachments.Add ActiveWorkbook.FullName
         'You can add other files also like this
         '.Attachments.Add ("C:\test.txt")
        .Display
    End With
    On Error GoTo 0
    
    With Application
        .EnableEvents = True
        .ScreenUpdating = True
    End With
    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub

Function RangetoHTML(rng As Range)
' Changed by Ron de Bruin 28-Oct-2006
' Working in Office 2000-2016
    Dim fso As Object
    Dim ts As Object
    Dim TempFile As String
    Dim TempWB As Workbook

    TempFile = Environ$("temp") & "\" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"

    'Copy the range and create a new workbook to past the data in
    rng.Copy
    Set TempWB = Workbooks.Add(1)
    With TempWB.Sheets(1)
        .Cells(1).PasteSpecial Paste:=8
        .Cells(1).PasteSpecial xlPasteValues, , False, False
        .Cells(1).PasteSpecial xlPasteFormats, , False, False
        .Cells(1).Select
        Application.CutCopyMode = False
        On Error Resume Next
        .DrawingObjects.Visible = True
        .DrawingObjects.Delete
        On Error GoTo 0
    End With

    'Publish the sheet to a htm file
    With TempWB.PublishObjects.Add( _
         SourceType:=xlSourceRange, _
         Filename:=TempFile, _
         Sheet:=TempWB.Sheets(1).Name, _
         Source:=TempWB.Sheets(1).UsedRange.Address, _
         HtmlType:=xlHtmlStatic)
        .Publish (True)
    End With

    'Read all data from the htm file into RangetoHTML
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
    RangetoHTML = ts.readall
    ts.Close
    RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", _
                          "align=left x:publishsource=")

    'Close TempWB
    TempWB.Close savechanges:=False

    'Delete the htm file we used in this function
    Kill TempFile

    Set ts = Nothing
    Set fso = Nothing
    Set TempWB = Nothing
End Function

