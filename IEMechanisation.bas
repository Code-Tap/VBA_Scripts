Attribute VB_Name = "Module4"
Private Sub delay(seconds As Long)
    Dim endTime As Date
    endTime = DateAdd("s", seconds, Now())
    Do While Now() < endTime
        DoEvents
    Loop
End Sub
Sub confirm()
'we define the essential variables
Dim IE As Object
Dim form As Variant, Button As Object

Set IE = CreateObject("InternetExplorer.Application")
With IE

.Visible = True
.Navigate ("https://webtms.tracktrans.net/webtms/#login")
MsgBox ("Logged in ready to proceed")
.Navigate ("https://webtms.tracktrans.net/webtms/#tms_customer_comms")
MsgBox ("please clear dates")
Sheets("Inspector").Range("C2").Select
Do Until ActiveCell = ""
If ActiveCell = "" Then If MsgBox("shall we rest page?", vbYesNo + vbQuestion) = vbYes Then GoTo Cleanup
'If ActiveCell.Offset(0, 1) = "y" Then GoTo Nextone
sepc = ActiveCell.Value
amxdate = ActiveCell.Offset(0, -2).Value
amxdate = Format(amxdate, "dd/mm/yyyy")
IE.Document.getelementsbyclassname("gwt-TextBox GIOQPGIKN GIOQPGIOT GIOQPGIFP").Item.innerText = sepc
'987LLlkj!!$$
IE.Document.getelementsbyclassname("gwt-Button GIOQPGIFP").Item.Click
Do While IE.Busy: DoEvents: Loop
Do Until (IE.readyState = 4 And Not IE.Busy)
Loop
delay 3
On Error GoTo -1
On Error Resume Next
On Error GoTo ErrorHandler
If IE.Document.getelementsbyclassname("gwt-HTML").Item(4).innerText = "No data matching your search" Then
IE.Document.getelementsbyclassname("close").Item.Click
'MsgBox ("Close small window")
GoTo Aborted
End If
ErrorHandler:

sitedate = IE.Document.getelementsbyclassname("center").Item.innerText
sitedate = Format(sitedate, "dd/mm/yyyy")
If amxdate = sitedate Then GoTo Nextone
If MsgBox("aborted/reshchduled?" & vbNewLine & "Amx date: " & amxdate & " site date: " & sitedate, vbYesNo + vbQuestion) = vbYes Then
Aborted:
ActiveCell.Offset(0, 4).Value = "Del"
End If




Nextone:
ActiveCell.Offset(1, 0).Select

Loop
Cleanup:
MsgBox ("Cleared")

End With
' cleaning up memory
IE.Quit
Set IE = Nothing


MsgBox "yippy"

End Sub

'Function HLink(rng As Range) As String
'extract URL from hyperlink
'posted by Rick Rothstein
 ' If rng(1).Hyperlinks.Count Then HLink = rng.Hyperlinks(1).Address
'End Function

Sub updateusernames()
'we define the essential variables
Dim IE As Object
Dim form As Variant, Button As Object

'add the "Microsoft Internet Controls" reference in your VBA Project indirectly
Set IE = CreateObject("InternetExplorer.Application")
With IE

.Visible = True
.Navigate ("http://inspect.portal.manheim.co.uk/Login.aspx")

'assigning the vinput variables to the html elements of the form
Do While IE.Busy: DoEvents: Loop
   'Do Until IE.ReadyState = READYSTATE_COMPLETE: DoEvents: Loop
MsgBox ("Shall we continue1")
IE.Document.getelementsbyname("ctl00$ContentPlaceHolder1$Login1$UserName").Item.innerText = <USERNAME>
IE.Document.getelementsbyname("ctl00_ContentPlaceHolder1_Login1_Password").Item.innerText = <PASSWORD>
'IE.Document.getElementBytype("submit").Click
IE.Document.getelementsbyname("ctl00$ContentPlaceHolder1$Login1$LoginButton").Item.Click
Do While IE.Busy: DoEvents: Loop
    'Do Until IE.ReadyState = READYSTATE_COMPLETE: DoEvents: Loop

Sheets("Sheet5").Range("A2").Select
Do Until ActiveCell = ""
If ActiveCell = "" Then If MsgBox("shall we rest page?", vbYesNo + vbQuestion) = vbYes Then GoTo Cleanup
If ActiveCell.Offset(0, 8) = "done" Then GoTo Nextone
Userlink = ActiveCell.Value
.Navigate (Userlink)
If MsgBox("Shall we update usertype", vbYesNo + vbQuestion) = vbNo Then GoTo Nextone


IE.Document.getelementsbyname("ctl00_ContentPlaceHolder1_btnEditUserRoles").Item.Click
'Do While IE.Busy: DoEvents: Loop
'delay 1

'If usertype = "Admin" Then
'MsgBox (IE.document.getelementsbyname("ctl00_ContentPlaceHolder1_UpdatePanel3").Item.innertext)
MsgBox ("Shall we continue2")
IE.Document.getelementsbyname("ctl00$ContentPlaceHolder1$UserRoleList$UserRoleList_3").Item.Click
'Else
'IE.document.getelementsbyname("ctl00$ContentPlaceHolder1$UserRoleList$UserRoleList_4").Item.Click
'End If
IE.Document.getelementsbyname("ctl00$ContentPlaceHolder1$btnUpdateUserRoles").Item.Click
'Do While IE.Busy: DoEvents: Loop
'delay 1
MsgBox ("Shall we continue3")
Nextone:
ActiveCell.Offset(0, 8).Value = "done"
ActiveCell.Offset(0, 9).Value = Now
ActiveCell.Offset(1, 0).Select

'Else
'ActiveCell.Offset(1, 0).Select
'End If
Loop
Cleanup:
MsgBox ("Cleared")

End With
' cleaning up memory
IE.Quit
Set IE = Nothing


MsgBox "yippy"

End Sub
