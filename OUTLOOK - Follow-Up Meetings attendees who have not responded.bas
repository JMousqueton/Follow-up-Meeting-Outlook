' Resend a meeting request to participants who have not yet answered
' Original idea @GuZeFR

Attribute VB_Name = "Module1"
Function GetCurrentItem() As Object
    Dim objApp As Outlook.Application

    Set objApp = Application
    On Error Resume Next
    Select Case TypeName(objApp.ActiveWindow)
        Case "Explorer"
            Set GetCurrentItem = objApp.ActiveExplorer.Selection.Item(1)
        Case "Inspector"
            Set GetCurrentItem = objApp.ActiveInspector.CurrentItem
    End Select

    Set objApp = Nothing
End Function
Sub SendEmailtoNoRepsonse()

Dim objApp As Outlook.Application
Dim objItem As Object
Dim objAttendees As Outlook.Recipients
Dim objAttendeeReq As String
Dim objOrganizer As String
Dim dtStart As Date
Dim dtEnd As Date
Dim strSubject As String
Dim strLocation As String
Dim strMeetStatus As String
Dim strCopyData As String



'start Core Macro
On Error Resume Next

Set objApp = CreateObject("Outlook.Application")
Set objItem = GetCurrentItem()
Set objAttendees = objItem.Recipients

' Is it an appointment
If objItem.Class <> 26 Then
  MsgBox "This only works with meetings."
  GoTo EndClean:
End If

' Get the data
dtStart = objItem.Start
dtEnd = objItem.End
strSubject = objItem.Subject
strLocation = objItem.Location
objOrganizer = objItem.Organizer
objAttendeeReq = ""

' Get The Attendee List
For x = 1 To objAttendees.Count

 ' 0 = no response, 2 = tentative, 3 = accepted, 4 = declined,
  If objAttendees(x).MeetingResponseStatus = 0 Then
    If objAttendees(x) <> objItem.Organizer Then
     objAttendeeReq = objAttendeeReq & "; " & objAttendees(x).Address
  End If
  End If

Next

 strCopyData = "Hello," & vbCrLf & "This is an automated email." & vbCrLf & _
 "I have not received a response regarding your participation to the meeting below. Please decline or accept the invitation." & vbCrLf & _
 "Have a nice day," & vbCrLf & "----- Original Appointment -----" & vbCrLf & _
 "Organizer: " & objOrganizer & vbCrLf & "Subject:  " & strSubject & _
 vbCrLf & "Where:   " & strLocation & vbCrLf & "When:    " & _
 dtStart & vbCrLf & "Ends:      " & dtEnd

Dim objOutlookRecip As Outlook.Recipient
Set listattendees = Application.CreateItem(olMailItem)
  listattendees.Body = strCopyData
  listattendees.Subject = "Please respond to: " & strSubject
  listattendees.To = objAttendeeReq

For Each objOutlookRecip In listattendees.Recipients
objOutlookRecip.Resolve
Next

  listattendees.Display

EndClean:
Set objApp = Nothing
Set objItem = Nothing
Set objAttendees = Nothing
End Sub
