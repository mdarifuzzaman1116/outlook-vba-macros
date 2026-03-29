' =====================================================================
' CreateTurnoverReport
' Repo: mdarifuzzaman1116/outlook-vba-macros
' Folder: outlook-vba
' Last updated: 2026-03-29
' Description: Searches marifuzzaman@oti.nyc.gov Turnover Reports folder
'              for the latest Service Desk Turnover Report email, opens
'              it in compose mode with updated date, recipients, and
'              optional DRAFT mode when no supervisor is available.
' =====================================================================

Public Sub CreateTurnoverReport()

    On Error GoTo ErrorHandler

    Dim ns As Outlook.NameSpace
    Dim rootFolder As Outlook.MAPIFolder
    Dim items As Outlook.items
    Dim olMail As Outlook.MailItem
    Dim olNewMail As Outlook.MailItem
    Dim timeRange As String
    Dim currentDate As Date
    Dim newDate As Date
    Dim i As Long
    Dim cleanBody As String
    Dim cutPos As Long
    Dim markerPos As Long
    Dim isDraft As Boolean

    currentDate = Now
    timeRange = "1900-0700"

    If TimeValue(currentDate) >= TimeValue("19:00:00") Then
        newDate = DateValue(currentDate) + 1
    Else
        newDate = DateValue(currentDate)
    End If

    Set ns = Application.GetNamespace("MAPI")

    On Error Resume Next
    Set rootFolder = ns.Folders("marifuzzaman@oti.nyc.gov").Folders("Inbox").Folders("Turnover Reports")
    On Error GoTo ErrorHandler

    If rootFolder Is Nothing Then
        MsgBox "Could not find the Turnover Reports folder.", vbExclamation
        Exit Sub
    End If

    Set items = rootFolder.items
    items.Sort "[ReceivedTime]", True

    Set olMail = Nothing
    For i = 1 To items.Count
        If TypeOf items(i) Is Outlook.MailItem Then
            If InStr(1, items(i).Subject, "Service Desk Turnover Report", vbTextCompare) > 0 Then
                Set olMail = items(i)
                Exit For
            End If
        End If
    Next i

    If olMail Is Nothing Then
        MsgBox "No email with 'Service Desk Turnover Report' was found in the Turnover Reports folder.", vbExclamation
        GoTo CleanUp
    End If

    ' Step 1 -- Confirm source email
    Dim confirm As Integer
    confirm = MsgBox( _
        "Found this email:" & vbCrLf & vbCrLf & _
        "From:     " & olMail.SenderName & vbCrLf & _
        "Subject:  " & olMail.Subject & vbCrLf & _
        "Received: " & Format(olMail.ReceivedTime, "ddd m/d/yyyy h:mm AM/PM") & vbCrLf & vbCrLf & _
        "Continue with this email?", _
        vbQuestion + vbYesNo, "Confirm Source Email")

    If confirm = vbNo Then GoTo CleanUp

    ' Step 2 -- Ask if this is a Draft (no supervisor available)
    Dim draftAnswer As Integer
    draftAnswer = MsgBox( _
        "Is there no supervisor available?" & vbCrLf & vbCrLf & _
        "Click YES to send as DRAFT to Incident Coordinators." & vbCrLf & _
        "Click NO to send normally to SIMT.", _
        vbQuestion + vbYesNo, "Send as Draft?")

    isDraft = (draftAnswer = vbYes)

    ' Grab original HTML before Forward touches it
    cleanBody = olMail.HTMLBody

    ' Strip signature and thread from lower half only
    Dim sigMarkers(3) As String
    sigMarkers(0) = "Best Regards"
    sigMarkers(1) = "-----Original Message"
    sigMarkers(2) = "Citywide Service Desk Portal"
    sigMarkers(3) = "Office of Technology &amp; Innovation"

    cutPos = 0
    Dim m As Integer
    For m = 0 To 3
        markerPos = InStr(1, cleanBody, sigMarkers(m), vbTextCompare)
        If markerPos > 0 Then
            If markerPos > Len(cleanBody) / 2 Then
                If cutPos = 0 Or markerPos < cutPos Then
                    cutPos = markerPos
                End If
            End If
        End If
    Next m

    If cutPos > 0 Then
        cleanBody = Left(cleanBody, cutPos - 1) & "</body></html>"
    End If

    ' Strip leading blank lines
    cleanBody = Replace(cleanBody, "<body>" & vbCrLf & "<p>&nbsp;</p>", "<body>")
    cleanBody = Replace(cleanBody, "<body><p>&nbsp;</p>", "<body>")
    cleanBody = Replace(cleanBody, "<body> <p>&nbsp;</p>", "<body>")
    cleanBody = Replace(cleanBody, "<body><br>", "<body>")
    cleanBody = Replace(cleanBody, "<body><br />", "<body>")
    cleanBody = Replace(cleanBody, "<body><br/>", "<body>")
    cleanBody = Replace(cleanBody, "<o:p>&nbsp;</o:p>", "")

    ' Replace main report date strings
    cleanBody = Replace(cleanBody, Format(olMail.ReceivedTime, "dddd, mmmm d, yyyy"), Format(newDate, "dddd, mmmm d, yyyy"))
    cleanBody = Replace(cleanBody, Format(olMail.ReceivedTime, "m/d/yyyy"), Format(newDate, "m/d/yyyy"))
    cleanBody = Replace(cleanBody, Format(DateValue(olMail.ReceivedTime), "dddd, mmmm d, yyyy"), Format(newDate, "dddd, mmmm d, yyyy"))

    ' Use Forward to open in compose mode
    Set olNewMail = olMail.Forward

    ' Clear all auto-populated recipients
    Dim j As Long
    For j = olNewMail.Recipients.Count To 1 Step -1
        olNewMail.Recipients.Item(j).Delete
    Next j

    ' Set subject and recipients based on draft mode
    If isDraft Then
        olNewMail.Subject = "DRAFT - Service Desk Turnover Report - " & _
                            Format(newDate, "dddd, mmmm d, yyyy") & " " & timeRange
        olNewMail.To = "incidentcoordinators@oti.nyc.gov"
        olNewMail.CC = ""
    Else
        olNewMail.Subject = "Service Desk Turnover Report - " & _
                            Format(newDate, "dddd, mmmm d, yyyy") & " " & timeRange
        olNewMail.To = "simt@oti.nyc.gov"
        olNewMail.CC = "awilliams@oti.nyc.gov; jmorrisroe@oti.nyc.gov"
    End If

    ' GetInspector forces Outlook to initialize compose window before body is set
    Dim oInspector As Outlook.Inspector
    Set oInspector = olNewMail.GetInspector

    ' Assign clean body once
    olNewMail.HTMLBody = cleanBody

    olNewMail.Display

CleanUp:
    Set ns = Nothing
    Set rootFolder = Nothing
    Set items = Nothing
    Set olMail = Nothing
    Set olNewMail = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred." & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbCritical
    Resume CleanUp

End Sub
