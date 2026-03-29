' =====================================================================
' SearchAndOpenEmailInReplyAllMode
' Repo: mdarifuzzaman1116/outlook-vba-macros
' Folder: outlook-vba
' Last updated: 2026-03-29
' Description: Searches OTI Change Maintenance Notifications Sent Items
'              for a CHG number or subject line, opens the most recent
'              match in ReplyAll mode, cleans recipients, fixes subject,
'              and warns if the change is already COMPLETED.
' =====================================================================

Sub SearchAndOpenEmailInReplyAllMode()

    On Error GoTo ErrorHandler

    Dim myNamespace As Outlook.NameSpace
    Dim myFolder As Outlook.folder
    Dim mySearch As String
    Dim mySearchResult As Outlook.items
    Dim searchCriteria As String
    Dim myMail As Outlook.MailItem
    Dim ReplyMail As Outlook.MailItem
    Dim subjectWithoutRE As String
    Dim rec As Outlook.Recipient
    Dim bccRecipients As String
    Dim i As Long
    Dim bccArr As Variant

    Set myNamespace = Application.GetNamespace("MAPI")

    DoEvents

    mySearch = InputBox("Enter the subject line or change number of the email you want to search for:", "Email Search")

    If Trim(mySearch) = "" Then
        MsgBox "No search term entered. Exiting.", vbInformation
        Exit Sub
    End If

    Set myFolder = Nothing
    On Error Resume Next
    Set myFolder = myNamespace.Folders("OTI Change Maintenance Notifications").Folders("Sent Items")
    On Error GoTo ErrorHandler

    If myFolder Is Nothing Then
        MsgBox "OTI Change Maintenance Notifications mailbox not found.", vbExclamation
        Exit Sub
    End If

    If mySearch Like "CHG*" Then
        searchCriteria = "@SQL=" & Chr(34) & "urn:schemas:httpmail:subject" & Chr(34) & _
                         " LIKE '%" & mySearch & "%' AND " & _
                         Chr(34) & "urn:schemas:httpmail:date" & Chr(34) & _
                         " >= '" & Format$(Date - 30, "yyyy-mm-dd") & "' AND " & _
                         Chr(34) & "urn:schemas:httpmail:date" & Chr(34) & _
                         " <= '" & Format$(Date + 1, "yyyy-mm-dd") & "'"
    Else
        searchCriteria = "@SQL=" & Chr(34) & "urn:schemas:httpmail:subject" & Chr(34) & _
                         " LIKE '%" & mySearch & "%' AND " & _
                         Chr(34) & "urn:schemas:httpmail:date" & Chr(34) & _
                         " >= '" & Format$(Date - 30, "yyyy-mm-dd") & "' AND " & _
                         Chr(34) & "urn:schemas:httpmail:date" & Chr(34) & _
                         " <= '" & Format$(Date + 1, "yyyy-mm-dd") & "'"
    End If

    Set mySearchResult = myFolder.Items.Restrict(searchCriteria)

    If mySearchResult Is Nothing Then
        MsgBox "Search returned no results for '" & mySearch & "'.", vbInformation
        Exit Sub
    End If

    mySearchResult.Sort "[SentOn]", True

    If mySearchResult.Count > 0 Then

        Set myMail = mySearchResult.GetFirst

        If myMail Is Nothing Then
            MsgBox "Could not retrieve the email item.", vbExclamation
            Exit Sub
        End If

        ' -------------------------------------------------------
        ' CHECK IF ALREADY COMPLETED
        ' -------------------------------------------------------
        If InStr(1, myMail.Subject, "<COMPLETED>", vbTextCompare) > 0 Then

            Dim answer As Integer
            answer = MsgBox( _
                "This change is already marked COMPLETED." & vbCrLf & vbCrLf & _
                "Subject:  " & myMail.Subject & vbCrLf & _
                "Sent on: " & Format(myMail.SentOn, "dddd, mmmm d yyyy h:mm AM/PM") & vbCrLf & vbCrLf & _
                "Do you want to open it anyway?", _
                vbQuestion + vbYesNo, _
                "Already Completed")

            If answer = vbNo Then
                GoTo CleanUp
            End If

        End If
        ' -------------------------------------------------------

        Set ReplyMail = myMail.ReplyAll

        bccRecipients = myMail.BCC

        For i = ReplyMail.Recipients.Count To 1 Step -1
            Set rec = ReplyMail.Recipients.Item(i)
            If rec.Type = olTo Or rec.Type = olCC Then
                rec.Delete
            End If
        Next i

        If Trim(bccRecipients) <> "" Then
            bccArr = Split(bccRecipients, ";")
            For i = LBound(bccArr) To UBound(bccArr)
                If Trim(bccArr(i)) <> "" Then
                    Dim newRec As Outlook.Recipient
                    Set newRec = ReplyMail.Recipients.Add(Trim(bccArr(i)))
                    newRec.Type = olBCC
                End If
            Next i
        End If

        subjectWithoutRE = Trim(myMail.Subject)
        Do While InStr(1, UCase(subjectWithoutRE), "RE:") = 1
            subjectWithoutRE = Trim(Mid(subjectWithoutRE, Len("RE:") + 1))
        Loop

        subjectWithoutRE = Replace(subjectWithoutRE, "<START>", "<COMPLETED>", , , vbTextCompare)

        ReplyMail.Subject = subjectWithoutRE
        ReplyMail.SentOnBehalfOfName = "OTIChangeMaintenanceNotifications@oti.nyc.gov"
        ReplyMail.HTMLBody = myMail.HTMLBody
        ReplyMail.Display

    Else
        MsgBox "Looks like the Initial Notification for '" & mySearch & "' was not sent.", vbInformation
    End If

CleanUp:
    Set myNamespace = Nothing
    Set myFolder = Nothing
    Set mySearchResult = Nothing
    Set myMail = Nothing
    Set ReplyMail = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred." & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbCritical
    Resume CleanUp

End Sub
