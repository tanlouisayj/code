Attribute VB_Name = "SLS_FeeAutomation"
Option Explicit

#If VBA7 Then
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

' ====================================================================
'  SLS FEE FORM AUTOMATION  (Outlook VBA)
'
'  What it does
'   1. Sweep (daily): finds emails with "SLS Fee Form" in the subject,
'      saves the PDF, reads it through Word, and writes one row per
'      fee line into an Excel log. Processed emails get the Outlook
'      category "SLS Logged" so they are never logged twice.
'   2. Digest (weekly by default): emails a summary of the fee forms dated this
'      month and in future months, grouped by month:
'        - fees to be booked to P&L
'        - fees sitting in Accounts Payable (GL 20122122); dropped once a later form
'          for the same borrower books the amount to P&L (part-cleared AP shows what is left)
'
'  Macros you can run by hand (Alt+F8 in Outlook):
'   TestOnePdf     - test the PDF reading on one saved form
'   RunSweepNow    - run the sweep now
'   SendDigest     - build the digest email now
'   DigestIfDue    - send the digest if its scheduled time has passed and it hasn't gone out
'   SetupSchedule  - create the daily/weekly reminders (run once)
'   ResetLoggedFlags - make the sweep re-log recent emails (after starting a new log)
' ====================================================================

' ------------------------- SETTINGS: EDIT THESE -------------------------

' Outlook folder holding the fee form emails.
'   ""                       = your Inbox
'   "Inbox\SLS Fee Forms"    = a subfolder of your Inbox
Public Const MAIL_FOLDER As String = ""

' Shared mailbox display name, exactly as shown in Outlook ("" = your own mailbox)
Public Const MAILBOX_NAME As String = ""

Public Const SUBJECT_KEYWORD As String = "SLS Fee Form"   ' looked for in the subject AND in attachment names
Public Const LOOKBACK_DAYS As Long = 30          ' how far back each sweep looks

' Folder where the PDFs are saved (created automatically)
Public Const PDF_FOLDER As String = "C:\SLS Fee Forms\PDFs\"

' Optional: the SharePoint/OneDrive web address of that SAME folder. If set, the "Source PDF"
' column becomes a clickable link to each PDF online. PDF_FOLDER must be a OneDrive/SharePoint
' synced folder for this to work, e.g.
' "https://yourcompany-my.sharepoint.com/personal/you_company_com/Documents/Fee Forms/SLS Fee Forms"
Public Const PDF_WEB_FOLDER As String = ""

' The Excel log. Any of these work:
'   - a local path:            "C:\SLS Fee Forms\SLS Fee Log.xlsx"
'   - a OneDrive-synced path:  "C:\Users\you\OneDrive - Company\Team Site - Documents\SLS Fee Log.xlsx"
'   - the SharePoint web address of an EXISTING file, as given by Excel's
'     File > Info > Copy path, e.g. "https://yourcompany.sharepoint.com/sites/Team/Shared Documents/SLS Fee Log.xlsx"
Public Const LOG_PATH As String = "C:\SLS Fee Forms\SLS Fee Log.xlsx"
Public Const LOG_SHEET As String = "FeeLog"

' Digest email
Public Const DIGEST_TO As String = "name1@company.com; name2@company.com"
Public Const DIGEST_CC As String = ""
Public Const DIGEST_AUTO_SEND As Boolean = True   ' True = sends by itself; False = opens a draft to check
Public Const DIGEST_SKIP_IF_EMPTY As Boolean = True ' True = don't auto-send an email with nothing in it
Public Const DIGEST_MONTHS_BACK As Long = 0       ' 0 = fee forms dated this month onwards; 1 = also last month
Public Const LOG_LINK As String = ""              ' optional: overrides the automatic link to the log

' Schedule (used by SetupSchedule)
Public Const SWEEP_TIME As String = "08:30"       ' daily sweep, Mon-Fri
Public Const DIGEST_TIME As String = "09:00"
Public Const DIGEST_DAY_MASK As Long = 2          ' Mon=2 Tue=4 Wed=8 Thu=16 Fri=32 (add to combine)

' ----------------------------- END SETTINGS -----------------------------

Public Const SWEEP_SUBJECT As String = "SLS Fee Sweep"
Public Const DIGEST_SUBJECT As String = "SLS Fee Digest"

Private Const CAT_DONE As String = "SLS Logged"
Private Const CAT_REVIEW As String = "SLS Review"
Private Const AP_GL As String = "20122122"
Private Const APP_KEY As String = "SLSFeeAutomation"
Private Const XL_UP As Long = -4162
Private Const COLS As Long = 15

Private Function LogHeaders() As Variant
    LogHeaders = Array("Logged On", "Email Received", "Form Date", "Segment", _
        "Borrower", "Deal Details", "Syndications OIC/2IC", "GL A/c", "GL Name", "Ccy", "Amount", _
        "Status", "Check", "Source PDF", "Email Subject")
End Function

' ========================================================================
'  SWEEP
' ========================================================================

Public Sub RunSweepNow()
    DoSweep False
End Sub

' Called when Outlook starts: sweeps only if today's sweep hasn't run yet
Public Sub SweepIfDue()
    If GetSetting(APP_KEY, "Run", "LastSweep", "") <> Format(Date, "yyyy-mm-dd") Then DoSweep True
End Sub

Public Sub DoSweep(ByVal silent As Boolean)
    Dim fld As Object, itms As Object, itm As Object
    Dim i As Long, cutoff As Date, errMsg As String
    Dim wdApp As Object, xl As Object, wb As Object, ws As Object
    Dim xlCreated As Boolean, wbOpened As Boolean
    Dim nForms As Long, nRows As Long, nReview As Long

    Set fld = GetFeeFolder()
    If fld Is Nothing Then Exit Sub

    On Error GoTo Fail
    EnsureFolder PdfDir()
    cutoff = Now - LOOKBACK_DAYS
    Set itms = fld.Items
    itms.Sort "[ReceivedTime]", True

    For i = 1 To itms.Count
        Set itm = itms(i)
        If TypeName(itm) = "MailItem" Then
            If itm.ReceivedTime < cutoff Then Exit For
            If IsFeeFormMail(itm) Then
                If InStr(1, itm.Categories, CAT_DONE, vbTextCompare) = 0 Then
                    If wdApp Is Nothing Then Set wdApp = StartWord()
                    If ws Is Nothing Then Set ws = OpenLog(xl, wb, xlCreated, wbOpened)
                    If Not ProcessMail(itm, wdApp, ws, nRows) Then nReview = nReview + 1
                    nForms = nForms + 1
                End If
            End If
        End If
    Next i
    GoTo Cleanup

Fail:
    errMsg = Err.Description
    Resume Cleanup

Cleanup:
    On Error Resume Next
    If Not wb Is Nothing Then
        wb.Save
        If wbOpened Then wb.Close False
    End If
    If xlCreated Then xl.Quit
    If Not wdApp Is Nothing Then wdApp.Quit 0
    On Error GoTo 0

    If Len(errMsg) > 0 Then
        MsgBox "The SLS fee sweep stopped with an error:" & vbCrLf & errMsg, vbExclamation, "SLS fee sweep"
        Exit Sub
    End If

    SaveSetting APP_KEY, "Run", "LastSweep", Format(Date, "yyyy-mm-dd")
    If (Not silent) Or nReview > 0 Then
        MsgBox nForms & " new fee form email(s) processed, " & nRows & " row(s) added to the log." & _
            IIf(nReview > 0, vbCrLf & vbCrLf & nReview & " email(s) need checking - see the 'Check' column " & _
            "in the log and the '" & CAT_REVIEW & "' category in Outlook.", ""), _
            IIf(nReview > 0, vbExclamation, vbInformation), "SLS fee sweep"
    End If
End Sub

' True if the keyword is in the email subject OR in the name of any attachment
Private Function IsFeeFormMail(ByVal m As Object) As Boolean
    Dim att As Object
    If SubjectMatches(m) Then
        IsFeeFormMail = True
        Exit Function
    End If
    On Error Resume Next
    For Each att In m.Attachments
        If NameMatches(att.FileName) Then
            IsFeeFormMail = True
            Exit Function
        End If
    Next att
End Function

Private Function SubjectMatches(ByVal m As Object) As Boolean
    SubjectMatches = (InStr(1, m.Subject, SUBJECT_KEYWORD, vbTextCompare) > 0)
End Function

' Attachment name match that ignores spaces, underscores and hyphens
' (so "SLS_Fee_Form_ABC.pdf" and "SLS Fee Form - ABC.pdf" both match)
Private Function NameMatches(ByVal fileName As String) As Boolean
    NameMatches = (InStr(1, SquashName(fileName), SquashName(SUBJECT_KEYWORD), vbTextCompare) > 0)
End Function

Private Function SquashName(ByVal s As String) As String
    SquashName = Replace(Replace(Replace(s, " ", ""), "_", ""), "-", "")
End Function

' Saves each PDF on the email, logs it, and tags the email. Returns False if anything needs checking.
Private Function ProcessMail(ByVal m As Object, ByVal wdApp As Object, ByVal ws As Object, _
                             ByRef nRows As Long) As Boolean
    Dim att As Object, p As String, ok As Boolean, anyPdf As Boolean, subjOk As Boolean
    ok = True
    subjOk = SubjectMatches(m)
    For Each att In m.Attachments
        If LCase$(Right$(att.FileName, 4)) = ".pdf" And (subjOk Or NameMatches(att.FileName)) Then
            anyPdf = True
            p = PdfDir() & Format(m.ReceivedTime, "yyyymmdd-hhnnss") & "_" & SafeName(att.FileName)
            att.SaveAsFile p
            If Not LogForm(p, m, wdApp, ws, nRows) Then ok = False
        End If
    Next att
    If Not anyPdf Then ok = False
    AddCategory m, CAT_DONE
    If Not ok Then AddCategory m, CAT_REVIEW
    m.Save
    ProcessMail = ok
End Function

' Reads one PDF and writes its fee lines to the log. Returns False if the Check column has a note.
Private Function LogForm(ByVal pdfPath As String, ByVal m As Object, ByVal wdApp As Object, _
                         ByVal ws As Object, ByRef nRows As Long) As Boolean
    Dim f As Object, ln As Variant, fd As Variant, chk As String

    Set f = ParseFeeForm(pdfPath, wdApp)
    chk = f("check")
    If IsDate(f("date")) Then fd = CDate(f("date")) Else fd = f("date")

    If f("lines").Count = 0 Then
        WriteRow ws, Array(Now, m.ReceivedTime, fd, f("segment"), f("borrower"), _
            f("deal"), f("oic"), "", "", "", "", "Needs review", chk, pdfPath, m.Subject)
        LinkSource ws, pdfPath
        nRows = nRows + 1
    Else
        For Each ln In f("lines")
            WriteRow ws, Array(Now, m.ReceivedTime, fd, f("segment"), f("borrower"), _
                f("deal"), f("oic"), ln(0), ln(1), ln(2), ln(3), FeeStatus(CStr(ln(0))), chk, _
                pdfPath, m.Subject)
            LinkSource ws, pdfPath
            nRows = nRows + 1
        Next ln
    End If
    LogForm = (Len(chk) = 0)
End Function

' Removes the "SLS Logged" / "SLS Review" categories from recent fee form emails,
' so the next sweep logs them again (use after starting a fresh log file)
Public Sub ResetLoggedFlags()
    Dim fld As Object, itms As Object, itm As Object, i As Long, n As Long, cutoff As Date
    If MsgBox("This removes the '" & CAT_DONE & "' and '" & CAT_REVIEW & "' categories from fee form emails " & _
        "received in the last " & LOOKBACK_DAYS & " days, so the next sweep logs them again." & vbCrLf & vbCrLf & _
        "Only do this after starting a fresh log file. Continue?", vbYesNo + vbQuestion, _
        "SLS fee automation") <> vbYes Then Exit Sub
    Set fld = GetFeeFolder()
    If fld Is Nothing Then Exit Sub
    cutoff = Now - LOOKBACK_DAYS
    Set itms = fld.Items
    itms.Sort "[ReceivedTime]", True
    For i = 1 To itms.Count
        Set itm = itms(i)
        If TypeName(itm) = "MailItem" Then
            If itm.ReceivedTime < cutoff Then Exit For
            If IsFeeFormMail(itm) Then
                If Len(itm.Categories) > 0 Then
                    RemoveCategory itm, CAT_DONE
                    RemoveCategory itm, CAT_REVIEW
                    itm.Save
                    n = n + 1
                End If
            End If
        End If
    Next i
    SaveSetting APP_KEY, "Run", "LastSweep", ""
    MsgBox "Reset " & n & " email(s). Run RunSweepNow to log them again.", vbInformation, "SLS fee automation"
End Sub

Private Sub RemoveCategory(ByVal m As Object, ByVal cat As String)
    Dim parts() As String, i As Long, out As String
    If Len(m.Categories) = 0 Then Exit Sub
    parts = Split(Replace(m.Categories, ";", ","), ",")
    For i = 0 To UBound(parts)
        If StrComp(Trim$(parts(i)), cat, vbTextCompare) <> 0 Then
            If Len(out) > 0 Then out = out & ", "
            out = out & Trim$(parts(i))
        End If
    Next i
    m.Categories = out
End Sub

Private Function FeeStatus(ByVal gl As String) As String
    If gl = AP_GL Then
        FeeStatus = "Upcoming - pending syndication"
    ElseIf gl = "GST" Then
        FeeStatus = "GST"
    Else
        FeeStatus = "Consolidated - P&L"
    End If
End Function

' ========================================================================
'  PDF READING (opens the PDF in Word, which converts it to editable tables)
' ========================================================================

' Returns a Dictionary: date, borrower, deal, oic, segment, lines (Collection of
' Array(GL, GL name, Ccy, Amount)), check (text; empty = all good)
Private Function ParseFeeForm(ByVal pdfPath As String, ByVal wdApp As Object) As Object
    Dim res As Object, doc As Object, t As Object, c As Object
    Dim hdr As Object, fee As Object
    Dim vals As Collection, lines As Collection, rc As Collection
    Dim rowMap As Object, k As Variant
    Dim s As String, rowTxt As String, j As Long, glIdx As Long
    Dim amt As Double, sv As Double, hasTotal As Boolean, anyNum As Boolean
    Dim rr As Collection, ccyL As Collection, amtL As Collection
    Dim sums As Object, totals As Object, p As Long, cc As String, s2 As String
    Dim check As String, errD As String
    Dim lbl As String, v As String

    Set res = CreateObject("Scripting.Dictionary")
    Set lines = New Collection
    Set vals = New Collection
    res("date") = "": res("borrower") = "": res("deal") = "": res("oic") = ""
    res("segment") = ""

    On Error GoTo Fail
    Set doc = OpenPdfInWord(wdApp, pdfPath)

    ' The whole "To:" line, e.g. "IBG - Australia: Mark Koh / Damien Tham"
    ParseToLine doc.Content.Text, res

    ' Find the header table (Date / Borrower...) and the fee table (GL Name...)
    For Each t In doc.Tables
        s = LCase$(t.Range.Text)
        If hdr Is Nothing And InStr(s, "borrower") > 0 Then Set hdr = t
        If fee Is Nothing And InStr(s, "gl name") > 0 Then Set fee = t
    Next t

    ' Header table, row by row: Date / Borrower Name / Deal Details, then the
    ' Syndications OIC/2IC names (the box under Deal Details, which may have no label)
    If Not hdr Is Nothing Then
        Set rowMap = CreateObject("Scripting.Dictionary")
        For Each c In hdr.Range.Cells
            If Not rowMap.Exists(c.RowIndex) Then rowMap.Add c.RowIndex, New Collection
            rowMap(c.RowIndex).Add CleanCell(c.Range.Text)
        Next c
        For Each k In rowMap.Keys
            Set rc = rowMap(k)
            lbl = "": v = ""
            If rc.Count >= 2 Then
                lbl = LCase$(rc(1)): v = rc(rc.Count)
            ElseIf rc.Count = 1 Then
                v = rc(1)
            End If
            If Left$(lbl, 4) = "date" Then
                res("date") = v
            ElseIf InStr(lbl, "borrower") > 0 Then
                res("borrower") = v
            ElseIf InStr(lbl, "deal") > 0 Then
                res("deal") = v
            ElseIf Len(v) > 0 Then
                If Len(res("oic")) = 0 Then res("oic") = v Else res("oic") = res("oic") & "; " & v
            End If
        Next k
    Else
        check = AddNote(check, "Header table not found")
    End If

    ' Fee table: group cells by row (works even with merged cells).
    ' A Ccy / Amount box can stack two currencies, so each box is split into its lines
    ' and every currency becomes its own row in the log.
    If Not fee Is Nothing Then
        Set rowMap = CreateObject("Scripting.Dictionary")
        Set sums = CreateObject("Scripting.Dictionary")
        Set totals = CreateObject("Scripting.Dictionary")
        For Each c In fee.Range.Cells
            If Not rowMap.Exists(c.RowIndex) Then rowMap.Add c.RowIndex, New Collection
            rowMap(c.RowIndex).Add c.Range.Text
        Next c

        For Each k In rowMap.Keys
            Set rr = rowMap(k)
            If rr.Count >= 3 Then
                glIdx = 0
                rowTxt = ""
                For j = 1 To rr.Count
                    s2 = CleanCell(rr(j))
                    rowTxt = rowTxt & " " & LCase$(s2)
                    If glIdx = 0 And IsGLAccount(s2) Then glIdx = j
                Next j

                ' Last box = amount(s), second-last box = currency(ies); "-" boxes are skipped
                Set ccyL = SplitCell(rr(rr.Count - 1))
                Set amtL = SplitCell(rr(rr.Count))
                anyNum = False
                For p = 1 To amtL.Count
                    If TryNumber(amtL(p), amt) Then anyNum = True
                Next p

                If anyNum Then
                    If glIdx > 0 And glIdx + 3 > rr.Count Then
                        check = AddNote(check, "Row for GL " & CleanCell(rr(glIdx)) & " not read clearly")
                    ElseIf ccyL.Count <> amtL.Count And ccyL.Count <> 1 Then
                        check = AddNote(check, "Currency and amount lines don't match in one row")
                    Else
                        For p = 1 To amtL.Count
                            If TryNumber(amtL(p), amt) Then
                                If ccyL.Count = 1 Then cc = ccyL(1) Else cc = ccyL(p)
                                cc = UCase$(Trim$(cc))
                                If amt <> 0 Then
                                    If glIdx > 0 Then
                                        lines.Add Array(CleanCell(rr(glIdx)), CleanCell(rr(glIdx + 1)), cc, amt)
                                        sums(cc) = sums(cc) + amt
                                    ElseIf InStr(rowTxt, "total") > 0 Then
                                        totals(cc) = amt
                                        hasTotal = True
                                    ElseIf InStr(rowTxt, "gst") > 0 Then
                                        lines.Add Array("GST", "GST", cc, amt)
                                        sums(cc) = sums(cc) + amt
                                    End If
                                End If
                            End If
                        Next p
                    End If
                End If
            End If
        Next k
    Else
        check = AddNote(check, "Fee table not found")
    End If

    doc.Close 0
    Set doc = Nothing

    ' Sanity checks
    If lines.Count = 0 Then check = AddNote(check, "No fee amounts found")
    If hasTotal Then
        For Each k In totals.Keys
            sv = 0
            If sums.Exists(k) Then sv = sums(k)
            If Abs(sv - totals(k)) > 0.01 Then check = AddNote(check, k & " lines add to " & _
                Format(sv, "#,##0.00") & " but form total is " & Format(totals(k), "#,##0.00"))
        Next k
        For Each k In sums.Keys
            If Not totals.Exists(k) Then check = AddNote(check, "No " & k & " total on the form")
        Next k
    ElseIf lines.Count > 0 Then
        check = AddNote(check, "Total not found")
    End If
    If Not IsDate(res("date")) Then check = AddNote(check, "Date not read")
    If Len(res("borrower")) = 0 Then check = AddNote(check, "Borrower not read")
    If Len(res("oic")) = 0 Then check = AddNote(check, "Syndications OIC/2IC not read")
    If Len(res("segment")) = 0 Then check = AddNote(check, "To line not read")

    Set res("lines") = lines
    res("check") = check
    Set ParseFeeForm = res
    Exit Function

Fail:
    errD = Err.Description
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close 0
    Set res("lines") = lines
    res("check") = AddNote(check, "Could not read PDF: " & errD)
    Set ParseFeeForm = res
End Function

Private Sub ParseToLine(ByVal s As String, ByVal res As Object)
    Dim paras() As String, i As Long, p As String
    paras = Split(s, vbCr)
    For i = 0 To UBound(paras)
        p = LTrim$(Replace(paras(i), Chr(160), " "))
        If LCase$(Left$(p, 3)) = "to:" Then
            res("segment") = CleanCell(Mid$(p, 4))
            ' If the text sits on the next line, take that instead
            If Len(res("segment")) = 0 And i < UBound(paras) Then res("segment") = CleanCell(paras(i + 1))
            Exit Sub
        End If
    Next i
End Sub

' Splits one table box into its separate lines (a box can stack two currencies or amounts)
Private Function SplitCell(ByVal raw As String) As Collection
    Dim out As New Collection, parts() As String, i As Long, s As String
    raw = Replace(raw, Chr(13) & Chr(7), "")
    raw = Replace(raw, Chr(7), "")
    raw = Replace(raw, Chr(13), "|")
    raw = Replace(raw, Chr(11), "|")
    raw = Replace(raw, vbLf, "|")
    parts = Split(raw, "|")
    For i = 0 To UBound(parts)
        s = CleanCell(parts(i))
        If Len(s) > 0 Then out.Add s
    Next i
    ' Two values wrapped onto one line, e.g. "AUD USD" or "450,000 200,000"
    If out.Count = 1 Then
        If InStr(out(1), " ") > 0 Then
            parts = Split(out(1), " ")
            Set out = New Collection
            For i = 0 To UBound(parts)
                If Len(parts(i)) > 0 Then out.Add parts(i)
            Next i
        End If
    End If
    Set SplitCell = out
End Function

' Removes Word cell markers, line breaks and footnote markers like (3)
Private Function CleanCell(ByVal s As String) As String
    s = Replace(s, Chr(13) & Chr(7), "")
    s = Replace(s, Chr(7), "")
    s = Replace(s, Chr(13), " ")
    s = Replace(s, Chr(11), " ")
    s = Replace(s, Chr(160), " ")
    s = Replace(s, vbTab, " ")
    With CreateObject("VBScript.RegExp")
        .Global = True
        .Pattern = "\s*\(\d\)"
        s = .Replace(s, "")
        .Pattern = "\s{2,}"
        s = .Replace(s, " ")
    End With
    CleanCell = Trim$(s)
End Function

Private Function IsHeaderLabel(ByVal s As String) As Boolean
    Select Case LCase$(s)
        Case "date", "borrower name", "borrower", "deal details", "deal detail"
            IsHeaderLabel = True
    End Select
End Function

Private Function IsGLAccount(ByVal s As String) As Boolean
    If Len(s) >= 7 And Len(s) <= 10 Then IsGLAccount = (s Like String(Len(s), "#"))
End Function

Private Function TryNumber(ByVal s As String, ByRef d As Double) As Boolean
    s = Replace(Replace(s, ",", ""), " ", "")
    If Len(s) > 0 And IsNumeric(s) Then d = CDbl(s): TryNumber = True
End Function

' Test the PDF reading on one saved form, without writing anything
Public Sub TestOnePdf()
    Dim p As String, wdApp As Object, f As Object, ln As Variant, msg As String
    p = InputBox("Paste the full path of a saved SLS fee form PDF" & vbCrLf & _
        "(tip: Shift + right-click the file > Copy as path):", "Test PDF reading")
    p = Trim$(Replace(p, """", ""))
    If Len(p) = 0 Then Exit Sub

    Set wdApp = StartWord()
    Set f = ParseFeeForm(p, wdApp)
    wdApp.Quit 0

    msg = "Segment: " & f("segment") & vbCrLf & _
          "Date: " & f("date") & vbCrLf & _
          "Borrower: " & f("borrower") & vbCrLf & _
          "Deal: " & f("deal") & vbCrLf & _
          "Syndications OIC/2IC: " & f("oic") & vbCrLf & vbCrLf & "Fee lines:" & vbCrLf
    For Each ln In f("lines")
        msg = msg & "   " & ln(0) & " | " & ln(1) & " | " & ln(2) & " " & Format(ln(3), "#,##0.00") & vbCrLf
    Next ln
    msg = msg & vbCrLf & "Check: " & IIf(Len(f("check")) = 0, "OK", f("check"))
    MsgBox msg, vbInformation, "SLS fee form test"
End Sub

' ========================================================================
'  DIGEST EMAIL
' ========================================================================

' Sends the digest if its scheduled time has passed and it hasn't gone out since.
' Called at Outlook start-up and by the scheduled reminder, so opening Outlook late
' (e.g. at 9:05) still sends it, and it is never sent twice for the same slot.
Public Sub DigestIfDue()
    Dim last As String, dueAt As Date, d As Long, dt As Date, bit As Long
    last = GetSetting(APP_KEY, "Run", "LastDigest", "")
    If Len(last) = 0 Then
        ' First time: start counting from now, so nothing is sent by surprise
        SaveSetting APP_KEY, "Run", "LastDigest", Format(Now, "yyyy-mm-dd hh:nn:ss")
        Exit Sub
    End If
    ' Find the most recent scheduled slot that has already passed
    For d = 0 To 7
        dt = Date - d
        bit = CLng(2 ^ (Weekday(dt, vbSunday) - 1))      ' Sun=1, Mon=2, Tue=4 ...
        If (DIGEST_DAY_MASK And bit) <> 0 Then
            dueAt = dt + TimeValue(DIGEST_TIME)
            If dueAt <= Now Then
                If CDate(last) < dueAt Then SendDigest
                Exit For
            End If
        End If
    Next d
End Sub

Public Sub SendDigest()
    Dim xl As Object, wb As Object, ws As Object, xlCreated As Boolean, wbOpened As Boolean
    Dim data As Variant, r As Long, lastRow As Long, fromDate As Date
    Dim upcoming As New Collection, toPL As New Collection
    Dim st As String, bk As String, mail As Object, body As String, errMsg As String
    Dim logUrl As String, fromTxt As String
    Dim remain() As Double, nPL As Long, plKey() As String, plDate() As Date, plAmt() As Double
    Dim i As Long, take As Double, leftAmt As Double, ck As String

    On Error GoTo Fail
    If DIGEST_AUTO_SEND And InStr(1, DIGEST_TO, "name1@company.com", vbTextCompare) > 0 Then
        Err.Raise vbObjectError + 1002, "SLS_FeeAutomation", _
            "DIGEST_TO still has the example addresses. Put the real recipients in the settings first."
    End If

    ' Only fee forms dated from the start of this month (or earlier, if DIGEST_MONTHS_BACK > 0)
    fromDate = DateSerial(Year(Date), Month(Date) - DIGEST_MONTHS_BACK, 1)
    fromTxt = Format(fromDate, "mmm yyyy")

    Set ws = OpenLog(xl, wb, xlCreated, wbOpened)
    lastRow = ws.Cells(ws.Rows.Count, 2).End(XL_UP).Row
    If lastRow >= 2 Then
        data = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, COLS)).Value
        ' AP fees are reduced by later P&L fees for the same borrower and currency.
        ' P&L amounts booked on a LATER form are applied against the AP lines, oldest first.
        ' A part-cleared AP line shows only what is left; a fully cleared one drops off.
        ReDim remain(1 To UBound(data, 1))
        ReDim plKey(1 To UBound(data, 1))
        ReDim plDate(1 To UBound(data, 1))
        ReDim plAmt(1 To UBound(data, 1))
        For r = 1 To UBound(data, 1)
            If IsNumeric(data(r, 11)) Then remain(r) = CDbl(data(r, 11))
            If Left$(CStr(data(r, 12)), 12) = "Consolidated" And IsDate(data(r, 2)) And IsNumeric(data(r, 11)) Then
                nPL = nPL + 1
                plKey(nPL) = LCase$(Trim$(CStr(data(r, 5)))) & "|" & UCase$(Trim$(CStr(data(r, 10))))
                plDate(nPL) = CDate(data(r, 2))
                plAmt(nPL) = CDbl(data(r, 11))
            End If
        Next r
        For r = 1 To UBound(data, 1)
            If Left$(CStr(data(r, 12)), 8) = "Upcoming" And IsDate(data(r, 2)) And remain(r) > 0 Then
                ck = LCase$(Trim$(CStr(data(r, 5)))) & "|" & UCase$(Trim$(CStr(data(r, 10))))
                leftAmt = remain(r)
                For i = 1 To nPL
                    If leftAmt <= 0.005 Then Exit For
                    If plKey(i) = ck And plAmt(i) > 0.005 And plDate(i) > CDate(data(r, 2)) Then
                        take = plAmt(i)
                        If take > leftAmt Then take = leftAmt
                        leftAmt = leftAmt - take
                        plAmt(i) = plAmt(i) - take
                    End If
                Next i
                remain(r) = leftAmt
            End If
        Next r

        For r = 1 To UBound(data, 1)
            If InDigestWindow(data(r, 3), fromDate) Then
                st = CStr(data(r, 12))
                If Left$(st, 8) = "Upcoming" Then
                    If remain(r) > 0.005 Or Not IsNumeric(data(r, 11)) Then
                        upcoming.Add DigestLine(data, r, remain(r))
                    End If
                ElseIf Left$(st, 12) = "Consolidated" Then
                    toPL.Add DigestLine(data, r)
                End If
            End If
        Next r
    End If

    ' Nothing to report: don't send an empty email by itself
    If DIGEST_AUTO_SEND And DIGEST_SKIP_IF_EMPTY And upcoming.Count = 0 And toPL.Count = 0 Then
        SaveSetting APP_KEY, "Run", "LastDigest", Format(Now, "yyyy-mm-dd hh:nn:ss")
        GoTo Cleanup
    End If

    body = "<div style=""font-family:Calibri,Arial,sans-serif;font-size:11pt"">" & _
        "<p>Hi all,</p><p>Here are the SLS loan fees on fee forms dated " & fromTxt & " onwards.</p>" & _
        BuildSection(toPL, "Fees to be booked to P&L (" & fromTxt & " onwards)", "") & _
        BuildSection(upcoming, "Fees in Accounts Payable, pending syndication (" & fromTxt & " onwards)", "")
    logUrl = LogLinkUrl(wb)
    If Len(logUrl) > 0 Then body = body & "<p>Full log: <a href=""" & logUrl & """>SLS Fee Log</a></p>"
    body = body & "<p style=""color:#777777;font-size:9pt"">Generated automatically from SLS fee forms.</p></div>"

    Set mail = Application.CreateItem(0)
    mail.To = DIGEST_TO
    mail.CC = DIGEST_CC
    mail.Subject = "SLS fee form summary (" & fromTxt & " onwards) - To book to P&L: " & toPL.Count & _
        " | In AP: " & upcoming.Count
    mail.HTMLBody = body
    If DIGEST_AUTO_SEND Then mail.Send Else mail.Display
    SaveSetting APP_KEY, "Run", "LastDigest", Format(Now, "yyyy-mm-dd hh:nn:ss")
    GoTo Cleanup

Fail:
    errMsg = Err.Description
    Resume Cleanup

Cleanup:
    On Error Resume Next
    If Not wb Is Nothing Then
        If wbOpened Then wb.Close False
    End If
    If xlCreated Then xl.Quit
    On Error GoTo 0
    If Len(errMsg) > 0 Then MsgBox "The SLS fee digest stopped with an error:" & vbCrLf & errMsg, _
        vbExclamation, "SLS fee digest"
End Sub

' True if the form date is on/after fromDate (a form whose date couldn't be read is kept in)
Private Function InDigestWindow(ByVal v As Variant, ByVal fromDate As Date) As Boolean
    If IsDate(v) Then InDigestWindow = (CDate(v) >= fromDate) Else InDigestWindow = True
End Function

' Link to the Excel log for the digest. Uses LOG_LINK if you set one; otherwise the web
' address Excel reports for a OneDrive/SharePoint-synced file; otherwise a network-drive link.
Private Function LogLinkUrl(ByVal wb As Object) As String
    Dim u As String
    If Len(LOG_LINK) > 0 Then
        LogLinkUrl = LOG_LINK
        Exit Function
    End If
    On Error Resume Next
    u = wb.FullName
    On Error GoTo 0
    If LCase$(Left$(u, 4)) = "http" Then
        LogLinkUrl = u
    ElseIf LCase$(Left$(LOG_PATH, 4)) = "http" Then
        LogLinkUrl = LOG_PATH
    ElseIf Left$(LOG_PATH, 2) = "\\" Then
        LogLinkUrl = "file:///" & Replace(Replace(LOG_PATH, "\", "/"), " ", "%20")
    End If
End Function

' Array(form date, borrower, segment, OIC/2IC, GL name, ccy, amount, sort key, month label)
Private Function DigestLine(d As Variant, ByVal r As Long, Optional ByVal amtOverride As Double = -1) As Variant
    Dim fd As String, amt As Double, sk As Double, ml As String
    If IsDate(d(r, 3)) Then
        fd = Format(CDate(d(r, 3)), "dd-mmm-yyyy")
        sk = CDbl(CDate(d(r, 3)))
        ml = Format(CDate(d(r, 3)), "mmmm yyyy")
    Else
        fd = CStr(d(r, 3))
        sk = 99999
        ml = "Form date not read"
    End If
    If IsNumeric(d(r, 11)) Then amt = CDbl(d(r, 11))
    If amtOverride >= 0 And IsNumeric(d(r, 11)) Then amt = amtOverride
    DigestLine = Array(fd, CStr(d(r, 5)), CStr(d(r, 4)), CStr(d(r, 7)), CStr(d(r, 9)), CStr(d(r, 10)), amt, sk, ml)
End Function

' One table: rows sorted by form date, grouped under a heading per month,
' with totals per currency for each month (and overall if there is more than one month)
Private Function BuildSection(ByVal items As Collection, ByVal title As String, ByVal note As String) As String
    Dim h As String, it As Variant, tmp As Variant, arr() As Variant
    Dim i As Long, j As Long, n As Long, nMonths As Long, cur As String, cc As String
    Dim monTot As Object, allTot As Object

    h = "<h3 style=""font-family:Calibri,Arial,sans-serif;color:#c0392b;margin:18px 0 4px"">" & HtmlEsc(title) & "</h3>"
    If Len(note) > 0 Then h = h & "<p style=""font-size:10pt;color:#555555;margin:0 0 6px"">" & HtmlEsc(note) & "</p>"
    If items.Count = 0 Then
        BuildSection = h & "<p>None.</p>"
        Exit Function
    End If

    ' Sort by form date (oldest first)
    n = items.Count
    ReDim arr(1 To n)
    For i = 1 To n
        arr(i) = items(i)
    Next i
    For i = 2 To n
        tmp = arr(i)
        j = i - 1
        Do While j >= 1
            If arr(j)(7) <= tmp(7) Then Exit Do
            arr(j + 1) = arr(j)
            j = j - 1
        Loop
        arr(j + 1) = tmp
    Next i

    Set allTot = CreateObject("Scripting.Dictionary")
    h = h & "<table style=""border-collapse:collapse;font-family:Calibri,Arial,sans-serif;font-size:10.5pt"">"
    h = h & "<tr>" & HtmlCell("Form date", True) & HtmlCell("Borrower", True) & HtmlCell("Segment", True) & _
        HtmlCell("Syndications OIC/2IC", True) & HtmlCell("Fee", True) & HtmlCell("Ccy", True) & HtmlCell("Amount", True, True) & "</tr>"

    For i = 1 To n
        it = arr(i)
        If CStr(it(8)) <> cur Then
            If Len(cur) > 0 Then h = h & TotalRows(monTot, "Total " & cur)
            cur = CStr(it(8))
            nMonths = nMonths + 1
            Set monTot = CreateObject("Scripting.Dictionary")
            h = h & "<tr><td colspan=""7"" style=""border:1px solid #bfbfbf;padding:4px 8px;font-weight:bold;" & _
                "background:#fde8e4"">" & HtmlEsc(cur) & "</td></tr>"
        End If
        h = h & "<tr>" & HtmlCell(CStr(it(0))) & HtmlCell(CStr(it(1))) & HtmlCell(CStr(it(2))) & _
            HtmlCell(CStr(it(3))) & HtmlCell(CStr(it(4))) & HtmlCell(CStr(it(5))) & _
            HtmlCell(Format(it(6), "#,##0.00"), False, True) & "</tr>"
        cc = CStr(it(5))
        monTot(cc) = monTot(cc) + it(6)
        allTot(cc) = allTot(cc) + it(6)
    Next i
    h = h & TotalRows(monTot, "Total " & cur)
    If nMonths > 1 Then h = h & TotalRows(allTot, "Total, all months")
    BuildSection = h & "</table>"
End Function

' Total rows, one per currency
Private Function TotalRows(ByVal tot As Object, ByVal label As String) As String
    Dim k As Variant, h As String
    For Each k In tot.Keys
        h = h & "<tr><td colspan=""5"" style=""border:1px solid #bfbfbf;padding:4px 8px;font-weight:bold;" & _
            "background:#f2f2f2"">" & HtmlEsc(label & " - " & CStr(k)) & "</td>" & _
            HtmlCell(CStr(k), False, False, True) & _
            HtmlCell(Format(tot(k), "#,##0.00"), False, True, True) & "</tr>"
    Next k
    TotalRows = h
End Function

Private Function HtmlCell(ByVal s As String, Optional ByVal isHead As Boolean = False, _
                          Optional ByVal alignRight As Boolean = False, Optional ByVal isBold As Boolean = False) As String
    Dim st As String
    st = "border:1px solid #bfbfbf;padding:4px 8px;"
    If isHead Then st = st & "background:#c0392b;color:#ffffff;font-weight:bold;"
    If isBold Then st = st & "font-weight:bold;background:#f2f2f2;"
    If alignRight Then st = st & "text-align:right;"
    HtmlCell = "<td style=""" & st & """>" & HtmlEsc(s) & "</td>"
End Function

Private Function HtmlEsc(ByVal s As String) As String
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    HtmlEsc = s
End Function

' ========================================================================
'  SCHEDULE: creates two recurring calendar reminders that trigger the macros
' ========================================================================

Public Sub SetupSchedule()
    CreateRecurring SWEEP_SUBJECT, SWEEP_TIME, 62          ' Mon-Fri
    CreateRecurring DIGEST_SUBJECT, DIGEST_TIME, DIGEST_DAY_MASK
    MsgBox "Done. Two recurring reminders were added to your calendar:" & vbCrLf & _
        "  - '" & SWEEP_SUBJECT & "' every weekday at " & SWEEP_TIME & vbCrLf & _
        "  - '" & DIGEST_SUBJECT & "' weekly at " & DIGEST_TIME & vbCrLf & vbCrLf & _
        "They are marked Free and dismiss themselves. Don't delete them.", vbInformation, "SLS fee automation"
End Sub

Private Sub CreateRecurring(ByVal subj As String, ByVal t As String, ByVal dayMask As Long)
    Dim cal As Object, found As Object, i As Long, appt As Object, rp As Object

    ' Remove any older copy so you can rerun this safely
    Set cal = Application.GetNamespace("MAPI").GetDefaultFolder(9)   ' Calendar
    Set found = cal.Items.Restrict("[Subject] = '" & subj & "'")
    For i = found.Count To 1 Step -1
        found(i).Delete
    Next i

    Set appt = Application.CreateItem(1)      ' Appointment
    appt.Subject = subj
    appt.Body = "Used by the SLS fee automation macro. Please don't delete."
    appt.BusyStatus = 0                       ' Free
    appt.ReminderSet = True
    appt.ReminderMinutesBeforeStart = 0
    appt.ReminderPlaySound = False

    Set rp = appt.GetRecurrencePattern
    rp.RecurrenceType = 1                     ' Weekly
    rp.DayOfWeekMask = dayMask
    rp.Interval = 1
    rp.PatternStartDate = Date + 1
    rp.NoEndDate = True
    rp.StartTime = TimeValue(t)
    rp.Duration = 5
    appt.Save
End Sub

' ========================================================================
'  HELPERS
' ========================================================================

Private Function GetFeeFolder() As Object
    Dim ns As Object, f As Object, parts() As String, i As Long
    Set ns = Application.GetNamespace("MAPI")
    On Error GoTo NotFound
    If Len(MAILBOX_NAME) = 0 Then
        Set f = ns.GetDefaultFolder(6).Parent              ' root of your mailbox
    Else
        Set f = ns.Folders(MAILBOX_NAME)
    End If
    If Len(MAIL_FOLDER) = 0 Then
        Set f = f.Folders(ns.GetDefaultFolder(6).Name)     ' Inbox
    Else
        parts = Split(MAIL_FOLDER, "\")
        For i = 0 To UBound(parts)
            Set f = f.Folders(parts(i))
        Next i
    End If
    Set GetFeeFolder = f
    Exit Function
NotFound:
    MsgBox "Couldn't find the Outlook folder '" & MAIL_FOLDER & "'" & _
        IIf(Len(MAILBOX_NAME) > 0, " in mailbox '" & MAILBOX_NAME & "'", "") & _
        ". Check MAIL_FOLDER / MAILBOX_NAME at the top of the SLS_FeeAutomation module.", _
        vbExclamation, "SLS fee automation"
    Set GetFeeFolder = Nothing
End Function

Private Function OpenLog(ByRef xl As Object, ByRef wb As Object, ByRef xlCreated As Boolean, _
                         ByRef wbOpened As Boolean) As Object
    Dim w As Object, fname As String, ws As Object, h As Variant, fso As Object, isUrl As Boolean

    Set fso = CreateObject("Scripting.FileSystemObject")
    isUrl = (LCase$(Left$(LOG_PATH, 4)) = "http")      ' a SharePoint web address
    If isUrl Then fname = UrlFileName(LOG_PATH) Else fname = fso.GetFileName(LOG_PATH)

    ' Use our own hidden Excel, so we never clash with a workbook you have open
    Set xl = CreateObject("Excel.Application")
    xlCreated = True
    xl.Visible = False
    xl.DisplayAlerts = False

    If isUrl Or fso.FileExists(LOG_PATH) Then
        Set wb = OpenWorkbookRetry(xl, LOG_PATH)
    Else
        EnsureFolder fso.GetParentFolderName(LOG_PATH)
        Set wb = xl.Workbooks.Add
        wb.SaveAs LOG_PATH, 51                         ' .xlsx
    End If
    wbOpened = True
    If wb.ReadOnly Then
        Err.Raise vbObjectError + 1001, "SLS_FeeAutomation", _
            "The Excel log opened as read-only. Close it everywhere (Excel, browser, other people) and try again."
    End If

    On Error Resume Next
    Set ws = wb.Worksheets(LOG_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(wb.Worksheets(1))
        ws.Name = LOG_SHEET
    End If

    If Len(CStr(ws.Cells(1, 1).Value)) = 0 Then
        h = LogHeaders()
        ws.Range(ws.Cells(1, 1), ws.Cells(1, COLS)).Value = h
        With ws.Range(ws.Cells(1, 1), ws.Cells(1, COLS))
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(192, 57, 43)
        End With
        ws.Range("A:B").NumberFormat = "dd-mmm-yyyy hh:mm"
        ws.Range("C:C").NumberFormat = "dd-mmm-yyyy"
        ws.Range("H:H").NumberFormat = "@"
        ws.Range("K:K").NumberFormat = "#,##0.00"
        ws.Range("A:O").ColumnWidth = 16
        ws.Range("D:D").ColumnWidth = 40
        ws.Range("F:F").ColumnWidth = 50
        ws.Range("G:G").ColumnWidth = 28
        ws.Range("M:M").ColumnWidth = 30
    ElseIf CStr(ws.Cells(1, 13).Value) <> "Check" Then
        Err.Raise vbObjectError + 1000, "SLS_FeeAutomation", _
            "The Excel log still has the old column layout. Rename or move the old log file " & _
            "(or clear the sheet) and run again, so a log with the new columns is created."
    End If
    Set OpenLog = ws
End Function

' File name at the end of a web address, e.g. ".../SLS%20Fee%20Log.xlsx" -> "SLS Fee Log.xlsx"
Private Function UrlFileName(ByVal u As String) As String
    Dim p As Long
    p = InStr(u, "?")
    If p > 0 Then u = Left$(u, p - 1)
    p = InStrRev(u, "/")
    If p > 0 Then u = Mid$(u, p + 1)
    UrlFileName = Replace(u, "%20", " ")
End Function

' Opens the log, retrying if Excel is still busy starting up ("Call was rejected by callee")
Private Function OpenWorkbookRetry(ByVal xl As Object, ByVal p As String) As Object
    Dim tries As Long, n As Long, d As String
    For tries = 1 To 6
        On Error Resume Next
        Set OpenWorkbookRetry = xl.Workbooks.Open(p)
        n = Err.Number: d = Err.Description
        On Error GoTo 0
        If n = 0 Then Exit Function
        If n <> -2147418111 And n <> -2147417846 Then Exit For
        Sleep 2000
    Next tries
    Err.Raise n, "OpenWorkbookRetry", d
End Function

' Opens a PDF in Word, retrying if Word is still busy
Private Function OpenPdfInWord(ByVal wdApp As Object, ByVal p As String) As Object
    Dim tries As Long, n As Long, d As String
    For tries = 1 To 6
        On Error Resume Next
        Set OpenPdfInWord = wdApp.Documents.Open(p, False, True, False)
        n = Err.Number: d = Err.Description
        On Error GoTo 0
        If n = 0 Then Exit Function
        If n <> -2147418111 And n <> -2147417846 Then Exit For
        Sleep 2000
    Next tries
    Err.Raise n, "OpenPdfInWord", d
End Function

' If PDF_WEB_FOLDER is set, turns the Source PDF cell of the row just written into a
' clickable link to the file on SharePoint/OneDrive (showing the file name)
Private Sub LinkSource(ByVal ws As Object, ByVal pdfPath As String)
    Dim r As Long, fn As String, u As String
    If Len(PDF_WEB_FOLDER) = 0 Then Exit Sub
    On Error Resume Next
    r = ws.Cells(ws.Rows.Count, 2).End(XL_UP).Row
    fn = Mid$(pdfPath, InStrRev(pdfPath, "\") + 1)
    u = Replace(PDF_WEB_FOLDER, " ", "%20")
    Do While Right$(u, 1) = "/" Or Right$(u, 1) = "\"
        u = Left$(u, Len(u) - 1)
    Loop
    u = u & "/" & EncodeUrlPart(fn)
    ws.Hyperlinks.Add Anchor:=ws.Cells(r, 14), Address:=u, TextToDisplay:=fn
    On Error GoTo 0
End Sub

Private Function EncodeUrlPart(ByVal s As String) As String
    s = Replace(s, "%", "%25")
    s = Replace(s, " ", "%20")
    s = Replace(s, "#", "%23")
    s = Replace(s, "?", "%3F")
    s = Replace(s, "&", "%26")
    s = Replace(s, "+", "%2B")
    EncodeUrlPart = s
End Function

Private Sub WriteRow(ByVal ws As Object, ByVal v As Variant)
    Dim r As Long
    r = ws.Cells(ws.Rows.Count, 2).End(XL_UP).Row + 1
    ws.Range(ws.Cells(r, 1), ws.Cells(r, UBound(v) + 1)).Value = v
End Sub

Private Function StartWord() As Object
    Dim w As Object, i As Long
    ' Stop Word's "converting your PDF" pop-up, which would block a hidden Word
    On Error Resume Next
    CreateObject("WScript.Shell").RegWrite _
        "HKCU\Software\Microsoft\Office\16.0\Word\Options\DisableConvertPdfWarning", 1, "REG_DWORD"
    On Error GoTo 0
    Set w = CreateObject("Word.Application")
    On Error Resume Next
    For i = 1 To 6                                     ' Word may still be starting up
        Err.Clear
        w.Visible = False
        If Err.Number = 0 Then Exit For
        Sleep 1000
    Next i
    Err.Clear
    w.DisplayAlerts = 0
    On Error GoTo 0
    Set StartWord = w
End Function

Private Sub AddCategory(ByVal m As Object, ByVal cat As String)
    If InStr(1, m.Categories, cat, vbTextCompare) = 0 Then
        If Len(m.Categories) > 0 Then m.Categories = m.Categories & ", " & cat Else m.Categories = cat
    End If
End Sub

Private Function AddNote(ByVal cur As String, ByVal note As String) As String
    If Len(cur) = 0 Then AddNote = note Else AddNote = cur & "; " & note
End Function

Private Function PdfDir() As String
    If Right$(PDF_FOLDER, 1) = "\" Then PdfDir = PDF_FOLDER Else PdfDir = PDF_FOLDER & "\"
End Function

Private Function SafeName(ByVal s As String) As String
    Dim ch As Variant
    For Each ch In Array("\", "/", ":", "*", "?", """", "<", ">", "|")
        s = Replace(s, ch, "_")
    Next ch
    SafeName = s
End Function

Private Sub EnsureFolder(ByVal p As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Right$(p, 1) = "\" Then p = Left$(p, Len(p) - 1)
    If Len(p) = 0 Then Exit Sub
    If fso.FolderExists(p) Then Exit Sub
    EnsureFolder fso.GetParentFolderName(p)
    fso.CreateFolder p
End Sub
