Attribute VB_Name = "DextopModule"
Option Explicit

Public db As Database

Public EditTestEntry As Boolean
Public EditMasterRegisters As Boolean
Public sCurrentUserCode As String
Public sCurrentUsername As String
Public dCurrentFinancialFromDate As Date
Public dCurrentFinancialToDate As Date

 
Public Sub initialisePublicVariables()
    Set db = OpenDatabase(App.Path & "\Storage.mdb", False, False, "MS Access;PWD=12345abcde")
End Sub

Public Function isLoadingFirstTime() As Boolean
    If Dir$(App.Path & "\Load.inf") = "" Then
        On Error Resume Next
        'Create all necessary folders
        MkDir App.Path & "\Reports"
        MkDir App.Path & "\Pdf"
        MkDir App.Path & "\Logo"
        'Create the file
        Open App.Path & "\Load.inf" For Output As #1
        Close #1
    End If
    
    
End Function


Public Function getFinancialCode(dDate As Date) As Long
    getFinancialCode = IIf(dDate >= DateValue(Day(dCurrentFinancialFromDate) & "," & Month(dCurrentFinancialFromDate) & "," & Year(dDate)), Year(dDate), Year(dDate) - 1)
End Function

Public Function getFinancialStartDate(dDate As Date) As Date
    getFinancialStartDate = IIf(dDate >= DateValue(Day(dCurrentFinancialFromDate) & "," & Month(dCurrentFinancialFromDate) & "," & Year(dDate)), DateValue(Day(dCurrentFinancialFromDate) & "," & Month(dCurrentFinancialFromDate) & "," & Year(dDate)), DateValue(Day(dCurrentFinancialFromDate) & "," & Month(dCurrentFinancialFromDate) & "," & (Year(dDate) - 1)))
End Function

Public Function getFinancialEndDate(dDate As Date) As Date
    getFinancialEndDate = IIf(dDate >= DateValue(Day(dCurrentFinancialFromDate) & "," & Month(dCurrentFinancialFromDate) & "," & Year(dDate)), DateValue(Day(dCurrentFinancialFromDate - 1) & "," & Month(dCurrentFinancialFromDate - 1) & "," & Year(dDate) + 1), DateValue(Day(dCurrentFinancialFromDate - 1) & "," & Month(dCurrentFinancialFromDate - 1) & "," & Year(dDate)))
End Function

Public Function isBillDateInFinancialYear(dTransactionDate As Date) As Boolean
Dim rs As Recordset
Dim isNotCorrect As Boolean

    If dTransactionDate < dCurrentFinancialFromDate Or dTransactionDate > dCurrentFinancialToDate Then
        isNotCorrect = False
    Else
        isNotCorrect = True
    End If
    isBillDateInFinancialYear = isNotCorrect
End Function

Public Sub SendSMS(sMobileNo As String, sPatient As String)
    Dim strURL As String, sResult As String, sMessage As String
    On Error GoTo Err
    
    Dim WinHttpReq As Object
    Set WinHttpReq = CreateObject("Msxml2.XMLHTTP")
        
    sMessage = Left("Hello " & sPatient & ", your medical test report is ready to collect.", 160)
    
    strURL = "http://bulksms.mysmsmantra.com:8080/WebSMS/SMSAPI.jsp?username=chanakyalab&password=381077558&sendername=CDLREP&mobileno=91" & sMobileNo & "&message=" & sMessage
    'Replace the Message & MobileNo with the no to whom you want to send SMS
    With WinHttpReq
        .Open "GET", strURL, False
        .Send
        sResult = .responseText
    End With
    
    MsgBox "Status of Sent SMS  is=" & sResult
     
    Exit Sub
Err:
    MsgBox ("Error Sending SMS " & vbCrLf & Err.Description)
End Sub

Sub SendeMail(sMailID As String, sPatient As String, sPdf As String)
    Dim iMsg As Object
    Dim iConf As Object
    Dim Flds As Variant


On Error GoTo Err

    Set iMsg = CreateObject("CDO.Message")
    Set iConf = CreateObject("CDO.Configuration")

    iConf.Load -1
    Set Flds = iConf.Fields
    
     With Flds
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "chanakyadiagnostics"
        .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "chanakya1234"
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.gmail.com" 'smtp mail server
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25 'stmp server
        .Update
    End With

    With iMsg

        Set .Configuration = iConf
        .To = sMailID
        .From = "chanakyadiagnostics@gmail.com"
        .Subject = "Medical Test Report"
        .HTMLBody = "<pre><font size=""3"" face=""verdana"" color=""black"">Dear " & sPatient & ", <br>"
        .HTMLBody = .HTMLBody & "    Your medical test report is ready. It is attached in this email for your reference.<br>"
        .HTMLBody = .HTMLBody & "Regards,<br>"
        .HTMLBody = .HTMLBody & "   Chanakya Diagnostics Laboratory<br>"
        .HTMLBody = .HTMLBody & "   Sonari-785690<br>"
        .HTMLBody = .HTMLBody & "   Ph. 7399199668<br></pre>"
        .AddAttachment (sPdf)
        .Send
    End With
    
    Set iMsg = Nothing
    Set iConf = Nothing
    
    MsgBox "Email is Send Successfully !"
    Exit Sub
    
Err:
    MsgBox ("Sorry there is some error while sending email" & vbCrLf & Err.Description)
End Sub
