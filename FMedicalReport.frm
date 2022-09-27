VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FMedicalReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medical  Report"
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15270
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial Narrow"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FMedicalReport.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FMedicalReport.frx":000C
   ScaleHeight     =   8040
   ScaleWidth      =   15270
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CShowDetailed 
      Height          =   505
      Left            =   360
      Picture         =   "FMedicalReport.frx":1FEC4E
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7320
      Width           =   1350
   End
   Begin VB.CommandButton CToExcel 
      Height          =   525
      Left            =   1770
      Picture         =   "FMedicalReport.frx":2010B0
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7305
      Width           =   1365
   End
   Begin VB.CommandButton CClose 
      Cancel          =   -1  'True
      Height          =   505
      Left            =   13605
      Picture         =   "FMedicalReport.frx":203512
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7290
      Width           =   1365
   End
   Begin MSFlexGridLib.MSFlexGrid MGrid 
      Height          =   3885
      Left            =   135
      TabIndex        =   10
      Top             =   2625
      Width           =   15030
      _ExtentX        =   26511
      _ExtentY        =   6853
      _Version        =   393216
      Rows            =   0
      Cols            =   0
      FixedRows       =   0
      FixedCols       =   0
      BackColorBkg    =   16777215
      GridColorFixed  =   8421504
      FocusRect       =   0
      SelectionMode   =   1
      BorderStyle     =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.DTPicker DTPFrom 
      Height          =   345
      Left            =   1725
      TabIndex        =   0
      Top             =   120
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   609
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   89915395
      CurrentDate     =   40458
   End
   Begin MSComCtl2.DTPicker DTPTo 
      Height          =   345
      Left            =   1725
      TabIndex        =   1
      Top             =   555
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   609
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   89915395
      CurrentDate     =   40458
   End
   Begin MSForms.Label Label6 
      Height          =   420
      Left            =   10395
      TabIndex        =   18
      Top             =   510
      Width           =   1455
      VariousPropertyBits=   8388627
      Caption         =   "Department"
      Size            =   "2566;741"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox CoDepartment 
      Height          =   405
      Left            =   12060
      TabIndex        =   3
      Top             =   480
      Width           =   3000
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "5292;706"
      MatchEntry      =   1
      ShowDropButtonWhen=   1
      BorderColor     =   8421504
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label5 
      Height          =   420
      Left            =   10410
      TabIndex        =   17
      Top             =   1740
      Width           =   1455
      VariousPropertyBits=   8388627
      Caption         =   "Patient"
      Size            =   "2566;741"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TPatient 
      Height          =   420
      Left            =   12060
      TabIndex        =   6
      Top             =   1650
      Width           =   3000
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "5292;741"
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox CoTest 
      Height          =   405
      Left            =   12060
      TabIndex        =   4
      Top             =   870
      Width           =   3000
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "5292;706"
      MatchEntry      =   1
      ShowDropButtonWhen=   1
      BorderColor     =   8421504
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label2 
      Height          =   420
      Left            =   10395
      TabIndex        =   16
      Top             =   900
      Width           =   1455
      VariousPropertyBits=   8388627
      Caption         =   "Test"
      Size            =   "2566;741"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox CoDoctor 
      Height          =   405
      Left            =   12060
      TabIndex        =   2
      Top             =   90
      Width           =   3000
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "5292;706"
      MatchEntry      =   1
      ShowDropButtonWhen=   1
      BorderColor     =   8421504
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label7 
      Height          =   420
      Left            =   10410
      TabIndex        =   15
      Top             =   150
      Width           =   1455
      VariousPropertyBits=   8388627
      Caption         =   "Doctor"
      Size            =   "2566;741"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox CoSubTest 
      Height          =   405
      Left            =   12060
      TabIndex        =   5
      Top             =   1260
      Width           =   3000
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "5292;706"
      MatchEntry      =   1
      ShowDropButtonWhen=   1
      BorderColor     =   8421504
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label1 
      Height          =   420
      Left            =   10410
      TabIndex        =   14
      Top             =   1335
      Width           =   1455
      VariousPropertyBits=   8388627
      Caption         =   "Sub Test"
      Size            =   "2566;741"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label3 
      Height          =   330
      Left            =   195
      TabIndex        =   13
      Top             =   150
      Width           =   1080
      VariousPropertyBits=   8388627
      Caption         =   "From"
      Size            =   "1905;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label4 
      Height          =   330
      Left            =   165
      TabIndex        =   12
      Top             =   540
      Width           =   1080
      VariousPropertyBits=   8388627
      Caption         =   "To"
      Size            =   "1905;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.OLE OLEExcel 
      Height          =   975
      Left            =   5025
      TabIndex        =   11
      Top             =   -60
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "FMedicalReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sDoctorCode() As String
Dim sDepartmentCode() As String
Dim sTestCode() As String
Dim sSubTestCode() As String

Private Sub CClose_Click()
    Unload Me
End Sub

Private Sub getTest()
Dim rs As Recordset
    
    CoTest.Clear
    
     Set rs = db.OpenRecordset("Select TestRegister.Code,TestRegister.TestName From TestRegister Where (TestRegister.Type='AGroup')  Order By TestRegister.TestName")
    
    Do While rs.EOF = False
        rs.MoveNext
    Loop
    If rs.RecordCount > 0 Then
        rs.MoveFirst
    End If
    ReDim sTestCode(rs.RecordCount + 1) As String
    While rs.EOF = False
        CoTest.AddItem UCase("" & rs!TestName)
        sTestCode(CoTest.ListCount) = "" & rs!Code
        rs.MoveNext
    Wend
    rs.Close
End Sub

Private Sub getSubTest()
Dim rs As Recordset
    
    CoSubTest.Clear
    
    If (CoTest.ListIndex > -1) Then
        Set rs = db.OpenRecordset("Select TestRegister.Code,TestRegister.TestName From TestRegister Where (TestRegister.GroupCode='" & sTestCode(CoTest.ListIndex + 1) & "') And (TestRegister.Type = 'BItem' ) Order By TestRegister.TestName")
    Else
        Set rs = db.OpenRecordset("Select TestRegister.Code,TestRegister.TestName From TestRegister Where (TestRegister.Type = 'BItem' ) Order By TestRegister.TestName")
    End If
    Do While rs.EOF = False
        rs.MoveNext
    Loop
    If rs.RecordCount > 0 Then
        rs.MoveFirst
    End If
    ReDim sSubTestCode(rs.RecordCount + 1) As String
    While rs.EOF = False
        CoSubTest.AddItem UCase("" & rs!TestName)
        sSubTestCode(CoSubTest.ListCount) = "" & rs!Code
        rs.MoveNext
    Wend
    rs.Close
End Sub

Private Sub getDoctor()
Dim rs As Recordset
    
    CoDoctor.Clear
    
    Set rs = db.OpenRecordset("Select DoctorMaster.DoctorCode,DoctorMaster.DoctorName From DoctorMaster Order By DoctorMaster.DoctorName")
    Do While rs.EOF = False
        rs.MoveNext
    Loop
    If rs.RecordCount > 0 Then
        rs.MoveFirst
    End If
    
    ReDim sDoctorCode(rs.RecordCount) As String
    While rs.EOF = False
        CoDoctor.AddItem UCase("" & rs!DoctorName)
        sDoctorCode(CoDoctor.ListCount) = "" & rs!DoctorCode
        rs.MoveNext
    Wend
    rs.Close
End Sub

Private Sub getDepartment()
Dim rs As Recordset
    
    CoDepartment.Clear
    
    Set rs = db.OpenRecordset("Select Department.DepartmentCode,Department.DepartmentName From Department Order By Department.DepartmentName")
    Do While rs.EOF = False
        rs.MoveNext
    Loop
    If rs.RecordCount > 0 Then
        rs.MoveFirst
    End If
    
    ReDim sDepartmentCode(rs.RecordCount) As String
    While rs.EOF = False
        CoDepartment.AddItem UCase("" & rs!DepartmentName)
        sDepartmentCode(CoDepartment.ListCount) = "" & rs!DepartmentCode
        rs.MoveNext
    Wend
    rs.Close
End Sub


Private Sub CoTest_Change()
    getSubTest
End Sub

Private Sub CShowDetailed_Click()
Dim gDate As Single, gBillNo As Single, gDoctor As Single, gPatient As Single, gAge As Single, gSex As Single, gEmail As Single, gMobile As Single, gDepartment As Single, gTest As Single, gSubTest As Single, gNormalValue As Single, gTestValue As Single
Dim rs As Recordset

    'INITIALISING GRID
    gDate = 0
    gBillNo = 1
    gDoctor = 2
    gPatient = 3
    gAge = 4
    gSex = 5
    gEmail = 6
    gMobile = 7
    gDepartment = 8
    gTest = 9
    gSubTest = 10
    gNormalValue = 11
    gTestValue = 12

    MGrid.Clear
    MGrid.Rows = 1 'FOR SKIPING ERROR
    MGrid.Cols = 1 'FOR SKIPING ERROR
    MGrid.FixedCols = 0
    MGrid.FixedRows = 0
    MGrid.Cols = 13
    MGrid.Rows = 1
    MGrid.ColWidth(gDate) = 1400
    MGrid.ColWidth(gBillNo) = 1200
    MGrid.ColWidth(gDoctor) = 2000
    MGrid.ColWidth(gPatient) = 2000
    MGrid.ColWidth(gAge) = 1000
    MGrid.ColWidth(gSex) = 1000
    MGrid.ColWidth(gEmail) = 2000
    MGrid.ColWidth(gMobile) = 2000
    MGrid.ColWidth(gDepartment) = 2000
    MGrid.ColWidth(gTest) = 2000
    MGrid.ColWidth(gSubTest) = 2000
    MGrid.ColWidth(gNormalValue) = 1500
    MGrid.ColWidth(gTestValue) = 1500

    MGrid.col = gDate
    MGrid.CellAlignment = vbAlignRight
    MGrid.CellFontBold = True
    MGrid.CellFontName = "Cosmic San"
    MGrid.CellFontSize = 10
    MGrid.TextMatrix(0, gDate) = "Date"

    MGrid.col = gBillNo
    MGrid.CellAlignment = vbAlignRight
    MGrid.CellFontBold = True
    MGrid.CellFontName = "Cosmic San"
    MGrid.CellFontSize = 10
    MGrid.TextMatrix(0, gBillNo) = "Bill No"

    MGrid.col = gDoctor
    MGrid.CellAlignment = vbAlignRight
    MGrid.CellFontBold = True
    MGrid.CellFontName = "Cosmic San"
    MGrid.CellFontSize = 10
    MGrid.TextMatrix(0, gDoctor) = "Doctor"

    MGrid.col = gPatient
    MGrid.CellAlignment = vbAlignRight
    MGrid.CellFontBold = True
    MGrid.CellFontName = "Cosmic San"
    MGrid.CellFontSize = 10
    MGrid.TextMatrix(0, gPatient) = "Patient"
    
    MGrid.col = gAge
    MGrid.CellAlignment = vbAlignRight
    MGrid.CellFontBold = True
    MGrid.CellFontName = "Cosmic San"
    MGrid.CellFontSize = 10
    MGrid.TextMatrix(0, gAge) = "Age"
    
    MGrid.col = gSex
    MGrid.CellAlignment = vbAlignRight
    MGrid.CellFontBold = True
    MGrid.CellFontName = "Cosmic San"
    MGrid.CellFontSize = 10
    MGrid.TextMatrix(0, gSex) = "Sex"
    
    MGrid.col = gEmail
    MGrid.CellAlignment = vbAlignRight
    MGrid.CellFontBold = True
    MGrid.CellFontName = "Cosmic San"
    MGrid.CellFontSize = 10
    MGrid.TextMatrix(0, gEmail) = "Email"
    
    MGrid.col = gMobile
    MGrid.CellAlignment = vbAlignRight
    MGrid.CellFontBold = True
    MGrid.CellFontName = "Cosmic San"
    MGrid.CellFontSize = 10
    MGrid.TextMatrix(0, gMobile) = "Mobile"

    MGrid.col = gDepartment
    MGrid.CellAlignment = vbAlignRight
    MGrid.CellFontBold = True
    MGrid.CellFontName = "Cosmic San"
    MGrid.CellFontSize = 10
    MGrid.TextMatrix(0, gDepartment) = "Department"

    MGrid.col = gTest
    MGrid.CellAlignment = vbAlignRight
    MGrid.CellFontBold = True
    MGrid.CellFontName = "Cosmic San"
    MGrid.CellFontSize = 10
    MGrid.TextMatrix(0, gTest) = "Test"

    MGrid.col = gSubTest
    MGrid.CellAlignment = vbAlignRight
    MGrid.CellFontBold = True
    MGrid.CellFontName = "Cosmic San"
    MGrid.CellFontSize = 10
    MGrid.TextMatrix(0, gSubTest) = "Sub Test"

    MGrid.col = gNormalValue
    MGrid.CellAlignment = vbAlignRight
    MGrid.CellFontBold = True
    MGrid.CellFontName = "Cosmic San"
    MGrid.CellFontSize = 10
    MGrid.TextMatrix(0, gNormalValue) = "Normal Value"


    MGrid.col = gTestValue
    MGrid.CellAlignment = vbAlignRight
    MGrid.CellFontBold = True
    MGrid.CellFontName = "Cosmic San"
    MGrid.CellFontSize = 10
    MGrid.TextMatrix(0, gTestValue) = "Test Value"

    MGrid.RowHeightMin = 350

    'SHOW DATAS ON GRID

    Me.Caption = "Medical Report"
    
    Set rs = db.OpenRecordset("Select Transaction.BillDate , Transaction.BillNo, DoctorMaster.DoctorName, Transaction.Patient, Transaction.Age, Transaction.Sex, Transaction.email, Transaction.Mobile, Department.DepartmentName,T1.TestName As Test,T2.TestName As SubTest,T2.DefaultValue,Units.UnitName,Transaction.TestValue From Transaction,DoctorMaster,Department,Units,TestRegister As T1,TestRegister As T2 Where (Transaction.BillDate Between cDate('" & DTPFrom.Value & "') And cDate('" & DTPTo.Value & "')) And (DoctorMaster.DoctorCode=Transaction.DoctorCode) And (Department.DepartmentCode=T1.DepartmentCode) And (Units.Code=T2.UnitCode) And (T1.Code=Transaction.TestCode) And (T2.Code=Transaction.SubTestCode) " _
    & IIf(CoDoctor.ListIndex = -1, "", " And (Transaction.DoctorCode='" & sDoctorCode(CoDoctor.ListIndex + 1) & "') ") _
    & IIf(CoTest.ListIndex = -1, "", " And (Transaction.TestCode='" & sTestCode(CoTest.ListIndex + 1) & "') ") _
    & IIf(CoSubTest.ListIndex = -1, "", " And (Transaction.SubTestCode='" & sSubTestCode(CoSubTest.ListIndex + 1) & "') ") _
    & IIf(CoDepartment.ListIndex = -1, "", " And (T1.DepartmentCode='" & sDepartmentCode(CoDepartment.ListIndex + 1) & "') ") _
    & IIf(Len(Trim(TPatient.Text)) = 0, "", " And ((Transaction.Patient Like '" & Trim(TPatient.Text) & "') " _
    & " Or (Transaction.Age Like " & Val(TPatient.Text) & ") " _
    & " Or (Transaction.Sex Like '" & Trim(TPatient.Text) & "') " _
    & " Or (Transaction.email Like '" & Trim(TPatient.Text) & "') " _
    & " Or (Transaction.Mobile Like '" & Trim(TPatient.Text) & "')) ") _
    & " Order By BillDate,Val(BillNo)")
    
    While rs.EOF = False
        MGrid.AddItem Format("" & rs!BillDate, "dd-MM-yyyy") & vbTab & "" & rs!BillNo & vbTab & "" & rs!DoctorName & vbTab & rs!Patient & vbTab & rs!Age & vbTab & rs!Sex & vbTab & rs!email & vbTab & rs!Mobile & vbTab & rs!DepartmentName & vbTab & rs!Test & vbTab & rs!SubTest & vbTab & rs!DefaultValue & " " & rs!UnitName & vbTab & rs!TestValue & " " & rs!UnitName
        rs.MoveNext
    Wend
    rs.Close
    
End Sub


Private Sub CToExcel_Click()
On Error GoTo ErrHandler
Dim oExcel As Object, oExcelSheet As Object
Dim lReturnValue As Long
Dim lRowCount As Long, lColCount As Long

    If MGrid.Rows = 0 Then
        MsgBox "Empty Data!", vbInformation
        Exit Sub
    End If

    OLEExcel.CreateEmbed vbNullString, "Excel.Sheet"

    lRowCount = MGrid.Rows
    lColCount = MGrid.Cols
    ReDim xData(1 To lRowCount + 2, 1 To lColCount) As Variant
    Dim i As Long, j As Long

    Set oExcel = OLEExcel.object
    Set oExcelSheet = oExcel.Sheets(1)

    For i = 1 To lRowCount
       For j = 1 To lColCount
          xData(i, j) = MGrid.TextMatrix(i - 1, j - 1)
       Next j
    Next i

    oExcelSheet.Range("A3:M" & lRowCount + 4).Value = xData

    oExcelSheet.Cells(1, 1).Value = Me.Caption & " From " & Format(DTPFrom.Value, "dd-MM-yyyy") & " To " & Format(DTPTo.Value, "dd-MM-yyyy")

    oExcelSheet.Range("A1:M" & lRowCount + 4).Select
    oExcel.Application.Selection.AutoFormat

On Error Resume Next

    Kill App.Path & "\Reports\" & Me.Caption & " Of " & Format(Date, "dd-MMM-yyyy") & ".xlsx"

    oExcel.SaveAs App.Path & "\Reports\" & Me.Caption & " Of " & Format(Date, "dd-MMM-yyyy") & ".xlsx"
    Set oExcel = Nothing
    Set oExcelSheet = Nothing

    lReturnValue = Shell(App.Path & "\EXCEL.exe - """ & App.Path & "\Reports\" & Me.Caption & " Of " & Format(Date, "dd-MMM-yyyy") & ".xlsx""", vbNormalFocus)

    OLEExcel.Close
    OLEExcel.Delete

    Dim xlTmp As Excel.Application
    Set xlTmp = New Excel.Application
    xlTmp.DisplayFullScreen = True
    xlTmp.Visible = True
    xlTmp.Workbooks.Open App.Path & "\Reports\" & Me.Caption & " Of " & Format(Date, "dd-MMM-yyyy") & ".xlsx"

    MsgBox "Successfully Exported !", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
End Sub

Private Sub DTPFrom_Change()
    MGrid.Rows = 0
End Sub

Private Sub DTPTo_Change()
    MGrid.Rows = 0
End Sub
'
'Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'    If (KeyCode = vbKeyS And ((Shift And 7) = 2)) Then
'        CShowSummary_Click
'    ElseIf (KeyCode = vbKeyD And ((Shift And 7) = 2)) Then
'        CShowDetailed_Click
'    ElseIf (KeyCode = vbKeyX And ((Shift And 7) = 2)) Then
'        CToExcel_Click
'    ElseIf (KeyCode = vbKeyC And ((Shift And 7) = 2)) Then
'        CClose_Click
'    End If
'End Sub

Private Sub Form_Load()
    DTPFrom.Value = Date
    DTPTo.Value = Date
    
    getTest
    getDoctor
    getDepartment
    getSubTest

End Sub

Private Sub MGrid_DblClick()
    If MGrid.Rows > 0 Then
        FMedicalTest.Show
        FMedicalTest.TTransactionNo = MGrid.TextMatrix(MGrid.Row, 1)
        FMedicalTest.getTransactionDetails
                
        If EditTestEntry = False Then
            FMedicalTest.CSave.Enabled = False
        End If
    End If
End Sub
