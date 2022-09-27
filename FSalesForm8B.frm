VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FSalesForm8B 
   BackColor       =   &H00EFEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sales - Form8B"
   ClientHeight    =   9255
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12705
   ControlBox      =   0   'False
   Icon            =   "FSalesForm8B.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FSalesForm8B.frx":628A
   ScaleHeight     =   9255
   ScaleWidth      =   12705
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CAddItem 
      Height          =   500
      Left            =   810
      Picture         =   "FSalesForm8B.frx":204ECC
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6570
      Width           =   1365
   End
   Begin VB.CommandButton CRemoveItem 
      Height          =   500
      Left            =   2250
      Picture         =   "FSalesForm8B.frx":20732E
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6555
      Width           =   1365
   End
   Begin VB.CommandButton CClear 
      Height          =   500
      Left            =   3690
      Picture         =   "FSalesForm8B.frx":209790
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6555
      Width           =   1365
   End
   Begin VB.CommandButton CNew 
      Height          =   500
      Left            =   315
      Picture         =   "FSalesForm8B.frx":20BBF2
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8670
      Width           =   1365
   End
   Begin VB.CommandButton CPrint 
      Height          =   500
      Left            =   1770
      Picture         =   "FSalesForm8B.frx":20E054
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   8670
      Width           =   1365
   End
   Begin VB.CommandButton CSave 
      Height          =   500
      Left            =   9585
      Picture         =   "FSalesForm8B.frx":2104B6
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   8670
      Width           =   1365
   End
   Begin VB.CommandButton CClose 
      Cancel          =   -1  'True
      Height          =   500
      Left            =   11025
      Picture         =   "FSalesForm8B.frx":212918
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   8670
      Width           =   1365
   End
   Begin VB.CommandButton CDelete 
      Height          =   500
      Left            =   4530
      Picture         =   "FSalesForm8B.frx":214D7A
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   135
      Width           =   1365
   End
   Begin MSFlexGridLib.MSFlexGrid MGrid 
      Height          =   3285
      Left            =   600
      TabIndex        =   28
      Top             =   2130
      Width           =   11475
      _ExtentX        =   20241
      _ExtentY        =   5794
      _Version        =   393216
      Rows            =   0
      FixedRows       =   0
      FixedCols       =   0
      BackColorBkg    =   16777215
      GridColorFixed  =   12632256
      FocusRect       =   0
      ScrollBars      =   2
      SelectionMode   =   1
      BorderStyle     =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.DTPicker DTPDate 
      Height          =   420
      Left            =   2850
      TabIndex        =   1
      Top             =   135
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   741
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   93388803
      CurrentDate     =   40544
   End
   Begin MSForms.Label Label1 
      Height          =   330
      Left            =   4380
      TabIndex        =   31
      Top             =   1740
      Width           =   3480
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Sub Test"
      Size            =   "6138;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.ComboBox CoSubTest 
      Height          =   390
      Left            =   4575
      TabIndex        =   30
      Top             =   5505
      Width           =   3090
      VariousPropertyBits=   746604571
      BackColor       =   16777215
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "5450;688"
      MatchEntry      =   1
      ShowDropButtonWhen=   1
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label28 
      Height          =   405
      Left            =   285
      TabIndex        =   29
      Top             =   150
      Width           =   690
      VariousPropertyBits=   8388627
      Caption         =   "Bill No"
      Size            =   "1217;714"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox TValue 
      Height          =   390
      Left            =   7680
      TabIndex        =   5
      Top             =   5505
      Width           =   1530
      VariousPropertyBits=   746604571
      BackColor       =   16777215
      BorderStyle     =   1
      Size            =   "2699;688"
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.OLE OLEExcel 
      Height          =   975
      Left            =   6870
      TabIndex        =   26
      Top             =   615
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSForms.Label Label8 
      Height          =   330
      Left            =   9135
      TabIndex        =   24
      Top             =   1755
      Width           =   1170
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Unit"
      Size            =   "2064;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label LUnit 
      Height          =   390
      Left            =   9210
      TabIndex        =   23
      Top             =   5550
      Width           =   945
      ForeColor       =   4210752
      BackColor       =   8421504
      VariousPropertyBits=   8388627
      Caption         =   "Unit"
      Size            =   "1667;688"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.ComboBox CoTest 
      Height          =   390
      Left            =   1470
      TabIndex        =   4
      Top             =   5505
      Width           =   3090
      VariousPropertyBits=   746604571
      BackColor       =   16777215
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "5450;688"
      MatchEntry      =   1
      ShowDropButtonWhen=   1
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TCost 
      Height          =   390
      Left            =   10170
      TabIndex        =   6
      Top             =   5505
      Width           =   1530
      VariousPropertyBits=   746604571
      BackColor       =   16777215
      BorderStyle     =   1
      Size            =   "2699;688"
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label11 
      Height          =   330
      Left            =   10140
      TabIndex        =   22
      Top             =   1740
      Width           =   1560
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Cost"
      Size            =   "2752;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label LSlNo 
      Height          =   390
      Left            =   630
      TabIndex        =   21
      Top             =   5505
      Width           =   555
      ForeColor       =   4210752
      BackColor       =   8421504
      VariousPropertyBits=   8388627
      Caption         =   "SLNo"
      Size            =   "979;688"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Shape Shape1 
      Height          =   4440
      Left            =   585
      Top             =   1665
      Width           =   11505
   End
   Begin MSForms.TextBox TTransactionNo 
      Height          =   435
      Left            =   1230
      TabIndex        =   0
      Top             =   135
      Width           =   1590
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "2805;767"
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label3 
      Height          =   405
      Left            =   7395
      TabIndex        =   20
      Top             =   360
      Width           =   1335
      VariousPropertyBits=   8388627
      Caption         =   "Doctor"
      Size            =   "2355;714"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox CoDoctor 
      Height          =   420
      Left            =   8655
      TabIndex        =   3
      Top             =   300
      Width           =   3210
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "5662;741"
      MatchEntry      =   1
      ShowDropButtonWhen=   1
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TAddress 
      Height          =   420
      Left            =   8655
      TabIndex        =   27
      Top             =   705
      Width           =   3210
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "5662;741"
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label LTotalAmount 
      Height          =   570
      Left            =   8655
      TabIndex        =   19
      Top             =   6360
      Width           =   3780
      ForeColor       =   64
      VariousPropertyBits=   8388627
      Caption         =   "Grand Amount"
      Size            =   "6667;1005"
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   525
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label Label13 
      Height          =   330
      Left            =   645
      TabIndex        =   18
      Top             =   1740
      Width           =   555
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Sl No"
      Size            =   "979;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label14 
      Height          =   330
      Left            =   1260
      TabIndex        =   17
      Top             =   1740
      Width           =   3480
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Test"
      Size            =   "6138;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label16 
      Height          =   330
      Left            =   7815
      TabIndex        =   16
      Top             =   1740
      Width           =   1170
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Value"
      Size            =   "2064;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label4 
      Height          =   405
      Left            =   345
      TabIndex        =   15
      Top             =   630
      Width           =   1335
      VariousPropertyBits=   8388627
      Caption         =   "Narration"
      Size            =   "2355;714"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TNarration 
      Height          =   420
      Left            =   1230
      TabIndex        =   2
      Top             =   555
      Width           =   3180
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "5609;741"
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label2 
      Height          =   525
      Index           =   0
      Left            =   585
      TabIndex        =   25
      Top             =   1650
      Width           =   11520
      BackColor       =   15724527
      Size            =   "20320;926"
      Picture         =   "FSalesForm8B.frx":2171DC
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "FSalesForm8B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sDoctorCode() As String
Dim sTestCode() As String, sSubTestCode() As String, sBillingName() As String, sGroupCode() As String
Dim gSerialNo As Single, gTest As Single, gValue As Single, gSubTest As Single, gUnit As Single, gCost As Single, gTestCode As Single, gSubTestCode As Single
Dim dCost As Double
    
Private Sub CAddItem_Click()
Dim lYN As Long, r As Long

    If CoTest.ListIndex = -1 Then
        MsgBox "Please Select a Test !", vbInformation
        CoTest.SetFocus
        Exit Sub
    End If
    
     If CoSubTest.ListIndex = -1 Then
        MsgBox "Please Select a Sub Test !", vbInformation
        CoSubTest.SetFocus
        Exit Sub
    End If
    
    
    If Val(LSlNo.Caption) > MGrid.Rows Then 'Add
        MGrid.AddItem ""
        MGrid.TextMatrix(MGrid.Rows - 1, gSerialNo) = Val(LSlNo.Caption)
        MGrid.TextMatrix(MGrid.Rows - 1, gTest) = Trim(CoTest.Text)
        MGrid.TextMatrix(MGrid.Rows - 1, gSubTest) = Trim(CoSubTest.Text)
        MGrid.TextMatrix(MGrid.Rows - 1, gValue) = Val(TValue.Text)
        MGrid.TextMatrix(MGrid.Rows - 1, gUnit) = LUnit.Caption
        MGrid.TextMatrix(MGrid.Rows - 1, gCost) = Format(Val(TCost.Text), "0.00")
        MGrid.TextMatrix(MGrid.Rows - 1, gTestCode) = sTestCode(CoTest.ListIndex + 1)
        
        MGrid.TextMatrix(MGrid.Rows - 1, gSubTestCode) = sSubTestCode(CoSubTest.ListIndex + 1)
        
    Else
        r = Val(LSlNo.Caption)
        MGrid.TextMatrix(r - 1, gSerialNo) = Val(LSlNo.Caption)
        MGrid.TextMatrix(r - 1, gTest) = Trim(CoTest.Text)
        MGrid.TextMatrix(r - 1, gSubTest) = Trim(CoSubTest.Text)
        MGrid.TextMatrix(r - 1, gValue) = Val(TValue.Text)
        MGrid.TextMatrix(r - 1, gUnit) = LUnit.Caption
        MGrid.TextMatrix(r - 1, gCost) = Format(Val(TCost.Text), "0.00")
        MGrid.TextMatrix(r - 1, gTestCode) = sTestCode(CoTest.ListIndex + 1)
        
        MGrid.TextMatrix(r - 1, gSubTestCode) = sSubTestCode(CoSubTest.ListIndex + 1)
       
    End If
    
    clearEditControls
    LTotalAmount.Caption = Format(getGrandTotal, "0.00")
    CoTest.SetFocus
End Sub

Private Sub CClear_Click()
    MGrid.Rows = 0
    LTotalAmount.Caption = Format(getGrandTotal, "0.00")
End Sub

Private Sub CClose_Click()
    Unload Me
End Sub

Private Sub MGridInitialise()
'INITIALISES MGRID
        'SETTING CONSTANTS
    gSerialNo = 0
    gTest = 1
    gSubTest = 2
    gValue = 3
    gUnit = 4
    gCost = 5
    gTestCode = 6
    gSubTestCode = 7
        
    MGrid.Clear
    MGrid.Rows = 1 'FOR SKIPING ERROR
    MGrid.Cols = 1 'FOR SKIPING ERROR
    MGrid.FixedCols = 0
    MGrid.FixedRows = 0
    MGrid.Cols = 8
    MGrid.Rows = 0
    MGrid.ColWidth(gSerialNo) = 900
    MGrid.ColWidth(gTest) = 3100
    MGrid.ColWidth(gSubTest) = 3100
    MGrid.ColWidth(gValue) = 1000
    MGrid.ColWidth(gUnit) = 800
    MGrid.ColWidth(gCost) = 1000
    MGrid.ColWidth(gTestCode) = 0
    MGrid.ColWidth(gSubTestCode) = 0
    
    MGrid.ColAlignment(gTest) = vbLeftJustify
    MGrid.ColAlignment(gSubTest) = vbLeftJustify
    MGrid.ColAlignment(gUnit) = vbLeftJustify
        
    MGrid.RowHeightMin = 350
End Sub

Private Function getNewTransactionNo() As String
Dim rs As Recordset, sTransactionNo As String
    
    Set rs = db.OpenRecordset("Select Max(Val( Transaction.BillNo)) As TNo From Transaction Where ( Transaction.BillType = 'T' ) And (Transaction.FinancialCode='" & getFinancialCode(DTPDate.Value) & "')")
    If rs.RecordCount > 0 Then
        sTransactionNo = Val("" & rs!TNo) + 1
    Else
        sTransactionNo = 1
    End If
    rs.Close
    
    getNewTransactionNo = sTransactionNo
End Function

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
        CoDoctor.AddItem "" & rs!DoctorName
        sDoctorCode(CoDoctor.ListCount) = "" & rs!DoctorCode
        rs.MoveNext
    Wend
    rs.Close
End Sub

Private Sub getTest()
Dim rs As Recordset
    
    CoTest.Clear
    
    Set rs = db.OpenRecordset("Select TestRegister.Code,TestRegister.TestName From TestRegister Where (TestRegister.Type = 'AGroup' ) Order By TestRegister.TestName")
    
    Do While rs.EOF = False
        rs.MoveNext
    Loop
    If rs.RecordCount > 0 Then
        rs.MoveFirst
    End If
    ReDim sTestCode(rs.RecordCount + 1) As String
    While rs.EOF = False
        CoTest.AddItem "" & rs!TestName
        sTestCode(CoTest.ListCount) = "" & rs!Code
        rs.MoveNext
    Wend
    rs.Close
End Sub

Private Sub getSubTest()
Dim rs As Recordset
    
    CoSubTest.Clear
    
    Set rs = db.OpenRecordset("Select TestRegister.Code,TestRegister.TestName From TestRegister Where (TestRegister.Type = 'BItem' ) And (TestRegister.GroupCode = '" & sTestCode(CoTest.ListIndex + 1) & "' ) Order By TestRegister.TestName")
    
    Do While rs.EOF = False
        rs.MoveNext
    Loop
    If rs.RecordCount > 0 Then
        rs.MoveFirst
    End If
    ReDim sSubTestCode(rs.RecordCount + 1) As String
    While rs.EOF = False
        CoSubTest.AddItem "" & rs!TestName
        sSubTestCode(CoSubTest.ListCount) = "" & rs!Code
        rs.MoveNext
    Wend
    rs.Close
End Sub

Private Sub clearControls()
    DTPDate.Value = Date
    TNarration.Text = ""
    CoDoctor.Text = ""
    TAddress.Text = ""
    MGrid.Rows = 0
    LSlNo.Caption = MGrid.Rows + 1
    CoTest.Text = ""
    CoSubTest.Text = ""
    TValue.Text = ""
    TCost.Text = ""
    LTotalAmount.Caption = Format(getGrandTotal, "0.00")
End Sub

Private Sub clearEditControls()
    LSlNo.Caption = MGrid.Rows + 1
    CoTest.Text = ""
    CoSubTest.Text = ""
    TValue.Text = ""
    TCost.Text = ""
End Sub

Private Function getGrandTotal() As Double
Dim dGrandTotal As Double, r As Long
    
    r = 0
    dGrandTotal = 0
    While r < MGrid.Rows
        dGrandTotal = dGrandTotal + Val(MGrid.TextMatrix(r, gCost))
        r = r + 1
    Wend
    getGrandTotal = Round(dGrandTotal, 0)
    LTotalAmount.Caption = Format(dGrandTotal, "0.00")
End Function

Private Sub CDelete_Click()
Dim rs As Recordset, lYN As Long, bFound As Boolean
    bFound = False
    If (MsgBox("Do you want to Delete the Bill ?", vbDefaultButton2 Or vbYesNo) = vbYes) Then
        Set rs = db.OpenRecordset("Select Transaction.* From Transaction Where (Transaction.BillNo = '" & Trim(TTransactionNo.Text) & "' ) And (Transaction.BillType = 'T' ) And (Transaction.FinancialCode='" & getFinancialCode(DTPDate.Value) & "')")
        While rs.EOF = False
            bFound = True
            rs.Delete
            rs.MoveNext
        Wend
        rs.Close
        
        If bFound Then
            MsgBox "Successfully Deleted !", vbInformation
            clearControls
            TTransactionNo.Text = getNewTransactionNo
        Else
            MsgBox "Bill Not Found !", vbInformation
        End If
    End If
End Sub


Private Sub CNew_Click()
    clearControls
    TTransactionNo.Text = getNewTransactionNo
End Sub

Private Sub CoTest_Change()
    getSubTest
End Sub

Private Sub CoTest_GotFocus()
    CoTest.SelStart = 0
    CoTest.SelLength = Len(CoTest.Text)
End Sub

Private Sub CoTest_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
Dim r As Long
    If KeyCode = 113 Then
        FTestRegister.Show vbModal
        getTest
    End If
End Sub

Private Sub CoDoctor_GotFocus()
    CoDoctor.SelStart = 0
    CoDoctor.SelLength = Len(CoDoctor.Text)
End Sub

Private Sub CoDoctor_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 113 Then
        FDoctorRegister.Show vbModal
        getDoctor
    End If
End Sub

Private Sub CRemoveItem_Click()
Dim r As Long
    If MGrid.Rows > 0 Then
        If MGrid.Rows = 1 Then
            MGrid.Rows = 0
            clearEditControls
        Else
            MGrid.RemoveItem (MGrid.Row)
            r = 0
            While r < MGrid.Rows
                MGrid.TextMatrix(r, gSerialNo) = r + 1
                r = r + 1
            Wend
            clearEditControls
        End If
        LTotalAmount.Caption = Format(getGrandTotal, "0.00")
    Else
    
    End If
End Sub

Private Sub CSave_Click()
Dim rs As Recordset
Dim r As Long, lYN As Long, sStatus As String

    If Val(TTransactionNo.Text) = 0 Then
        MsgBox "Please Enter Valid Transaction No !", vbInformation
        TTransactionNo.SetFocus
        Exit Sub
    End If
    
    If CoDoctor.ListIndex = -1 Then
        MsgBox "Please Select a Doctor !", vbInformation
        CoDoctor.SetFocus
        Exit Sub
        
    End If
    
    If MGrid.Rows = 0 Then
        MsgBox "No Items Entered !", vbInformation
        CoTest.SetFocus
        Exit Sub
    End If
        
    'SAVES DATA TO Transaction TABLE
    Set rs = db.OpenRecordset("Select Transaction.* From Transaction Where (Transaction.BillNo = '" & Trim(TTransactionNo.Text) & "' ) And (Transaction.BillType = 'T' ) And (Transaction.FinancialCode='" & getFinancialCode(DTPDate.Value) & "')")
    If rs.RecordCount > 0 Then  'Edit
         
        'SAVES DATA TO TransactionRegister ReadyMade
        While rs.EOF = False
            rs.Delete
            rs.MoveNext
        Wend
    End If
    
    r = 0
    While r < MGrid.Rows
        rs.AddNew
        rs!BillNo = Trim(TTransactionNo.Text)
        rs!BillType = "T"
        rs!BillDate = DTPDate.Value
        rs!BillTime = Format(Time, "HH:MM AMPM")
        rs!Narration = Trim(TNarration.Text)
        rs!DoctorCode = IIf(CoDoctor.ListIndex = -1, "", sDoctorCode(CoDoctor.ListIndex + 1))
        rs!Doctor = Trim(CoDoctor.Text)
        rs!SerialNo = Val(MGrid.TextMatrix(r, gSerialNo))
        rs!TestCode = Trim(MGrid.TextMatrix(r, gTestCode))
        rs!SubTestCode = Trim(MGrid.TextMatrix(r, gSubTestCode))
        rs!Cost = Val(MGrid.TextMatrix(r, gCost))
        rs!TestValue = Val(MGrid.TextMatrix(r, gValue))
        rs!FinancialCode = getFinancialCode(DTPDate.Value)
        rs.Update
        r = r + 1
    Wend
    rs.Close
    
    MsgBox "Successfully Saved !", vbInformation
    
    lYN = MsgBox("Do you want to take Print ?", vbYesNo)
    If lYN = vbYes Then
        printSaleBill Trim(TTransactionNo.Text)
    Else
        
    End If
    
    clearControls
    TTransactionNo.Text = getNewTransactionNo
    TTransactionNo.SetFocus
End Sub

Private Sub DTPDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        TNarration.SetFocus
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyN And ((Shift And 7) = 2)) Then
        CNew_Click
    ElseIf (KeyCode = vbKeyD And ((Shift And 7) = 2)) Then
        CDelete_Click
    ElseIf (KeyCode = vbKeyS And ((Shift And 7) = 2)) Then
        CSave_Click
    ElseIf (KeyCode = vbKeyX And ((Shift And 7) = 2)) Then
        CClose_Click
    ElseIf (KeyCode = vbKeyP And ((Shift And 7) = 2)) Then
        CPrint_Click
    ElseIf (KeyCode = vbKeyA And ((Shift And 7) = 2)) Then
        CAddItem_Click
    ElseIf (KeyCode = vbKeyR And ((Shift And 7) = 2)) Then
        CRemoveItem_Click
    ElseIf (KeyCode = vbKeyL And ((Shift And 7) = 2)) Then
        CClear_Click
    End If
End Sub

Private Sub Form_Load()
    MGridInitialise
    
    getTest
    getDoctor
    clearControls
    TTransactionNo.Text = getNewTransactionNo

End Sub

Private Sub MGrid_Click()
Dim r As Long, i As Long

    If MGrid.Rows > 0 Then
        r = MGrid.Row
        LSlNo.Caption = Val(MGrid.TextMatrix(r, gSerialNo))
        CoTest.Text = Trim(MGrid.TextMatrix(r, gTest))
        CoSubTest.Text = Trim(MGrid.TextMatrix(r, gSubTest))
        TValue.Text = Val(MGrid.TextMatrix(r, gValue))
        TCost.Text = Val(MGrid.TextMatrix(r, gCost))
    Else
    End If
End Sub

Private Sub MGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        CoTest.SetFocus
    End If
End Sub

Private Sub TAddress_GotFocus()
    TAddress.SelStart = 0
    TAddress.SelLength = Len(TAddress.Text)
End Sub

Private Sub TValue_GotFocus()
    TValue.SelStart = 0
    TValue.SelLength = Len(TValue.Text)
End Sub

Private Sub TCost_GotFocus()
    TCost.SelStart = 0
    TCost.SelLength = Len(TCost.Text)
End Sub

Private Sub CoSubTest_GotFocus()
    CoSubTest.SelStart = 0
    CoSubTest.SelLength = Len(CoSubTest.Text)
End Sub

Private Sub TTransactionNo_GotFocus()
    TTransactionNo.SelStart = 0
    TTransactionNo.SelLength = Len(TTransactionNo.Text)
End Sub

Private Sub TTransactionNo_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then
        clearControls
        getTransactionDetails
    End If
End Sub

Public Sub getTransactionDetails()
Dim rs As Recordset, r As Long
        
    Set rs = db.OpenRecordset("Select Units.UnitName,Transaction.*,TR1.TestName As SubTest,TR2.TestName As Test From TestRegister As TR1,TestRegister As TR2,Transaction,Units Where (Transaction.BillNo = '" & Trim(TTransactionNo.Text) & "' ) And (Transaction.BillType = 'T' ) And (TR1.Code = Transaction.SubTestCode ) And (TR2.Code = Transaction.TestCode ) And (Units.Code = TR1.UnitCode ) And (Transaction.FinancialCode='" & getFinancialCode(DTPDate.Value) & "') Order By Transaction.SerialNo")
    MGrid.Rows = 0
    If rs.RecordCount > 0 Then
        DTPDate.Value = rs!BillDate
        CoDoctor.Text = "" & rs!Doctor
        TNarration.Text = "" & rs!Narration
        r = 0
        rs.MoveFirst
        While rs.EOF = False
            MGrid.AddItem ""
            
            MGrid.TextMatrix(r, gSerialNo) = "" & rs!SerialNo
            MGrid.TextMatrix(r, gTest) = "" & rs!Test
            MGrid.TextMatrix(r, gSubTest) = "" & rs!SubTest
            MGrid.TextMatrix(r, gValue) = Val("" & rs!TestValue)
            MGrid.TextMatrix(r, gUnit) = "" & rs!UnitName
            MGrid.TextMatrix(r, gCost) = Format(Val("" & rs!Cost), "0.00")
            MGrid.TextMatrix(r, gTestCode) = "" & rs!TestCode
            MGrid.TextMatrix(r, gSubTestCode) = "" & rs!SubTestCode
        
            r = r + 1
            rs.MoveNext
        Wend
        rs.Close
    Else
        rs.Close
    End If
    
    LSlNo.Caption = MGrid.Rows + 1
    LTotalAmount.Caption = Format(getGrandTotal, "0.00")
    
End Sub

Private Sub CPrint_Click()
    printSaleBill Trim(TTransactionNo.Text)
End Sub

Public Sub printSaleBill(sBillNo As String)
Dim rs As Recordset
Dim SerialNo As Double, dGrossValue As Double, dDiscount As Double, dNetValue As Double, dTaxValue As Double, dTotalValue As Double
Dim dTotalGrossVaue As Double, dTotalDiscount As Double, dTotalNetValue As Double, dTotalTaxValue As Double, dGrandValue As Double, dTotalQuantity As Double
Dim d1PercenCoSubTest As Double, d5PercenCoSubTest As Double, d14_5PercenCoSubTest As Double

On Error GoTo GoOut
    
    Open "LPT1:" For Output As #1
    Set rs = db.OpenRecordset("Select TestRegister.ItemName,Units.UnitName,Transaction.BillNo,Transaction.BillDate,Transaction.ItemDiscount,Transaction.SerialNo,Transaction.Quantity,Transaction.Tax,Transaction.MRP,Transaction.Retail,Transaction.WholeSale,Transaction.Other,Transaction.SaleRate,DoctorMaster.* From TestRegister,Transaction,DoctorMaster,Units Where (Units.Code=TestRegister.SaleUnitCode) And (Transaction.BillNo = '" & sBillNo & "' ) And (TestRegister.Code =Transaction.ItemCode ) And (Transaction.BillType = 'SB' ) And (DoctorMaster.DoctorCode = Transaction.DoctorCode) Order By Val(Transaction.SerialNo)")
    
    Print #1, Chr(27) & "j" & Chr(216)
    Print #1, Chr(27) & "j" & Chr(216)
    Print #1, Chr(27) & "!" & Chr(4) & ""
    Print #1, Chr(27) & "!" & Chr(45) & "                                 THOUFEEQ AGENCIES " & Chr(27) & "!" & Chr(0)
    Print #1, Chr(27) & "!" & Chr(4) & "Tin:32100561606                                           KOTT, ALINCHUVADU, TIRUR-1 " & Chr(27) & "!" & Chr(0)
    Print #1, Chr(27) & "!" & Chr(4) & "Dated: 1-9-07                                       Ph:0494 2421219,9847145323,9605445062" & Chr(27) & "!" & Chr(0)
    Print #1, Chr(27) & "!" & Chr(4) & "                                                  THE KERALA VALUE ADDED TAX RULES 2005" & Chr(27) & "!" & Chr(0)
    Print #1, Chr(27) & "!" & Chr(4) & "                                                               FORM NO.8B " & Chr(27) & "!" & Chr(0)
    Print #1, Chr(27) & "!" & Chr(4) & "                                                       RETAIL INVOICE(CASH/CREDIT) " & Chr(27) & "!" & Chr(0)
    Print #1,
    If rs.RecordCount > 0 Then
    
        'Header
        Print #1, Chr(27) & "!" & Chr(4) & Left("Invoice No: " & rs!BillNo & Space(68), 68) & "                              Date : " & Right(Space(11) & Format(rs!BillDate, "dd-MM-yy"), 11) & Chr(27) & "!" & Chr(0)
        Print #1, Chr(27) & "!" & Chr(4) & "Name & Address of Purchasing Dealer:" & Left(UCase("" & rs!DoctorName & ", " & rs!Address1) & ", " & rs!Address2 & Space(80), 80) & Chr(27) & "!" & Chr(0)
        Print #1, Chr(27) & "!" & Chr(4) & Left("Tel.No : " & rs!Phone & Space(28), 28) & Left("TIN No : " & rs!TinNo & Space(28), 28) & Left("CST Reg.No : " & rs!Address3 & Space(28), 28) & Chr(27) & "!" & Chr(0)
        Print #1, Chr(27) & "!" & Chr(4) & "----------------------------------------------------------------------------------------------------------------------------" & Chr(27) & "!" & Chr(0)
        Print #1, Chr(27) & "!" & Chr(4) & "Sl.No       Commodity Item             Tax      MRP       Rate     Qty     G.Value  Discount  Net Value   TaxAmt    Total   " & Chr(27) & "!" & Chr(0)
        Print #1, Chr(27) & "!" & Chr(4) & "----------------------------------------------------------------------------------------------------------------------------" & Chr(27) & "!" & Chr(0)
        
        'Details
        
        
        
        While rs.EOF = False
                        
            dGrossValue = rs!SaleRate * Abs(Val(rs!Quantity))
            dDiscount = Val("" & rs!ItemDiscount)
            dNetValue = rs!SaleRate * Abs(Val(rs!Quantity)) - dDiscount
            dTaxValue = dNetValue * Val("" & rs!Tax) / 100
            dTotalValue = dNetValue + dTaxValue
            
            dTotalGrossVaue = dTotalGrossVaue + dGrossValue
            dTotalDiscount = dTotalDiscount + dDiscount
            dTotalNetValue = dTotalNetValue + dNetValue
            dTotalTaxValue = dTotalTaxValue + dTaxValue
            dGrandValue = dGrandValue + dTotalValue
            dTotalQuantity = dTotalQuantity + Abs(Val("" & rs!Quantity))
            
            If (Val("" & rs!Tax) = 1) Then
                d1PercenCoSubTest = d1PercenCoSubTest + dTaxValue
            ElseIf (Val("" & rs!Tax) = 5) Then
                d5PercenCoSubTest = d5PercenCoSubTest + dTaxValue
            ElseIf (Val("" & rs!Tax) = 14.5) Then
                d14_5PercenCoSubTest = d14_5PercenCoSubTest + dTaxValue
            End If
            
            SerialNo = SerialNo + 1
            Print #1, Chr(27) & "!" & Chr(4) & " " & Left(SerialNo & Space(5), 5) & " " & Left(rs!ItemName & Space(30), 30) & " " & Right(Space(6) & Format(Val("" & rs!Tax), "0.00") & "%", 6) & " " & Right(Space(9) & Format(rs!MRP, "0.00"), 9) & " " & Right(Space(9) & Format(Val(rs!SaleRate), "0.00"), 9) & " " & Right(Space(8) & Abs(Val(rs!Quantity)) & " " & rs!UnitName, 8) & " " & Right(Space(9) & Format(dGrossValue, "0.00"), 9) & " " & Right(Space(9) & Format(dDiscount), 9) & " " & Right(Space(9) & Format(dNetValue, "0.00"), 9) & " " & Right(Space(9) & Format(dTaxValue, "0.00"), 9) & " " & Right(Space(9) & Format(dTotalValue, "0.00"), 9) & " " & Chr(27) & "!" & Chr(0)
            
            rs.MoveNext
        Wend
        
        Print #1,
        Print #1,
        Print #1, Chr(27) & "!" & Chr(4) & "----------------------------------------------------------------------------------------------------------------------------" & Chr(27) & "!" & Chr(0)
        Print #1, Chr(27) & "!" & Chr(4) & "Tax Details" & Chr(27) & "!" & Chr(0)
        Print #1, Chr(27) & "!" & Chr(4) & "1%      =  " & Right(Space(10) & Format(d1PercenCoSubTest, "0.00"), 10) & "                                                                       Total         :" & Format(dTotalGrossVaue, "0.00") & Chr(27) & "!" & Chr(0)
        Print #1, Chr(27) & "!" & Chr(4) & "5%      =  " & Right(Space(10) & Format(d5PercenCoSubTest, "0.00"), 10) & "                                                                        VAT           :" & Format(dTotalTaxValue, "0.00") & Chr(27) & "!" & Chr(0)
        Print #1, Chr(27) & "!" & Chr(4) & "14.5%   =  " & Right(Space(10) & Format(d14_5PercenCoSubTest, "0.00"), 10) & "                                                                        Round Off     :" & Format(Round((dGrandValue), 0) - (dGrandValue), "0.00") & Chr(27) & "!" & Chr(0)
        Print #1, Chr(27) & "!" & Chr(4) & "                                                                                             Grand Total   :" & Chr(27) & "!" & Chr(44) & Format(Round(dGrandValue, 0), "0.00") & Chr(27) & "!" & Chr(0) & Chr(27) & "!" & Chr(0)
                  
    End If
    
    
    'Print #1, Chr(27) & "!" & Chr(4) & "(" & NumberToWords(Round(dGrandValue, 0)) & ")" & Chr(27) & "!" & Chr(0)
    Print #1, Chr(27) & "!" & Chr(4) & "----------------------------------------------------------------------------------------------------------------------------" & Chr(27) & "!" & Chr(0)
    Print #1, Chr(27) & "!" & Chr(4) & "                                                                                                                 Authorised Signatory" & Chr(27) & "!" & Chr(0)
    Print #1,
    Print #1,
    Print #1,
    Print #1,
    Print #1,
    Print #1,
    Print #1,
    Print #1,
    Print #1,
    Print #1,
    Print #1,
    Print #1,
    Print #1,
    Print #1,
    
    Close #1

    Exit Sub
GoOut:
    MsgBox Err.Description & " : Please Check Printer !", vbInformation
End Sub

