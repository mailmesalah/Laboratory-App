VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FMedicalTest 
   BackColor       =   &H00EFEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test Result Entry"
   ClientHeight    =   9255
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12705
   ControlBox      =   0   'False
   Icon            =   "FMedicalTest.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FMedicalTest.frx":628A
   ScaleHeight     =   9255
   ScaleWidth      =   12705
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CSendeMail 
      BackColor       =   &H8000000D&
      Caption         =   "Send eMail"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   6285
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   8670
      Width           =   1365
   End
   Begin VB.CommandButton CSendSMS 
      BackColor       =   &H8000000D&
      Caption         =   "Send SMS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   4860
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   8670
      Width           =   1365
   End
   Begin VB.CommandButton CAddItem 
      Height          =   500
      Left            =   630
      Picture         =   "FMedicalTest.frx":204ECC
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7680
      Width           =   1365
   End
   Begin VB.CommandButton CRemoveItem 
      Height          =   500
      Left            =   2070
      Picture         =   "FMedicalTest.frx":20732E
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7665
      Width           =   1365
   End
   Begin VB.CommandButton CClear 
      Height          =   500
      Left            =   3510
      Picture         =   "FMedicalTest.frx":209790
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   7665
      Width           =   1365
   End
   Begin VB.CommandButton CNew 
      Height          =   500
      Left            =   315
      Picture         =   "FMedicalTest.frx":20BBF2
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   8670
      Width           =   1365
   End
   Begin VB.CommandButton CPrint 
      Height          =   500
      Left            =   1770
      Picture         =   "FMedicalTest.frx":20E054
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   8670
      Width           =   1365
   End
   Begin VB.CommandButton CSave 
      Height          =   500
      Left            =   9585
      Picture         =   "FMedicalTest.frx":2104B6
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   8670
      Width           =   1365
   End
   Begin VB.CommandButton CClose 
      Cancel          =   -1  'True
      Height          =   500
      Left            =   11025
      Picture         =   "FMedicalTest.frx":212918
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   8670
      Width           =   1365
   End
   Begin VB.CommandButton CDelete 
      Height          =   500
      Left            =   4530
      Picture         =   "FMedicalTest.frx":214D7A
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   135
      Width           =   1365
   End
   Begin MSFlexGridLib.MSFlexGrid MGrid 
      Height          =   3285
      Left            =   420
      TabIndex        =   33
      Top             =   2775
      Width           =   11475
      _ExtentX        =   20241
      _ExtentY        =   5794
      _Version        =   393216
      Rows            =   0
      FixedRows       =   0
      FixedCols       =   0
      BackColorBkg    =   16777215
      GridColorFixed  =   12632256
      ScrollTrack     =   -1  'True
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
      Format          =   95551491
      CurrentDate     =   40544
   End
   Begin MSForms.ComboBox CoSex 
      Height          =   420
      Left            =   10410
      TabIndex        =   6
      Top             =   555
      Width           =   1455
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "2566;741"
      MatchEntry      =   1
      ShowDropButtonWhen=   1
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label10 
      Height          =   285
      Left            =   9990
      TabIndex        =   41
      Top             =   600
      Width           =   375
      VariousPropertyBits=   8388627
      Caption         =   "Sex"
      Size            =   "661;503"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label9 
      Height          =   405
      Left            =   7530
      TabIndex        =   40
      Top             =   1440
      Width           =   750
      VariousPropertyBits=   8388627
      Caption         =   "Mobile"
      Size            =   "1323;714"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TMobile 
      Height          =   420
      Left            =   8655
      TabIndex        =   8
      Top             =   1365
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
   Begin MSForms.Label Label7 
      Height          =   405
      Left            =   7530
      TabIndex        =   39
      Top             =   1020
      Width           =   885
      VariousPropertyBits=   8388627
      Caption         =   "email"
      Size            =   "1561;714"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox Temail 
      Height          =   420
      Left            =   8655
      TabIndex        =   7
      Top             =   960
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
   Begin MSForms.Label Label6 
      Height          =   405
      Left            =   7530
      TabIndex        =   38
      Top             =   600
      Width           =   990
      VariousPropertyBits=   8388627
      Caption         =   "Age"
      Size            =   "1746;714"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TAge 
      Height          =   420
      Left            =   8655
      TabIndex        =   5
      Top             =   555
      Width           =   1185
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "2090;741"
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label5 
      Height          =   405
      Left            =   7530
      TabIndex        =   37
      Top             =   195
      Width           =   1020
      VariousPropertyBits=   8388627
      Caption         =   "Patient"
      Size            =   "1799;714"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label LDepartment 
      Height          =   330
      Left            =   4200
      TabIndex        =   36
      Top             =   6900
      Width           =   3480
      ForeColor       =   12582912
      VariousPropertyBits=   8388627
      Caption         =   "Department"
      Size            =   "6138;582"
      FontName        =   "Arial Rounded MT Bold"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label1 
      Height          =   330
      Left            =   4200
      TabIndex        =   35
      Top             =   2385
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
      Left            =   4410
      TabIndex        =   10
      Top             =   6150
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
      TabIndex        =   34
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
      Left            =   7530
      TabIndex        =   11
      Top             =   6150
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
      Height          =   480
      Left            =   -150
      TabIndex        =   32
      Top             =   60
      Visible         =   0   'False
      Width           =   420
   End
   Begin MSForms.Label Label8 
      Height          =   330
      Left            =   8955
      TabIndex        =   30
      Top             =   2400
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
      Left            =   9030
      TabIndex        =   29
      Top             =   6195
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
      Left            =   1290
      TabIndex        =   9
      Top             =   6150
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
      Left            =   9990
      TabIndex        =   12
      Top             =   6150
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
      Left            =   9960
      TabIndex        =   28
      Top             =   2385
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
      Left            =   450
      TabIndex        =   27
      Top             =   6150
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
      Left            =   405
      Top             =   2310
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
      Left            =   330
      TabIndex        =   26
      Top             =   1005
      Width           =   735
      VariousPropertyBits=   8388627
      Caption         =   "Doctor"
      Size            =   "1296;714"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox CoDoctor 
      Height          =   420
      Left            =   1230
      TabIndex        =   3
      Top             =   960
      Width           =   3180
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "5609;741"
      MatchEntry      =   1
      ShowDropButtonWhen=   1
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TPatient 
      Height          =   420
      Left            =   8655
      TabIndex        =   4
      Top             =   150
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
      Left            =   8475
      TabIndex        =   25
      Top             =   7005
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
      Left            =   465
      TabIndex        =   24
      Top             =   2385
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
      Left            =   1080
      TabIndex        =   23
      Top             =   2385
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
      Left            =   7635
      TabIndex        =   22
      Top             =   2385
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
      Left            =   330
      TabIndex        =   21
      Top             =   600
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
      Left            =   405
      TabIndex        =   31
      Top             =   2295
      Width           =   11505
      BackColor       =   4194304
      Size            =   "20294;926"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "FMedicalTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sDoctorCode() As String
Dim sTestCode() As String, sGroupCode() As String
Dim sSubTestCode() As String, sDefaultValue() As String, sDefaultCost() As Double, sUnitName() As String, sDepartment() As String
Dim gSerialNo As Single, gTest As Single, gValue As Single, gSubTest As Single, gUnit As Single, gCost As Single, gTestCode As Single, gSubTestCode As Single
Dim dCost As Double
    
Public Sub setViewable()
    CSave.Enabled = False
End Sub

    
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
        MGrid.TextMatrix(MGrid.Rows - 1, gValue) = TValue.Text
        MGrid.TextMatrix(MGrid.Rows - 1, gUnit) = LUnit.Caption
        MGrid.TextMatrix(MGrid.Rows - 1, gCost) = Format(Val(TCost.Text), "0.00")
        MGrid.TextMatrix(MGrid.Rows - 1, gTestCode) = sTestCode(CoTest.ListIndex + 1)
        MGrid.TextMatrix(MGrid.Rows - 1, gSubTestCode) = sSubTestCode(CoSubTest.ListIndex + 1)
        
    Else
        r = Val(LSlNo.Caption)
        MGrid.TextMatrix(r - 1, gSerialNo) = Val(LSlNo.Caption)
        MGrid.TextMatrix(r - 1, gTest) = Trim(CoTest.Text)
        MGrid.TextMatrix(r - 1, gSubTest) = Trim(CoSubTest.Text)
        MGrid.TextMatrix(r - 1, gValue) = TValue.Text
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
    MGrid.ColWidth(gValue) = 1550
    MGrid.ColWidth(gUnit) = 930
    MGrid.ColWidth(gCost) = 1550
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
    
    Set rs = db.OpenRecordset("Select TestRegister.Code,TestRegister.TestName,Department.DepartmentName From TestRegister,Department Where (Department.DepartmentCode=TestRegister.DepartmentCode) And (TestRegister.Type = 'AGroup' ) Order By TestRegister.TestName")
    
    Do While rs.EOF = False
        rs.MoveNext
    Loop
    If rs.RecordCount > 0 Then
        rs.MoveFirst
    End If
    ReDim sTestCode(rs.RecordCount + 1) As String
    ReDim sDepartment(rs.RecordCount + 1) As String
    While rs.EOF = False
        CoTest.AddItem "" & rs!TestName
        sDepartment(CoTest.ListCount) = "" & rs!DepartmentName
        sTestCode(CoTest.ListCount) = "" & rs!Code
        rs.MoveNext
    Wend
    rs.Close
End Sub

Private Sub getSubTest()
Dim rs As Recordset
    
    CoSubTest.Clear
    
    Set rs = db.OpenRecordset("Select TestRegister.*,(Select Units.UnitName From Units Where (Units.Code=TestRegister.UnitCode)) As UnitName From TestRegister Where (TestRegister.Type = 'BItem' ) And (TestRegister.GroupCode = '" & sTestCode(CoTest.ListIndex + 1) & "' ) Order By TestRegister.TestName")
    
    Do While rs.EOF = False
        rs.MoveNext
    Loop
    If rs.RecordCount > 0 Then
        rs.MoveFirst
    End If
    ReDim sSubTestCode(rs.RecordCount + 1) As String
    ReDim sDefaultValue(rs.RecordCount + 1) As String
    ReDim sDefaultCost(rs.RecordCount + 1) As Double
    ReDim sUnitName(rs.RecordCount + 1) As String
    While rs.EOF = False
        CoSubTest.AddItem "" & rs!TestName
        sSubTestCode(CoSubTest.ListCount) = "" & rs!Code
        sDefaultCost(CoSubTest.ListCount) = "" & rs!Cost
        sDefaultValue(CoSubTest.ListCount) = "" & rs!DefaultValue
        sUnitName(CoSubTest.ListCount) = "" & rs!UnitName
        rs.MoveNext
    Wend
    rs.Close
End Sub

Private Sub clearControls()
    DTPDate.Value = Date
    LDepartment.Caption = ""
    TNarration.Text = ""
    CoDoctor.Text = ""
    TPatient.Text = ""
    TAge.Text = ""
    CoSex.Text = ""
    TMobile.Text = ""
    Temail.Text = ""
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
    'TValue.Text = ""
    'TCost.Text = ""
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

Private Sub CoSubTest_Change()
    TValue.Text = sDefaultValue(CoSubTest.ListIndex + 1)
    LUnit.Caption = sUnitName(CoSubTest.ListIndex + 1)
    TCost.Text = Format(Val(sDefaultCost(CoSubTest.ListIndex + 1)), "0.00")
End Sub

Private Sub CoSubTest_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = 113 And EditMasterRegisters Then
        FTestRegister.Show vbModal
        getSubTest
    End If
End Sub

Private Sub CoTest_Change()
    getSubTest
    LDepartment.Caption = sDepartment(CoTest.ListIndex + 1)
End Sub

Private Sub CoTest_GotFocus()
    CoTest.SelStart = 0
    CoTest.SelLength = Len(CoTest.Text)
End Sub

Private Sub CoTest_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 113 And EditMasterRegisters Then
        FTestRegister.Show vbModal
        getTest
    End If
End Sub

Private Sub CoDoctor_GotFocus()
    CoDoctor.SelStart = 0
    CoDoctor.SelLength = Len(CoDoctor.Text)
End Sub

Private Sub CoDoctor_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 113 And EditMasterRegisters Then
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
    
    If Len(Trim(TPatient.Text)) = 0 Then
        MsgBox "Please Enter Patient !", vbInformation
        TPatient.SetFocus
        Exit Sub
    End If
    
    If Val(TAge.Text) = 0 Then
        MsgBox "Please Enter Age !", vbInformation
        TAge.SetFocus
        Exit Sub
    End If
    
    If CoSex.ListIndex = -1 Then
        MsgBox "Please Select Sex !", vbInformation
        CoSex.SetFocus
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
        rs!Patient = Trim(TPatient.Text)
        rs!Age = Val(TAge.Text)
        rs!Sex = CoSex.Text
        rs!email = Trim(Temail.Text)
        rs!Mobile = Trim(TMobile.Text)
        rs!SerialNo = Val(MGrid.TextMatrix(r, gSerialNo))
        rs!TestCode = Trim(MGrid.TextMatrix(r, gTestCode))
        rs!SubTestCode = Trim(MGrid.TextMatrix(r, gSubTestCode))
        rs!Cost = Val(MGrid.TextMatrix(r, gCost))
        rs!TestValue = MGrid.TextMatrix(r, gValue)
        rs!FinancialCode = getFinancialCode(DTPDate.Value)
        rs.Update
        r = r + 1
    Wend
    rs.Close
    
    MsgBox "Successfully Saved !", vbInformation
    
    lYN = MsgBox("Do you want to Print Bill Report ?", vbYesNo)
    If lYN = vbYes Then
        printBillDetails
    End If
    
    lYN = MsgBox("Do you want to Print Medical Report ?", vbYesNo)
    If lYN = vbYes Then
        printMedicalDetails
    End If
    
    If Trim(Temail.Text) <> "" Then
        lYN = MsgBox("Do you want to Send email to Patient ?", vbYesNo)
        If lYN = vbYes Then
            createPDFReport
            SendeMail Trim(Temail.Text), Trim(TPatient.Text), App.Path & "\Pdf\Medical Report.pdf"
            Kill App.Path & "\Pdf\Medical Report.pdf"
        End If
        
    End If
    
    If Trim(TMobile.Text) <> "" Then
        lYN = MsgBox("Do you want to Send SMS to Patient ?", vbYesNo)
        If lYN = vbYes Then
            SendSMS Trim(TMobile.Text), Trim(TPatient.Text)
        End If
    End If
    
    clearControls
    TTransactionNo.Text = getNewTransactionNo
    If TTransactionNo.Enabled = True Then
        TTransactionNo.SetFocus
    Else
        CoDoctor.SetFocus
    End If
End Sub

Private Sub CSendeMail_Click()
Dim lYN As Long
Dim rs As Recordset
        
    If Trim(Temail.Text) <> "" Then
        lYN = MsgBox("Do you want to Send email to Patient ?", vbYesNo)
        If lYN = vbYes Then
            createPDFReport
            Set rs = db.OpenRecordset("Select Transaction.Patient,Transaction.email From Transaction Where (Transaction.BillNo = '" & Trim(TTransactionNo.Text) & "' ) And (Transaction.BillType = 'T' ) And (Transaction.FinancialCode='" & getFinancialCode(DTPDate.Value) & "')")
            If rs.RecordCount > 0 Then
                SendeMail Trim("" & rs!email), Trim("" & rs!Patient), App.Path & "\Pdf\Medical Report.pdf"
                rs.Close
            End If
            
            Kill App.Path & "\Pdf\Medical Report.pdf"
        End If
    End If
    
    
End Sub

Private Sub CSendSMS_Click()
Dim lYN As Long
Dim rs As Recordset
    
    If Trim(TMobile.Text) <> "" Then
        lYN = MsgBox("Do you want to Send SMS to Patient ?", vbYesNo)
        If lYN = vbYes Then
            Set rs = db.OpenRecordset("Select Transaction.Patient,Transaction.Mobile From Transaction Where (Transaction.BillNo = '" & Trim(TTransactionNo.Text) & "' ) And (Transaction.BillType = 'T' ) And (Transaction.FinancialCode='" & getFinancialCode(DTPDate.Value) & "')")
            If rs.RecordCount > 0 Then
                SendSMS Trim("" & rs!Mobile), Trim("" & rs!Patient)
                rs.Close
            End If
        End If
    End If
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
    CoSex.AddItem "Male"
    CoSex.AddItem "Female"
    CoSex.AddItem "Other"
    
    getTest
    getDoctor
    clearControls
    TTransactionNo.Text = getNewTransactionNo
    
    If EditTestEntry = False Then
        TTransactionNo.Enabled = False
    End If

End Sub

Private Sub MGrid_Click()
Dim r As Long, i As Long

    If MGrid.Rows > 0 Then
        r = MGrid.Row
        LSlNo.Caption = Val(MGrid.TextMatrix(r, gSerialNo))
        CoTest.Text = Trim(MGrid.TextMatrix(r, gTest))
        CoSubTest.Text = Trim(MGrid.TextMatrix(r, gSubTest))
        TValue.Text = MGrid.TextMatrix(r, gValue)
        TCost.Text = Val(MGrid.TextMatrix(r, gCost))
    Else
    End If
End Sub

Private Sub MGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        CoTest.SetFocus
    End If
End Sub

Private Sub Temail_LostFocus()
    Temail.Text = Left(Trim(Temail.Text), 255)
End Sub

Private Sub TPatient_GotFocus()
    TPatient.SelStart = 0
    TPatient.SelLength = Len(TPatient.Text)
End Sub

Private Sub TAge_GotFocus()
    TAge.SelStart = 0
    TAge.SelLength = Len(TAge.Text)
End Sub

Private Sub CoSex_GotFocus()
    CoSex.SelStart = 0
    CoSex.SelLength = Len(CoSex.Text)
End Sub

Private Sub Temail_GotFocus()
    Temail.SelStart = 0
    Temail.SelLength = Len(Temail.Text)
End Sub

Private Sub TMobile_GotFocus()
    TMobile.SelStart = 0
    TMobile.SelLength = Len(TMobile.Text)
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
        
    Set rs = db.OpenRecordset("Select (Select Units.UnitName From Units Where (Units.Code = TR1.UnitCode )) As UnitName,Transaction.*,TR1.TestName As SubTest,TR2.TestName As Test From TestRegister As TR1,TestRegister As TR2,Transaction Where (Transaction.BillNo = '" & Trim(TTransactionNo.Text) & "' ) And (Transaction.BillType = 'T' ) And (TR1.Code = Transaction.SubTestCode ) And (TR2.Code = Transaction.TestCode ) And (Transaction.FinancialCode='" & getFinancialCode(DTPDate.Value) & "') Order By Transaction.SerialNo")
    MGrid.Rows = 0
    If rs.RecordCount > 0 Then
        DTPDate.Value = rs!BillDate
        CoDoctor.Text = "" & rs!Doctor
        TNarration.Text = "" & rs!Narration
        TPatient.Text = "" & rs!Patient
        TAge.Text = Val("" & rs!Age)
        CoSex.Text = "" & rs!Sex
        Temail.Text = "" & rs!email
        TMobile.Text = "" & rs!Mobile
        
        r = 0
        rs.MoveFirst
        While rs.EOF = False
            MGrid.AddItem ""
            
            MGrid.TextMatrix(r, gSerialNo) = "" & rs!SerialNo
            MGrid.TextMatrix(r, gTest) = "" & rs!Test
            MGrid.TextMatrix(r, gSubTest) = "" & rs!SubTest
            MGrid.TextMatrix(r, gValue) = "" & rs!TestValue
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
Dim lYN As Long

    lYN = MsgBox("Do you want to Print Bill Report ?", vbYesNo)
    If lYN = vbYes Then
        printBillDetails
    End If
    
    lYN = MsgBox("Do you want to Print Medical Report ?", vbYesNo)
    If lYN = vbYes Then
        printMedicalDetails
    End If
End Sub


Private Sub createPDFReport()
' Create a simple PDF file using the mjwPDF class
Dim objPDF As New mjwPDF
Dim rs As Recordset, r As Long
Dim x As Double, y As Double, x1 As Double, y1 As Double
Dim myPicture As IPictureDisp
Dim count As Long

On Error GoTo TryAgain

Start:
count = count + 1

    Set rs = db.OpenRecordset("Select (Select Units.UnitName From Units Where (Units.Code = TR1.UnitCode )) As UnitName,Transaction.*,TR1.TestName As SubTest,TR1.DefaultValue,TR2.TestName As Test,Department.DepartmentName From TestRegister As TR1,TestRegister As TR2,Transaction,Department Where (Transaction.BillNo = '" & Trim(TTransactionNo.Text) & "' ) And (Transaction.BillType = 'T' ) And (Department.DepartmentCode=TR1.DepartmentCode) And (TR1.Code = Transaction.SubTestCode ) And (TR2.Code = Transaction.TestCode ) And (Transaction.FinancialCode='" & getFinancialCode(DTPDate.Value) & "') Order By Transaction.SerialNo")
    
    If rs.RecordCount > 0 Then
    
        objPDF.PDFTitle = "Medical Report"
        objPDF.PDFFileName = App.Path & "\Pdf\Medical Report.pdf"
        objPDF.PDFLoadAfm = App.Path & "\Fonts"
        objPDF.PDFSetUnit = UNIT_PT
        'objPDF.PDFView = True
        objPDF.PDFBeginDoc
          
                    
        y = 5
        x1 = (objPDF.PDFGetPageWidth / 2) - 110
        objPDF.PDFImage App.Path & "\Logo\Logo.jpg", x1, y
        y = 150
                
        x1 = 80
        objPDF.PDFSetFont FONT_ARIAL, 22, FONT_BOLD
        objPDF.PDFSetTextColor = vbBlue
        objPDF.PDFTextOut "CHANAKYA DIAGNOSTIC LABORATORY", x1, y
        'Header Creation
        y = y + 20

        objPDF.PDFSetFont FONT_ARIAL, 11, FONT_BOLD
        objPDF.PDFTextOut "BHOJO ROAD, SONARI- 785690, ASSAM, PH. 7399199668", 150, y


        objPDF.PDFSetFont FONT_ARIAL, 10, FONT_NORMAL
        objPDF.PDFSetTextColor = vbBlack
        y = y + 30
        x = 10
        objPDF.PDFDrawLine x, y, objPDF.PDFGetPageWidth - x, y
        y = y + 30
        objPDF.PDFTextOut "Report No", x, y

        x1 = 65
        objPDF.PDFSetFont FONT_ARIAL, 10, FONT_BOLD
        objPDF.PDFTextOut "" & rs!BillNo, x1, y - 3
        objPDF.PDFSetFont FONT_ARIAL, 10, FONT_NORMAL
        objPDF.PDFDrawLine 60, y, 200, y

        objPDF.PDFTextOut "Name of Patient", 205, y

        x1 = 290
        objPDF.PDFSetFont FONT_ARIAL, 10, FONT_BOLD
        objPDF.PDFTextOut "" & rs!Patient, x1, y - 3
        objPDF.PDFSetFont FONT_ARIAL, 10, FONT_NORMAL
        objPDF.PDFDrawLine 283, y, objPDF.PDFGetPageWidth - x, y

        y = y + 30
        objPDF.PDFTextOut "Age", x, y

        x1 = 45
        objPDF.PDFSetFont FONT_ARIAL, 10, FONT_BOLD
        objPDF.PDFTextOut Val("" & rs!Age), x1, y - 3
        objPDF.PDFSetFont FONT_ARIAL, 10, FONT_NORMAL
        objPDF.PDFDrawLine 40, y, 120, y

        objPDF.PDFTextOut "Gender", 125, y

        x1 = 170
        objPDF.PDFSetFont FONT_ARIAL, 10, FONT_BOLD
        objPDF.PDFTextOut "" & rs!Sex, x1, y - 3
        objPDF.PDFSetFont FONT_ARIAL, 10, FONT_NORMAL
        objPDF.PDFDrawLine 165, y, 250, y

        x1 = 255
        objPDF.PDFTextOut "Contact No", x1, y

        x1 = 315
        objPDF.PDFSetFont FONT_ARIAL, 10, FONT_BOLD
        objPDF.PDFTextOut "" & rs!Mobile, x1, y - 3
        objPDF.PDFSetFont FONT_ARIAL, 10, FONT_NORMAL
        objPDF.PDFDrawLine 310, y, objPDF.PDFGetPageWidth - x, y

        y = y + 30
        objPDF.PDFTextOut "Ref. by ", x, y

        x1 = 55
        objPDF.PDFSetFont FONT_ARIAL, 10, FONT_BOLD
        objPDF.PDFTextOut "" & rs!Doctor, x1, y - 3
        objPDF.PDFSetFont FONT_ARIAL, 10, FONT_NORMAL
        objPDF.PDFDrawLine 50, y, 400, y

        x1 = 405
        objPDF.PDFTextOut "Date", x1, y

        x1 = 435
        objPDF.PDFSetFont FONT_ARIAL, 10, FONT_BOLD
        objPDF.PDFTextOut Format("" & rs!BillDate, "dd-MM-yyyy"), x1, y - 3
        objPDF.PDFSetFont FONT_ARIAL, 10, FONT_NORMAL
        objPDF.PDFDrawLine 430, y, objPDF.PDFGetPageWidth - x, y

        y = y + 30
        objPDF.PDFDrawLine x, y, objPDF.PDFGetPageWidth - x, y
        
        'Table creation
        y = y + 60

        'Box for departmeant
        y1 = y - 3
        objPDF.PDFDrawLine x, y1, objPDF.PDFGetPageWidth - x, y1
        objPDF.PDFDrawLine x, y1 + 30, objPDF.PDFGetPageWidth - x, y1 + 30
        objPDF.PDFDrawLine x, y1, x, y1 + 30
        objPDF.PDFDrawLine objPDF.PDFGetPageWidth - x, y1, objPDF.PDFGetPageWidth - x, y1 + 30
        'Box for heading
        y1 = y1 + 30
        objPDF.PDFDrawLine x, y1, objPDF.PDFGetPageWidth - x, y1
        objPDF.PDFDrawLine x, y1 + 30, objPDF.PDFGetPageWidth - x, y1 + 30
        objPDF.PDFDrawLine x, y1, x, y1 + 30
        objPDF.PDFDrawLine objPDF.PDFGetPageWidth - x, y1, objPDF.PDFGetPageWidth - x, y1 + 30
        x1 = Val(objPDF.PDFGetPageWidth - x) / 2
        objPDF.PDFDrawLine x1, y1, x1, y1 + 30
        x1 = (Val(objPDF.PDFGetPageWidth - x) / 2) + (Val(objPDF.PDFGetPageWidth - x) / 4)
        objPDF.PDFDrawLine x1, y1, x1, y1 + 30

        
        Dim dep As String, inves As String
        dep = ""
        inves = ""

        r = 0
        rs.MoveFirst
        While rs.EOF = False

            If y > objPDF.PDFGetPageHeight - 200 Then

                objPDF.PDFDrawLine x, y1, x, y
                objPDF.PDFDrawLine objPDF.PDFGetPageWidth - x, y1, objPDF.PDFGetPageWidth - x, y
                x1 = Val(objPDF.PDFGetPageWidth - x) / 2
                objPDF.PDFDrawLine x1, y1, x1, y
                x1 = (Val(objPDF.PDFGetPageWidth - x) / 2) + (Val(objPDF.PDFGetPageWidth - x) / 4)
                objPDF.PDFDrawLine x1, y1, x1, y
                x1 = (Val(objPDF.PDFGetPageWidth - x) / 4)
                objPDF.PDFDrawLine x1, y1 + 30, x1, y
                objPDF.PDFDrawLine x, y, objPDF.PDFGetPageWidth - x, y

                objPDF.PDFEndPage
                objPDF.PDFNewPage
                'Printer.EndDoc

                y = 100

                'Box for departmeant
                y1 = y - 3
                objPDF.PDFDrawLine x, y1, objPDF.PDFGetPageWidth - x, y1
                objPDF.PDFDrawLine x, y1 + 30, objPDF.PDFGetPageWidth - x, y1 + 30
                objPDF.PDFDrawLine x, y1, x, y1 + 30
                objPDF.PDFDrawLine objPDF.PDFGetPageWidth - x, y1, objPDF.PDFGetPageWidth - x, y1 + 30
                'Box for heading
                y1 = y1 + 30
                objPDF.PDFDrawLine x, y1, objPDF.PDFGetPageWidth - x, y1
                objPDF.PDFDrawLine x, y1 + 30, objPDF.PDFGetPageWidth - x, y1 + 30
                objPDF.PDFDrawLine x, y1, x, y1 + 30
                objPDF.PDFDrawLine objPDF.PDFGetPageWidth - x, y1, objPDF.PDFGetPageWidth - x, y1 + 30
                x1 = Val(objPDF.PDFGetPageWidth - x) / 2
                objPDF.PDFDrawLine x1, y1, x1, y1 + 30
                x1 = (Val(objPDF.PDFGetPageWidth - x) / 2) + (Val(objPDF.PDFGetPageWidth - x) / 4)
                objPDF.PDFDrawLine x1, y1, x1, y1 + 30

                'Department
                objPDF.PDFSetFont FONT_ARIAL, 12, FONT_BOLD
                x1 = (Val(objPDF.PDFGetPageWidth - x) / 2) - (Val(objPDF.PDFGetStringWidth(dep, "", 12)) / 2)
                objPDF.PDFTextOut dep, x1, y + 15
                objPDF.PDFSetFont FONT_ARIAL, 10, FONT_NORMAL
                
                y = y + 30
                
                'First header of the department
                x1 = (Val(objPDF.PDFGetPageWidth - x) / 4) - (Val(objPDF.PDFGetStringWidth("Investigation", "", 10)) / 2)
                objPDF.PDFTextOut "Investigation", x1, y + 15
                
                x1 = ((Val(objPDF.PDFGetPageWidth - x) / 2) + (Val(objPDF.PDFGetPageWidth - x) / 8)) - (Val(objPDF.PDFGetStringWidth("Result", "", 10)) / 2)
                objPDF.PDFTextOut "Result", x1, y + 15
               
                x1 = ((Val(objPDF.PDFGetPageWidth - x) / 2) + (3 * (Val(objPDF.PDFGetPageWidth - x) / 8))) - (Val(objPDF.PDFGetStringWidth("Bio. Ref. Interval", "", 10)) / 2)
                objPDF.PDFTextOut "Bio. Ref. Interval", x1, y + 15
               
                y = y + 30
                'First Row including the Test
                x1 = (Val(objPDF.PDFGetPageWidth - x) / 8) - (Val(objPDF.PDFGetStringWidth(Trim("" & rs!Test), "", 10)) / 2)
                objPDF.PDFTextOut Trim("" & rs!Test), x1, y + 15
              
                x1 = ((Val(objPDF.PDFGetPageWidth - x) / 4) + (Val(objPDF.PDFGetPageWidth - x) / 16)) - (Val(objPDF.PDFGetStringWidth(Trim("" & rs!SubTest), "", 10)) / 2)
                objPDF.PDFTextOut Trim("" & rs!SubTest), x1, y + 15
               
                x1 = ((Val(objPDF.PDFGetPageWidth - x) / 2) + (Val(objPDF.PDFGetPageWidth - x) / 8)) - (Val(objPDF.PDFGetStringWidth(Trim("" & rs!TestValue & " " & rs!UnitName), "", 10)) / 2)
                objPDF.PDFTextOut Trim("" & rs!TestValue & " " & rs!UnitName), x1, y + 15
                
                x1 = ((Val(objPDF.PDFGetPageWidth - x) / 2) + (3 * (Val(objPDF.PDFGetPageWidth - x) / 8))) - (Val(objPDF.PDFGetStringWidth(Trim("" & rs!DefaultValue & " " & rs!UnitName), "", 10)) / 2)
                objPDF.PDFTextOut Trim("" & rs!DefaultValue & " " & rs!UnitName), x1, y + 15
            End If

            If dep <> Trim("" & rs!DepartmentName) Then

                If dep <> "" Then

                    objPDF.PDFDrawLine x, y1, x, y
                    objPDF.PDFDrawLine objPDF.PDFGetPageWidth - x, y1, objPDF.PDFGetPageWidth - x, y
                    x1 = Val(objPDF.PDFGetPageWidth - x) / 2
                    objPDF.PDFDrawLine x1, y1, x1, y
                    x1 = (Val(objPDF.PDFGetPageWidth - x) / 2) + (Val(objPDF.PDFGetPageWidth - x) / 4)
                    objPDF.PDFDrawLine x1, y1, x1, y
                    x1 = (Val(objPDF.PDFGetPageWidth - x) / 4)
                    objPDF.PDFDrawLine x1, y1 + 30, x1, y
                    objPDF.PDFDrawLine x, y, objPDF.PDFGetPageWidth - x, y
                    y = y + 60
                End If

                dep = Trim("" & rs!DepartmentName)
                inves = Trim("" & rs!Test)

                'Box for departmeant
                y1 = y - 3
                objPDF.PDFDrawLine x, y1, objPDF.PDFGetPageWidth - x, y1
                objPDF.PDFDrawLine x, y1 + 30, objPDF.PDFGetPageWidth - x, y1 + 30
                objPDF.PDFDrawLine x, y1, x, y1 + 30
                objPDF.PDFDrawLine objPDF.PDFGetPageWidth - x, y1, objPDF.PDFGetPageWidth - x, y1 + 30
                'Box for heading
                y1 = y1 + 30
                objPDF.PDFDrawLine x, y1, objPDF.PDFGetPageWidth - x, y1
                objPDF.PDFDrawLine x, y1 + 30, objPDF.PDFGetPageWidth - x, y1 + 30
                objPDF.PDFDrawLine x, y1, x, y1 + 30
                objPDF.PDFDrawLine objPDF.PDFGetPageWidth - x, y1, objPDF.PDFGetPageWidth - x, y1 + 30
                x1 = Val(objPDF.PDFGetPageWidth - x) / 2
                objPDF.PDFDrawLine x1, y1, x1, y1 + 30
                x1 = (Val(objPDF.PDFGetPageWidth - x) / 2) + (Val(objPDF.PDFGetPageWidth - x) / 4)
                objPDF.PDFDrawLine x1, y1, x1, y1 + 30

                'Department
                objPDF.PDFSetFont FONT_ARIAL, 12, FONT_BOLD
                x1 = (Val(objPDF.PDFGetPageWidth - x) / 2) - (Val(objPDF.PDFGetStringWidth(dep, "", 12)) / 2)
                objPDF.PDFTextOut dep, x1, y + 15
                objPDF.PDFSetFont FONT_ARIAL, 10, FONT_NORMAL
                
                y = y + 30
                
                'First header of the department
                x1 = (Val(objPDF.PDFGetPageWidth - x) / 4) - (Val(objPDF.PDFGetStringWidth("Investigation", "", 10)) / 2)
                objPDF.PDFTextOut "Investigation", x1, y + 15
                
                x1 = ((Val(objPDF.PDFGetPageWidth - x) / 2) + (Val(objPDF.PDFGetPageWidth - x) / 8)) - (Val(objPDF.PDFGetStringWidth("Result", "", 10)) / 2)
                objPDF.PDFTextOut "Result", x1, y + 15
               
                x1 = ((Val(objPDF.PDFGetPageWidth - x) / 2) + (3 * (Val(objPDF.PDFGetPageWidth - x) / 8))) - (Val(objPDF.PDFGetStringWidth("Bio. Ref. Interval", "", 10)) / 2)
                objPDF.PDFTextOut "Bio. Ref. Interval", x1, y + 15
               
                y = y + 30
                'First Row including the Test
                x1 = (Val(objPDF.PDFGetPageWidth - x) / 8) - (Val(objPDF.PDFGetStringWidth(Trim("" & rs!Test), "", 10)) / 2)
                objPDF.PDFTextOut Trim("" & rs!Test), x1, y + 15
              
                x1 = ((Val(objPDF.PDFGetPageWidth - x) / 4) + (Val(objPDF.PDFGetPageWidth - x) / 16)) - (Val(objPDF.PDFGetStringWidth(Trim("" & rs!SubTest), "", 10)) / 2)
                objPDF.PDFTextOut Trim("" & rs!SubTest), x1, y + 15
               
                x1 = ((Val(objPDF.PDFGetPageWidth - x) / 2) + (Val(objPDF.PDFGetPageWidth - x) / 8)) - (Val(objPDF.PDFGetStringWidth(Trim("" & rs!TestValue & " " & rs!UnitName), "", 10)) / 2)
                objPDF.PDFTextOut Trim("" & rs!TestValue & " " & rs!UnitName), x1, y + 15
                
                x1 = ((Val(objPDF.PDFGetPageWidth - x) / 2) + (3 * (Val(objPDF.PDFGetPageWidth - x) / 8))) - (Val(objPDF.PDFGetStringWidth(Trim("" & rs!DefaultValue & " " & rs!UnitName), "", 10)) / 2)
                objPDF.PDFTextOut Trim("" & rs!DefaultValue & " " & rs!UnitName), x1, y + 15
                
            ElseIf inves <> Trim("" & rs!Test) Then
                inves = Trim("" & rs!Test)

                objPDF.PDFDrawLine x, y - 3, objPDF.PDFGetPageWidth - x, y - 3

                x1 = (Val(objPDF.PDFGetPageWidth - x) / 8) - (Val(objPDF.PDFGetStringWidth(Trim("" & rs!Test), "", 10)) / 2)
                objPDF.PDFTextOut Trim("" & rs!Test), x1, y + 15
            End If

            x1 = ((Val(objPDF.PDFGetPageWidth - x) / 4) + (Val(objPDF.PDFGetPageWidth - x) / 16)) - (Val(objPDF.PDFGetStringWidth(Trim("" & rs!SubTest), "", 10)) / 2)
            objPDF.PDFTextOut Trim("" & rs!SubTest), x1, y + 15

            x1 = ((Val(objPDF.PDFGetPageWidth - x) / 2) + (Val(objPDF.PDFGetPageWidth - x) / 8)) - (Val(objPDF.PDFGetStringWidth(Trim("" & rs!TestValue & " " & rs!UnitName), "", 10)) / 2)
            objPDF.PDFTextOut Trim("" & rs!TestValue & " " & rs!UnitName), x1, y + 15
            
            x1 = ((Val(objPDF.PDFGetPageWidth - x) / 2) + (3 * (Val(objPDF.PDFGetPageWidth - x) / 8))) - (Val(objPDF.PDFGetStringWidth(Trim("" & rs!DefaultValue & " " & rs!UnitName), "", 10)) / 2)
            objPDF.PDFTextOut Trim("" & rs!DefaultValue & " " & rs!UnitName), x1, y + 15
            
            y = y + 30
            r = r + 1
            rs.MoveNext
        Wend
        rs.Close

        objPDF.PDFDrawLine x, y1, x, y
        objPDF.PDFDrawLine objPDF.PDFGetPageWidth - x, y1, objPDF.PDFGetPageWidth - x, y
        x1 = Val(objPDF.PDFGetPageWidth - x) / 2
        objPDF.PDFDrawLine x1, y1, x1, y
        x1 = (Val(objPDF.PDFGetPageWidth - x) / 2) + (Val(objPDF.PDFGetPageWidth - x) / 4)
        objPDF.PDFDrawLine x1, y1, x1, y
        x1 = (Val(objPDF.PDFGetPageWidth - x) / 4)
        objPDF.PDFDrawLine x1, y1 + 30, x1, y
        objPDF.PDFDrawLine x, y, objPDF.PDFGetPageWidth - x, y

        y = y + 60
        x1 = (Val(objPDF.PDFGetPageWidth - x) / 2) + (Val(objPDF.PDFGetPageWidth - x) / 4)
        objPDF.PDFTextOut "Signature", x1, y + 15
       

        ' End our PDF document (this will save it to the filename)
        objPDF.PDFEndDoc
        Set objPDF = Nothing
    Else
        rs.Close
    End If
    
    Exit Sub
    
TryAgain:
    If (count > 2) Then
        MsgBox "Error Printing pdf!"
        Exit Sub
    End If

    GoTo Start
End Sub

Public Sub printMedicalDetails()
Dim rs As Recordset, r As Long
Dim x As Long, y As Long, x1 As Long, y1 As Long
        
Dim myPicture As IPictureDisp
  'Load the picture into the variable
  Set myPicture = LoadPicture(App.Path & "\Logo\Logo.jpg")
            
        
    Set rs = db.OpenRecordset("Select (Select Units.UnitName From Units Where (Units.Code = TR1.UnitCode )) As UnitName,Transaction.*,TR1.TestName As SubTest,TR1.DefaultValue,TR2.TestName As Test,Department.DepartmentName From TestRegister As TR1,TestRegister As TR2,Transaction,Department Where (Transaction.BillNo = '" & Trim(TTransactionNo.Text) & "' ) And (Transaction.BillType = 'T' ) And (Department.DepartmentCode=TR1.DepartmentCode) And (TR1.Code = Transaction.SubTestCode ) And (TR2.Code = Transaction.TestCode ) And (Transaction.FinancialCode='" & getFinancialCode(DTPDate.Value) & "') Order By Transaction.SerialNo")
    
    If rs.RecordCount > 0 Then
    
        Printer.ScaleMode = 1
        
        y = 500
        x1 = 4400
        Printer.PaintPicture myPicture, x1, y
        y = 2500
                
        Printer.FontName = "Arial"
    
        
        'Header Creation
        Printer.FontSize = 22
        Printer.FontBold = True
        Printer.ForeColor = vbBlue
        Printer.FontUnderline = False
        Printer.CurrentX = (Val(Printer.Width) / 2) - (Val(Printer.TextWidth("CHANAKYA DIAGNOSTIC LABORATORY")) / 2)
        Printer.CurrentY = y
        Printer.Print "CHANAKYA DIAGNOSTIC LABORATORY"
        y = y + 700
    
        Printer.FontSize = 11
        Printer.CurrentX = (Val(Printer.Width) / 2) - (Val(Printer.TextWidth("BHOJO ROAD, SONARI- 785690, ASSAM, PH. 7399199668")) / 2)
        Printer.CurrentY = y
        Printer.Print "BHOJO ROAD, SONARI- 785690, ASSAM, PH. 7399199668"
    
        Printer.FontSize = 10
        Printer.FontBold = False
        Printer.ForeColor = vbBlack
        y = y + 700
        x = 500
        Printer.Line (x, y)-(Printer.Width - x, y)
                
        y = y + 300
        Printer.CurrentX = x
        Printer.CurrentY = y
        Printer.Print "Report No"
                
        x1 = Val(Printer.TextWidth("Report No") + x)
        Printer.FontBold = True
        Printer.CurrentX = x1 + 100
        Printer.CurrentY = y - 50
        Printer.Print "" & rs!BillNo
        Printer.FontBold = False
        Printer.Line (x1, y + 200)-(x1 + 3000, y + 200)
                
        Printer.CurrentX = x1 + 3000
        Printer.CurrentY = y
        Printer.Print "Name of Patient"
        
        x1 = x1 + Val(Printer.TextWidth("Name of Patient")) + 3000
        Printer.FontBold = True
        Printer.CurrentX = x1 + 100
        Printer.CurrentY = y - 50
        Printer.Print "" & rs!Patient
        Printer.FontBold = False
        Printer.Line (x1, y + 200)-(Printer.Width - x, y + 200)
        
        y = y + 500
        Printer.CurrentX = x
        Printer.CurrentY = y
        Printer.Print "Age"
        
        x1 = Val(Printer.TextWidth("Age")) + x
        Printer.FontBold = True
        Printer.CurrentX = x1 + 100
        Printer.CurrentY = y - 50
        Printer.Print Val("" & rs!Age)
        Printer.FontBold = False
        Printer.Line (x1, y + 200)-(x1 + 2000, y + 200)
        
        Printer.CurrentX = x1 + 2000
        Printer.CurrentY = y
        Printer.Print "Gender"
        
        x1 = x1 + Val(Printer.TextWidth("Gender")) + 2000
        Printer.FontBold = True
        Printer.CurrentX = x1 + 100
        Printer.CurrentY = y - 50
        Printer.Print "" & rs!Sex
        Printer.FontBold = False
        Printer.Line (x1, y + 200)-(x1 + 2000, y + 200)
        
        x1 = x1 + 2000
        Printer.CurrentX = x1
        Printer.CurrentY = y
        Printer.Print "Contact No"
        
        x1 = x1 + Val(Printer.TextWidth("Contact No"))
        Printer.FontBold = True
        Printer.CurrentX = x1 + 100
        Printer.CurrentY = y - 50
        Printer.Print "" & rs!Mobile
        Printer.FontBold = False
        Printer.Line (x1, y + 200)-(Printer.Width - x, y + 200)
        
        y = y + 500
        Printer.CurrentX = x
        Printer.CurrentY = y
        Printer.Print "Ref. by "
                
        x1 = Val(Printer.TextWidth("Ref. by ") + x)
        Printer.FontBold = True
        Printer.CurrentX = x1 + 100
        Printer.CurrentY = y - 50
        Printer.Print "" & rs!Doctor
        Printer.FontBold = False
        Printer.Line (x1, y + 200)-(x1 + 7000, y + 200)
        
        x1 = x1 + 7000
        Printer.CurrentX = x1
        Printer.CurrentY = y
        Printer.Print "Date"
        
        x1 = x1 + Val(Printer.TextWidth("Date"))
        Printer.FontBold = True
        Printer.CurrentX = x1 + 100
        Printer.CurrentY = y - 50
        Printer.Print Format("" & rs!BillDate, "dd-MM-yyyy")
        Printer.FontBold = False
        Printer.Line (x1, y + 200)-(Printer.Width - x, y + 200)
        
        y = y + 500
        Printer.Line (x, y + 200)-(Printer.Width - x, y + 200)
                    
        
        'Table creation
        y = y + 1000
        
        'Box for departmeant
        y1 = y - 100
        Printer.Line (x, y1)-(Printer.Width - x, y1)
        Printer.Line (x, y1 + 500)-(Printer.Width - x, y1 + 500)
        Printer.Line (x, y1)-(x, y1 + 500)
        Printer.Line (Printer.Width - x, y1)-(Printer.Width - x, y1 + 500)
        'Box for heading
        y1 = y1 + 500
        Printer.Line (x, y1)-(Printer.Width - x, y1)
        Printer.Line (x, y1 + 500)-(Printer.Width - x, y1 + 500)
        Printer.Line (x, y1)-(x, y1 + 500)
        Printer.Line (Printer.Width - x, y1)-(Printer.Width - x, y1 + 500)
        x1 = Val(Printer.Width - x) / 2
        Printer.Line (x1, y1)-(x1, y1 + 500)
        x1 = (Val(Printer.Width - x) / 2) + (Val(Printer.Width - x) / 4)
        Printer.Line (x1, y1)-(x1, y1 + 500)
        
        Dim dep As String, inves As String
        Dim nextPageStart As Boolean
        nextPageStart = False
        dep = ""
        inves = ""
        
        r = 0
        rs.MoveFirst
        While rs.EOF = False
            
            If y > Printer.Height - 3000 Then
                                                    
                Printer.Line (x, y1)-(x, y)
                Printer.Line (Printer.Width - x, y1)-(Printer.Width - x, y)
                x1 = Val(Printer.Width - x) / 2
                Printer.Line (x1, y1)-(x1, y)
                x1 = (Val(Printer.Width - x) / 2) + (Val(Printer.Width - x) / 4)
                Printer.Line (x1, y1)-(x1, y)
                x1 = (Val(Printer.Width - x) / 4)
                Printer.Line (x1, y1 + 500)-(x1, y)
                Printer.Line (x, y)-(Printer.Width - x, y)
                
                Printer.EndDoc
                nextPageStart = True
                
                y = 1000
                
                'Box for departmeant
                y1 = y - 100
                Printer.Line (x, y1)-(Printer.Width - x, y1)
                Printer.Line (x, y1 + 500)-(Printer.Width - x, y1 + 500)
                Printer.Line (x, y1)-(x, y1 + 500)
                Printer.Line (Printer.Width - x, y1)-(Printer.Width - x, y1 + 500)
                'Box for heading
                y1 = y1 + 500
                Printer.Line (x, y1)-(Printer.Width - x, y1)
                Printer.Line (x, y1 + 500)-(Printer.Width - x, y1 + 500)
                Printer.Line (x, y1)-(x, y1 + 500)
                Printer.Line (Printer.Width - x, y1)-(Printer.Width - x, y1 + 500)
                x1 = Val(Printer.Width - x) / 2
                Printer.Line (x1, y1)-(x1, y1 + 500)
                x1 = (Val(Printer.Width - x) / 2) + (Val(Printer.Width - x) / 4)
                Printer.Line (x1, y1)-(x1, y1 + 500)
                        
                'Department
                Printer.FontSize = 12
                Printer.FontBold = True
                Printer.CurrentX = (Val(Printer.Width - x) / 2) - (Val(Printer.TextWidth(dep)) / 2)
                Printer.CurrentY = y
                Printer.Print dep
                Printer.FontSize = 10
                Printer.FontBold = False
                
                y = y + 500
                'First header of the department
                Printer.CurrentX = (Val(Printer.Width - x) / 4) - (Val(Printer.TextWidth("Investigation")) / 2)
                Printer.CurrentY = y
                Printer.Print "Investigation"
                                
                Printer.CurrentX = ((Val(Printer.Width - x) / 2) + (Val(Printer.Width - x) / 8)) - (Val(Printer.TextWidth("Result")) / 2)
                Printer.CurrentY = y
                Printer.Print "Result"
                
                Printer.CurrentX = ((Val(Printer.Width - x) / 2) + (3 * (Val(Printer.Width - x) / 8))) - (Val(Printer.TextWidth("Bio. Ref. Interval")) / 2)
                Printer.CurrentY = y
                Printer.Print "Bio. Ref. Interval"
                
                y = y + 500
                'First Row including the Test
                Printer.CurrentX = (Val(Printer.Width - x) / 8) - (Val(Printer.TextWidth(Trim("" & rs!Test))) / 2)
                Printer.CurrentY = y
                Printer.Print Trim("" & rs!Test)
                
                Printer.CurrentX = ((Val(Printer.Width - x) / 4) + (Val(Printer.Width - x) / 16)) - (Val(Printer.TextWidth(Trim("" & rs!SubTest))) / 2)
                Printer.CurrentY = y
                Printer.Print Trim("" & rs!SubTest)
                
                Printer.CurrentX = ((Val(Printer.Width - x) / 2) + (Val(Printer.Width - x) / 8)) - (Val(Printer.TextWidth(Trim("" & rs!TestValue & " " & rs!UnitName))) / 2)
                Printer.CurrentY = y
                Printer.Print Trim("" & rs!TestValue & " " & rs!UnitName)
                
                Printer.CurrentX = ((Val(Printer.Width - x) / 2) + (3 * (Val(Printer.Width - x) / 8))) - (Val(Printer.TextWidth(Trim("" & rs!DefaultValue & " " & rs!UnitName))) / 2)
                Printer.CurrentY = y
                Printer.Print Trim("" & rs!DefaultValue & " " & rs!UnitName)
            End If

            If dep <> Trim("" & rs!DepartmentName) Then
                    
                If dep <> "" Then
                    'closing the box
                    If nextPageStart Then
                        y = y + 400
                    End If
                    
                    Printer.Line (x, y1)-(x, y)
                    Printer.Line (Printer.Width - x, y1)-(Printer.Width - x, y)
                    x1 = Val(Printer.Width - x) / 2
                    Printer.Line (x1, y1)-(x1, y)
                    x1 = (Val(Printer.Width - x) / 2) + (Val(Printer.Width - x) / 4)
                    Printer.Line (x1, y1)-(x1, y)
                    x1 = (Val(Printer.Width - x) / 4)
                    Printer.Line (x1, y1 + 500)-(x1, y)
                    Printer.Line (x, y)-(Printer.Width - x, y)
                    
                    If nextPageStart Then
                        nextPageStart = False
                        y = y - 400
                    End If
                    y = y + 1000
                End If
                
                dep = Trim("" & rs!DepartmentName)
                inves = Trim("" & rs!Test)
                    
                'Box for departmeant
                y1 = y - 100
                Printer.Line (x, y1)-(Printer.Width - x, y1)
                Printer.Line (x, y1 + 500)-(Printer.Width - x, y1 + 500)
                Printer.Line (x, y1)-(x, y1 + 500)
                Printer.Line (Printer.Width - x, y1)-(Printer.Width - x, y1 + 500)
                'Box for heading
                y1 = y1 + 500
                Printer.Line (x, y1)-(Printer.Width - x, y1)
                Printer.Line (x, y1 + 500)-(Printer.Width - x, y1 + 500)
                Printer.Line (x, y1)-(x, y1 + 500)
                Printer.Line (Printer.Width - x, y1)-(Printer.Width - x, y1 + 500)
                x1 = Val(Printer.Width - x) / 2
                Printer.Line (x1, y1)-(x1, y1 + 500)
                x1 = (Val(Printer.Width - x) / 2) + (Val(Printer.Width - x) / 4)
                Printer.Line (x1, y1)-(x1, y1 + 500)
                        
                'Department
                Printer.FontSize = 12
                Printer.FontBold = True
                Printer.CurrentX = (Val(Printer.Width - x) / 2) - (Val(Printer.TextWidth(dep)) / 2)
                Printer.CurrentY = y
                Printer.Print dep
                Printer.FontSize = 10
                Printer.FontBold = False
                
                y = y + 500
                'First header of the department
                Printer.CurrentX = (Val(Printer.Width - x) / 4) - (Val(Printer.TextWidth("Investigation")) / 2)
                Printer.CurrentY = y
                Printer.Print "Investigation"
                                
                Printer.CurrentX = ((Val(Printer.Width - x) / 2) + (Val(Printer.Width - x) / 8)) - (Val(Printer.TextWidth("Result")) / 2)
                Printer.CurrentY = y
                Printer.Print "Result"
                
                Printer.CurrentX = ((Val(Printer.Width - x) / 2) + (3 * (Val(Printer.Width - x) / 8))) - (Val(Printer.TextWidth("Bio. Ref. Interval")) / 2)
                Printer.CurrentY = y
                Printer.Print "Bio. Ref. Interval"
                
                y = y + 500
                'First Row including the Test
                Printer.CurrentX = (Val(Printer.Width - x) / 8) - (Val(Printer.TextWidth(Trim("" & rs!Test))) / 2)
                Printer.CurrentY = y
                Printer.Print Trim("" & rs!Test)
                
                Printer.CurrentX = ((Val(Printer.Width - x) / 4) + (Val(Printer.Width - x) / 16)) - (Val(Printer.TextWidth(Trim("" & rs!SubTest))) / 2)
                Printer.CurrentY = y
                Printer.Print Trim("" & rs!SubTest)
                
                Printer.CurrentX = ((Val(Printer.Width - x) / 2) + (Val(Printer.Width - x) / 8)) - (Val(Printer.TextWidth(Trim("" & rs!TestValue & " " & rs!UnitName))) / 2)
                Printer.CurrentY = y
                Printer.Print Trim("" & rs!TestValue & " " & rs!UnitName)
                
                Printer.CurrentX = ((Val(Printer.Width - x) / 2) + (3 * (Val(Printer.Width - x) / 8))) - (Val(Printer.TextWidth(Trim("" & rs!DefaultValue & " " & rs!UnitName))) / 2)
                Printer.CurrentY = y
                Printer.Print Trim("" & rs!DefaultValue & " " & rs!UnitName)
            ElseIf inves <> Trim("" & rs!Test) Then
                inves = Trim("" & rs!Test)
                
                Printer.Line (x, y - 100)-(Printer.Width - x, y - 100)
                
                Printer.CurrentX = (Val(Printer.Width - x) / 8) - (Val(Printer.TextWidth(Trim("" & rs!Test))) / 2)
                Printer.CurrentY = y
                Printer.Print Trim("" & rs!Test)
            End If

            Printer.CurrentX = ((Val(Printer.Width - x) / 4) + (Val(Printer.Width - x) / 16)) - (Val(Printer.TextWidth(Trim("" & rs!SubTest))) / 2)
            Printer.CurrentY = y
            Printer.Print Trim("" & rs!SubTest)
            
            Printer.CurrentX = ((Val(Printer.Width - x) / 2) + (Val(Printer.Width - x) / 8)) - (Val(Printer.TextWidth(Trim("" & rs!TestValue & " " & rs!UnitName))) / 2)
            Printer.CurrentY = y
            Printer.Print Trim("" & rs!TestValue & " " & rs!UnitName)
            
            Printer.CurrentX = ((Val(Printer.Width - x) / 2) + (3 * (Val(Printer.Width - x) / 8))) - (Val(Printer.TextWidth(Trim("" & rs!DefaultValue & " " & rs!UnitName))) / 2)
            Printer.CurrentY = y
            Printer.Print Trim("" & rs!DefaultValue & " " & rs!UnitName)
            
            y = y + 500
            r = r + 1
            rs.MoveNext
        Wend
        rs.Close
        
        Printer.Line (x, y1)-(x, y)
        Printer.Line (Printer.Width - x, y1)-(Printer.Width - x, y)
        x1 = Val(Printer.Width - x) / 2
        Printer.Line (x1, y1)-(x1, y)
        x1 = (Val(Printer.Width - x) / 2) + (Val(Printer.Width - x) / 4)
        Printer.Line (x1, y1)-(x1, y)
        x1 = (Val(Printer.Width - x) / 4)
        Printer.Line (x1, y1 + 500)-(x1, y)
        Printer.Line (x, y)-(Printer.Width - x, y)
        
        y = y + 1000
        Printer.FontBold = False
        x1 = (Val(Printer.Width - x) / 2) + (Val(Printer.Width - x) / 4)
        Printer.CurrentX = x1
        Printer.CurrentY = y
        Printer.Print "Signature"
        
        
        Printer.EndDoc
    Else
        rs.Close
    End If
    
    
End Sub

Public Sub printBillDetails()
Dim rs As Recordset, r As Long
Dim x As Long, y As Long, x1 As Long, y1 As Long
        
Dim myPicture As IPictureDisp
  'Load the picture into the variable
  Set myPicture = LoadPicture(App.Path & "\Logo\Logo.jpg")
        
    Set rs = db.OpenRecordset("Select (Select Units.UnitName From Units Where (Units.Code = TR1.UnitCode )) As UnitName,Transaction.*,TR1.TestName As SubTest,TR2.TestName As Test,Department.DepartmentName From TestRegister As TR1,TestRegister As TR2,Transaction,Department Where (Transaction.BillNo = '" & Trim(TTransactionNo.Text) & "' ) And (Transaction.BillType = 'T' ) And (Department.DepartmentCode=TR1.DepartmentCode) And (TR1.Code = Transaction.SubTestCode ) And (TR2.Code = Transaction.TestCode ) And (Transaction.FinancialCode='" & getFinancialCode(DTPDate.Value) & "') Order By Transaction.SerialNo")
    
    If rs.RecordCount > 0 Then
        Printer.ScaleMode = 1
        
        y = 500
        x1 = 4400
        Printer.PaintPicture myPicture, x1, y
        y = 2500
                
        Printer.FontName = "Arial"
    
        'header creation
        Printer.FontSize = 22
        Printer.FontBold = True
        Printer.ForeColor = vbBlue
        Printer.FontUnderline = False
        Printer.CurrentX = (Val(Printer.Width) / 2) - (Val(Printer.TextWidth("CHANAKYA DIAGNOSTIC LABORATORY")) / 2)
        Printer.CurrentY = y
        Printer.Print "CHANAKYA DIAGNOSTIC LABORATORY"
        y = y + 700
    
        Printer.FontSize = 11
        Printer.CurrentX = (Val(Printer.Width) / 2) - (Val(Printer.TextWidth("BHOJO ROAD, SONARI- 785690, ASSAM, PH. 7399199668")) / 2)
        Printer.CurrentY = y
        Printer.Print "BHOJO ROAD, SONARI- 785690, ASSAM, PH. 7399199668"
    
        Printer.FontSize = 10
        Printer.FontBold = False
        Printer.ForeColor = vbBlack
        y = y + 700
        x = 500
        Printer.Line (x, y)-(Printer.Width - x, y)
                
        y = y + 300
        Printer.CurrentX = x
        Printer.CurrentY = y
        Printer.Print "Bill No"
                
        x1 = Val(Printer.TextWidth("Bill No") + x)
        Printer.FontBold = True
        Printer.CurrentX = x1 + 100
        Printer.CurrentY = y - 50
        Printer.Print "" & rs!BillNo
        Printer.FontBold = False
        Printer.Line (x1, y + 200)-(x1 + 3000, y + 200)
                
        Printer.CurrentX = x1 + 3000
        Printer.CurrentY = y
        Printer.Print "Name of Patient"
        
        x1 = x1 + Val(Printer.TextWidth("Name of Patient")) + 3000
        Printer.FontBold = True
        Printer.CurrentX = x1 + 100
        Printer.CurrentY = y - 50
        Printer.Print "" & rs!Patient
        Printer.FontBold = False
        Printer.Line (x1, y + 200)-(Printer.Width - x, y + 200)
        
        y = y + 500
        Printer.CurrentX = x
        Printer.CurrentY = y
        Printer.Print "Age"
        
        x1 = Val(Printer.TextWidth("Age")) + x
        Printer.FontBold = True
        Printer.CurrentX = x1 + 100
        Printer.CurrentY = y - 50
        Printer.Print Val("" & rs!Age)
        Printer.FontBold = False
        Printer.Line (x1, y + 200)-(x1 + 2000, y + 200)
        
        Printer.CurrentX = x1 + 2000
        Printer.CurrentY = y
        Printer.Print "Gender"
        
        x1 = x1 + Val(Printer.TextWidth("Gender")) + 2000
        Printer.FontBold = True
        Printer.CurrentX = x1 + 100
        Printer.CurrentY = y - 50
        Printer.Print "" & rs!Sex
        Printer.FontBold = False
        Printer.Line (x1, y + 200)-(x1 + 2000, y + 200)
        
        x1 = x1 + 2000
        Printer.CurrentX = x1
        Printer.CurrentY = y
        Printer.Print "Contact No"
        
        x1 = x1 + Val(Printer.TextWidth("Contact No"))
        Printer.FontBold = True
        Printer.CurrentX = x1 + 100
        Printer.CurrentY = y - 50
        Printer.Print "" & rs!Mobile
        Printer.FontBold = False
        Printer.Line (x1, y + 200)-(Printer.Width - x, y + 200)
        
        y = y + 500
        Printer.CurrentX = x
        Printer.CurrentY = y
        Printer.Print "Ref. by "
                
        x1 = Val(Printer.TextWidth("Ref. by ") + x)
        Printer.FontBold = True
        Printer.CurrentX = x1 + 100
        Printer.CurrentY = y - 50
        Printer.Print "" & rs!Doctor
        Printer.FontBold = False
        Printer.Line (x1, y + 200)-(x1 + 7000, y + 200)
        
        x1 = x1 + 7000
        Printer.CurrentX = x1
        Printer.CurrentY = y
        Printer.Print "Date"
        
        x1 = x1 + Val(Printer.TextWidth("Date"))
        Printer.FontBold = True
        Printer.CurrentX = x1 + 100
        Printer.CurrentY = y - 50
        Printer.Print Format("" & rs!BillDate, "dd-MM-yyyy")
        Printer.FontBold = False
        Printer.Line (x1, y + 200)-(Printer.Width - x, y + 200)
        
        y = y + 500
        Printer.Line (x, y + 200)-(Printer.Width - x, y + 200)
        
        
        'Table creation
        y = y + 1000
        'Outer Box
        Printer.Line (x, y)-(Printer.Width - x, y)
        Printer.Line (x, y + 500)-(Printer.Width - x, y + 500)
        Printer.Line (x, y)-(x, y + 500)
        Printer.Line (Printer.Width - x, y)-(Printer.Width - x, y + 500)
        'Inner Lines
        Printer.Line (x + 1000, y)-(x + 1000, y + 500)
        x1 = (Printer.Width - (x + 1000 + x + 2000)) / 2
        Printer.Line (x + 1000 + x1, y)-(x + 1000 + x1, y + 500)
        Printer.Line (x + 1000 + x1 + x1, y)-(x + 1000 + x1 + x1, y + 500)
        Printer.Line (Printer.Width - (x + 2000), y)-(Printer.Width - (x + 2000), y + 500)
        'Printing the Values inside the blocks
        x1 = x + 100
        y1 = y + 100
        Printer.FontBold = True
        Printer.CurrentX = x1
        Printer.CurrentY = y1
        Printer.Print "Sl.No"
        x1 = x1 + 1000
        Printer.CurrentX = x1
        Printer.CurrentY = y1
        Printer.Print "Investigation"
        x1 = x1 + (Printer.Width - (x + 1000 + x + 2000)) / 2
        Printer.CurrentX = x1
        Printer.CurrentY = y1
        Printer.Print "Department"
        x1 = x1 + (Printer.Width - (x + 1000 + x + 2000)) / 2
        Printer.CurrentX = x1
        Printer.CurrentY = y1
        Printer.Print "Cost"
        Printer.FontBold = False
        
        Dim dTotal As Double
        dTotal = 0
        
        r = 0
        rs.MoveFirst
        While rs.EOF = False
            'Adding total cost
            dTotal = dTotal + Val("" & rs!Cost)
            
            If y > Printer.Height - 3000 Then
                y = 2000
                Printer.EndDoc
            End If
            
            y = y + 500
            'Outer Box
            Printer.Line (x, y)-(Printer.Width - x, y)
            Printer.Line (x, y + 500)-(Printer.Width - x, y + 500)
            Printer.Line (x, y)-(x, y + 500)
            Printer.Line (Printer.Width - x, y)-(Printer.Width - x, y + 500)
            'Inner Lines
            Printer.Line (x + 1000, y)-(x + 1000, y + 500)
            x1 = (Printer.Width - (x + 1000 + x + 2000)) / 2
            Printer.Line (x + 1000 + x1, y)-(x + 1000 + x1, y + 500)
            Printer.Line (x + 1000 + x1 + x1, y)-(x + 1000 + x1 + x1, y + 500)
            Printer.Line (Printer.Width - (x + 2000), y)-(Printer.Width - (x + 2000), y + 500)
            'Printing the Values inside the blocks
            x1 = x + 100
            y1 = y + 100
            Printer.CurrentX = x1
            Printer.CurrentY = y1
            Printer.Print "" & rs!SerialNo
            x1 = x1 + 1000
            Printer.CurrentX = x1
            Printer.CurrentY = y1
            Printer.Print "" & rs!SubTest
            x1 = x1 + (Printer.Width - (x + 1000 + x + 2000)) / 2
            Printer.CurrentX = x1
            Printer.CurrentY = y1
            Printer.Print "" & rs!DepartmentName
            x1 = x1 + (Printer.Width - (x + 1000 + x + 2000)) / 2
            Printer.CurrentX = x1
            Printer.CurrentY = y1
            Printer.Print Format(Val("" & rs!Cost), "0.00")
             
            r = r + 1
            rs.MoveNext
        Wend
        rs.Close
        
        'Total
        y = y + 1000
        x1 = x + (Printer.Width - (x + 1000 + x + 2000))
        Printer.FontBold = True
        Printer.CurrentX = x1
        Printer.CurrentY = y
        Printer.Print "Total = " & Format(dTotal, "0.00")
        
        y = y + 500
        Printer.FontBold = False
        Printer.CurrentX = x1
        Printer.CurrentY = y
        Printer.Print "Signature"
        
        
        Printer.EndDoc
    Else
        rs.Close
    End If
    
    
End Sub

