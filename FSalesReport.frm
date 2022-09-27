VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FSalesReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sales Report"
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
   Icon            =   "FSalesReport.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FSalesReport.frx":000C
   ScaleHeight     =   8040
   ScaleWidth      =   15270
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CShowDetailed 
      Height          =   505
      Left            =   1815
      Picture         =   "FSalesReport.frx":1FEC4E
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7320
      Width           =   1350
   End
   Begin VB.CommandButton CShowSummary 
      Height          =   505
      Left            =   375
      Picture         =   "FSalesReport.frx":2010B0
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7320
      Width           =   1350
   End
   Begin VB.CommandButton CToExcel 
      Height          =   525
      Left            =   3225
      Picture         =   "FSalesReport.frx":203512
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7305
      Width           =   1365
   End
   Begin VB.CommandButton CClose 
      Cancel          =   -1  'True
      Height          =   505
      Left            =   13605
      Picture         =   "FSalesReport.frx":205974
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7290
      Width           =   1365
   End
   Begin MSFlexGridLib.MSFlexGrid MGrid 
      Height          =   4890
      Left            =   135
      TabIndex        =   9
      Top             =   2055
      Width           =   15030
      _ExtentX        =   26511
      _ExtentY        =   8625
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
      Format          =   93388803
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
      Format          =   93388803
      CurrentDate     =   40458
   End
   Begin MSForms.Label Label5 
      Height          =   420
      Left            =   10410
      TabIndex        =   17
      Top             =   1440
      Width           =   1455
      VariousPropertyBits=   8388627
      Caption         =   "Patient"
      Size            =   "2566;741"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TAddress 
      Height          =   420
      Left            =   12060
      TabIndex        =   16
      Top             =   1395
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
      TabIndex        =   3
      Top             =   525
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
      TabIndex        =   15
      Top             =   540
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
      TabIndex        =   14
      Top             =   135
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
      TabIndex        =   4
      Top             =   960
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
      TabIndex        =   13
      Top             =   990
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
      Left            =   165
      TabIndex        =   12
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
      TabIndex        =   11
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
      TabIndex        =   10
      Top             =   -60
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "FSalesReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sDoctorCode() As String
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


Private Sub CoTest_Change()
    getSubTest
End Sub

Private Sub CShowDetailed_Click()
Dim gDate As Single, gBillNo As Single, gDealer As Single, gItem As Single, gTax As Single, gRate As Single, gQuantity As Single, gGrossValue As Single, gTaxAmount As Single, gTotalAmount As Single
Dim rs As Recordset
    
    'INITIALISING GRID
    gDate = 0
    gBillNo = 1
    gDealer = 2
    gItem = 3
    gTax = 4
    gRate = 5
    gQuantity = 6
    gGrossValue = 7
    gTaxAmount = 8
    gTotalAmount = 9
    
    MGrid.Clear
    MGrid.Rows = 1 'FOR SKIPING ERROR
    MGrid.Cols = 1 'FOR SKIPING ERROR
    MGrid.FixedCols = 0
    MGrid.FixedRows = 0
    MGrid.Cols = 10
    MGrid.Rows = 1
    MGrid.ColWidth(gDate) = 1400
    MGrid.ColWidth(gBillNo) = 1200
    MGrid.ColWidth(gDealer) = 2000
    MGrid.ColWidth(gItem) = 2000
    MGrid.ColWidth(gTax) = 1500
    MGrid.ColWidth(gRate) = 1500
    MGrid.ColWidth(gQuantity) = 1500
    MGrid.ColWidth(gGrossValue) = 1500
    MGrid.ColWidth(gTaxAmount) = 1500
    MGrid.ColWidth(gTotalAmount) = 1500
    
    MGrid.Col = gDate
    MGrid.CellAlignment = vbAlignRight
    MGrid.CellFontBold = True
    MGrid.CellFontName = "Cosmic San"
    MGrid.CellFontSize = 10
    MGrid.TextMatrix(0, gDate) = "Date"
    
    MGrid.Col = gBillNo
    MGrid.CellAlignment = vbAlignRight
    MGrid.CellFontBold = True
    MGrid.CellFontName = "Cosmic San"
    MGrid.CellFontSize = 10
    MGrid.TextMatrix(0, gBillNo) = "Bill No"
    
    MGrid.Col = gDealer
    MGrid.CellAlignment = vbAlignRight
    MGrid.CellFontBold = True
    MGrid.CellFontName = "Cosmic San"
    MGrid.CellFontSize = 10
    MGrid.TextMatrix(0, gDealer) = "Doctor"
    
    MGrid.Col = gItem
    MGrid.CellAlignment = vbAlignRight
    MGrid.CellFontBold = True
    MGrid.CellFontName = "Cosmic San"
    MGrid.CellFontSize = 10
    MGrid.TextMatrix(0, gItem) = "Item"
    
    MGrid.Col = gTax
    MGrid.CellAlignment = vbAlignRight
    MGrid.CellFontBold = True
    MGrid.CellFontName = "Cosmic San"
    MGrid.CellFontSize = 10
    MGrid.TextMatrix(0, gTax) = "Tax"
    
    MGrid.Col = gRate
    MGrid.CellAlignment = vbAlignRight
    MGrid.CellFontBold = True
    MGrid.CellFontName = "Cosmic San"
    MGrid.CellFontSize = 10
    MGrid.TextMatrix(0, gRate) = "Sale Rate"
    
    MGrid.Col = gQuantity
    MGrid.CellAlignment = vbAlignRight
    MGrid.CellFontBold = True
    MGrid.CellFontName = "Cosmic San"
    MGrid.CellFontSize = 10
    MGrid.TextMatrix(0, gQuantity) = "Quantity"
    
    MGrid.Col = gGrossValue
    MGrid.CellAlignment = vbAlignRight
    MGrid.CellFontBold = True
    MGrid.CellFontName = "Cosmic San"
    MGrid.CellFontSize = 10
    MGrid.TextMatrix(0, gGrossValue) = "Gross Value"
    
    MGrid.Col = gTaxAmount
    MGrid.CellAlignment = vbAlignRight
    MGrid.CellFontBold = True
    MGrid.CellFontName = "Cosmic San"
    MGrid.CellFontSize = 10
    MGrid.TextMatrix(0, gTaxAmount) = "Tax Amount"
    
    MGrid.Col = gTotalAmount
    MGrid.CellAlignment = vbAlignRight
    MGrid.CellFontBold = True
    MGrid.CellFontName = "Cosmic San"
    MGrid.CellFontSize = 10
    MGrid.TextMatrix(0, gTotalAmount) = "Total Amount"
    
    MGrid.RowHeightMin = 350
    
    'SHOW DATAS ON GRID
    
    Me.Caption = "Sales Form 8 Report - Detailed"
    
    Dim dQuantity, dGrossValue, dTaxAmount, dTotalAmount As Double

    'Set rs = db.OpenRecordset("Select BillDate,BillNo,Doctor,Tax,Transaction.Quantity,Transaction.SaleRate,ItemName From Transaction,TestRegister Where (BillType = 'S8' ) And (BillDate Between cDate('" & DTPFrom.Value & "') And cDate('" & DTPTo.Value & "')) And (TestRegister.Code=Transaction.ItemCode) Group By BillDate,BillNo,Doctor,Tax,Transaction.Quantity,Transaction.SaleRate,ItemName Order By BillDate,Val(BillNo)")
    If CoItem.ListIndex = -1 And CoDoctor.ListIndex = -1 Then
        If CoTest.ListIndex = -1 Then
            Set rs = db.OpenRecordset("Select BillDate,BillNo,Doctor,Tax,Transaction.Quantity,Transaction.SaleRate,ItemName From Transaction,TestRegister Where (BillType In ('S8','SB') ) And (BillDate Between cDate('" & DTPFrom.Value & "') And cDate('" & DTPTo.Value & "')) And (TestRegister.Code=Transaction.ItemCode) Group By BillDate,BillNo,Doctor,Tax,Transaction.Quantity,Transaction.SaleRate,ItemName Order By BillDate,Val(BillNo)")
        Else
            Set rs = db.OpenRecordset("Select BillDate,BillNo,Doctor,Tax,Transaction.Quantity,Transaction.SaleRate,ItemName From Transaction,TestRegister Where (BillType In ('S8','SB') ) And (BillDate Between cDate('" & DTPFrom.Value & "') And cDate('" & DTPTo.Value & "')) And (TestRegister.Code=Transaction.ItemCode) And (Transaction.ItemCode In (Select TestRegister.Code From TestRegister Where(TestRegister.TestCode='" & sTestCode(CoTest.ListIndex + 1) & "'))) Group By BillDate,BillNo,Doctor,Tax,Transaction.Quantity,Transaction.SaleRate,ItemName Order By BillDate,Val(BillNo)")
        End If
    ElseIf CoItem.ListIndex = -1 And CoDoctor.ListIndex > -1 Then
        If CoTest.ListIndex = -1 Then
            Set rs = db.OpenRecordset("Select BillDate,BillNo,Doctor,Tax,Transaction.Quantity,Transaction.SaleRate,ItemName From Transaction,TestRegister Where (BillType In ('S8','SB') ) And (BillDate Between cDate('" & DTPFrom.Value & "') And cDate('" & DTPTo.Value & "')) And (TestRegister.Code=Transaction.ItemCode) And (Transaction.DoctorCode ='" & sDoctorCode(CoDoctor.ListIndex + 1) & "') Group By BillDate,BillNo,Doctor,Tax,Transaction.Quantity,Transaction.SaleRate,ItemName Order By BillDate,Val(BillNo)")
        Else
            Set rs = db.OpenRecordset("Select BillDate,BillNo,Doctor,Tax,Transaction.Quantity,Transaction.SaleRate,ItemName From Transaction,TestRegister Where (BillType In ('S8','SB') ) And (BillDate Between cDate('" & DTPFrom.Value & "') And cDate('" & DTPTo.Value & "')) And (TestRegister.Code=Transaction.ItemCode) And (Transaction.ItemCode In (Select TestRegister.Code From TestRegister Where(TestRegister.TestCode='" & sTestCode(CoTest.ListIndex + 1) & "'))) And (Transaction.DoctorCode ='" & sDoctorCode(CoDoctor.ListIndex + 1) & "') Group By BillDate,BillNo,Doctor,Tax,Transaction.Quantity,Transaction.SaleRate,ItemName Order By BillDate,Val(BillNo)")
        End If
    ElseIf CoItem.ListIndex > -1 And CoDoctor.ListIndex = -1 Then
        Set rs = db.OpenRecordset("Select BillDate,BillNo,Doctor,Tax,Transaction.Quantity,Transaction.SaleRate,ItemName From Transaction,TestRegister Where (BillType In ('S8','SB') ) And (BillDate Between cDate('" & DTPFrom.Value & "') And cDate('" & DTPTo.Value & "')) And (TestRegister.Code=Transaction.ItemCode) And (Transaction.ItemCode ='" & sSubTestCode(CoItem.ListIndex + 1) & "') Group By BillDate,BillNo,Doctor,Tax,Transaction.Quantity,Transaction.SaleRate,ItemName Order By BillDate,Val(BillNo)")
    ElseIf CoItem.ListIndex > -1 And CoDoctor.ListIndex > -1 Then
        Set rs = db.OpenRecordset("Select BillDate,BillNo,Doctor,Tax,Transaction.Quantity,Transaction.SaleRate,ItemName From Transaction,TestRegister Where (BillType In ('S8','SB') ) And (BillDate Between cDate('" & DTPFrom.Value & "') And cDate('" & DTPTo.Value & "')) And (TestRegister.Code=Transaction.ItemCode) And (Transaction.ItemCode ='" & sSubTestCode(CoItem.ListIndex + 1) & "') And (Transaction.DoctorCode ='" & sDoctorCode(CoDoctor.ListIndex + 1) & "') Group By BillDate,BillNo,Doctor,Tax,Transaction.Quantity,Transaction.SaleRate,ItemName Order By BillDate,Val(BillNo)")
    End If
    
    While rs.EOF = False
        dQuantity = dQuantity + Abs(Val("" & rs!Quantity))
        dGrossValue = dGrossValue + Val("" & rs!SaleRate) * Abs(Val("" & rs!Quantity))
        dTaxAmount = dTaxAmount + (Val("" & rs!SaleRate) * Abs(Val("" & rs!Quantity))) * (Val("" & rs!Tax) / 100)
        dTotalAmount = dTotalAmount + Val("" & rs!SaleRate) * Abs(Val("" & rs!Quantity)) + (Val("" & rs!SaleRate) * Abs(Val("" & rs!Quantity))) * (Val("" & rs!Tax) / 100)
        
        MGrid.AddItem Format("" & rs!BillDate, "dd-MM-yyyy") & vbTab & "" & rs!BillNo & vbTab & "" & rs!Doctor & vbTab & rs!ItemName & vbTab & Format(Val("" & rs!Tax), "0.00") & vbTab & Format(Val("" & rs!SaleRate), "0.00") & vbTab & Format(Abs(Val("" & rs!Quantity)), "0.000") & vbTab & Format(Val("" & rs!SaleRate) * Abs(Val("" & rs!Quantity)), "0.00") & vbTab & Format((Val("" & rs!SaleRate) * Abs(Val("" & rs!Quantity))) * (Val("" & rs!Tax) / 100), "0.00") & vbTab & Format(((Val("" & rs!SaleRate) * Abs(Val("" & rs!Quantity))) * (Val("" & rs!Tax) / 100)) + (Val("" & rs!SaleRate) * Abs(Val("" & rs!Quantity))), "0.00")
        rs.MoveNext
    Wend
    rs.Close

    MGrid.AddItem vbTab & vbTab & vbTab & "Total" & vbTab & vbTab & vbTab & Format(Abs(dQuantity), "0.000") & vbTab & Format(dGrossValue, "0.00") & vbTab & Format(dTaxAmount, "0.00") & vbTab & Format(dTotalAmount, "0.00")
End Sub

Private Sub CShowSummary_Click()
Dim gDate As Single, gBillNo As Single, gDealer As Single, gNarration As Single, gTotalAmount As Single, gTaxAmount As Single, gBillAmount As Single, gExtraCharge As Single, gDiscount As Single, gAdvance As Single, gBalance As Single
Dim rs As Recordset
    
    'INITIALISING GRID
    gDate = 0
    gBillNo = 1
    gDealer = 2
    gNarration = 3
    gTotalAmount = 4
    gTaxAmount = 5
    gBillAmount = 6
    gExtraCharge = 7
    gDiscount = 8
    gAdvance = 9
    gBalance = 10
    
    MGrid.Clear
    MGrid.Rows = 1 'FOR SKIPING ERROR
    MGrid.Cols = 1 'FOR SKIPING ERROR
    MGrid.FixedCols = 0
    MGrid.FixedRows = 0
    MGrid.Cols = 11
    MGrid.Rows = 1
    MGrid.ColWidth(gDate) = 1300
    MGrid.ColWidth(gBillNo) = 1200
    MGrid.ColWidth(gDealer) = 2000
    MGrid.ColWidth(gNarration) = 2000
    MGrid.ColWidth(gTotalAmount) = 1300
    MGrid.ColWidth(gTaxAmount) = 1300
    MGrid.ColWidth(gBillAmount) = 1300
    MGrid.ColWidth(gExtraCharge) = 1300
    MGrid.ColWidth(gDiscount) = 1300
    MGrid.ColWidth(gAdvance) = 1300
    MGrid.ColWidth(gBalance) = 1300
    
    MGrid.Col = gDate
    MGrid.CellAlignment = vbAlignRight
    MGrid.CellFontBold = True
    MGrid.CellFontName = "Cosmic San"
    MGrid.CellFontSize = 10
    MGrid.TextMatrix(0, gDate) = "Date"
    
    MGrid.Col = gBillNo
    MGrid.CellAlignment = vbAlignRight
    MGrid.CellFontBold = True
    MGrid.CellFontName = "Cosmic San"
    MGrid.CellFontSize = 10
    MGrid.TextMatrix(0, gBillNo) = "Bill No"
    
    MGrid.Col = gDealer
    MGrid.CellAlignment = vbAlignRight
    MGrid.CellFontBold = True
    MGrid.CellFontName = "Cosmic San"
    MGrid.CellFontSize = 10
    MGrid.TextMatrix(0, gDealer) = "Doctor"
    
    MGrid.Col = gNarration
    MGrid.CellAlignment = vbAlignRight
    MGrid.CellFontBold = True
    MGrid.CellFontName = "Cosmic San"
    MGrid.CellFontSize = 10
    MGrid.TextMatrix(0, gNarration) = "Narration"
    
    MGrid.Col = gTotalAmount
    MGrid.CellAlignment = vbAlignRight
    MGrid.CellFontBold = True
    MGrid.CellFontName = "Cosmic San"
    MGrid.CellFontSize = 10
    MGrid.TextMatrix(0, gTotalAmount) = "Total Amount"
    
    MGrid.Col = gTaxAmount
    MGrid.CellAlignment = vbAlignRight
    MGrid.CellFontBold = True
    MGrid.CellFontName = "Cosmic San"
    MGrid.CellFontSize = 10
    MGrid.TextMatrix(0, gTaxAmount) = "Tax Amount"
    
    MGrid.Col = gBillAmount
    MGrid.CellAlignment = vbAlignRight
    MGrid.CellFontBold = True
    MGrid.CellFontName = "Cosmic San"
    MGrid.CellFontSize = 10
    MGrid.TextMatrix(0, gBillAmount) = "Bill Amount"
    
    MGrid.Col = gExtraCharge
    MGrid.CellAlignment = vbAlignRight
    MGrid.CellFontBold = True
    MGrid.CellFontName = "Cosmic San"
    MGrid.CellFontSize = 10
    MGrid.TextMatrix(0, gExtraCharge) = "Extra Charge"
    
    MGrid.Col = gDiscount
    MGrid.CellAlignment = vbAlignRight
    MGrid.CellFontBold = True
    MGrid.CellFontName = "Cosmic San"
    MGrid.CellFontSize = 10
    MGrid.TextMatrix(0, gDiscount) = "Discount"
    
    MGrid.Col = gAdvance
    MGrid.CellAlignment = vbAlignRight
    MGrid.CellFontBold = True
    MGrid.CellFontName = "Cosmic San"
    MGrid.CellFontSize = 10
    MGrid.TextMatrix(0, gAdvance) = "Advance"
    
    MGrid.Col = gBalance
    MGrid.CellAlignment = vbAlignRight
    MGrid.CellFontBold = True
    MGrid.CellFontName = "Cosmic San"
    MGrid.CellFontSize = 10
    MGrid.TextMatrix(0, gBalance) = "Balance"
    
    MGrid.RowHeightMin = 350

    
    'SHOW DATAS ON GRID
    
    Me.Caption = "Sales Report - Summary"
    
    Dim dAdvance, dDiscount, dExtraCharge, dTaxAmount, dBasicAmount, dBillAmount, dBalance As Double
    
    'Set rs = db.OpenRecordset("Select BillDate,BillNo,Doctor,Narration,Advance,Discount,ExtraCharges,Tax,Sum(Abs(Transaction.Quantity)*Transaction.SaleRate) As TotalAmount From Transaction Where (BillType = 'S8' ) And (BillDate Between cDate('" & DTPFrom.Value & "') And cDate('" & DTPTo.Value & "') ) Group By BillDate,BillNo,Doctor,Narration,Advance,Discount,ExtraCharges,Tax Order By BillDate,Val(BillNo)")
    If CoItem.ListIndex = -1 And CoDoctor.ListIndex = -1 Then
        If CoTest.ListIndex = -1 Then
            Set rs = db.OpenRecordset("Select BillDate,BillNo,Doctor,Narration,Advance,Discount,ExtraCharges,Sum((Abs(Transaction.Quantity)*Transaction.SaleRate)*(Transaction.Tax/100)) As TaxAmount,Sum(Abs(Transaction.Quantity)*Transaction.SaleRate) As TotalAmount From Transaction Where (BillType In ('S8','SB') ) And (BillDate Between cDate('" & DTPFrom.Value & "') And cDate('" & DTPTo.Value & "') ) Group By BillDate,BillNo,Doctor,Narration,Advance,Discount,ExtraCharges Order By BillDate,Val(BillNo)")
        Else
            Set rs = db.OpenRecordset("Select BillDate,BillNo,Doctor,Narration,Advance,Discount,ExtraCharges,Sum((Abs(Transaction.Quantity)*Transaction.SaleRate)*(Transaction.Tax/100)) As TaxAmount,Sum(Abs(Transaction.Quantity)*Transaction.SaleRate) As TotalAmount From Transaction Where (BillType In ('S8','SB') ) And (BillDate Between cDate('" & DTPFrom.Value & "') And cDate('" & DTPTo.Value & "') ) And (Transaction.ItemCode In (Select TestRegister.Code From TestRegister Where(TestRegister.TestCode='" & sTestCode(CoTest.ListIndex + 1) & "'))) Group By BillDate,BillNo,Doctor,Narration,Advance,Discount,ExtraCharges Order By BillDate,Val(BillNo)")
        End If
    ElseIf CoItem.ListIndex = -1 And CoDoctor.ListIndex > -1 Then
        
        If CoTest.ListIndex = -1 Then
            Set rs = db.OpenRecordset("Select BillDate,BillNo,Doctor,Narration,Advance,Discount,ExtraCharges,Sum((Abs(Transaction.Quantity)*Transaction.SaleRate)*(Transaction.Tax/100)) As TaxAmount,Sum(Abs(Transaction.Quantity)*Transaction.SaleRate) As TotalAmount From Transaction Where (BillType In ('S8','SB') ) And (BillDate Between cDate('" & DTPFrom.Value & "') And cDate('" & DTPTo.Value & "') ) And (Transaction.DoctorCode ='" & sDoctorCode(CoDoctor.ListIndex + 1) & "') Group By BillDate,BillNo,Doctor,Narration,Advance,Discount,ExtraCharges Order By BillDate,Val(BillNo)")
        Else
            Set rs = db.OpenRecordset("Select BillDate,BillNo,Doctor,Narration,Advance,Discount,ExtraCharges,Sum((Abs(Transaction.Quantity)*Transaction.SaleRate)*(Transaction.Tax/100)) As TaxAmount,Sum(Abs(Transaction.Quantity)*Transaction.SaleRate) As TotalAmount From Transaction Where (BillType In ('S8','SB') ) And (BillDate Between cDate('" & DTPFrom.Value & "') And cDate('" & DTPTo.Value & "') ) And (Transaction.ItemCode In (Select TestRegister.Code From TestRegister Where(TestRegister.TestCode='" & sTestCode(CoTest.ListIndex + 1) & "'))) And (Transaction.DoctorCode ='" & sDoctorCode(CoDoctor.ListIndex + 1) & "') Group By BillDate,BillNo,Doctor,Narration,Advance,Discount,ExtraCharges Order By BillDate,Val(BillNo)")
        End If
     
    ElseIf CoItem.ListIndex > -1 And CoDoctor.ListIndex = -1 Then
        Set rs = db.OpenRecordset("Select BillDate,BillNo,Doctor,Narration,Advance,Discount,ExtraCharges,Sum((Abs(Transaction.Quantity)*Transaction.SaleRate)*(Transaction.Tax/100)) As TaxAmount,Sum(Abs(Transaction.Quantity)*Transaction.SaleRate) As TotalAmount From Transaction Where (BillType In ('S8','SB') ) And (BillDate Between cDate('" & DTPFrom.Value & "') And cDate('" & DTPTo.Value & "') ) And (Transaction.ItemCode ='" & sSubTestCode(CoItem.ListIndex + 1) & "') Group By BillDate,BillNo,Doctor,Narration,Advance,Discount,ExtraCharges Order By BillDate,Val(BillNo)")
    ElseIf CoItem.ListIndex > -1 And CoDoctor.ListIndex > -1 Then
        Set rs = db.OpenRecordset("Select BillDate,BillNo,Doctor,Narration,Advance,Discount,ExtraCharges,Sum((Abs(Transaction.Quantity)*Transaction.SaleRate)*(Transaction.Tax/100)) As TaxAmount,Sum(Abs(Transaction.Quantity)*Transaction.SaleRate) As TotalAmount From Transaction Where (BillType In ('S8','SB') ) And (BillDate Between cDate('" & DTPFrom.Value & "') And cDate('" & DTPTo.Value & "') ) And (Transaction.ItemCode ='" & sSubTestCode(CoItem.ListIndex + 1) & "') And (Transaction.DoctorCode ='" & sDoctorCode(CoDoctor.ListIndex + 1) & "') Group By BillDate,BillNo,Doctor,Narration,Advance,Discount,ExtraCharges Order By BillDate,Val(BillNo)")
    End If
    
    While rs.EOF = False
        dAdvance = dAdvance + Val("" & rs!Advance)
        dDiscount = dDiscount + Val("" & rs!Discount)
        dExtraCharge = dExtraCharge + Val("" & rs!ExtraCharges)
        dTaxAmount = dTaxAmount + Val("" & rs!TaxAmount)
        dBasicAmount = dBasicAmount + Val("" & rs!totalAmount)
        dBillAmount = dBillAmount + (Val("" & rs!TaxAmount)) + Val("" & rs!totalAmount)
        dBalance = dBalance + (Val("" & rs!totalAmount) + (Val("" & rs!TaxAmount)) + Val("" & rs!ExtraCharges) - (Val("" & rs!Discount) + Val("" & rs!Advance)))
        
        MGrid.AddItem Format("" & rs!BillDate, "dd-MM-yyyy") & vbTab & "" & rs!BillNo & vbTab & "" & rs!Doctor & vbTab & rs!Narration & vbTab & Format(Val("" & rs!totalAmount), "0.00") & vbTab & Format(Val("" & rs!TaxAmount), "0.00") & vbTab & Format(Val("" & rs!TaxAmount) + Val("" & rs!totalAmount), "0.00") & vbTab & Format(Val("" & rs!ExtraCharges), "0.00") & vbTab & Format(Val("" & rs!Discount), "0.00") & vbTab & Format(Val("" & rs!Advance), "0.00") & vbTab & Format((Val("" & rs!totalAmount) + (Val("" & rs!TaxAmount)) + Val("" & rs!ExtraCharges) - (Val("" & rs!Discount) + Val("" & rs!Advance))), "0.00")
        rs.MoveNext
    Wend
    rs.Close
    
    MGrid.AddItem "" & vbTab & "" & vbTab & "" & vbTab & "Total" & vbTab & Format(dBasicAmount, "0.00") & vbTab & Format(dTaxAmount, "0.00") & vbTab & Format(dBillAmount, "0.00") & vbTab & Format(dExtraCharge, "0.00") & vbTab & Format(dDiscount, "0.00") & vbTab & Format(dAdvance, "0.00") & vbTab & Format(dBalance, "0.00")
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
    
    oExcelSheet.Range("A3:J" & lRowCount + 4).Value = xData

    oExcelSheet.Cells(1, 1).Value = Me.Caption & " From " & Format(DTPFrom.Value, "dd-MM-yyyy") & " To " & Format(DTPTo.Value, "dd-MM-yyyy")

    oExcelSheet.Range("A1:J" & lRowCount + 4).Select
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

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyS And ((Shift And 7) = 2)) Then
        CShowSummary_Click
    ElseIf (KeyCode = vbKeyD And ((Shift And 7) = 2)) Then
        CShowDetailed_Click
    ElseIf (KeyCode = vbKeyX And ((Shift And 7) = 2)) Then
        CToExcel_Click
    ElseIf (KeyCode = vbKeyC And ((Shift And 7) = 2)) Then
        CClose_Click
    End If
End Sub

Private Sub Form_Load()
    DTPFrom.Value = Date
    DTPTo.Value = Date
    
    getTest
    getDoctor

End Sub

Private Sub MGrid_DblClick()
    If MGrid.Rows > 0 Then
        FMedicalTest.Show
        FMedicalTest.TTransactionNo = MGrid.TextMatrix(MGrid.Row, 1)
        FMedicalTest.getTransactionDetails
    End If
End Sub
