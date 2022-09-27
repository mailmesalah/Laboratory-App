VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FDoctorRegister 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Register"
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9930
   ControlBox      =   0   'False
   Icon            =   "FCustomerRegister.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FCustomerRegister.frx":000C
   ScaleHeight     =   7395
   ScaleWidth      =   9930
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CClose 
      Cancel          =   -1  'True
      Height          =   505
      Left            =   7950
      Picture         =   "FCustomerRegister.frx":1FEC4E
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6780
      Width           =   1365
   End
   Begin VB.CommandButton CSave 
      Height          =   505
      Left            =   6495
      Picture         =   "FCustomerRegister.frx":2010B0
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6780
      Width           =   1365
   End
   Begin VB.CommandButton CDeleteCustomer 
      Height          =   505
      Left            =   2055
      Picture         =   "FCustomerRegister.frx":203512
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6780
      Width           =   1365
   End
   Begin VB.CommandButton CAddNew 
      Height          =   505
      Left            =   630
      Picture         =   "FCustomerRegister.frx":205974
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6780
      Width           =   1365
   End
   Begin MSComctlLib.TreeView TrDoctors 
      Height          =   4800
      Left            =   270
      TabIndex        =   0
      Top             =   210
      Width           =   4035
      _ExtentX        =   7117
      _ExtentY        =   8467
      _Version        =   393217
      LabelEdit       =   1
      Style           =   7
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSForms.Label Label2 
      Height          =   345
      Left            =   4950
      TabIndex        =   4
      Top             =   1005
      Width           =   1170
      VariousPropertyBits=   8388627
      Caption         =   "Name"
      Size            =   "2064;609"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TDoctorName 
      Height          =   345
      Left            =   6150
      TabIndex        =   3
      Top             =   975
      Width           =   3360
      VariousPropertyBits=   746604571
      MaxLength       =   100
      BorderStyle     =   1
      Size            =   "5927;609"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label1 
      Height          =   345
      Left            =   4950
      TabIndex        =   2
      Top             =   630
      Width           =   1155
      VariousPropertyBits=   8388627
      Caption         =   "Code"
      Size            =   "2037;609"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TDoctorCode 
      Height          =   345
      Left            =   6150
      TabIndex        =   1
      Top             =   585
      Width           =   3360
      VariousPropertyBits=   746604575
      BorderStyle     =   1
      Size            =   "5927;609"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
End
Attribute VB_Name = "FDoctorRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CAddNew_Click()
    TDoctorCode.Text = getNewDoctorCode
    TDoctorName.SetFocus
End Sub
Public Function getNewDoctorCode() As String
Dim rs As Recordset, sCode As String
    
    Set rs = db.OpenRecordset("Select Max(Val(DoctorCode)) As CCode From DoctorMaster")
    If rs.RecordCount > 0 Then
        sCode = Val("" & rs!CCode) + 1
    Else
        sCode = "1"
    
    End If
    rs.Close
    
    getNewDoctorCode = sCode
End Function
Private Sub CClose_Click()
    Unload Me
End Sub

Private Sub clearTexts()
    TDoctorCode.Text = ""
    TDoctorName.Text = ""
        
End Sub

Private Sub CDeleteDoctor_Click()
Dim rs As Recordset

    If checkAlreadyUsed(Trim(TDoctorCode.Text)) Then
        MsgBox "The Doctor is Already Used , Please Remove it First.", vbInformation
        Exit Sub
    End If

    Set rs = db.OpenRecordset("Select * From DoctorMaster Where DoctorCode='" & Trim(TDoctorCode.Text) & "'")
    If rs.RecordCount > 0 Then
        rs.Delete
        MsgBox "Successfully Deleted !", vbInformation
    Else
        MsgBox "Doctor not Found !", vbInformation
    End If
    rs.Close
    
    clearTexts
    refreshTree
End Sub

Private Sub CSave_Click()
Dim rs As Recordset, sStatus As String, sAccountCode As String

    If (Trim(TDoctorCode.Text) = "" Or Trim(TDoctorName.Text) = "") Then
        MsgBox "Enter a Doctor !", vbInformation
        Exit Sub
    End If
        
    Set rs = db.OpenRecordset("Select * From DoctorMaster Where DoctorCode='" & Trim(TDoctorCode.Text) & "'")
    
    If rs.RecordCount > 0 Then
        rs.Edit
        sStatus = "Edited"
    Else
        sStatus = "Added"
        TDoctorCode.Text = getNewDoctorCode
        rs.AddNew
        rs!DoctorCode = Trim(TDoctorCode.Text)
        rs!AddedDate = Date
    End If
    
    'VALIDATE INPUT DATAS
    rs!DoctorName = Trim(TDoctorName.Text)
    rs!EditedDate = Date
    rs.Update
    rs.Close
        
    MsgBox "Successfully " & sStatus & " !", vbInformation
    
    refreshTree
    clearTexts
    
    CAddNew.SetFocus
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyN And ((Shift And 7) = 2)) Then
        CAddNew_Click
    ElseIf (KeyCode = vbKeyD And ((Shift And 7) = 2)) Then
        CDeleteDoctor_Click
    ElseIf (KeyCode = vbKeyS And ((Shift And 7) = 2)) Then
        CSave_Click
    ElseIf (KeyCode = vbKeyL And ((Shift And 7) = 2)) Then
        CClose_Click
    End If
End Sub

Private Sub Form_Load()

    CoStatus.AddItem "Enabled"
    CoStatus.AddItem "Disabled"
    
    refreshTree
End Sub

Private Sub refreshTree()
Dim rs As Recordset
    
    TrDoctors.Nodes.Clear
    
    Set rs = db.OpenRecordset("Select DoctorMaster.DoctorCode,DoctorMaster.DoctorName From DoctorMaster Order By DoctorMaster.DoctorName")
    
    While rs.EOF = False
        TrDoctors.Nodes.Add , , "C" & rs!DoctorCode, rs!DoctorName
        rs.MoveNext
    Wend
    rs.Close
End Sub

Private Sub TDoctorName_GotFocus()
    TDoctorName.SelStart = 0
    TDoctorName.SelLength = Len(TDoctorName.Text)
End Sub

Private Sub TrDoctors_Click()
Dim rs As Recordset
    If TrDoctors.Nodes.Count > 0 Then
        TDoctorCode.Text = Right(TrDoctors.SelectedItem.Key, Len(TrDoctors.SelectedItem.Key) - 1)
        TDoctorName.Text = TrDoctors.SelectedItem.Text
    End If
End Sub

Private Function checkAlreadyUsed(sCCode As String) As Boolean
Dim rs As Recordset
Dim bExist As Boolean
    bExist = False
    Set rs = db.OpenRecordset("Select Transaction.* From Transaction Where (Transaction.DoctorCode = '" & sCCode & "' )")
    If rs.RecordCount > 0 Then
        bExist = True
    End If
    rs.Close
    checkAlreadyUsed = bExist
End Function
