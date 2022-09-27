VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FManufacturer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Department"
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9795
   ControlBox      =   0   'False
   Icon            =   "FManufacturer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FManufacturer.frx":000C
   ScaleHeight     =   7395
   ScaleWidth      =   9795
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CClose 
      Cancel          =   -1  'True
      Height          =   505
      Left            =   7950
      Picture         =   "FManufacturer.frx":1FEC4E
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6435
      Width           =   1365
   End
   Begin VB.CommandButton CSave 
      Height          =   505
      Left            =   6510
      Picture         =   "FManufacturer.frx":2010B0
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6420
      Width           =   1365
   End
   Begin VB.CommandButton CDelete 
      Height          =   505
      Left            =   1950
      Picture         =   "FManufacturer.frx":203512
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6435
      Width           =   1365
   End
   Begin VB.CommandButton CAddNew 
      Height          =   505
      Left            =   510
      Picture         =   "FManufacturer.frx":205974
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6435
      Width           =   1365
   End
   Begin MSComctlLib.TreeView TrDepartment 
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
   Begin MSForms.TextBox TDepartmentCode 
      Height          =   345
      Left            =   6045
      TabIndex        =   1
      Top             =   615
      Width           =   3360
      VariousPropertyBits=   746604571
      MaxLength       =   10
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
      Height          =   405
      Left            =   4725
      TabIndex        =   8
      Top             =   645
      Width           =   1530
      VariousPropertyBits=   8388627
      Caption         =   "Code"
      Size            =   "2699;714"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TDepartmentName 
      Height          =   345
      Left            =   6045
      TabIndex        =   2
      Top             =   1050
      Width           =   3360
      VariousPropertyBits=   746604571
      MaxLength       =   30
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
   Begin MSForms.Label Label2 
      Height          =   405
      Left            =   4725
      TabIndex        =   7
      Top             =   1065
      Width           =   1530
      VariousPropertyBits=   8388627
      Caption         =   "Description"
      Size            =   "2699;714"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "FManufacturer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CAddNew_Click()
Dim rs As Recordset
    
    Set rs = db.OpenRecordset("Select Max(Val(Code)) As GCode From Department")
    If rs.RecordCount > 0 Then
        TDepartmentCode.Text = Val("" & rs!gCode) + 1
    Else
        TDepartmentCode.Text = 1
    End If
    rs.Close

    TDepartmentName.SetFocus
End Sub

Private Function getNewCodeForDepartment() As String
Dim rs As Recordset, sNewCode As String
    
    Set rs = db.OpenRecordset("Select Max(Val(Code)) As GCode From Department")
    If rs.RecordCount > 0 Then
        sNewCode = Val("" & rs!gCode) + 1
    Else
        sNewCode = 1
    End If
    rs.Close
    
    getNewCodeForDepartment = sNewCode
End Function

Private Sub CClose_Click()
    Unload Me
End Sub

Private Sub CDelete_Click()
Dim rs As Recordset
    
    If checkMediaAlreadyUsed(Trim(TDepartmentCode.Text)) Then
        MsgBox "The Department is Already Used , Please Remove it First.", vbInformation
        Exit Sub
    End If
    

    Set rs = db.OpenRecordset("Select * From Department Where Code='" & Trim(TDepartmentCode.Text) & "'")
    If rs.RecordCount > 0 Then
        rs.Delete
        MsgBox "Successfully Deleted !", vbInformation
    Else
        MsgBox "Item not Found !", vbInformation
    End If
    rs.Close
    
    clearTexts
    refreshTree
End Sub
Private Sub clearTexts()
    TDepartmentCode.Text = ""
    TDepartmentName.Text = ""
    TShortName.Text = ""
End Sub
Private Sub CSave_Click()
Dim rs As Recordset, sStatus As String
    If (Trim(TDepartmentCode.Text) = "" Or Trim(TDepartmentName.Text) = "") Then
        MsgBox "Enter any Item !", vbInformation
        TDepartmentName.SetFocus
        Exit Sub
    End If
    
    If (Trim(TShortName.Text) = "") Then
        MsgBox "Enter Short Name !", vbInformation
        TShortName.SetFocus
        Exit Sub
    End If
    
    Set rs = db.OpenRecordset("Select * From Department Where Code='" & Trim(TDepartmentCode.Text) & "'")
    
    If rs.RecordCount > 0 Then
        rs.Edit
        sStatus = "Edited"
    Else
        rs.AddNew
        sStatus = "Added"
    End If
    
    rs!Code = Trim(TDepartmentCode.Text)
    rs!DepartmentName = Trim(TDepartmentName.Text)
    rs!ShortName = Trim(TShortName.Text)
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
        CDelete_Click
    ElseIf (KeyCode = vbKeyS And ((Shift And 7) = 2)) Then
        CSave_Click
    ElseIf (KeyCode = vbKeyC And ((Shift And 7) = 2)) Then
        CClose_Click
    End If
End Sub

Private Sub Form_Load()
    TDepartmentCode.Text = getNewCodeForDepartment
    refreshTree
End Sub

Private Sub refreshTree()
Dim rs As Recordset
    
    TrDepartment.Nodes.Clear
    
    Set rs = db.OpenRecordset("Select Department.Code,Department.DepartmentName From Department Order By Department.DepartmentName")
    
    While rs.EOF = False
        TrDepartment.Nodes.Add , , "C" & rs!Code, UCase(rs!DepartmentName)
        rs.MoveNext
    Wend
    rs.Close
End Sub

Private Sub TDepartmentName_GotFocus()
    TDepartmentName.SelStart = 0
    TDepartmentName.SelLength = Len(TDepartmentName.Text)
End Sub

Private Sub TrDepartment_Click()
Dim rs As Recordset
    If TrDepartment.Nodes.Count > 0 Then
        TDepartmentCode.Text = Right(TrDepartment.SelectedItem.Key, Len(TrDepartment.SelectedItem.Key) - 1)
        TDepartmentName.Text = TrDepartment.SelectedItem.Text
        Set rs = db.OpenRecordset("Select Department.ShortName From Department Where (Department.Code = '" & Trim(TDepartmentCode.Text) & "' )")
        If rs.RecordCount > 0 Then
            TShortName.Text = "" & rs!ShortName
        End If
        rs.Close
    End If
End Sub

Private Function checkMediaAlreadyUsed(sMCode As String) As Boolean
Dim rs As Recordset
Dim bExist As Boolean
    bExist = False
    Set rs = db.OpenRecordset("Select ItemMaster.* From ItemMaster Where (ItemMaster.DepartmentCode = '" & sMCode & "' )")
    If rs.RecordCount > 0 Then
        bExist = True
    End If
    rs.Close
    
    checkMediaAlreadyUsed = bExist
End Function
