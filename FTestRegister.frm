VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FTestRegister 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test Register"
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9870
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial Narrow"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FTestRegister.frx":0000
   ScaleHeight     =   8520
   ScaleWidth      =   9870
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CAddGroup 
      BackColor       =   &H8000000D&
      Caption         =   "Add Test"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   505
      Left            =   525
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7290
      Width           =   1365
   End
   Begin VB.CommandButton CAddNew 
      BackColor       =   &H8000000D&
      Caption         =   "Add Sub Test"
      Height          =   505
      Left            =   525
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7905
      Width           =   1365
   End
   Begin VB.CommandButton CClose 
      Cancel          =   -1  'True
      Height          =   505
      Left            =   7995
      Picture         =   "FTestRegister.frx":1FEC42
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7905
      Width           =   1365
   End
   Begin VB.CommandButton CSave 
      Height          =   505
      Left            =   6540
      Picture         =   "FTestRegister.frx":2010A4
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7905
      Width           =   1365
   End
   Begin VB.CommandButton CDeleteTest 
      Height          =   505
      Left            =   1995
      Picture         =   "FTestRegister.frx":203506
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7905
      Width           =   1365
   End
   Begin VB.CommandButton CFindNext 
      CausesValidation=   0   'False
      Height          =   505
      Left            =   2805
      Picture         =   "FTestRegister.frx":205968
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5775
      Width           =   1365
   End
   Begin MSComctlLib.TreeView TrItems 
      Height          =   5400
      Left            =   240
      TabIndex        =   0
      Top             =   210
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   9525
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
   Begin MSComDlg.CommonDialog CoDialog 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSForms.ComboBox CoDepartment 
      Height          =   375
      Left            =   6405
      TabIndex        =   3
      Top             =   1425
      Width           =   3180
      VariousPropertyBits=   746604571
      MaxLength       =   50
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "5609;661"
      MatchEntry      =   1
      ShowDropButtonWhen=   1
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
      Index           =   1
      Left            =   4560
      TabIndex        =   19
      Top             =   1410
      Width           =   1320
      VariousPropertyBits=   8388627
      Caption         =   "Department"
      Size            =   "2328;714"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TDefaultValue 
      Height          =   390
      Left            =   6405
      TabIndex        =   4
      Top             =   2070
      Width           =   1590
      VariousPropertyBits=   746604571
      BackColor       =   16777215
      BorderStyle     =   1
      Size            =   "2805;688"
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox TCost 
      Height          =   390
      Left            =   6405
      TabIndex        =   6
      Top             =   2925
      Width           =   1590
      VariousPropertyBits=   746604571
      BackColor       =   16777215
      BorderStyle     =   1
      Size            =   "2805;688"
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label5 
      Height          =   405
      Left            =   4560
      TabIndex        =   18
      Top             =   2925
      Width           =   1320
      VariousPropertyBits=   8388627
      Caption         =   "Standard Rate"
      Size            =   "2328;714"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox CoUnit 
      Height          =   375
      Left            =   6405
      TabIndex        =   5
      Top             =   2505
      Width           =   1590
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "2805;661"
      MatchEntry      =   1
      ShowDropButtonWhen=   1
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label4 
      Height          =   405
      Left            =   4560
      TabIndex        =   17
      Top             =   2505
      Width           =   1320
      VariousPropertyBits=   8388627
      Caption         =   "Unit"
      Size            =   "2328;714"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label3 
      Height          =   405
      Left            =   4560
      TabIndex        =   16
      Top             =   2100
      Width           =   1320
      VariousPropertyBits=   8388627
      Caption         =   "Default Value"
      Size            =   "2328;714"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label1 
      Height          =   405
      Left            =   4560
      TabIndex        =   15
      Top             =   300
      Width           =   1320
      VariousPropertyBits=   8388627
      Caption         =   "Code"
      Size            =   "2328;714"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label2 
      Height          =   405
      Index           =   0
      Left            =   4560
      TabIndex        =   14
      Top             =   735
      Width           =   1320
      VariousPropertyBits=   8388627
      Caption         =   "Test Name"
      Size            =   "2328;714"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox CoTest 
      Height          =   375
      Left            =   6405
      TabIndex        =   2
      Top             =   750
      Width           =   3180
      VariousPropertyBits=   746604571
      MaxLength       =   50
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "5609;661"
      MatchEntry      =   1
      ShowDropButtonWhen=   1
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.TextBox TItemCode 
      Height          =   375
      Left            =   6405
      TabIndex        =   1
      Top             =   330
      Width           =   3180
      VariousPropertyBits=   746604575
      BorderStyle     =   1
      Size            =   "5609;661"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.TextBox TFind 
      Height          =   315
      Left            =   240
      TabIndex        =   10
      Top             =   5865
      Width           =   2520
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "4445;556"
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
End
Attribute VB_Name = "FTestRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sItemCode() As String, sUnitCode() As String, sDepartmentCode() As String
Dim bCreateNewGroup As Boolean

Private Sub getUnits()
Dim rs As Recordset
    
    
    CoUnit.Clear
    
    Set rs = db.OpenRecordset("Select Units.Code,Units.UnitName From Units Order By Units.UnitName")
    ReDim sUnitCode(rs.RecordCount + 1) As String
    While rs.EOF = False
        CoUnit.AddItem ("" & rs!UnitName)
        sUnitCode(CoUnit.ListCount) = "" & rs!Code
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
        CoDepartment.AddItem ("" & rs!DepartmentName)
        sDepartmentCode(CoDepartment.ListCount) = "" & rs!DepartmentCode
        rs.MoveNext
    Wend
    rs.Close
End Sub

Private Sub getItems()
Dim rs As Recordset
    
    CoTest.Clear
    
    Set rs = db.OpenRecordset("Select TestRegister.TestName From TestRegister  Order By TestRegister.TestName")
    Do While rs.EOF = False
        rs.MoveNext
    Loop
    If rs.RecordCount > 0 Then
        rs.MoveFirst
    End If
    
    While rs.EOF = False
        CoTest.AddItem ("" & rs!TestName)
        rs.MoveNext
    Wend
    rs.Close
End Sub

Private Sub CAddGroup_Click()
    clearControls
    TItemCode = getNewItemCode
    CoTest.SetFocus
    bCreateNewGroup = True
    enableDisableControlsOnAdd
End Sub

Private Sub CAddNew_Click()
    If (TrItems.Nodes.count = 0) Then
        MsgBox "Please Create a Group First !", vbInformation
        Exit Sub
    End If
    
    If (TrItems.SelectedItem Is Nothing) Then
        MsgBox "Please Create a Group First !", vbInformation
        Exit Sub
    End If
    
    If Left(Trim(TrItems.SelectedItem.Key), 1) = "B" Then
        MsgBox "Please Select any Group to create Account !", vbInformation
        Exit Sub
    End If
    clearControls
    bCreateNewGroup = False
    TItemCode = getNewItemCode
    getDepartmentOf
    enableDisableControlsOnAdd
    CoTest.SetFocus
End Sub

Private Sub CClose_Click()
    Unload Me
End Sub

Private Sub CDeleteTest_Click()
Dim rs As Recordset
    
    If Trim(TItemCode.Text) = "" Then
        MsgBox "Please Select Any Test to Delete !", vbInformation
        Exit Sub
    End If
        
    If checkAlreadyUsed(Trim(TItemCode.Text)) Then
        MsgBox "The Test is Already Used !", vbInformation
        Exit Sub
    End If
    
    If checkIfParentNode(Trim(TItemCode.Text)) Then
        MsgBox "The Test has Sub Test in It, Delete them first  !", vbInformation
        Exit Sub
    End If
    
    Set rs = db.OpenRecordset("Select TestRegister.* From TestRegister Where (TestRegister.Code = '" & Trim(TItemCode.Text) & "' )")
    If rs.RecordCount > 0 Then
        rs.Delete
        rs.Close
    Else
        rs.Close
        MsgBox "The Test doesn't Exist !", vbInformation
        Exit Sub
    End If
    
    MsgBox "Successfully Deleted the Test !", vbInformation
    
    refreshTree
    clearControls
End Sub

Private Sub enableDisableControlsOnAdd()
    
    If bCreateNewGroup = True Then
    
        CoTest.Enabled = True
        CoDepartment.Enabled = True
        CoUnit.Enabled = False
        TDefaultValue.Enabled = False
        TCost.Enabled = False
    Else
    
        CoTest.Enabled = True
        CoDepartment.Enabled = False
        CoUnit.Enabled = True
        TDefaultValue.Enabled = True
        TCost.Enabled = True
    End If
End Sub

Private Sub CFindNext_Click()
Static lFindIndex As Long
Static sFindWord As String
    
    If Trim(TFind.Text) <> sFindWord Then
        lFindIndex = 1
    Else
        lFindIndex = lFindIndex + 1
    End If
    
    sFindWord = Trim(TFind.Text)
    
    Do While lFindIndex <= TrItems.Nodes.count
        
        If InStr(1, LCase(TrItems.Nodes.Item(lFindIndex)), LCase(sFindWord), vbTextCompare) > 0 Then
            TrItems.Nodes.Item(lFindIndex).Selected = True
            getDetailsOfItem
            TrItems.SetFocus
            Exit Do
        End If
        lFindIndex = lFindIndex + 1
    Loop
    
    If lFindIndex > TrItems.Nodes.count Then
        MsgBox "No more Items !", vbInformation
        lFindIndex = 1
        Exit Sub
    End If
End Sub

Private Sub CoDepartment_GotFocus()
    CoDepartment.SelStart = 0
    CoDepartment.SelLength = Len(CoDepartment.Text)
End Sub

Private Sub CoDepartment_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 113 Then
        FDepartment.Show vbModal
        getDepartment
    End If
End Sub

Private Sub CoUnit_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 113 Then
        FUnits.Show vbModal
        getUnits
    End If
End Sub

Private Sub CSave_Click()
Dim rs As Recordset, sStatus As String, sBarCode As String, sParenttype As String, sParentCode As String
Dim r As Long
    If Trim(TItemCode.Text) = "" Then
        MsgBox "Please Select a Test to Edit or click Add New button To add new Test", vbInformation
        Exit Sub
    ElseIf CoTest.Text = "" Then
        MsgBox "Please Enter a Test !", vbInformation
        CoTest.SetFocus
        Exit Sub
    ElseIf CoDepartment.ListIndex = -1 Then
        MsgBox "Please Select a Department !", vbInformation
        CoDepartment.SetFocus
        Exit Sub
    End If
    
    'Determines GroupCode
    If TrItems.Nodes.count > 0 Then
        If (bCreateNewGroup) Then
            sParenttype = ""
            sParentCode = ""
        Else
            If CoUnit.ListIndex = -1 Then
                MsgBox "Please Select a Unit !", vbInformation
                CoUnit.SetFocus
                Exit Sub
            End If
            sParenttype = Trim(Left(TrItems.SelectedItem.Key, 1))
            sParentCode = Trim(Right(TrItems.SelectedItem.Key, Len(TrItems.SelectedItem.Key) - 1))
        End If
    Else
        sParenttype = ""
        sParentCode = ""
    End If
    
    Set rs = db.OpenRecordset("Select TestRegister.* From TestRegister Where (TestRegister.Code = '" & Trim(TItemCode.Text) & "' )")
    If rs.RecordCount > 0 Then
        sStatus = "Edited"
        rs.Edit
        rs!EditedDate = Date
    Else
        sStatus = "Added"
        TItemCode.Text = getNewItemCode()
        rs.AddNew
        rs!Code = Trim(TItemCode.Text)
        rs!Type = IIf(Trim(sParenttype) = "", "AGroup", "BItem")
        rs!GroupCode = sParentCode
        rs!AddedDate = Date
        rs!EditedDate = Date
    End If
    rs!TestName = CoTest.Text
    rs!DepartmentCode = sDepartmentCode(CoDepartment.ListIndex + 1)
    rs!UnitCode = sUnitCode(CoUnit.ListIndex + 1)
    rs!DefaultValue = TDefaultValue.Text
    rs!Cost = Val(TCost.Text)
    rs.Update
    rs.Close
    
    
    MsgBox "Successfully " & sStatus & " !", vbInformation
    
    refreshTree
    getItems
    clearControls
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyF And ((Shift And 7) = 2)) Then
        CFindNext_Click
    ElseIf (KeyCode = vbKeyN And ((Shift And 7) = 2)) Then
        CAddNew_Click
    ElseIf (KeyCode = vbKeyD And ((Shift And 7) = 2)) Then
        CDeleteTest_Click
    ElseIf (KeyCode = vbKeyS And ((Shift And 7) = 2)) Then
        CSave_Click
    ElseIf (KeyCode = vbKeyL And ((Shift And 7) = 2)) Then
        CClose_Click
    End If
End Sub

Private Sub Form_Load()
    
    clearControls
    refreshTree
    getItems
    getUnits
    getDepartment
End Sub

Private Sub refreshTree()
Dim rs As Recordset
    TrItems.Nodes.Clear
    
    Set rs = db.OpenRecordset("Select TestRegister.Code,TestRegister.TestName,TestRegister.Type,TestRegister.GroupCode From TestRegister Order By TestRegister.Type,TestRegister.TestName")
    While rs.EOF = False
        If Trim(rs!Type) = "AGroup" Then
            TrItems.Nodes.Add , , "A" & rs!Code, UCase(rs!TestName)
        ElseIf Trim(rs!Type) = "BItem" Then
             TrItems.Nodes.Add "A" & rs!GroupCode, tvwChild, "B" & rs!Code, UCase(rs!TestName)
        End If
        rs.MoveNext
    Wend
    rs.Close
End Sub

Private Sub TItemCode_GotFocus()
    TItemCode.SelStart = 0
    TItemCode.SelLength = Len(TItemCode.Text)
End Sub

Private Sub CoTest_GotFocus()
    CoTest.SelStart = 0
    CoTest.SelLength = Len(CoTest.Text)
End Sub

Private Sub TFind_GotFocus()
    TFind.SelStart = 0
    TFind.SelLength = Len(TFind.Text)
End Sub

Private Sub enableDisableControls()
    
    If TrItems.SelectedItem Is Nothing Then
        Exit Sub
    End If
    
    If Left(TrItems.SelectedItem.Key, 1) = "A" Then
    
        CoTest.Enabled = True
        CoDepartment.Enabled = True
        CoUnit.Enabled = False
        TDefaultValue.Enabled = False
        TCost.Enabled = False
        bCreateNewGroup = True
        
    ElseIf Left(TrItems.SelectedItem.Key, 1) = "B" Then
    
        CoTest.Enabled = True
        CoDepartment.Enabled = False
        CoUnit.Enabled = True
        TDefaultValue.Enabled = True
        TCost.Enabled = True
        bCreateNewGroup = False
    End If
End Sub

Private Sub TrItems_NodeClick(ByVal Node As MSComctlLib.Node)
    If TrItems.Nodes.count > 0 Then
        enableDisableControls
        getDetailsOfItem
    End If
End Sub

Private Sub getDetailsOfItem()
Dim rs As Recordset, r As Long, sCategory As String

    If (Left(TrItems.SelectedItem.Key, 1) = "A") Then
        Set rs = db.OpenRecordset("Select '' As UnitName,Department.DepartmentName,TestRegister.DefaultValue,TestRegister.Code,TestRegister.Type,TestRegister.TestName,TestRegister.Cost From TestRegister,Department Where (Department.DepartmentCode=TestRegister.DepartmentCode) And (TestRegister.Code = '" & Trim(Right(TrItems.SelectedItem.Key, Len(TrItems.SelectedItem.Key) - 1)) & "' )")
    ElseIf (Left(TrItems.SelectedItem.Key, 1) = "B") Then
        Set rs = db.OpenRecordset("Select (Select Units.UnitName From Units Where (Units.Code = TestRegister.UnitCode))As UnitName,Department.DepartmentName,TestRegister.DefaultValue,TestRegister.Code,TestRegister.TestName,TestRegister.Cost From TestRegister,Department Where (Department.DepartmentCode=TestRegister.DepartmentCode) And (TestRegister.Code = '" & Trim(Right(TrItems.SelectedItem.Key, Len(TrItems.SelectedItem.Key) - 1)) & "' )")
    Else
        Exit Sub
    End If
    
    If rs.RecordCount > 0 Then
        TItemCode.Text = "" & rs!Code
        CoTest.Text = "" & rs!TestName
        CoDepartment.Text = "" & rs!DepartmentName
        CoUnit.Text = "" & rs!UnitName
        TDefaultValue.Text = "" & rs!DefaultValue
        TCost.Text = Val("" & rs!Cost)
        rs.Close
        
    Else
        clearControls
    End If
    
End Sub

Private Sub getDepartmentOf()
Dim rs As Recordset, r As Long, sCategory As String

    If (Left(TrItems.SelectedItem.Key, 1) = "A") Then
        Set rs = db.OpenRecordset("Select Department.DepartmentName From TestRegister,Department Where (Department.DepartmentCode=TestRegister.DepartmentCode) And (TestRegister.Code = '" & Trim(Right(TrItems.SelectedItem.Key, Len(TrItems.SelectedItem.Key) - 1)) & "' )")
    Else
        Exit Sub
    End If
    
    If rs.RecordCount > 0 Then
        CoDepartment.Text = "" & rs!DepartmentName
        rs.Close
        
    Else
     
    End If
    
End Sub

Private Sub clearControls()
    TItemCode.Text = ""
    CoTest.Text = ""
    CoDepartment.Text = ""
    CoUnit.Text = ""
    TDefaultValue.Text = ""
    TCost.Text = ""
End Sub

Private Function getNewItemCode() As String
Dim rs As Recordset, sItemCode As String
    
    Set rs = db.OpenRecordset("Select Max(Val(TestRegister.Code))As ACode From TestRegister")
    If rs.RecordCount > 0 Then
        sItemCode = Val("" & rs!ACode) + 1
    Else
        sItemCode = "1"
    
    End If
    rs.Close
    
    getNewItemCode = sItemCode
End Function

Private Function checkAlreadyUsed(sMCode As String) As Boolean
Dim rs As Recordset
Dim bExist As Boolean
    bExist = False
    Set rs = db.OpenRecordset("Select Transaction.* From Transaction Where (Transaction.TestCode = '" & sMCode & "' ) Or (Transaction.SubTestCode = '" & sMCode & "' )")
    If rs.RecordCount > 0 Then
        bExist = True
    End If
    rs.Close
    
    checkAlreadyUsed = bExist
End Function

Private Function checkIfParentNode(sMCode As String) As Boolean
Dim rs As Recordset
Dim bExist As Boolean
    bExist = False
    Set rs = db.OpenRecordset("Select TestRegister.* From TestRegister Where (TestRegister.GroupCode = '" & sMCode & "' )")
    If rs.RecordCount > 0 Then
        bExist = True
    End If
    rs.Close
    
    checkIfParentNode = bExist
End Function
