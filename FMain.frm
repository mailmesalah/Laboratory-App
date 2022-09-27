VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FMain 
   Caption         =   "Lab App by Lychee Technologies LLP"
   ClientHeight    =   7335
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   11190
   Icon            =   "FMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7335
   ScaleWidth      =   11190
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   780
      Left            =   4020
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2805
      Visible         =   0   'False
      Width           =   1860
   End
   Begin MSComDlg.CommonDialog CoDialog 
      Left            =   945
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu MTransactions 
      Caption         =   "Transactions"
      Begin VB.Menu MTMedicalTest 
         Caption         =   "Test Result Entry"
      End
      Begin VB.Menu Separator0 
         Caption         =   "-"
      End
      Begin VB.Menu MRBillReport 
         Caption         =   "Bill Report"
      End
      Begin VB.Menu MRMedicalReport 
         Caption         =   "Medical Report"
      End
      Begin VB.Menu Separator1 
         Caption         =   "-"
      End
   End
   Begin VB.Menu MSettings 
      Caption         =   "Settings"
      Begin VB.Menu MSUserAccounts 
         Caption         =   "User Accounts"
      End
      Begin VB.Menu Separator2 
         Caption         =   "-"
      End
   End
   Begin VB.Menu MAbout 
      Caption         =   "About"
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long

Private Sub Form_Load()

    setUserAccess
    'If Date > DateValue("08/30/2015") Then
    '    MsgBox "Your Trial Period has Expired !, Please Contact Lychee Technologies." & Format(DateValue("08/30/2015"), "dd-MM-yyyy")
    '    End
    'End If
    
End Sub

Public Sub initialisePublicVariables()
    Set db = OpenDatabase(App.Path & "\Storage.mdb", False, False, "MS Access;PWD=12345abcde")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub


Private Sub MAbout_Click()
    FAboutUs.Show
End Sub

Private Sub MRBillReport_Click()
    FBillReport.Show
End Sub

Private Sub MRMedicalReport_Click()
    FMedicalReport.Show
End Sub

Private Sub MSUserAccounts_Click()
    FUserAccounts.Show
End Sub

Private Sub MTMedicalTest_Click()
    FMedicalTest.Show
End Sub

Private Sub setUserAccess()
Dim rs As Recordset, r As Long
    
    Set rs = db.OpenRecordset("Select Rights.RightDescription,Rights.MapName,Rights.Status,Users.RightCode From Rights,Users Where (Users.Code = '" & sCurrentUserCode & "' ) And (Rights.Code = Users.RightCode ) Order By Val(Rights.MenuCode) Desc")
    If rs.RecordCount > 0 Then
        If Trim(rs!RightDescription) = "Administrator" Then
            'SHOW ALL
            r = 0
            Do While r < Me.Controls.count
                'SKIPPING MENU DIVIDERS
                If Left(Me.Controls(r).Name, 1) = "M" And Len(Me.Controls(r).Name) > 5 Then
                    Me.Controls(r).Visible = True
                End If
                r = r + 1
            Loop
            
            MAbout.Visible = True
            
        ElseIf Trim(rs!RightDescription) = "None" Then
            'SHOW NONE
            r = 0
            Do While r < Me.Controls.count
                If Left(Me.Controls(r).Name, 1) = "M" And Len(Me.Controls(r).Name) > 5 Then
                    Me.Controls(r).Visible = False
                End If
                r = r + 1
            Loop
            
            MAbout.Visible = True
                        
        Else
            While rs.EOF = False
    
                If Left(rs!MapName, 1) = "B" Then
                    Select Case rs!MapName
                    Case "BTestEdit":
                        EditTestEntry = rs!Status
                    Case "BRegisterEdit":
                        EditMasterRegisters = rs!Status
                    Case Else
                    End Select
                     
                Else
                    r = 0
                    Do While r < Me.Controls.count
                        If Trim(Me.Controls(r).Name) = Trim(rs!MapName) Then
                            Me.Controls(r).Visible = rs!Status
                            Exit Do
                      End If
                        r = r + 1
                    Loop
                End If
                rs.MoveNext
            Wend
            
            
            MAbout.Visible = True
        End If
    Else
        'SHOW NONE
        MAbout.Visible = True
    End If
    rs.Close
End Sub

