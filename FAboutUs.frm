VERSION 5.00
Begin VB.Form FAboutUs 
   BackColor       =   &H00EFEFEF&
   ClientHeight    =   7020
   ClientLeft      =   60
   ClientTop       =   165
   ClientWidth     =   10590
   ControlBox      =   0   'False
   Icon            =   "FAboutUs.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7020
   ScaleWidth      =   10590
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CClose 
      Cancel          =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   9060
      Picture         =   "FAboutUs.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6435
      Width           =   1365
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   4785
      Picture         =   "FAboutUs.frx":246E
      Top             =   3090
      Width           =   720
   End
End
Attribute VB_Name = "FAboutUs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CClose_Click()
    Unload Me
End Sub
