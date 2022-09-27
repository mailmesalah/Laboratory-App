VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "print"
      Height          =   390
      Left            =   1560
      TabIndex        =   0
      Top             =   1245
      Width           =   1110
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Dim objPDF As New mjwPDF
Dim x As Double, y As Double, x1 As Double, y1 As Double
Dim myPicture As IPictureDisp

        objPDF.PDFTitle = "Medical Report"
        objPDF.PDFFileName = App.Path & "\Pdf\Medical Report.pdf"
        objPDF.PDFLoadAfm = App.Path & "\Fonts"
        objPDF.PDFSetUnit = UNIT_PT
        'objPDF.PDFView = True
        objPDF.PDFBeginDoc
          
                    
        y = 5
        x1 = (objPDF.PDFGetPageWidth / 2) - 110
        objPDF.PDFImage App.Path & "\Logo.jpg", x1, y
        y = 150
                
        x1 = 80
        objPDF.PDFSetFont FONT_ARIAL, 22, FONT_BOLD
        objPDF.PDFSetTextColor = vbBlue
        objPDF.PDFTextOut "CHANAKYA DIAGNOSTIC LABORATORY", x1, y
        ' End our PDF document (this will save it to the filename)
        objPDF.PDFEndDoc
        Set objPDF = Nothing
        
        
End Sub
