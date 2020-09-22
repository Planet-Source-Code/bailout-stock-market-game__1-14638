VERSION 5.00
Begin VB.Form frmGraph 
   Caption         =   "Graph"
   ClientHeight    =   1245
   ClientLeft      =   4440
   ClientTop       =   4575
   ClientWidth     =   1620
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   1245
   ScaleWidth      =   1620
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   1215
      Left            =   0
      ScaleHeight     =   1155
      ScaleWidth      =   1515
      TabIndex        =   0
      Top             =   0
      Width           =   1575
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print"
      End
      Begin VB.Menu mnuSeperator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
      End
   End
End
Attribute VB_Name = "frmGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuClose_Click()
    Unload Me
End Sub

Private Sub mnuPrint_Click()
    PrintPictureToFitPage Printer, Picture1.Picture
    Printer.EndDoc
End Sub

Private Sub Picture1_Resize()
    frmGraph.Height = Picture1.Height + 700
    frmGraph.Width = Picture1.Width + 120
   
End Sub
