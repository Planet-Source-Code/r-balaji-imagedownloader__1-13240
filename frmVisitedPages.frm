VERSION 5.00
Begin VB.Form frmVisitedPages 
   Caption         =   "Visited Pages"
   ClientHeight    =   3480
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3480
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   3000
      Width           =   1095
   End
   Begin VB.ListBox lstVisitedPages 
      Height          =   2205
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   4095
   End
   Begin VB.Label lblTotalPages 
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "frmVisitedPages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim i As Integer
lstVisitedPages.Clear
For i = 1 To Form1.colVisitedPages.Count
    lstVisitedPages.AddItem Form1.colVisitedPages(i)
Next i
lblTotalPages.Caption = "Total Number of Links Visited are " & Form1.colVisitedPages.Count
End Sub
