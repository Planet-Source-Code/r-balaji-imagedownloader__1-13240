VERSION 5.00
Begin VB.Form frmStatistics 
   Caption         =   "Image Downloader- Statistics"
   ClientHeight    =   4545
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4260
   Icon            =   "frmStatistics.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   4545
   ScaleWidth      =   4260
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   315
      Left            =   1673
      TabIndex        =   6
      Top             =   4200
      Width           =   915
   End
   Begin VB.TextBox txtLastPageCount 
      Height          =   285
      Left            =   3120
      TabIndex        =   5
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox txtTotalImgsCount 
      Height          =   285
      Left            =   3120
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "DownLoaded Images"
      Height          =   3015
      Left            =   143
      TabIndex        =   2
      Top             =   1080
      Width           =   3975
      Begin VB.FileListBox fleImageList 
         Height          =   2430
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   3735
      End
   End
   Begin VB.Label Label2 
      Caption         =   " Images Downloaded in Last Page"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Total Number of Images Downloaded"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "frmStatistics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
    txtLastPageCount.Text = Form1.n_imgCountLastPage
    txtTotalImgsCount.Text = Form1.n_imgCountTotal
    fleImageList.Path = Form1.txtFilePath.Text
End Sub
