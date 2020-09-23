VERSION 5.00
Begin VB.Form frmChooseDir 
   Caption         =   "Select Directory"
   ClientHeight    =   3780
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3630
   LinkTopic       =   "Form2"
   ScaleHeight     =   3780
   ScaleWidth      =   3630
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select Download Directory"
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3255
      Begin VB.DirListBox Dir1 
         Height          =   1890
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   3015
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   3015
      End
   End
End
Attribute VB_Name = "frmChooseDir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Mid$(Dir1.Path, Len(Dir1.Path)) <> "\" Then
        Form1.txtFilePath.Text = Dir1.Path & "\"
    Else
        Form1.txtFilePath.Text = Dir1.Path
    End If
    Unload Me
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
    If Form1.txtFilePath.Text <> "" Then
        If Mid$(Form1.txtFilePath.Text, Len(Form1.txtFilePath.Text)) <> "\" Then
            Dir1.Path = Mid$(Form1.txtFilePath.Text, 1, Len(Form1.txtFilePath.Text) - 1)
        Else
            Dir1.Path = Form1.txtFilePath.Text
        End If
    End If
End Sub
