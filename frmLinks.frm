VERSION 5.00
Begin VB.Form frmLinks 
   Caption         =   "Choose the Links"
   ClientHeight    =   5115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5835
   LinkTopic       =   "Form2"
   ScaleHeight     =   5115
   ScaleWidth      =   5835
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdVisitedLinks 
      Caption         =   "..."
      Height          =   375
      Left            =   2640
      TabIndex        =   9
      ToolTipText     =   "Visited Pages"
      Top             =   4680
      Width           =   495
   End
   Begin VB.TextBox txtLink 
      Height          =   285
      Left            =   240
      TabIndex        =   8
      ToolTipText     =   "Selected Page"
      Top             =   120
      Width           =   5415
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3270
      TabIndex        =   7
      ToolTipText     =   "Cancel "
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   1470
      TabIndex        =   6
      ToolTipText     =   "Click when Selection is Over"
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton cmdUnselectall 
      Caption         =   "<<"
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      ToolTipText     =   "Unselect All Links"
      Top             =   2760
      Width           =   375
   End
   Begin VB.CommandButton cmdUnselect 
      Caption         =   "<"
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      ToolTipText     =   "Unselect Single Link"
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton cmdSelectAll 
      Caption         =   ">>"
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      ToolTipText     =   "Select All Links"
      Top             =   1800
      Width           =   375
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   ">"
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      ToolTipText     =   "Select Single Link"
      Top             =   1320
      Width           =   375
   End
   Begin VB.ListBox lstSelected 
      Height          =   3765
      Left            =   3120
      TabIndex        =   1
      ToolTipText     =   "Double Click the Item to Unselect"
      Top             =   840
      Width           =   2535
   End
   Begin VB.ListBox lstAvailable 
      Height          =   3765
      ItemData        =   "frmLinks.frx":0000
      Left            =   240
      List            =   "frmLinks.frx":0007
      TabIndex        =   0
      ToolTipText     =   "Double Click the Item to Select"
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "Selected Links"
      Height          =   255
      Left            =   3120
      TabIndex        =   11
      Top             =   480
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Available Links"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   480
      Width           =   2295
   End
End
Attribute VB_Name = "frmLinks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
Dim i As Integer
    While Form1.colPendingPages.Count > 0
        Form1.colPendingPages.Remove (1)
    Wend
    For i = 0 To lstSelected.ListCount - 1
        Form1.colPendingPages.Add (lstSelected.List(i))
    Next i
    Unload Me
End Sub

Private Sub cmdSelect_Click()
Dim i As Integer
i = 0
    While lstAvailable.SelCount > 0
        If lstAvailable.Selected(i) = True Then
            lstSelected.AddItem lstAvailable.List(i)
            lstAvailable.RemoveItem (i)
        End If
        i = i + 1
    Wend
End Sub

Private Sub cmdSelectAll_Click()
Dim i As Integer
    For i = 0 To (lstAvailable.ListCount - 1)
        lstSelected.AddItem lstAvailable.List(i)
    Next i
    lstAvailable.Clear
End Sub

Private Sub cmdUnselect_Click()
Dim i As Integer
i = 0
    While lstSelected.SelCount > 0
        If lstSelected.Selected(i) = True Then
            lstAvailable.AddItem lstSelected.List(i)
            lstSelected.RemoveItem (i)
        End If
        i = i + 1
    Wend
End Sub

Private Sub cmdUnselectall_Click()
Dim i As Integer
    For i = 0 To (lstSelected.ListCount - 1)
        lstAvailable.AddItem lstSelected.List(i)
    Next i
    lstSelected.Clear
End Sub

Private Sub cmdVisitedLinks_Click()
    frmVisitedPages.Show 1
End Sub

Private Sub Form_Load()
Dim i As Integer
    lstAvailable.Clear
    lstSelected.Clear
    For i = 1 To Form1.colLink.Count
        lstAvailable.AddItem Form1.colLink(i)
    Next i
    
    For i = 1 To Form1.colPendingPages.Count
        lstSelected.AddItem Form1.colPendingPages(i)
    Next i
End Sub

Private Sub lstAvailable_Click()
    txtLink.Text = lstAvailable.Text
End Sub

Private Sub lstAvailable_DblClick()
    cmdSelect_Click
End Sub

Private Sub lstSelected_Click()
    txtLink.Text = lstSelected.Text
End Sub

Private Sub lstSelected_DblClick()
    cmdUnselect_Click
End Sub
