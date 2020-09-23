VERSION 5.00
Begin VB.Form frmOptions 
   Caption         =   "Options"
   ClientHeight    =   5235
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   5235
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkPutInSysTray 
      Caption         =   "When I click Hide Button Show Icon alone in System Tray"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   2160
      Width           =   4575
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   3480
      TabIndex        =   11
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton cmdResume 
      Caption         =   "Resume"
      Height          =   375
      Left            =   2280
      TabIndex        =   10
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3120
      TabIndex        =   9
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Add links whose name like "
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   3120
      Width           =   4335
      Begin VB.TextBox txtURLlike 
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   4095
      End
   End
   Begin VB.CheckBox chkAlertOnImage 
      Caption         =   "Alert on completion of every image."
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   3975
   End
   Begin VB.CheckBox chkAlertPage 
      Caption         =   "Alert on completion of every page."
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   3855
   End
   Begin VB.CheckBox chkDLAll 
      Caption         =   "Download All Links "
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   3855
   End
   Begin VB.CommandButton cmdLinks 
      Caption         =   "Select links"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton cmdStats 
      Caption         =   "Statistics"
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   3840
      Width           =   975
   End
   Begin VB.CheckBox chkAutoDownload 
      Caption         =   "Download the Images in the page Automatically"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkAutoDownload_Click()
On Error Resume Next
    If chkAutoDownload.Value = 1 Then
        If Trim(Form1.txtFilePath.Text) = "" Then
            MsgBox ("Please Enter the file path...!!!")
            frmChooseDir.Show 1
            Exit Sub
        End If
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
On Error Resume Next
    If chkDLAll.Value = 1 Then
        Form1.m_bDownloadAll = True
        Form1.DownloadAll
    Else
        Form1.m_bDownloadAll = False
    End If

    If chkAlertOnImage.Value = 1 Then
        Form1.m_bAlertOnImageCompletion = True
    Else
        Form1.m_bAlertOnImageCompletion = False
    End If

    If chkAutoDownload.Value = 1 Then
        Form1.m_bAutoDownload = True
    Else
        Form1.m_bAutoDownload = False
    End If
    
    If chkAlertPage.Value = 1 Then
        Form1.m_bAlertOnPageCompletion = True
    Else
        Form1.m_bAlertOnPageCompletion = False
    End If
    If txtURLlike.Text <> "" Then
        Form1.m_strURLlike = txtURLlike.Text
    End If
    If chkPutInSysTray.Value = 1 Then
        Form1.m_bPutIconInSystemTray = True
    Else
        Form1.m_bPutIconInSystemTray = False
    End If

    Unload Me
End Sub

Public Sub cmdResume_Click()
    Dim strFileName As String
    Dim intFileHandle As Integer
    Dim strLinkName As String
    Dim strImageName As String
    On Error GoTo ErrHandler
    
    'resume available links
    intFileHandle = FreeFile
    strFileName = Form1.txtFilePath.Text & "AvailableLinks.idl"
    Open strFileName For Input As intFileHandle
    Do While Not EOF(intFileHandle)
        Input #intFileHandle, strLinkName
        Form1.colLink.Add strLinkName
    Loop
    Close intFileHandle
    
    'resume pending links
    intFileHandle = FreeFile
    strFileName = Form1.txtFilePath.Text & "PendingLinks.idl"
    Open strFileName For Input As intFileHandle
    Do While Not EOF(intFileHandle)
        Input #intFileHandle, strLinkName
        Form1.colPendingPages.Add strLinkName
    Loop
    Close intFileHandle
    
    'resume visited links
    intFileHandle = FreeFile
    strFileName = Form1.txtFilePath.Text & "VisitedLinks.idl"
    Open strFileName For Input As intFileHandle
    Do While Not EOF(intFileHandle)
        Input #intFileHandle, strLinkName
        Form1.colVisitedPages.Add strLinkName
    Loop
    Close intFileHandle
    
    'resume image list
    intFileHandle = FreeFile
    strFileName = Form1.txtFilePath.Text & "DLImages.idl"
    Open strFileName For Input As intFileHandle
    Do While Not EOF(intFileHandle)
        Input #intFileHandle, strImageName
        Form1.colImg.Add strImageName
    Loop
    Close intFileHandle
    Exit Sub
ErrHandler:
        Exit Sub
End Sub

Public Sub cmdSave_Click()
    Dim strFileName As String
    Dim intFileHandle As Integer
    On Error GoTo ErrHandler
    
    'save available links
    intFileHandle = FreeFile
    strFileName = Form1.txtFilePath.Text & "AvailableLinks.idl"
    Open strFileName For Output As intFileHandle
    For i = 1 To Form1.colLink.Count
        Write #intFileHandle, Form1.colLink(i)
    Next i
    Close intFileHandle
    
    'save pending links
    intFileHandle = FreeFile
    strFileName = Form1.txtFilePath.Text & "PendingLinks.idl"
    Open strFileName For Output As intFileHandle
    For i = 1 To Form1.colPendingPages.Count
        Write #intFileHandle, Form1.colPendingPages(i)
    Next i
    Close intFileHandle
    
    'save visited links
    intFileHandle = FreeFile
    strFileName = Form1.txtFilePath.Text & "VisitedLinks.idl"
    Open strFileName For Output As intFileHandle
    For i = 1 To Form1.colVisitedPages.Count
        Write #intFileHandle, Form1.colVisitedPages(i)
    Next i
    Close intFileHandle
    
    'save image lise
    intFileHandle = FreeFile
    strFileName = Form1.txtFilePath.Text & "DLImages.idl"
    Open strFileName For Output As intFileHandle
    For i = 1 To Form1.colVisitedPages.Count
        Write #intFileHandle, Form1.colImg(i)
    Next i
    Close intFileHandle
ErrHandler:
    Exit Sub

End Sub

Private Sub cmdStats_Click()
   frmStatistics.Show 1
End Sub
Private Sub cmdLinks_Click()
    frmLinks.Show 1
End Sub
Private Sub chkDLAll_Click()
    If chkDLAll.Value = 1 Then
        chkAutoDownload.Value = 1
    End If
End Sub

Private Sub Form_Load()
On Error Resume Next
    With Form1
        If .m_bAlertOnImageCompletion = True Then
            chkAlertOnImage.Value = 1
        Else
            chkAlertOnImage.Value = 0
        End If
        
        If .m_bAlertOnPageCompletion = True Then
            chkAlertPage.Value = 1
        Else
            chkAlertPage.Value = 0
        End If
        
        If .m_bDownloadAll = True Then
            chkDLAll.Value = 1
        Else
            chkDLAll.Value = 0
        End If
        
        If .m_bAutoDownload = True Then
            chkAutoDownload.Value = 1
        Else
            chkAutoDownload.Value = 0
        End If
        If .m_strURLlike <> "" Then
             txtURLlike.Text = .m_strURLlike
        End If
        If .m_bPutIconInSystemTray = True Then
            chkPutInSysTray.Value = 1
        End If
    End With
End Sub
