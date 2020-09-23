VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "Image Downloder"
   ClientHeight    =   4875
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   6045
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   6045
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1560
      Top             =   5040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdOptions 
      Caption         =   "Options"
      Height          =   375
      Left            =   4920
      TabIndex        =   16
      ToolTipText     =   "My Preferences"
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton cmdDir 
      Caption         =   "..."
      Height          =   375
      Left            =   5520
      TabIndex        =   15
      ToolTipText     =   "Click here to Select the Directory"
      Top             =   600
      Width           =   375
   End
   Begin VB.CommandButton cmdCancelDownload 
      Caption         =   "Cancel DL"
      Height          =   375
      Left            =   3960
      TabIndex        =   10
      ToolTipText     =   "Cancel the Download"
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Height          =   375
      Left            =   1080
      TabIndex        =   9
      ToolTipText     =   "Stop the browser"
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton cmdHide 
      Caption         =   "Hide"
      Height          =   375
      Left            =   2040
      TabIndex        =   8
      ToolTipText     =   "Hide the Browser"
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      ToolTipText     =   "start the browser"
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox txtURL 
      Height          =   405
      Left            =   1320
      TabIndex        =   6
      ToolTipText     =   "Enter the URL path "
      Top             =   120
      Width           =   4575
   End
   Begin VB.TextBox txtFilePath 
      Enabled         =   0   'False
      Height          =   405
      Left            =   1320
      TabIndex        =   5
      ToolTipText     =   "Select the Directory"
      Top             =   600
      Width           =   4095
   End
   Begin IDL.AsyncBitmap AsyncBitmap1 
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   5040
      Width           =   1095
      _ExtentX        =   2778
      _ExtentY        =   1931
   End
   Begin VB.PictureBox Picture2 
      Height          =   15
      Left            =   2520
      ScaleHeight     =   15
      ScaleWidth      =   135
      TabIndex        =   3
      Top             =   2280
      Width           =   135
   End
   Begin VB.PictureBox Picture1 
      Height          =   135
      Left            =   2640
      ScaleHeight     =   135
      ScaleWidth      =   15
      TabIndex        =   2
      Top             =   2160
      Width           =   15
   End
   Begin VB.CommandButton cmdDownload 
      Caption         =   "DownLoad"
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      ToolTipText     =   "Download the Images"
      Top             =   1080
      Width           =   975
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   5775
      ExtentX         =   10186
      ExtentY         =   3413
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Download Path"
      Height          =   255
      Left            =   -240
      TabIndex        =   14
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Target Url Path"
      Height          =   255
      Left            =   0
      TabIndex        =   13
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label lblDownLoadStatus 
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   4200
      Width           =   5775
   End
   Begin VB.Label lblStatus 
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   3600
      Width           =   5775
   End
   Begin VB.Menu mnuShowIDL 
      Caption         =   "Show Image Downloader"
      Visible         =   0   'False
      Begin VB.Menu mnuShow 
         Caption         =   "Show Image Downloader"
      End
      Begin VB.Menu mnuHypen1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStatistics 
         Caption         =   "Statistics"
      End
      Begin VB.Menu mnuChooseLinks 
         Caption         =   "Choose Links"
      End
      Begin VB.Menu mnuSaveStatus 
         Caption         =   "Save Status"
      End
      Begin VB.Menu mnuResumeStatus 
         Caption         =   "Resume Status"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbt 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim obj1 As New HTMLDocument    'to store the html document object
Dim imgObj As New HTMLImg       'to store the image object
Dim linkObj As New HTMLLinkElement  'to store the linkelement object
Dim strImgFileName As String
Dim strPath As String

Public m_strURLlike As String
Private n_imgCount As Integer   'to store the image number which is getting downloaded
Public n_imgCountTotal As Integer   'to count the total number of images downloaded
Public n_imgCountLastPage As Integer    'to count the total number of images downloaded in recent page

Public colImg As New Collection 'to store the image collection
Public colLink As New Collection    'to store the links
Public colVisitedPages As New Collection    'to keep track of visited pages\links
Public colPendingPages As New Collection    'to keep track of pending pages\links

'option variables
Public m_bPutIconInSystemTray As Boolean    'Get the option from user and store it
Public bCancelThisDownload As Boolean
Public m_bAutoDownload As Boolean
Public m_bDownloadAll As Boolean
Public m_bAlertOnPageCompletion As Boolean
Public m_bAlertOnImageCompletion As Boolean

'Declare a user-defined variable to pass to the Shell_NotifyIcon function.
Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

'The following constants are the messages sent to the Shell_NotifyIcon function to add, modify, or delete an icon from the taskbar status area.
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2

'The following constant is the message sent when a mouse event occurs within the rectangular boundaries of the icon in the taskbar status area.
Private Const WM_MOUSEMOVE = &H200

'The following constants are the flags that indicate the valid members of the NOTIFYICONDATA data type.
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

'The following constants are used to determine the mouse input on the the icon in the taskbar status area.
'Left-click constants.
Private Const WM_LBUTTONDBLCLK = &H203   'Double-click
Private Const WM_LBUTTONDOWN = &H201     'Button down
Private Const WM_LBUTTONUP = &H202       'Button up

'Right-click constants.
Private Const WM_RBUTTONDBLCLK = &H206   'Double-click
Private Const WM_RBUTTONDOWN = &H204     'Button down
Private Const WM_RBUTTONUP = &H205       'Button up

'Declare the API function call.
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

'Dimension a variable as the user-defined data type.
Dim nid As NOTIFYICONDATA

'function to add the icon to the taskbar status area.
Private Sub AddIcon()
   'Set the individual values of the NOTIFYICONDATA data type.
   nid.cbSize = Len(nid)
   nid.hWnd = Form1.hWnd
   nid.uId = vbNull
   nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
   nid.uCallBackMessage = WM_MOUSEMOVE
   nid.hIcon = Form1.Icon
   nid.szTip = "Image DownLoader" & vbNullChar

   'Call the Shell_NotifyIcon function to add the icon to the taskbar status area.
   Shell_NotifyIcon NIM_ADD, nid

End Sub
'function to delete the added icon from the taskbar status area
Private Sub DeleteIcon()
   'Click this button to delete the added icon from the taskbar
   'status area by calling the Shell_NotifyIcon function.
   Shell_NotifyIcon NIM_DELETE, nid
End Sub


'function to delete the added icon from the taskbar status area when the program ends.
Private Sub Form_Terminate()
   'Delete the added icon from the taskbar status area when the
   'program ends.
   Shell_NotifyIcon NIM_DELETE, nid
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
          'Event occurs when the mouse pointer is within the rectangular
          'boundaries of the icon in the taskbar status area.
          Dim mnuObj1 As Menu
          Dim msg As Long
          Dim sFilter As String
          msg = X / Screen.TwipsPerPixelX
          Select Case msg
             Case WM_LBUTTONDOWN
             Case WM_LBUTTONUP
             Case WM_LBUTTONDBLCLK
                'when the user clicks the icon, just display the form
                Form1.Show
             Case WM_RBUTTONDOWN
                'when the user rightclicks on the taskbar icon, show the option menu
                Form1.PopupMenu mnuShowIDL
             Case WM_RBUTTONUP
             Case WM_RBUTTONDBLCLK
          End Select
End Sub
'this event is called when the download is completed.
Private Sub AsyncBitmap1_DownloadCompleted()
    lblDownLoadStatus.Caption = ""
    If bCancelThisDownload = False Then
        DownloadNextFile
    End If
End Sub


Private Sub cmdCancelDownload_Click()
    bCancelThisDownload = True
    m_bDownloadAll = False
    WebBrowser1.Stop
End Sub

Private Sub cmdDir_Click()
    frmChooseDir.Show 1
End Sub



Private Sub cmdGo_Click()
    Dim strURL As String
    strURL = txtURL.Text
    If WebBrowser1.Busy = False Then
        cmdDownload.Enabled = False
        bCancelThisDownload = False
        WebBrowser1.Navigate strURL
    Else
        Beep
    End If
End Sub

Private Sub cmdHide_Click()
    If m_bPutIconInSystemTray = True Then
            cmdHide.Caption = "Hide"
            AddIcon
            Form1.Hide
    Else
        If WebBrowser1.Height = 30 Then
            WebBrowser1.Height = 1935
            cmdHide.Caption = "Hide"
            Form1.Height = 5235
        Else
            WebBrowser1.Height = 30
            cmdHide.Caption = "Show"
            Form1.Height = 1900
        End If
    End If
End Sub

Private Sub cmdOptions_Click()
    frmOptions.Show 1
End Sub

Private Sub cmdStop_Click()
    bCancelThisDownload = True
    WebBrowser1.Stop
End Sub

Private Sub cmdDownload_Click()
    bCancelThisDownload = False
    strPath = txtFilePath.Text
    If Trim(txtFilePath.Text) = "" Then
        MsgBox ("Please Enter the file path...!!!")
        Exit Sub
    End If
    If WebBrowser1.Busy = False Then
        Set obj1 = WebBrowser1.Document
        'download the links
        DownloadLinks
        colVisitedPages.Add txtURL.Text
        DownloadNextFile
    Else
        Beep
    End If
End Sub

Private Function ParseFileName(ByVal strURL As String)
    Dim nIndex As Integer
    nIndex = InStrRev(strURL, "/")
    If nIndex > 0 Then
        ParseFileName = Mid$(strURL, nIndex + 1)
    Else
        ParseFileName = ""
    End If
    
    'remove special characters
    nIndex = InStr(ParseFileName, "%20")
    While (nIndex > 0)
        ParseFileName = Mid$(ParseFileName, 1, nIndex - 1) & " " & Mid$(ParseFileName, nIndex + 3)
        nIndex = InStr(ParseFileName, "%20")
    Wend
End Function
Private Sub Form_Load()
    'Initialize the global variables
    n_imgCountTotal = 0
    n_imgCountLastPage = 0
    m_bDownloadAll = False
    m_bAlertOnPageCompletion = False
    m_bAlertOnImageCompletion = False
    m_bPutIconInSystemTray = False
    m_strURLlike = ""
End Sub


Private Sub mnuAbt_Click()
    frmAbout.Show 1
End Sub

Private Sub mnuChooseLinks_Click()
    frmLinks.Show 1
End Sub

Private Sub mnuExit_Click()
    DeleteIcon
    Unload Me
End Sub

Private Sub mnuResumeStatus_Click()
    frmOptions.cmdSave_Click
End Sub

Private Sub mnuSaveStatus_Click()
    frmOptions.cmdSave_Click
End Sub

Private Sub mnuShow_Click()
    Form1.Show
    DeleteIcon
End Sub

Private Sub mnuStatistics_Click()
    frmStatistics.Show 1
End Sub

Private Sub txtURL_Change()
    n_imgCount = 0
    n_imgCountLastPage = 0
End Sub

Private Sub WebBrowser1_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
    If txtURL.Text <> CStr(URL) Then
        bCancelThisDownload = False
        txtURL.Text = URL
    End If
End Sub

Private Function CheckForCancellation() As Boolean
    Dim i
    i = i + 100
End Function

Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    cmdDownload.Enabled = True
    If colPendingPages.Count > 0 Then
            colPendingPages.Remove (1)
            cmdDownload_Click
            Exit Sub
    End If
    If m_bAutoDownload = True Then
        cmdDownload_Click
        Exit Sub
    End If
End Sub
Private Sub WebBrowser1_StatusTextChange(ByVal Text As String)
    lblStatus.Caption = Text
End Sub

Private Sub DownloadNextFile()
If WebBrowser1.Busy = False Then
    If bCancelThisDownload = False Then
        If n_imgCount <= obj1.images.length - 1 Then
            Set imgObj = obj1.images(n_imgCount)
            n_imgCount = n_imgCount + 1
            strImgFileName = ParseFileName(imgObj.src)
            If (AddToImageCollection(strImgFileName) = True) Then
                If strImgFileName <> "" Then
                    n_imgCountTotal = n_imgCountTotal + 1
                    n_imgCountLastPage = n_imgCountLastPage + 1
                    lblDownLoadStatus.Caption = "Downloading the image " & strImgFileName
                    AsyncBitmap1.SaveToFileName = strPath & strImgFileName
                    AsyncBitmap1.PictureFromURL = imgObj.src
                    If m_bAlertOnImageCompletion = True Then
                        MsgBox ("The image " & strImgFileName & " is downloaded...")
                    End If
                End If
            Else
                'download the next image
                AsyncBitmap1_DownloadCompleted
            End If
        Else
            If m_bAlertOnPageCompletion = True Then
                MsgBox ("downloading of " & txtURL.Text & " is over")
            End If
    
            If colPendingPages.Count > 0 Then
                cmdDownload.Enabled = False
                DownloadAll
            End If
        End If
    End If
End If
End Sub
Private Function AddToImageCollection(strFileName) As Boolean
    'check whether the element already exists
    For i = 1 To colImg.Count
        If strFileName = colImg(i) Then
            AddToImageCollection = False
            Exit Function
        End If
    Next i
    colImg.Add (strFileName)
    AddToImageCollection = True
End Function
Public Sub DownloadLinks()
    If WebBrowser1.Busy = False Then
        For i = 0 To obj1.links.length - 1
            Set linkObj = obj1.links(i)
            If linkObj.protocol = "http:" Then
                strLink = CStr(linkObj)
                If m_strURLlike <> "" And InStr(1, strLink, m_strURLlike) > 0 Then
                    AddToLinkCollection (strLink)
                End If
            End If
        Next i
    End If
End Sub
Private Function AddToLinkCollection(strLinkName) As Boolean
    'check whether the element already exists
    For i = 1 To colLink.Count
        If strLinkName = colLink(i) Then
            AddToLinkCollection = False
            Exit Function
        End If
    Next i
    colLink.Add (strLinkName)
    
    If m_bDownloadAll = True Then
        'check whether the element already exists
        For i = 1 To colVisitedPages.Count
            If strLinkName = colVisitedPages(i) Then
                AddToLinkCollection = False
                Exit Function
            End If
        Next i
        colPendingPages.Add strLinkName
    End If
    AddToLinkCollection = True
End Function

Public Sub DownloadAll()
    If colPendingPages.Count > 0 Then
        txtURL = colPendingPages(1)
        cmdGo_Click
    End If
End Sub
