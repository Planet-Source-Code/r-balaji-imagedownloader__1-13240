VERSION 5.00
Begin VB.UserControl AsyncBitmap 
   ClientHeight    =   2055
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2340
   ScaleHeight     =   2055
   ScaleWidth      =   2340
   Begin VB.PictureBox picBitmap 
      AutoSize        =   -1  'True
      Height          =   1695
      Left            =   120
      ScaleHeight     =   1635
      ScaleWidth      =   1995
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      Height          =   135
      Left            =   480
      ScaleHeight     =   135
      ScaleWidth      =   15
      TabIndex        =   0
      Top             =   480
      Width           =   15
   End
End
Attribute VB_Name = "AsyncBitmap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim g_Counter As Integer
Private mstrPictureFromURL As String
Private mstrFileName As String
Private m_bDownloadCompleted As Boolean
Public Event DownloadCompleted()
Public Property Get DownloadCompleted() As Boolean
    DownloadCompleted = m_bDownloadCompleted
End Property
Public Property Let DownloadCompleted(ByVal newVal As Boolean)
    m_bDownloadCompleted = newVal
End Property
Public Property Get SaveToFileName() As String
   SaveToFileName = mstrPictureFromURL
End Property

Public Property Let SaveToFileName(ByVal NewString As String)
   mstrFileName = NewString
   PropertyChanged "SaveToFileName"
End Property
Public Property Get PictureFromURL() As String
   PictureFromURL = mstrPictureFromURL
End Property

Public Property Let PictureFromURL(ByVal NewString As String)
   On Error GoTo ErrHandler
   ' (Code to validate path or URL omitted.)
   mstrPictureFromURL = NewString
   If (Ambient.UserMode = True) And (NewString <> "") Then
      ' If program is in run mode, and the URL string
      ' is not empty, begin the download.
    'AsyncRead NewString, vbAsyncTypePicture, "PictureFromURL"
    m_bDownloadCompleted = False
    'AsyncRead NewString, vbAsyncTypeFile, "FileFromURL", 1
    AsyncRead NewString, vbAsyncTypeFile, mstrFileName & CStr(1000 * Rnd), 1
   End If
   Exit Property
ErrHandler:
    RaiseEvent DownloadCompleted
End Property

Private Sub UserControl_AsyncReadComplete(AsyncProp As AsyncProperty)
   On Error Resume Next
   'Select Case AsyncProp.PropertyName
    '  Case "PictureFromURL"
     '    Set Picture = AsyncProp.Value
      '   Debug.Print "Download complete"
    'Case "FileFromURL"
        CopyToFile (AsyncProp.Value)
   'End Select
End Sub
Private Sub picBitmap_Resize()
   ' If there's a Picture assigned, resize.
   If picBitmap.Picture <> 0 Then
      UserControl.Size picBitmap.Width, _
         picBitmap.Height
   End If
End Sub

Private Sub UserControl_Resize()
   If picBitmap.Picture = 0 Then
      picBitmap.Move 0, 0, ScaleWidth, ScaleHeight
   Else
      If (Width <> picBitmap.Width) _
            Or (Height <> picBitmap.Height) Then
         Size picBitmap.Width, picBitmap.Height
      End If
   End If
End Sub

Private Sub UserControl_InitProperties()
   ' Use Nothing as the default when initializing,
   '   reading, and writing the Picture property,
   '   so than an .frx file won't be needed if
   '   there's no picture.
   m_bDownloadCompleted = True
   Set Picture = Nothing
End Sub

Private Sub UserControl_ReadProperties( _
         PropBag As PropertyBag)
   Set Picture = _
      PropBag.ReadProperty("Picture", Nothing)
End Sub

Private Sub UserControl_WriteProperties( _
         PropBag As PropertyBag)
   PropBag.WriteProperty "Picture", Picture, Nothing
End Sub



Public Property Get Picture() As Picture
   Set Picture = picBitmap.Picture
End Property

Public Property Let Picture(ByVal NewPicture _
      As Picture)
   Set picBitmap.Picture = NewPicture
   PropertyChanged "Picture"
End Property

Public Property Set Picture(ByVal NewPicture _
      As Picture)
   Set picBitmap.Picture = NewPicture
   'PropertyChanged "Picture"
End Property

Private Function CopyToFile(ByVal filename As String)
Dim nHandle As Integer, nHandleRead As Integer
Dim sPath As String, sFile As String
Dim byteTmp As Byte
    nHandle = FreeFile
    Open mstrFileName For Binary Access Write As nHandle
     nHandleRead = FreeFile
    Open filename For Binary Access Read As nHandleRead
    Do While Not EOF(nHandleRead)
        Get nHandleRead, , byteTmp
        Put nHandle, , byteTmp
    Loop
    Close nHandle
    Close nHandleRead
    m_bDownloadCompleted = True
    RaiseEvent DownloadCompleted
End Function

