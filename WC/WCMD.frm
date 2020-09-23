VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H80000004&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "WCMD [WebCam Motion Detector]"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14850
   Icon            =   "WCMD.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   14850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Clear list"
      Height          =   495
      Left            =   10200
      TabIndex        =   12
      Top             =   3960
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   495
      Left            =   12840
      TabIndex        =   11
      Top             =   4560
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Start"
      Default         =   -1  'True
      Height          =   495
      Left            =   12840
      TabIndex        =   10
      Top             =   3960
      Width           =   1815
   End
   Begin VB.ListBox List1 
      Height          =   3570
      Left            =   10200
      TabIndex        =   9
      Top             =   120
      Width           =   4455
   End
   Begin MSComctlLib.Slider Slider2 
      Height          =   255
      Left            =   2040
      TabIndex        =   6
      Top             =   4680
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   450
      _Version        =   393216
      LargeChange     =   1
      Min             =   1
      Max             =   49
      SelStart        =   15
      Value           =   15
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   255
      Left            =   2040
      TabIndex        =   5
      Top             =   4080
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   450
      _Version        =   393216
      LargeChange     =   1
      Min             =   1
      Max             =   50
      SelStart        =   41
      Value           =   41
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   11760
      Top             =   5040
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   7680
      TabIndex        =   3
      Top             =   4800
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Take Normal State"
      Height          =   495
      Left            =   7200
      TabIndex        =   2
      Top             =   3960
      Width           =   2895
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      Height          =   3600
      Left            =   5280
      ScaleHeight     =   3540
      ScaleWidth      =   4740
      TabIndex        =   1
      Top             =   120
      Width           =   4800
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      DrawWidth       =   3
      Height          =   3600
      Left            =   120
      ScaleHeight     =   3540
      ScaleWidth      =   4740
      TabIndex        =   0
      Top             =   120
      Width           =   4800
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   5640
      Top             =   5040
   End
   Begin VB.Label Label3 
      Caption         =   "Motion sensitivity:"
      Height          =   255
      Left            =   720
      TabIndex        =   8
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Color sensitivity:"
      Height          =   255
      Left            =   720
      TabIndex        =   7
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   6600
      TabIndex        =   4
      Top             =   4800
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This program was created by Haik Haiotsyan
'Special thanks to Grenik Poghosyan
'If you have any questions mail me
'Mail:haik_111@yahoo.com
'Sorry for bad English
Private Declare Function capCreateCaptureWindow Lib "avicap32.dll" Alias "capCreateCaptureWindowA" (ByVal lpszWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal nID As Long) As Long
Private Sub Command1_Click()
Picture2.Picture = Picture1.Picture
End Sub
Private Sub Command2_Click()
STARTCAM

Load Progress
Progress.Show
Progress.start
Unload Progress

Command2.Enabled = False
Command3.Enabled = True


Picture1.AutoRedraw = True
Picture2.AutoRedraw = True
Timer1.Enabled = True
Timer2.Enabled = True

End Sub

Private Sub Command3_Click()
STOPCAM
ProgressBar1.Value = 0
Command3.Enabled = False
Command2.Enabled = True

Picture1.Picture = LoadPicture("nosignal.bmp")
Picture2.Picture = LoadPicture("nosignal.bmp")
Label1.Caption = "0%"
End Sub

Private Sub Command4_Click()
List1.Clear
End Sub

Private Sub Form_Load()

Picture1.Width = 320 * Screen.TwipsPerPixelX
Picture1.Height = 240 * Screen.TwipsPerPixelY
Picture2.Width = 320 * Screen.TwipsPerPixelX
Picture2.Height = 240 * Screen.TwipsPerPixelY

Picture1.Picture = LoadPicture("nosignal.bmp")
Picture2.Picture = LoadPicture("nosignal.bmp")
End Sub

Private Function Different(ByVal a As Long, ByVal b As Long) As Boolean
'Checks different of two colors
ar = a Mod 256: a = a \ 256
ag = a Mod 256: a = a \ 256
ab = a Mod 256: a = a \ 256

br = b Mod 256: b = b \ 256
bg = b Mod 256: b = b \ 256
bb = b Mod 256: b = b \ 256
sense = 255 - Slider1.Value * 5

Different = (Sqr((ar - br) * (ar - br) + (ag - bg) * (ag - bg) + (ab - bb) * (ab - bb)) > sense) 'formula for counting different
End Function

Private Sub Form_Unload(Cancel As Integer)
STOPCAM
SaveSetting "MotionDetect", "Param", "s1", Str(Slider1.Value)
SaveSetting "MotionDetect", "Param", "s2", Str(Slider2.Value)

End Sub

Private Sub Timer1_Timer()
'getting picture from camera
SendMessage mCapHwnd, GET_FRAME, 0, 0
SendMessage mCapHwnd, COPY, 0, 0
Picture1.Picture = Clipboard.GetData: Clipboard.Clear


stepp = 3 'Grid dense

Dim qan, qann As Long
qan = 0
qann = 0

For i = 1 To Picture1.Width / Screen.TwipsPerPixelX Step stepp
For j = 1 To Picture1.Height / Screen.TwipsPerPixelY Step stepp

If Different(Picture1.Point(i * stepp * Screen.TwipsPerPixelX, j * stepp * Screen.TwipsPerPixelY), Picture2.Point(Screen.TwipsPerPixelX * i * stepp, j * stepp * Screen.TwipsPerPixelY)) Then
Picture1.Circle (i * stepp * Screen.TwipsPerPixelX, Screen.TwipsPerPixelY * j * stepp), 1, RGB(255, 0, 0)
qann = qann + 1
End If

Next
Next
Label1.Caption = Int(qann * 100 / 910) & "%" 'Counting motion in pracentes
ProgressBar1.Value = Int(qann * 100 / 910)
End Sub

Sub STOPCAM()
DoEvents: SendMessage mCapHwnd, DISCONNECT, 0, 0
Timer1.Enabled = False
Timer2.Enabled = False
End Sub

Sub STARTCAM()
'Getting handle of camera window
mCapHwnd = capCreateCaptureWindow("WebcamCapture", 0, 0, 0, 320, 240, Me.hwnd, 0)
DoEvents
SendMessage mCapHwnd, CONNECT, 0, 0 'connecting to camera
SendMessage mCapHwnd, WM_CAP_DLG_VIDEOFORMAT, 0, 0 'Calling video format dialog
DoEvents
Slider1.Value = GetSetting("MotionDetect", "Param", "s1", "0")
Slider2.Value = GetSetting("MotionDetect", "Param", "s2", "0")

End Sub

Private Sub Timer2_Timer()
If ProgressBar1.Value > 100 - Slider2.Value * 2 Then
Beep
SavePicture Picture1.Picture, App.Path + "\Detected\" + Format(Date, "ddmmyyyy") + "__" + Format(Time, "hhmmss") + ".bmp"
List1.AddItem "Saved in " + Str(Time) + " " + Str(ProgressBar1.Value) + "%   -->  " + Format(Date, "ddmmyyyy") + "__" + Format(Time, "hhmmss") + ".bmp"
End If
End Sub
