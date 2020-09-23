VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Progress 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Wait while camera waked up..."
   ClientHeight    =   735
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   Icon            =   "Progress.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   735
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   1
      Max             =   15
   End
End
Attribute VB_Name = "Progress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub start()

For i = 1 To 15
DoEvents
SendMessage mCapHwnd, GET_FRAME, 0, 0
SendMessage mCapHwnd, COPY, 0, 0
Form1.Picture2.Picture = Clipboard.GetData: Clipboard.Clear
DoEvents
ProgressBar1.Value = ProgressBar1.Value + 1
Next
End Sub

Private Sub Form_Load()
ProgressBar1.Value = 0
End Sub
