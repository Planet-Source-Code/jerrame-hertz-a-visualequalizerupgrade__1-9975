VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Begin VB.Form VisualEQ 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   2472
   ClientLeft      =   1560
   ClientTop       =   3060
   ClientWidth     =   3768
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmVisEQ.frx":0000
   ScaleHeight     =   2472
   ScaleWidth      =   3768
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   216
      Top             =   2520
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      Height          =   2148
      Left            =   144
      Picture         =   "frmVisEQ.frx":DFEC
      ScaleHeight     =   2100
      ScaleWidth      =   3420
      TabIndex        =   0
      Top             =   175
      Width           =   3468
      Begin PicClip.PictureClip PictureClip2 
         Left            =   0
         Top             =   144
         _ExtentX        =   212
         _ExtentY        =   1588
         _Version        =   327680
         Rows            =   3
         Picture         =   "frmVisEQ.frx":21677
      End
      Begin PicClip.PictureClip PictureClip1 
         Left            =   -72
         Top             =   1800
         _ExtentX        =   1016
         _ExtentY        =   466
         _Version        =   327680
         Picture         =   "frmVisEQ.frx":22029
      End
      Begin VB.PictureBox Picture3 
         Height          =   1050
         Left            =   216
         ScaleHeight     =   1008
         ScaleWidth      =   2880
         TabIndex        =   2
         Top             =   732
         Width           =   2930
         Begin MSComctlLib.ProgressBar ProgressBar10 
            Height          =   900
            Left            =   2664
            TabIndex        =   3
            ToolTipText     =   "16kHz"
            Top             =   72
            Width           =   156
            _ExtentX        =   275
            _ExtentY        =   1588
            _Version        =   393216
            Appearance      =   0
            Orientation     =   1
            Scrolling       =   1
         End
         Begin MSComctlLib.ProgressBar ProgressBar9 
            Height          =   900
            Left            =   2376
            TabIndex        =   4
            ToolTipText     =   "12kHz"
            Top             =   72
            Width           =   156
            _ExtentX        =   275
            _ExtentY        =   1588
            _Version        =   393216
            Appearance      =   0
            Orientation     =   1
            Scrolling       =   1
         End
         Begin MSComctlLib.ProgressBar ProgressBar8 
            Height          =   900
            Left            =   2088
            TabIndex        =   5
            ToolTipText     =   "6kHz"
            Top             =   72
            Width           =   156
            _ExtentX        =   275
            _ExtentY        =   1588
            _Version        =   393216
            Appearance      =   0
            Orientation     =   1
            Scrolling       =   1
         End
         Begin MSComctlLib.ProgressBar ProgressBar7 
            Height          =   900
            Left            =   1800
            TabIndex        =   6
            ToolTipText     =   "3kHz"
            Top             =   72
            Width           =   156
            _ExtentX        =   275
            _ExtentY        =   1588
            _Version        =   393216
            Appearance      =   0
            Orientation     =   1
            Scrolling       =   1
         End
         Begin MSComctlLib.ProgressBar ProgressBar6 
            Height          =   900
            Left            =   1512
            TabIndex        =   7
            ToolTipText     =   "1kHz"
            Top             =   72
            Width           =   156
            _ExtentX        =   275
            _ExtentY        =   1588
            _Version        =   393216
            Appearance      =   0
            Orientation     =   1
            Scrolling       =   1
         End
         Begin MSComctlLib.ProgressBar ProgressBar5 
            Height          =   900
            Left            =   1224
            TabIndex        =   8
            ToolTipText     =   "31Hz"
            Top             =   72
            Width           =   156
            _ExtentX        =   275
            _ExtentY        =   1588
            _Version        =   393216
            Appearance      =   0
            Orientation     =   1
            Scrolling       =   1
         End
         Begin MSComctlLib.ProgressBar ProgressBar4 
            Height          =   900
            Left            =   936
            TabIndex        =   9
            ToolTipText     =   "62Hz"
            Top             =   72
            Width           =   156
            _ExtentX        =   275
            _ExtentY        =   1588
            _Version        =   393216
            Appearance      =   0
            Orientation     =   1
            Scrolling       =   1
         End
         Begin MSComctlLib.ProgressBar ProgressBar3 
            Height          =   900
            Left            =   648
            TabIndex        =   10
            ToolTipText     =   "125Hz"
            Top             =   72
            Width           =   156
            _ExtentX        =   275
            _ExtentY        =   1588
            _Version        =   393216
            Appearance      =   0
            Orientation     =   1
            Scrolling       =   1
         End
         Begin MSComctlLib.ProgressBar ProgressBar2 
            Height          =   900
            Left            =   360
            TabIndex        =   11
            ToolTipText     =   "250Hz"
            Top             =   72
            Width           =   156
            _ExtentX        =   275
            _ExtentY        =   1588
            _Version        =   393216
            Appearance      =   0
            Orientation     =   1
            Scrolling       =   1
         End
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   900
            Left            =   72
            TabIndex        =   12
            ToolTipText     =   "500Hz"
            Top             =   72
            Width           =   156
            _ExtentX        =   275
            _ExtentY        =   1588
            _Version        =   393216
            Appearance      =   0
            Orientation     =   1
            Scrolling       =   1
         End
         Begin VB.Line Line5 
            Index           =   2
            X1              =   0
            X2              =   2880
            Y1              =   950
            Y2              =   950
         End
         Begin VB.Line Line5 
            Index           =   1
            X1              =   0
            X2              =   2880
            Y1              =   504
            Y2              =   504
         End
         Begin VB.Line Line5 
            Index           =   0
            X1              =   0
            X2              =   2880
            Y1              =   72
            Y2              =   72
         End
      End
      Begin VB.Image AboutButton 
         Height          =   175
         Left            =   144
         Stretch         =   -1  'True
         Top             =   216
         Width           =   175
      End
      Begin VB.Image CloseButton 
         Height          =   175
         Left            =   3024
         Stretch         =   -1  'True
         ToolTipText     =   "Exit"
         Top             =   216
         Width           =   175
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00400000&
         BorderWidth     =   3
         Index           =   2
         X1              =   72
         X2              =   3312
         Y1              =   2016
         Y2              =   2016
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00400000&
         BorderWidth     =   3
         Index           =   1
         X1              =   72
         X2              =   3312
         Y1              =   504
         Y2              =   504
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   2
         X1              =   3240
         X2              =   3240
         Y1              =   648
         Y2              =   1872
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00404040&
         BorderWidth     =   3
         Index           =   1
         X1              =   132
         X2              =   132
         Y1              =   648
         Y2              =   1872
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   144
         X2              =   3240
         Y1              =   1872
         Y2              =   1872
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00400000&
         BorderWidth     =   3
         Index           =   0
         X1              =   72
         X2              =   3312
         Y1              =   72
         Y2              =   72
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00404040&
         BorderWidth     =   3
         Index           =   0
         X1              =   144
         X2              =   3240
         Y1              =   648
         Y2              =   648
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "Graphic Equalizer"
         BeginProperty Font 
            Name            =   "Ruach LET"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   504
         TabIndex        =   1
         Top             =   144
         Width           =   2460
      End
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00808080&
      BorderWidth     =   6
      X1              =   0
      X2              =   4104
      Y1              =   2465
      Y2              =   2465
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   8
      Index           =   3
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   2376
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00808080&
      BorderWidth     =   8
      Index           =   2
      X1              =   3744
      X2              =   3744
      Y1              =   72
      Y2              =   2376
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   6
      Index           =   0
      X1              =   0
      X2              =   4176
      Y1              =   40
      Y2              =   50
   End
End
Attribute VB_Name = "VisualEQ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim hmixer As Long                      ' mixer handle
Dim inputVolCtrl As MIXERCONTROL        ' waveout volume control
Dim outputVolCtrl As MIXERCONTROL       ' microphone volume control
Dim rc As Long                          ' return code
Dim OK As Boolean                       ' boolean return code
Dim mxcd As MIXERCONTROLDETAILS         ' control info
Dim vol As MIXERCONTROLDETAILS_SIGNED   ' control's signed value
Dim volume As Long                      ' volume value
Dim volHmem As Long                     ' Volume Buffer
Private VU As VULights                  ' Volume Unit Values
Private FreqNum As Frequency
Private Declare Sub ReleaseCapture Lib "user32" ()
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2

Private Sub VolVal(VolIs As Long, VolFreq As Double)
For FreqNum = 0 To 9
Next FreqNum
VolIs = volume * 327.67
VolFreq = VU.Freq(FreqNum)
VU.FreqVal = VolIs * VolFreq
End Sub

Private Sub LightsA()
' ProgressBar1
FreqNum = Freq500Hz
For VU.InOutLev = CDbl(VU.VolLev * 0.2) To FreqNum
Next VU.InOutLev
ProgressBar1.Value = VU.InOutLev
End Sub

Private Sub LightsB()
' ProgressBar2
FreqNum = Freq250Hz
For VU.InOutLev = CDbl(VU.VolLev * 0.4) To FreqNum
Next VU.InOutLev
ProgressBar2.Value = VU.InOutLev
End Sub

Private Sub LightsC()
' ProgressBar3
FreqNum = Freq125Hz
For VU.InOutLev = CDbl(VU.VolLev * 0.8) To FreqNum
Next VU.InOutLev
ProgressBar3.Value = VU.InOutLev
End Sub

Private Sub LightsD()
' ProgressBar4
FreqNum = Freq62Hz
For VU.InOutLev = CDbl(VU.VolLev * 1.61290322580645E-02) To FreqNum
Next VU.InOutLev
ProgressBar4.Value = VU.InOutLev
End Sub
Private Sub LightsE()
' ProgressBar5
FreqNum = Freq31Hz
For VU.InOutLev = CDbl(VU.VolLev * 0.032258064516129) To FreqNum
Next VU.InOutLev
ProgressBar5.Value = VU.InOutLev
End Sub

Private Sub LightsF()
' ProgressBar6
FreqNum = Freq1kHz
For VU.InOutLev = CDbl(VU.VolLev * 0.01) To FreqNum
Next VU.InOutLev
ProgressBar6.Value = VU.InOutLev
End Sub

Private Sub LightsG()
' ProgressBar7
FreqNum = Freq3kHz
For VU.InOutLev = CDbl(VU.VolLev * 0.03) To FreqNum
Next VU.InOutLev
ProgressBar7.Value = VU.InOutLev
End Sub

Private Sub LightsH()
' ProgressBar8
FreqNum = Freq6kHz
For VU.InOutLev = CDbl(VU.VolLev * 0.06) To FreqNum
Next VU.InOutLev
ProgressBar8.Value = VU.InOutLev
End Sub

Private Sub LightsI()
' ProgressBar9
FreqNum = Freq12kHz
For VU.InOutLev = CDbl(VU.VolLev * 0.12) To FreqNum
Next VU.InOutLev
ProgressBar9.Value = VU.InOutLev
End Sub

Private Sub LightsJ()
' ProgressBar10
FreqNum = Freq16kHz
For VU.InOutLev = CDbl(VU.VolLev * 0.16) To FreqNum
Next VU.InOutLev
ProgressBar10.Value = VU.InOutLev
End Sub



Private Sub AboutButton_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    AboutButton.Picture = PictureClip2.GraphicCell(0)
End Sub

Private Sub AboutButton_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    AboutButton.Picture = PictureClip2.GraphicCell(2)
'    frmAbout.Show
End Sub

Private Sub CloseButton_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    CloseButton.Picture = PictureClip1.GraphicCell(1)
End Sub

Private Sub CloseButton_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    CloseButton.Picture = PictureClip1.GraphicCell(0)
    End
End Sub

Private Sub Form_Load()
    PictureClip1.Cols = 2
    PictureClip1.Rows = 1
    CloseButton.Picture = PictureClip1.GraphicCell(0)
    AboutButton.Picture = PictureClip2.GraphicCell(2)
    Timer1.Interval = 6.25
   ' Open the mixer specified by DEVICEID
   rc = mixerOpen(hmixer, DEVICEID, 0, 0, 0)
   If ((MMSYSERR_NOERROR <> rc)) Then
       MsgBox "Couldn't open the mixer."
       Exit Sub
   End If
   ' Get the output volume meter
   OK = GetControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT, MIXERCONTROL_CONTROLTYPE_PEAKMETER, outputVolCtrl)
   If (OK = True) Then
   ' Set frequencies for the Volume Units
    ProgressBar1.Max = Frequency.Freq500Hz + 1
    ProgressBar1.Min = Frequency.Freq500Hz
    ProgressBar2.Max = Frequency.Freq250Hz + 1
    ProgressBar2.Min = Frequency.Freq250Hz
    ProgressBar3.Max = Frequency.Freq125Hz + 1
    ProgressBar3.Min = Frequency.Freq125Hz
    ProgressBar4.Max = Frequency.Freq62Hz + 1
    ProgressBar4.Min = Frequency.Freq62Hz
    ProgressBar5.Max = Frequency.Freq31Hz + 1
    ProgressBar5.Min = Frequency.Freq31Hz
    ProgressBar6.Max = Frequency.Freq1kHz + 1
    ProgressBar6.Min = Frequency.Freq1kHz
    ProgressBar7.Max = Frequency.Freq3kHz + 1
    ProgressBar7.Min = Frequency.Freq3kHz
    ProgressBar8.Max = Frequency.Freq6kHz + 1
    ProgressBar8.Min = Frequency.Freq6kHz
    ProgressBar9.Max = Frequency.Freq12kHz + 1
    ProgressBar9.Min = Frequency.Freq12kHz
    ProgressBar10.Max = Frequency.Freq16kHz + 1
    ProgressBar10.Min = Frequency.Freq16kHz
   Else
      MsgBox "Couldn't get waveout meter"
   End If
   ' Initialize mixercontrol structure
   mxcd.cbStruct = Len(mxcd)
   volHmem = GlobalAlloc(&H0, Len(volume))  ' Allocate a buffer for the volume value
   mxcd.paDetails = GlobalLock(volHmem)
   mxcd.cbDetails = Len(volume)
   mxcd.cChannels = 1
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim Ret&
    ReleaseCapture
    Ret& = SendMessage(Me.hwnd, &H112, &HF012, 0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If (fRecording = True) Then
       StopInput
   End If
   GlobalFree volHmem
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim Ret&
    ReleaseCapture
    Ret& = SendMessage(Me.hwnd, &H112, &HF012, 0)
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim Ret&
    ReleaseCapture
    Ret& = SendMessage(Me.hwnd, &H112, &HF012, 0)
End Sub

Private Sub Timer1_Timer()
    VU.VolLev = volume / 327.67
    If (volume < 0) Then volume = -volume
    ' Get the current output level
    If (1 = 1) Then
    mxcd.dwControlID = outputVolCtrl.dwControlID
    mxcd.item = outputVolCtrl.cMultipleItems
    rc = mixerGetControlDetails(hmixer, mxcd, MIXER_GETCONTROLDETAILSF_VALUE)
    CopyStructFromPtr volume, mxcd.paDetails, Len(volume)
    If (volume < 0) Then volume = -volume
    End If
    ActivateVolumeUnits
End Sub

Private Sub ActivateVolumeUnits()
    LightsA
    LightsB
    LightsC
    LightsD
    LightsE
    LightsF
    LightsG
    LightsH
    LightsI
    LightsJ
End Sub



