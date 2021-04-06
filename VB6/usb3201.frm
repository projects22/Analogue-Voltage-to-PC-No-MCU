VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "USB MCP3201 ADC "
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4980
   Icon            =   "usb3201.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   4980
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   4080
      TabIndex        =   8
      Text            =   "5"
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "START"
      Height          =   375
      Left            =   3480
      TabIndex        =   7
      Top             =   3600
      Width           =   1215
   End
   Begin MSCommLib.MSComm Comm1 
      Left            =   4200
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Frame Frame2 
      Caption         =   "Sampling"
      Height          =   1095
      Left            =   1800
      TabIndex        =   4
      Top             =   2880
      Width           =   1335
      Begin VB.OptionButton Option5 
         Caption         =   "0.1 Sec"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton Option4 
         Caption         =   "1 Sec"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Conversion"
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   2880
      Width           =   1335
      Begin VB.OptionButton Option2 
         Caption         =   "10 Bits"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "12 Bits"
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   360
      Top             =   2160
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000014&
      Caption         =   "0.000 V"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1140
      Left            =   105
      Top             =   1560
      Width           =   4800
   End
   Begin VB.Label Label2 
      Caption         =   "COM"
      Height          =   255
      Left            =   3480
      TabIndex        =   9
      Top             =   3000
      Width           =   495
   End
   Begin VB.Line Line7 
      X1              =   4200
      X2              =   4680
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line6 
      X1              =   3480
      X2              =   3840
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line5 
      X1              =   2760
      X2              =   3240
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line4 
      X1              =   2040
      X2              =   2400
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line3 
      X1              =   1320
      X2              =   1680
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line2 
      X1              =   480
      X2              =   840
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   2655
      Left            =   45
      Top             =   45
      Width           =   4905
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   500
      X2              =   2500
      Y1              =   2500
      Y2              =   2500
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Analogue to PC using ADC MCP3201 driven by serial to USB adapter
'Code by moty22.co.uk

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim bit, word, bitcount, dec  As Integer
Dim t, angl As Double

Private Sub Form_Load()
Timer1.Enabled = False

'analoge meter is constructed of 7 lines, line1 is the pointer. These instructions
'are for positioning the lines.
Line1.ZOrder (x)
Shape2.ZOrder (x)
Label5.ZOrder (z)
Line1.X2 = 2500
Line1.Y2 = 2500
Line1.X1 = 2500 - 2000 * Cos(40 * 3.14 / 180)      '0
Line1.Y1 = 2500 - 2000 * Sin(40 * 3.14 / 180)
Line2.X1 = 2500 - 2000 * Cos(40 * 3.14 / 180) 'Line star point
Line2.Y1 = 2500 - 2000 * Sin(40 * 3.14 / 180)
Line2.X2 = 2500 - 1800 * Cos(40 * 3.14 / 180) 'Line end point
Line2.Y2 = 2500 - 1800 * Sin(40 * 3.14 / 180)
Line3.X1 = 2500 - 2000 * Cos(140 * 3.14 / 180)
Line3.Y1 = 2500 - 2000 * Sin(140 * 3.14 / 180)
Line3.X2 = 2500 - 1800 * Cos(140 * 3.14 / 180)
Line3.Y2 = 2500 - 1800 * Sin(140 * 3.14 / 180)
Line4.X1 = 2500 - 2000 * Cos(60 * 3.14 / 180)
Line4.Y1 = 2500 - 2000 * Sin(60 * 3.14 / 180)
Line4.X2 = 2500 - 1800 * Cos(60 * 3.14 / 180)
Line4.Y2 = 2500 - 1800 * Sin(60 * 3.14 / 180)
Line5.X1 = 2500 - 2000 * Cos(80 * 3.14 / 180)
Line5.Y1 = 2500 - 2000 * Sin(80 * 3.14 / 180)
Line5.X2 = 2500 - 1800 * Cos(80 * 3.14 / 180)
Line5.Y2 = 2500 - 1800 * Sin(80 * 3.14 / 180)
Line6.X1 = 2500 - 2000 * Cos(100 * 3.14 / 180)
Line6.Y1 = 2500 - 2000 * Sin(100 * 3.14 / 180)
Line6.X2 = 2500 - 1800 * Cos(100 * 3.14 / 180)
Line6.Y2 = 2500 - 1800 * Sin(100 * 3.14 / 180)
Line7.X1 = 2500 - 2000 * Cos(120 * 3.14 / 180)
Line7.Y1 = 2500 - 2000 * Sin(120 * 3.14 / 180)
Line7.X2 = 2500 - 1800 * Cos(120 * 3.14 / 180)
Line7.Y2 = 2500 - 1800 * Sin(120 * 3.14 / 180)
End Sub
Private Sub Command1_Click()
If Comm1.PortOpen = True Then Comm1.PortOpen = False
Comm1.CommPort = Text1.Text
Comm1.Settings = "1200,n,8,1"
Comm1.PortOpen = True
Comm1.RTSEnable = False 'Switch on 5V for the ADC
Comm1.DTREnable = True 'ADC Clock low starts the convertion
Timer1.Enabled = True

End Sub
Private Sub Timer1_Timer()  'Timer 1 defines sampling rate.

If Option4 Then          'Selecting sampling rate
Timer1.Interval = 1000  '1 second sampling
ElseIf Option5 Then
Timer1.Interval = 100
End If

word = 0              'word is the collection of the ADC serial bits, starting with the MSB.
If Option1 Then       'Selecting number of bits to read.
bitcount = 11          ' Read 11 bits
dec = 4096             'When reading 8 bits full scale is 256 bits
ElseIf Option2 Then
bitcount = 9
dec = 1024
End If

'reading the bits
Comm1.RTSEnable = True    'CS low Enables ADC
Sleep 1
For i = 2 To 0 Step -1   'First 3 ADC Clocks are a flag for start readings.
Comm1.DTREnable = False
clock
Comm1.DTREnable = True
clock
Next

strt:                  'Loop for reading serial bits, reads MSB first, for 8 bits loops 8 times.
Comm1.DTREnable = False     'ADC Clock high passes the bit to Data output.
clock
bit = 0                'If bit is 0 then it adds 0 to word
If Comm1.CTSHolding = False Then
bit = 2 ^ bitcount     'If bit is 1 then it adds to word: 2 power of the order number of the bit.
End If
word = word + bit
Comm1.DTREnable = True     'ADC Clock low.
clock
bitcount = bitcount - 1
If bitcount > 0 Then    'Loop
GoTo strt
End If
Comm1.RTSEnable = False    'CS high disables ADC

'setting the meters
t = word / dec * 5  't is the digital meter reading: ADC reading divided by full scale bits reading multiplied by input voltage.
Label5.Caption = Format(t, "0.000") & " V"
angl = (40 + t * 20) * 3.1415 / 180  'calculate the angle of the analogue meter needle.
Line1.X1 = 2500 - 2000 * Cos(angl)
Line1.Y1 = 2500 - 2000 * Sin(angl)
End Sub

'setting the clock frequency,
'readings are more stable with clock at 2KHz or less.
Private Sub clock()
For i = 4000 To 0 Step -1
Next
End Sub
Private Sub Form_Unload(Cancel As Integer)
If Comm1.PortOpen = True Then   'port must be closed
Comm1.PortOpen = False
End If
End Sub
