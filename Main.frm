VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form Form1 
   Caption         =   "Bandwidth Meter v2.3"
   ClientHeight    =   4905
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   ScaleHeight     =   86.519
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   89.429
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.Slider Slider1 
      Height          =   255
      Left            =   1920
      TabIndex        =   16
      Top             =   0
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Min             =   40
      Max             =   255
      SelStart        =   255
      TickFrequency   =   40
      Value           =   255
   End
   Begin VB.Timer Timer5 
      Interval        =   3000
      Left            =   0
      Top             =   5280
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   700
      Left            =   960
      Top             =   5280
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   480
      Top             =   5280
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1440
      Top             =   5280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":015C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":02B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0414
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   1695
      Left            =   0
      OleObjectBlob   =   "Main.frx":0570
      TabIndex        =   17
      Top             =   1440
      Width           =   5055
   End
   Begin MSChart20Lib.MSChart MSChart2 
      Height          =   1815
      Left            =   0
      OleObjectBlob   =   "Main.frx":2D21
      TabIndex        =   18
      Top             =   3360
      Width           =   5055
   End
   Begin VB.Label Label12 
      Caption         =   "kBps"
      Height          =   255
      Left            =   4680
      TabIndex        =   21
      Top             =   3120
      Width           =   375
   End
   Begin VB.Label Label11 
      Caption         =   "kBps"
      Height          =   255
      Left            =   4680
      TabIndex        =   20
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   840
      TabIndex        =   19
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label7 
      Caption         =   "Incoming:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label9 
      Caption         =   "kBps"
      Height          =   255
      Left            =   1560
      TabIndex        =   8
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label RecordOutgoing 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   3960
      TabIndex        =   15
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label Label17 
      Caption         =   "Record Outgoing:"
      Height          =   255
      Left            =   2640
      TabIndex        =   14
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label RecordIncoming 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   3960
      TabIndex        =   13
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label16 
      Caption         =   "Record Incoming:"
      Height          =   255
      Left            =   2640
      TabIndex        =   12
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label10 
      Caption         =   "kBps"
      Height          =   255
      Left            =   1560
      TabIndex        =   11
      Top             =   3120
      Width           =   375
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   840
      TabIndex        =   10
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "Outgoing:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   3120
      Width           =   735
   End
   Begin VB.Line Line11 
      X1              =   2.117
      X2              =   84.667
      Y1              =   19.05
      Y2              =   19.05
   End
   Begin VB.Line Line10 
      X1              =   2.117
      X2              =   52.917
      Y1              =   16.933
      Y2              =   16.933
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Bytes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2520
      TabIndex        =   6
      Top             =   465
      Width           =   495
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Bytes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2520
      TabIndex        =   5
      Top             =   225
      Width           =   495
   End
   Begin VB.Line Line9 
      X1              =   16.933
      X2              =   16.933
      Y1              =   4.233
      Y2              =   12.7
   End
   Begin VB.Line Line8 
      X1              =   2.117
      X2              =   52.917
      Y1              =   8.467
      Y2              =   8.467
   End
   Begin VB.Line Line7 
      X1              =   2.117
      X2              =   52.917
      Y1              =   12.7
      Y2              =   12.7
   End
   Begin VB.Line Line6 
      X1              =   52.917
      X2              =   52.917
      Y1              =   4.233
      Y2              =   16.933
   End
   Begin VB.Line Line5 
      X1              =   31.75
      X2              =   52.917
      Y1              =   4.233
      Y2              =   4.233
   End
   Begin VB.Line Line4 
      X1              =   33.867
      X2              =   33.867
      Y1              =   0
      Y2              =   4.233
   End
   Begin VB.Line Line3 
      X1              =   2.117
      X2              =   33.867
      Y1              =   4.233
      Y2              =   4.233
   End
   Begin VB.Line Line2 
      X1              =   2.117
      X2              =   33.867
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      X1              =   2.117
      X2              =   2.117
      Y1              =   0
      Y2              =   16.933
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Received"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   -120
      TabIndex        =   4
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Sent"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   -120
      TabIndex        =   3
      Tag             =   "0"
      Top             =   240
      Width           =   975
   End
   Begin VB.Label lblType 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Transparency Control:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   1815
   End
   Begin VB.Label lblSent 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Tag             =   "0"
      Top             =   225
      Width           =   1335
   End
   Begin VB.Label lblRecv 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1080
      TabIndex        =   0
      Tag             =   "0"
      Top             =   465
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDOWN = &H201     'Button down
Private Const WM_LBUTTONUP = &H202       'Button up
Private Const WM_LBUTTONDBLCLK = &H203   'Double-click
Private Const WM_RBUTTONDOWN = &H204     'Button down
Private Const WM_RBUTTONUP = &H205       'Button up
Private Const WM_RBUTTONDBLCLK = &H206   'Double-click
Private Const MAXLEN_IFDESCR = 256
Private Const MAXLEN_PHYSADDR = 8
Private Const MAX_INTERFACE_NAME_LEN = 256
Private nid As NOTIFYICONDATA
Private m_objIpHelper As CIpHelper

Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_ALPHA = &H2

Private Declare Function GetWindowLong Lib "user32" _
  Alias "GetWindowLongA" (ByVal hwnd As Long, _
  ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong Lib "user32" _
   Alias "SetWindowLongA" (ByVal hwnd As Long, _
   ByVal nIndex As Long, ByVal dwNewLong As Long) _
   As Long

Private Declare Function SetLayeredWindowAttributes Lib _
    "user32" (ByVal hwnd As Long, ByVal crKey As Long, _
    ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Dim oOld As Long
Dim oNew As Long
Dim aOld As Long
Dim aNew As Long
Dim i As Long
Dim Incoming As Long
Dim Outgoing As Long
Dim temp1 As Long
Dim r1 As Single
Dim r2 As Single

Dim chartIndex(1 To 200)
Dim chart2Index(1 To 200)
Dim x As Long
Dim x1 As Long
Dim x2 As Long
Dim x3 As Long
Dim startPause As Boolean

Dim objInterface2 As CInterface
Dim obJHelper As CInterface
Dim tValue As Long
Dim aValue As Long


Private Sub Form_Load()
    SetOnTop Me.hwnd, True
    TranslucentForm Me, 255
    
    x = 1
    x1 = 1
    x2 = 2
    x3 = 1
    startPause = False
    
    '//Fill entire chart with zero's to give them data
    For x3 = 1 To 200
    chartIndex(x3) = 0
    Next x3
    MSChart1.ChartData = chartIndex
    MSChart2.ChartData = chartIndex
    
    '//Set all 200 to color red
    For x3 = 1 To 200
            With MSChart1.Plot.SeriesCollection(x3).DataPoints(-1)
              .Brush.FillColor.Red = 255
              .Brush.FillColor.Green = 0
              .Brush.FillColor.Blue = 0
            End With
            With MSChart2.Plot.SeriesCollection(x3).DataPoints(-1)
              .Brush.FillColor.Red = 30
              .Brush.FillColor.Green = 144
              .Brush.FillColor.Blue = 255
            End With
        Next x3
        
    Set m_objIpHelper = New CIpHelper
    With nid
        .cbSize = Len(nid)
        .hwnd = Me.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE
        .hIcon = ImageList1.ListImages(4).Picture
        nid.szTip = "Bytes received: " & Trim(Format(m_objIpHelper.BytesReceived, "###,###,###,###")) & vbCrLf & " Bytes sent: " & Trim(Format(m_objIpHelper.BytesSent, "###,###,###,###")) & vbNullChar
    End With
    
    Shell_NotifyIcon NIM_ADD, nid

r1 = 0
r2 = 0

End Sub

Private Sub UpdateInterfaceInfo()

Dim objInterface        As CInterface
Static st_objInterface  As CInterface
Static lngBytesRecv     As Long
Static lngBytesSent     As Long
Dim blnIsRecv           As Boolean
Dim blnIsSent           As Boolean

    If st_objInterface Is Nothing Then Set st_objInterface = New CInterface
    Set objInterface = m_objIpHelper.Interfaces(1)
   
    If Label1.Tag = 0 Then
        lblRecv.Tag = Trim(Format(m_objIpHelper.BytesReceived, "###,###,###,###"))
        lblSent.Tag = Trim(Format(m_objIpHelper.BytesSent, "###,###,###,###"))
        Label1.Tag = 1
    Else
        lblRecv.Caption = Trim(Format(m_objIpHelper.BytesReceived, "###,###,###,###")) - lblRecv.Tag
        lblSent.Caption = Trim(Format(m_objIpHelper.BytesSent, "###,###,###,###")) - lblSent.Tag
    End If

Set st_objInterface = objInterface

blnIsRecv = (m_objIpHelper.BytesReceived > lngBytesRecv)
blnIsSent = (m_objIpHelper.BytesSent > lngBytesSent)
    If blnIsRecv And blnIsSent Then
        nid.hIcon = ImageList1.ListImages(4).Picture
    ElseIf (Not blnIsRecv) And blnIsSent Then
        nid.hIcon = ImageList1.ListImages(3).Picture
    ElseIf blnIsRecv And (Not blnIsSent) Then
        nid.hIcon = ImageList1.ListImages(2).Picture
    ElseIf Not (blnIsRecv And blnIsSent) Then
        nid.hIcon = ImageList1.ListImages(1).Picture
    End If
    
lngBytesRecv = m_objIpHelper.BytesReceived
lngBytesSent = m_objIpHelper.BytesSent

nid.szTip = "Bytes received: " & Trim(Format(m_objIpHelper.BytesReceived, "###,###,###,###")) & vbCrLf & " Bytes sent: " & Trim(Format(m_objIpHelper.BytesSent, "###,###,###,###")) & vbNullChar
Shell_NotifyIcon NIM_MODIFY, nid

'##############################################
Set objInterface2 = New CInterface
Set obJHelper = m_objIpHelper.Interfaces(1)

oNew = Trim(Format(m_objIpHelper.BytesReceived, "###,###,###,###")) '//give bandwidth
aNew = Trim(Format(m_objIpHelper.BytesSent, "###,###,###,###")) '//give bandwidth

tValue = oNew - oOld
aValue = aNew - aOld

Label8.Caption = (tValue / 1000)
Label6.Caption = (aValue / 1000)

oOld = oNew
aOld = aNew

 '//********MSChart control*********
 
 'kbIndex, chartIndex, x, x1, x2, x3
 'ProgressBar1.Value
 
If x <= 199 Then
    x = x + 1
    chartIndex(x) = chartIndex(200)
    chart2Index(x) = chart2Index(200)
Else
    For x1 = 1 To 199
    chartIndex(x1) = chartIndex(x2)
    chart2Index(x1) = chart2Index(x2)
    x2 = x2 + 1
    Next x1
    x1 = 1
    x2 = 2
    
    chartIndex(x) = chartIndex(200)
    chart2Index(x) = chart2Index(200)
End If

If startPause = True Then
chartIndex(200) = (tValue / 1000)
chart2Index(200) = (aValue / 1000)
Else
startPause = True
End If

MSChart1.ChartData = chartIndex
MSChart2.ChartData = chart2Index
 '//********END MSChart control*********
 '###############################################

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Shell_NotifyIcon NIM_DELETE, nid
End Sub

Private Sub Form_Terminate()
Shell_NotifyIcon NIM_DELETE, nid
End Sub

Private Sub Form_Unload(Cancel As Integer)
Shell_NotifyIcon NIM_DELETE, nid
End Sub

Private Sub Slider1_Scroll()
TranslucentForm Me, Slider1.Value
End Sub

Private Sub Timer1_Timer()
Call UpdateInterfaceInfo
End Sub

Private Function bandwidth()

  
End Function

Private Sub Timer2_Timer()
bandwidth
End Sub

Private Sub Timer3_Timer()
If r1 < Label8.Caption Then
 r1 = Label8.Caption
 RecordIncoming.Caption = r1
 Else
End If

If r2 < Label6.Caption Then
 r2 = Label6.Caption
 RecordOutgoing.Caption = r2
 Else
End If

End Sub

Private Sub Timer5_Timer()
Timer3.Enabled = True
Timer5.Enabled = False
End Sub

Public Function TranslucentForm(frm As Form, TranslucenceLevel As Byte) As Boolean
SetWindowLong frm.hwnd, GWL_EXSTYLE, WS_EX_LAYERED
SetLayeredWindowAttributes frm.hwnd, 0, TranslucenceLevel, LWA_ALPHA
TranslucentForm = Err.LastDllError = 0
End Function
