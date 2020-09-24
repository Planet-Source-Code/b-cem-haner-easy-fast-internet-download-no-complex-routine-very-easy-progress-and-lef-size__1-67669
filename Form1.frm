VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "Easy fast download routine"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7080
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   7080
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "CANCEL"
      Enabled         =   0   'False
      Height          =   555
      Left            =   3465
      TabIndex        =   10
      Top             =   2340
      Width           =   1755
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   315
      Left            =   3180
      TabIndex        =   9
      Top             =   870
      Width           =   3675
      _ExtentX        =   6482
      _ExtentY        =   556
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.TextBox Text1 
      Height          =   345
      Left            =   240
      TabIndex        =   1
      Text            =   "http://download.microsoft.com/download/vb60pro/install/6/win98me/en-us/vbrun60.exe"
      Top             =   315
      Width           =   6615
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   375
      Top             =   2190
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "DOWNLOAD FILE"
      Height          =   555
      Left            =   1590
      TabIndex        =   0
      Top             =   2340
      Width           =   1755
   End
   Begin VB.Label HedLab 
      Caption         =   "Left size : "
      Height          =   225
      Index           =   2
      Left            =   240
      TabIndex        =   8
      Top             =   1500
      Width           =   1230
   End
   Begin VB.Label StatLab 
      Caption         =   "%"
      Height          =   225
      Index           =   3
      Left            =   1620
      TabIndex        =   7
      Top             =   1875
      Width           =   915
   End
   Begin VB.Label StatLab 
      Caption         =   "[left]"
      Height          =   225
      Index           =   2
      Left            =   1620
      TabIndex        =   6
      Top             =   1515
      Width           =   915
   End
   Begin VB.Label StatLab 
      Caption         =   "[current]"
      Height          =   225
      Index           =   1
      Left            =   1635
      TabIndex        =   5
      Top             =   1185
      Width           =   915
   End
   Begin VB.Label StatLab 
      Caption         =   "[total]"
      Height          =   225
      Index           =   0
      Left            =   1635
      TabIndex        =   4
      Top             =   870
      Width           =   915
   End
   Begin VB.Label HedLab 
      Caption         =   "Current size : "
      Height          =   225
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   1170
      Width           =   1230
   End
   Begin VB.Label HedLab 
      Caption         =   "File size :"
      Height          =   225
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   855
      Width           =   915
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Size As Long, Remaining As Long, FFile As Integer, Chunk() As Byte
Dim FileName As String, NowSize As Long, Yuzde As Integer
    FileName = Right(Text1.Text, Len(Text1.Text) - InStrRev(Text1.Text, "/"))
    Inet1.Execute Trim(Text1.Text), "GET"
    Do While Inet1.StillExecuting
        DoEvents
    Loop
    ProgressBar1.Max = 100: Command2.Enabled = True
    Size = CLng(Inet1.GetHeader("Content-Length"))
    Remaining = Size
    NowSize = 0
    StatLab(0).Caption = Size
    FFile = FreeFile
    Open "c:\" & FileName For Binary Access Write As #FFile
    Do Until Remaining = 0
        If Form1.Tag = "cancel" Then Inet1.Cancel: MsgBox "file download aborted": End
        If Remaining > 1024 Then
            Chunk = Inet1.GetChunk(1024, icByteArray)
            Remaining = Remaining - 1024
        Else
            Chunk = Inet1.GetChunk(Remaining, icByteArray)
            Remaining = 0
        End If
        NowSize = Size - Remaining
        Yuzde = CInt((100 / Size) * NowSize)
        StatLab(1).Caption = NowSize
        StatLab(2).Caption = Size - NowSize
        StatLab(3).Caption = "% " & Yuzde
        ProgressBar1.Value = Yuzde
        Put #FFile, , Chunk
    Loop
    Close #FFile
    MsgBox "File downloaded"
End Sub

Private Sub Command2_Click()
Form1.Tag = "cancel"
End Sub
