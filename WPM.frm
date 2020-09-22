VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3540
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5430
   LinkTopic       =   "Form1"
   ScaleHeight     =   3540
   ScaleWidth      =   5430
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TestTxt 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3120
      Left            =   15
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   375
      Width           =   5385
   End
   Begin VB.Timer MinTimer 
      Left            =   0
      Top             =   45
   End
   Begin VB.Label TPrompt 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   15
      TabIndex        =   1
      Tag             =   "1"
      Top             =   15
      Width           =   5385
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private SpCount As Integer, MinFlag As Boolean
Private Sub Form_Load()
Dim msg$
MinTimer.Interval = 0 ' Turn off the timer.
msg$ = "Type the lines as they appears and press return "
msg$ = msg$ & "when you are finished typing each line."
msg$ = msg$ & Chr(13) & Chr(10)
msg$ = msg$ & "Continue to type until the results appear. Type "
msg$ = msg$ & "as quickly and accuratly "
msg$ = msg$ & Chr(13) & Chr(10)
msg$ = msg$ & "as possible."
MsgBox msg$, vbInformation, "Instructions"
With TPrompt
 .Caption = "The sly fox darted quickly through the "
 .Caption = .Caption & "woods to"
End With
End Sub

Private Sub MinTimer_Timer()
MinTimer.Interval = 0
MinFlag = True
End Sub

Private Sub TestTxt_Change()
If MinTimer.Interval = 0 Then
MinTimer.Interval = 60000
End If
End Sub


Private Sub TestTxt_KeyPress(KeyAscii As Integer)
If KeyAscii = 32 And MinFlag = False Then SpCount = SpCount + 1
If KeyAscii = 13 And MinFlag = False Then
KeyAscii = Asc(" ")
Select Case TPrompt.Tag
Case 1
TPrompt.Caption = "grandma's house, so he could try"
TPrompt.Caption = TPrompt.Caption & " her cookies."
Case 2
TPrompt.Caption = "Her grandson was only 8. He "
TPrompt.Caption = TPrompt.Caption & "doesn't know how to steal."
Case 3
TPrompt.Caption = "Grandma gave the poor fox "
TPrompt.Caption = TPrompt.Caption & "a cookie."
End Select
TPrompt.Tag = TPrompt.Tag + 1
End If
If KeyAscii = 13 And MinFlag = True Then
If SpCount < 10 Then lev$ = "Please try again!"
If SpCount > 10 And SpCount < 40 Then lev$ = "Good Job, you're in the average range!"
If SpCount > 40 Then lev$ = "Excelent work, you came above average!"
a$ = "The sly fox darted quickly through the woods to"
b$ = "grandma's house, so he could try her cookies."
c$ = "Her grandson was only 8. He doesn't know how to steal."
d$ = "Grandma gave the poor fox a cookie."
If TPrompt.Tag >= 2 Then Key$ = a$ & " " & b$
If TPrompt.Tag >= 3 Then Key$ = Key$ & " " & c$
If TPrompt.Tag = 4 Then Key$ = Key$ & " " & d$
If Len(TestTxt.Text) < Len(Key$) Then
erlett$ = "Incomplete"
Else
For i% = 1 To Len(Key$)
If Mid(TestTxt.Text, i%, 1) <> Mid(Key$, i%, 1) Then _
ernumm% = ernumm% + 1
Next i%
End If
msg$ = "Congratulations! You're done!"
msg$ = msg$ & Chr(13) & Chr(10)
msg$ = msg$ & Chr(13) & Chr(10)
msg$ = msg$ & lev$
msg$ = msg$ & Chr(13) & Chr(10)
msg$ = msg$ & "WPM:" & Str(SpCount)
msg$ = msg$ & Chr(13) & Chr(10)
msg$ = msg$ & "ERR: " & IIf(erlett$ = "", ernumm%, erlett$)
MsgBox msg$, , "Results"
End If
End Sub
