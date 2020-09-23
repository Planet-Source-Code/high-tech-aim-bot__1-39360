VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   2085
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   990
      Width           =   4560
   End
   Begin VB.CommandButton Command1 
      Caption         =   "On"
      Height          =   330
      Left            =   2250
      TabIndex        =   2
      Top             =   180
      Width           =   1500
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   700
      Left            =   135
      Top             =   1080
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   495
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   585
      Width           =   3480
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   495
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   135
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'code by High Tech
'aim: telnet guru
'email: techx@mailwire.net
'website a) http://customsoftware.cjb.net
'website b) http://heyyouvisitme.cjb.net

Private Phrase(10) As String

Private Sub Command1_Click()
If Command1.Caption = "On" Then
Command1.Caption = "Off"
Timer1.Enabled = True
Else
Command1.Caption = "On"
Timer1.Enabled = False
End If
End Sub

Private Sub Text2_Change()
If Not Text1 = "TelnetGuru" Then
Text3 = Text3 & Text1 & ":" & Text2 & vbCrLf
If InStr(Text2, "what") Or InStr(Text2, "why") Or InStr(Text2, "how") And InStr(Text2, "?") Then
        Phrase(0) = "You always ask :-/"
        Phrase(1) = "Did you have to ask?"
        Phrase(2) = "Because."
        Phrase(3) = "Wouldnt you like to know!"
        Phrase(4) = "Dunno."
        Phrase(5) = "uhh wha?"
        Phrase(6) = "beats me!"
        Phrase(7) = "Well, Do tell you the truth, I dont know. ;-)"
    ElseIf Right(Text2, 1) = "." And Len(Text2) > 2 Then
        Phrase(0) = "Oh yeah?"
        Phrase(1) = "Why?"
        Phrase(2) = "I agree."
        Phrase(3) = "wtf"
        Phrase(4) = "huh?"
        Phrase(5) = "High Tech's Bot is a pimp ;-)"
        Phrase(6) = "hmmm"
        Phrase(7) = "Hm, Is that fact?"
    ElseIf Text2 = "hehe" Or Text2 = "lol" Or Text2 = "rofl" Or Text2 = "lmao" Then
        Phrase(0) = "LoL, that is kinda funny."
        Phrase(1) = "ROFL"
        Phrase(2) = "hehe"
        Phrase(3) = "what so funny?"
        Phrase(4) = ":-d Yay"
        Phrase(5) = "*laughs too*"
        Phrase(6) = "I'm glad your amused. ;-)"
        Phrase(7) = "=-o!!! Funny!"
    ElseIf Text2 = "hi" Or Text2 = "hello" Or Text2 = "greetings." Then
        Phrase(0) = "Hello! How are you?"
        Phrase(1) = "Greetings!"
        Phrase(2) = "Hello " & Text1 & ", nice to meet you"
        Phrase(3) = "Hey " & Text1 & ", whats up?"
        Phrase(4) = "*drum roll* Its " & Text1 & "!"
        Phrase(5) = "hihi"
        Phrase(6) = "whats up?"
        Phrase(7) = "Hey, How are you? The time is " & Time
    ElseIf InStr(Text2, "and") Then
        Phrase(0) = Right(Text2, Len(Text2) - InStr(Text2, "and ")) & "?"
        Phrase(1) = "Why " & Right(Text2, Len(Text2) - InStr(Text2, "and ")) & "?"
        Phrase(2) = "*yawns* thats fun"
        Phrase(3) = "You think too much."
        Phrase(4) = "is that so?"
        Phrase(5) = "Yep. ;-)"
        Phrase(6) = "noo way!"
        Phrase(7) = "serious?"
    ElseIf InStr(Text2, "are you") Then
        Phrase(0) = "No, but you are " & Right(Text2, Len(Text2) - InStr(Text2, "are you")) & "."
        Phrase(1) = "mind your own business ;-)"
        Phrase(2) = "Haha, Only if you are (and maybe not then either)."
        Phrase(3) = "Possibly."
        Phrase(4) = "Nope."
        Phrase(5) = "Yep. ;-)"
        Phrase(6) = "I can be anything for a dollar :-)"
        Phrase(7) = "Only if santa can really fit down a chimney."
    Else
        Phrase(0) = "Well, thats good =)"
        Phrase(1) = "I know."
        Phrase(2) = "Oh yeah?!"
        Phrase(3) = "*laughs*"
        Phrase(4) = Text2 & "?"
        Phrase(5) = "Cool"
        Phrase(6) = "so what's new?"
        Phrase(7) = "w00tage"
    End If

Randomize Timer
DoEvents
Werd = "(B) " & Phrase(Int(Rnd * 7))
ReturnIM Werd
Text3 = Text3 & Werd & vbCrLf
'IM_Close
End If
End Sub

Private Sub Text3_Change()
Text3.SelStart = Len(Text3)
End Sub

Private Sub Timer1_Timer()
Text1 = IM_GetLastSN
Text2 = LCase(IM_GetLastWords)
End Sub
