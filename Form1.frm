VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Text Find"
   ClientHeight    =   3855
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7230
   LinkTopic       =   "Form1"
   ScaleHeight     =   3855
   ScaleWidth      =   7230
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Replace with"
      Height          =   495
      Left            =   1560
      TabIndex        =   11
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2640
      TabIndex        =   9
      Top             =   2760
      Width           =   2175
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   3120
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Load text file"
      Height          =   495
      Left            =   4440
      TabIndex        =   8
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   495
      Left            =   5880
      TabIndex        =   7
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clear"
      Height          =   495
      Left            =   3000
      TabIndex        =   6
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   4920
      TabIndex        =   3
      Top             =   2760
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   2760
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Find word or letter"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   3240
      Width           =   1215
   End
   Begin RichTextLib.RichTextBox rtb 
      Height          =   2175
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   3836
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"Form1.frx":0000
   End
   Begin VB.Label Label3 
      Caption         =   "Replace with"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   10
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Found or Replaced"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      TabIndex        =   5
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Search word or letter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim i As Long
Dim letter As String
Dim rep As String
Private Sub Command1_Click()
Dim front As Integer
Dim back As Integer
Dim counter As Long

letter = Text1.Text

i = InStr(1, rtb.Text, letter)

If letter <> "" Then

If Len(letter) > 1 And i = 1 Then
back = Asc(Mid(rtb.Text, i + Len(letter), 1))
If (back < 65 Or back > 90 And back < 97 Or back > 122) Then
rtb.SelStart = 0
rtb.SelLength = Len(letter)
rtb.SelColor = vbRed
counter = 1
End If
End If

i = InStr(Len(letter) + 1, rtb.Text, letter)

If Len(letter) > 1 And i > 1 Then

Do While i
If i = Len(rtb.Text) - Len(letter) + 1 Then
highlight
counter = counter + 1
Exit Do
End If
front = Asc(Mid(rtb.Text, i - 1, 1))
back = Asc(Mid(rtb.Text, i + Len(letter), 1))
If (front < 65 Or front > 90 And front < 97 Or front > 122) And (back < 65 Or back > 90 And back < 97 Or back > 122) Then
highlight
counter = counter + 1
End If
i = InStr(i + 1, rtb.Text, letter)
Loop

ElseIf Len(letter) = 1 Then
i = InStr(1, rtb.Text, letter)
Do While i
highlight
i = InStr(i + 1, rtb.Text, letter)
counter = counter + 1
Loop
End If

If counter = 0 Then
MsgBox Text1.Text & "  not found", vbInformation, "Findings"
Text1.Text = ""
Text1.SetFocus
End If

Else
MsgBox "Cant search for nothing", vbInformation, "character needed"
counter = 0
Text1.SetFocus
End If

Text2.Text = counter & "  """ & letter & """" & "   Found"

End Sub



Private Sub highlight()
rtb.SelStart = i - 1
rtb.SelLength = Len(letter)
rtb.SelColor = vbRed
End Sub

Private Sub Command2_Click()
rtb.SelStart = 0
rtb.SelLength = Len(rtb.Text)
rtb.SelColor = vbBlack
rtb.SelStart = 0

Text2.Text = ""
Text3.Text = ""

Text1.Text = ""
Text1.SetFocus
End Sub


Private Sub Command3_Click()
Unload Me
End Sub


Private Sub Command4_Click()
On Error GoTo err

cd.Filter = "Text files (*.txt)|*.txt"
cd.ShowOpen
cd.CancelError = True

Open cd.FileName For Input As #1
rtb.Text = Input$(LOF(1), 1)
Close #1

Text1.SetFocus

err:
Exit Sub
End Sub


Private Sub Command5_Click()
Dim counter As Integer
Dim back As Integer
Dim front As Integer

letter = Text1.Text
rep = Text3.Text

If letter <> "" And rep <> "" Then


If Len(letter) = 1 Then

   i = InStr(1, rtb.Text, letter)
   Do While i
   counter = counter + 1
   i = InStr(i + 1, rtb.Text, letter)
   Loop
   rtb.Text = Replace(rtb.Text, letter, rep)
   Text2.Text = counter & "  """ & letter & """" & "   Replaced"
   If counter = 0 Then
   MsgBox Text1.Text & "  not found", vbInformation, "Replaced"
   End If
End If

i = InStr(1, rtb.Text, letter)

If Len(letter) > 1 And i = 1 Then

 back = Asc(Mid(rtb.Text, i + Len(letter), 1))
If (back < 65 Or back > 90 And back < 97 Or back > 122) Then
rtb.SelStart = 0
rtb.SelLength = Len(letter)
rtb.SelText = Replace(rtb.SelText, rtb.SelText, rep)
counter = 1
End If
End If

i = InStr(Len(letter) + 1, rtb.Text, letter)

If Len(letter) > 1 And i > 1 Then

Do While i
If i = Len(rtb.Text) - Len(letter) + 1 Then
rtb.SelStart = i - 1
rtb.SelLength = Len(letter)
rtb.SelText = Replace(rtb.SelText, rtb.SelText, rep)
counter = counter + 1
Exit Do
End If
front = Asc(Mid(rtb.Text, i - 1, 1))
back = Asc(Mid(rtb.Text, i + Len(letter), 1))
If (front < 65 Or front > 90 And front < 97 Or front > 122) And (back < 65 Or back > 90 And back < 97 Or back > 122) Then
rtb.SelStart = i - 1
rtb.SelLength = Len(letter)
rtb.SelText = Replace(rtb.SelText, rtb.SelText, rep)
counter = counter + 1
End If
i = InStr(i + 1, rtb.Text, letter)
Loop
End If
   
Text2.Text = counter & "  """ & letter & """" & "   Replaced"
 
If counter = 0 Then
   MsgBox Text1.Text & "  not found", vbInformation, "Replaced"
   End If

Else
MsgBox "Both, Search word or letter and" & vbCrLf & "Replace with textboxes must be filled", vbInformation, "Replacing"
End If

End Sub


