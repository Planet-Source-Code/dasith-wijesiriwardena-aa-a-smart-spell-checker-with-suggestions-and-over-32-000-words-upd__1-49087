VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmYourProgGoesHere 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Your program would replace this one."
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6405
   Icon            =   "frmForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   6405
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3360
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Check"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   1
      Top             =   2880
      Width           =   1695
   End
   Begin VB.TextBox txtBody 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   6135
   End
   Begin VB.Label lblMsg 
      Caption         =   "Status : Click Check to check spelling."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2880
      Width           =   4335
   End
End
Attribute VB_Name = "frmYourProgGoesHere"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit

Private Sub Find_Words()
On Error Resume Next
Dim exists As Boolean
Dim n As Long
Dim awrds As Integer

frmSpelling.lstWrong.Clear
frmSpelling.Visible = False

DoEvents

ds1 = Timer

' Checks the text for words and puts them in the word listbox
Dim ds() As String
ds = Split(Replace(txtBody.Text, vbCrLf, " "), " ")

For n = 0 To UBound(ds) 'still working on a better way to ignore web addrresses and etc....
    
    ds(n) = Replace(ds(n), ".", "")

    Do Until Right$(ds(n), 1) <> "."
        ds(n) = Left(ds(n), Len(ds(n)) - 1)
    Loop

    ds(n) = Replace(ds(n), "http://", "")
    ds(n) = Replace(ds(n), "www.", "")
    ds(n) = Replace(ds(n), ".com", "")
    ds(n) = Replace(ds(n), ".net", "")
    ds(n) = Replace(ds(n), ".org", "")
    ds(n) = Replace(ds(n), ",", "")
    ds(n) = Replace(ds(n), "?", "")
    ds(n) = Replace(ds(n), """", "")
    
    If IsNumeric(ds(n)) = True Then
        ds(n) = ""
    End If
    
    ds(n) = Trim(ds(n))
Next

'checks the words to make sure they are spelled correctly
For n = 0 To UBound(ds)

    ProgressBar1.Value = (n * 100 / UBound(ds))
    DoEvents

    Select Case UCase(Left(ds(n), 1))
         Case "A": awrds = 0
         Case "B": awrds = 1
         Case "C": awrds = 2
         Case "D": awrds = 3
         Case "E": awrds = 4
         Case "F": awrds = 5
         Case "G": awrds = 6
         Case "H": awrds = 7
         Case "I": awrds = 8
         Case "J": awrds = 9
         Case "K": awrds = 10
         Case "L": awrds = 11
         Case "M": awrds = 12
         Case "N": awrds = 13
         Case "O": awrds = 14
         Case "P": awrds = 15
         Case "Q": awrds = 16
         Case "R": awrds = 17
         Case "S": awrds = 18
         Case "T": awrds = 19
         Case "U": awrds = 20
         Case "V": awrds = 21
         Case "W": awrds = 22
         Case "X": awrds = 23
         Case "Y": awrds = 24
         Case "Z": awrds = 25
      End Select
   
      exists = alphabetWords(awrds).Exist(ds(n))
  
      Select Case exists
         Case False
         If ds(n) <> vbNullString Then
           frmSpelling.lstWrong.AddItem (ds(n))
         End If
         Case True
         exists = False
      End Select
Next n
 
ProgressBar1.Value = 100

Me.lblMsg.Caption = "Status : " & (UBound(ds) + 1) & " Word(s) completed in " & Str(Round(Timer, 4) - Round(ds1, 4)) & " seconds."

If frmSpelling.lstWrong.ListCount > 0 Then
    frmSpelling.Visible = True
    frmSpelling.List1.Clear
    frmSpelling.cmbSort.SelText = ""
End If
    
If frmSpelling.Visible = False Then
    'Unload frmSpelling
    MsgBox "Spell checker is complete", vbInformation + vbOKOnly, "Spell check"
End If

End Sub

Private Sub cmdFind_Click()
    Call Find_Words
End Sub

Private Sub Form_Load()
  Me.Caption = Full_list_count & " Words in database ...."
End Sub

Private Sub Form_Unload(Cancel As Integer)

On Error Resume Next
Dim dsyes As Integer
dsyes = MsgBox("Do You want to clear the memory before exiting ?", vbYesNo)

If dsyes = vbYes Then

Me.lblMsg.Caption = "Status : Please wailt while the program exits !"
DoEvents

closer

End If

  End
End Sub
