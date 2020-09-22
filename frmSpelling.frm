VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSpelling 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Spelling"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7785
   ControlBox      =   0   'False
   Icon            =   "frmSpelling.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   7785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   255
      Left            =   1800
      TabIndex        =   13
      Top             =   1200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   1200
      TabIndex        =   12
      Top             =   1200
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   4200
      TabIndex        =   11
      Top             =   2800
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Frame Frame1 
      Caption         =   "Meaning"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   950
      Left            =   4200
      TabIndex        =   8
      Top             =   1800
      Width           =   3375
      Begin VB.Label Label1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   3135
         WordWrap        =   -1  'True
      End
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1620
      Left            =   4200
      TabIndex        =   7
      Top             =   120
      Width           =   3375
   End
   Begin VB.ComboBox cmbSort 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmSpelling.frx":0442
      Left            =   120
      List            =   "frmSpelling.frx":0449
      Locked          =   -1  'True
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2400
      Width           =   3975
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
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
      Left            =   2880
      TabIndex        =   5
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdSuggestion 
      Caption         =   "Suggestions"
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
      Left            =   2880
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add to list"
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
      Left            =   2880
      TabIndex        =   3
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdSkip 
      Caption         =   "Skip it!"
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
      Left            =   2880
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.ListBox lstWrong 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "Status :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2800
      Width           =   7455
   End
   Begin VB.Label lblCheck 
      Caption         =   "Spell check found a word that was spelled incorrectly!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "frmSpelling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strsug As String

Private Sub cmdAdd_Click()
  Dim word As String
  Dim awrds As Integer
  
  On Error Resume Next
  
  If lstWrong.ListCount = 0 Then Exit Sub
  If lstWrong.ListIndex < 0 Then Exit Sub
    
    word = lstWrong.Text
    
    Select Case UCase(Left$(word, 1))
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
       
    alphabetWords(awrds).Add word, " N/A"
       
    'cmbSort.AddItem word
    
    intFileNum = FreeFile
    
    Open App.Path & "\word.lst" For Append As #intFileNum
        'adds the correct words to the file
        Print #intFileNum, word
    Close #intFileNum
    
    lstWrong.RemoveItem lstWrong.ListIndex
    
   MsgBox ("The change might not take effect until the next time you start the program !")

End Sub

Private Sub cmdQuit_Click()
 
    lstWrong.Clear
    Me.Visible = False

End Sub

Private Sub cmdSkip_Click()

If lstWrong.ListIndex >= 0 Then
    i = lstWrong.ListIndex
    lstWrong.RemoveItem (i)
  Else
    If lstWrong.ListCount = 0 Then
        cmdSkip.Enabled = False
    End If
End If

End Sub

Private Sub cmdSuggestion_Click()
On Error Resume Next

Dim i As Long
Dim awrds As Integer

Dim ds As String
Dim score As Long

Dim highval As Long
Dim lowval As Long

Dim lst_in As String

Dim fullcount As Long
fullcount = Full_list_count - 1

Dim Fulllist_item

'this does not work yet!
'MsgBox "Suggestion: (This function is not complete)" & strsug, vbInformation + vbOKOnly, "Found!"
If lstWrong.SelCount = 0 Then
    lstWrong.ListIndex = 0
End If

List1.Clear
Label2.Caption = "Status: Searching.."

ds = lstWrong.List(lstWrong.ListIndex)
    
    Select Case UCase(Left$(ds, 1))
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
    
 Dim fullcount1 As Long
 Dim fullcount2 As Long
    
 fullcount1 = 0
 fullcount2 = 0
    
 For i = 0 To awrds - 1
    fullcount1 = fullcount1 + alphabetWords(i).Count
 Next i
 
 fullcount2 = fullcount1 + alphabetWords(awrds).Count
 
 fullcount = fullcount2 - fullcount1
    
    For i = fullcount1 To fullcount2
    'For i = 0 To fullcount
    
        Fulllist_item = Full_list(i)
            
        If List1.ListCount >= 10 Then
            ProgressBar1.Value = 100
            DoEvents
            Exit For
        End If
    
        score = 0
        'Label2.Caption = "Status: Searching.." & (i + 1) & " of " & cmbSort.ListCount
        ProgressBar1.Value = (i * 100) / fullcount
        DoEvents
    
        If InStr(1, ds, Fulllist_item, vbTextCompare) > 0 Or InStr(1, Fulllist_item, ds, vbTextCompare) > 0 Then
            If Abs(Len(ds) - Len(Fulllist_item)) <= 2 Then
                If Not InStr(1, lst_in, LCase(Fulllist_item)) > 0 Then
                    List1.AddItem (Fulllist_item)
                    lst_in = lst_in & Fulllist_item
                End If
            End If
        Else
        
            For x = 1 To Len(ds)
                If InStr(1, Mid(Fulllist_item, x, 2), Mid(ds, x, 1), vbTextCompare) > 0 Then
                    score = score + 1
                End If
            Next x
            
            highval = (Len(ds) * 0.8) ' (len(ds) * 0.5) - 1
            
            If highval < 2 Then
                highval = 2
            End If
            
            lowval = highval - 1
            
            If score > highval Then ' score >= highval
               If score >= Len(ds) - 2 Then
                    If Abs(Len(ds) - Len(Fulllist_item)) <= 1 Then
                        If Not InStr(1, lst_in, LCase(Fulllist_item)) > 0 Then
                            List1.AddItem (Fulllist_item)
                            lst_in = lst_in & Fulllist_item
                        End If
                    End If
               End If
            ElseIf score > lowval Then
               For x = 1 To Len(ds) - 1 ' score >= lowval
                If Abs(Len(ds) - Len(Fulllist_item)) = 0 Then ' turn off if neccaserry
                    If InStr(1, Fulllist_item, Mid(ds, x, lowval), vbTextCompare) Then
                        If Not InStr(1, lst_in, LCase(Fulllist_item)) > 0 Then
                            List1.AddItem (Fulllist_item)
                            lst_in = lst_in & Fulllist_item
                        End If
                        Exit For
                    End If
                End If ' turn off if neccaserry
               Next
            End If
        End If
        
    Next i

If List1.ListCount = 0 Then
    
    'ds1 = MsgBox("No Suggestions Found ! Do you want to search again with lower filter settings ?", vbYesNo + vbMsgBoxRtlReading)
    
    'If ds1 = vbNo Then
        'Label2.Caption = "Status: No Suggestions !"
    'Else
        Call Command2_Click
    'End If
    
Else
    Label2.Caption = "Status: " & List1.ListCount & " suggestions found !"
End If

End Sub

Private Sub Command1_Click()

Label2.Caption = "Status: Searching !"

On Error Resume Next

Dim ds As String
Dim score As Long

Dim highval As Long
Dim lowval As Long

Dim lst_in As String

Dim fullcount As Long
fullcount = Full_list_count - 1

Dim Fulllist_item

'this does not work yet!
'MsgBox "Suggestion: (This function is not complete)" & strsug, vbInformation + vbOKOnly, "Found!"

If lstWrong.SelCount = 0 Then
    lstWrong.ListIndex = 0
End If

List1.Clear
Label2.Caption = "Status: Searching.."
DoEvents

ds = lstWrong.List(lstWrong.ListIndex)

    For i = 0 To fullcount
    
        Fulllist_item = Full_list(i)
    
        If List1.ListCount >= 10 Then
            ProgressBar1.Value = 100
            Exit For
        End If
    
        score = 0
        'Label2.Caption = "Status: Searching.." & (i + 1) & " of " & cmbSort.ListCount
        ProgressBar1.Value = (i * 100) / fullcount
        DoEvents
    
        If InStr(1, ds, Fulllist_item, vbTextCompare) > 0 Or InStr(1, Fulllist_item, ds, vbTextCompare) > 0 Then
            If Abs(Len(ds) - Len(Fulllist_item)) <= 2 Then
                If Not InStr(1, lst_in, LCase(Fulllist_item)) > 0 Then
                    List1.AddItem (Fulllist_item)
                    lst_in = lst_in & Fulllist_item
                End If
            End If
        Else
        
            For x = 1 To Len(ds)
                If InStr(1, Mid(Fulllist_item, x, 2), Mid(ds, x, 1), vbTextCompare) > 0 Then
                    score = score + 1
                End If
            Next
            
            highval = (Len(ds) * 0.7) ' (len(ds) * 0.5) - 1
            
            If highval < 2 Then
                highval = 2
            End If
            
            lowval = highval - 1
            
            If score > highval Then ' score >= highval
               If score >= Len(ds) - 2 Then
                    If Abs(Len(ds) - Len(Fulllist_item)) <= 1 Then
                        If Not InStr(1, lst_in, LCase(Fulllist_item)) > 0 Then
                            List1.AddItem (Fulllist_item)
                            lst_in = lst_in & Fulllist_item
                        End If
                    End If
               End If
            ElseIf score > lowval Then
               For x = 1 To Len(ds) - 1 ' score >= lowval
                If Abs(Len(ds) - Len(Fulllist_item)) = 0 Then ' turn off if neccaserry
                    If InStr(1, Fulllist_item, Mid(ds, x, lowval), vbTextCompare) Then
                        If Not InStr(1, lst_in, LCase(Fulllist_item)) > 0 Then
                            List1.AddItem (Fulllist_item)
                            lst_in = lst_in & Fulllist_item
                        End If
                        Exit For
                    End If
                End If ' turn off if neccaserry
               Next
            End If
        End If
        
    Next

If List1.ListCount = 0 Then
    Label2.Caption = "Status: No Suggestions !"
    'Call Command1_Click
Else
    Label2.Caption = "Status: " & List1.ListCount & " suggestions found !"
End If

End Sub

Private Sub Command2_Click()
On Error Resume Next

Dim i As Long
Dim awrds As Integer

Dim ds As String
Dim score As Long

Dim highval As Long
Dim lowval As Long

Dim lst_in As String

Dim fullcount As Long
fullcount = Full_list_count - 1

Dim Fulllist_item

'this does not work yet!
'MsgBox "Suggestion: (This function is not complete)" & strsug, vbInformation + vbOKOnly, "Found!"
If lstWrong.SelCount = 0 Then
    lstWrong.ListIndex = 0
End If

List1.Clear
Label2.Caption = "Status: Searching.."

ds = lstWrong.List(lstWrong.ListIndex)
    
    Select Case UCase(Left$(ds, 1))
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
    
 Dim fullcount1 As Long
 Dim fullcount2 As Long
    
 fullcount1 = 0
 fullcount2 = 0
    
 For i = 0 To awrds - 1
    fullcount1 = fullcount1 + alphabetWords(i).Count
 Next i
 
 fullcount2 = fullcount1 + alphabetWords(awrds).Count
 
 fullcount = fullcount2 - fullcount1
    
    For i = fullcount1 To fullcount2
    'For i = 0 To fullcount
    
        Fulllist_item = Full_list(i)
            
        If List1.ListCount >= 10 Then
            ProgressBar1.Value = 100
            DoEvents
            Exit For
        End If
    
        score = 0
        'Label2.Caption = "Status: Searching.." & (i + 1) & " of " & cmbSort.ListCount
        ProgressBar1.Value = (i * 100) / fullcount
        DoEvents
    
        If InStr(1, ds, Fulllist_item, vbTextCompare) > 0 Or InStr(1, Fulllist_item, ds, vbTextCompare) > 0 Then
            If Abs(Len(ds) - Len(Fulllist_item)) <= 2 Then
                If Not InStr(1, lst_in, LCase(Fulllist_item)) > 0 Then
                    List1.AddItem (Fulllist_item)
                    lst_in = lst_in & Fulllist_item
                End If
            End If
        Else
        
            For x = 1 To Len(ds)
                If InStr(1, Mid(Fulllist_item, x, 2), Mid(ds, x, 1), vbTextCompare) > 0 Then
                    score = score + 1
                End If
            Next x
            
            highval = (Len(ds) * 0.6) ' (len(ds) * 0.5) - 1
            
            If highval < 2 Then
                highval = 2
            End If
            
            lowval = highval - 1
            
            If score > highval Then ' score >= highval
               If score >= Len(ds) - 2 Then
                    If Abs(Len(ds) - Len(Fulllist_item)) <= 1 Then
                        If Not InStr(1, lst_in, LCase(Fulllist_item)) > 0 Then
                            List1.AddItem (Fulllist_item)
                            lst_in = lst_in & Fulllist_item
                        End If
                    End If
               End If
            ElseIf score > lowval Then
               For x = 1 To Len(ds) - 1 ' score >= lowval
                If Abs(Len(ds) - Len(Fulllist_item)) = 0 Then ' turn off if neccaserry
                    If InStr(1, Fulllist_item, Mid(ds, x, lowval), vbTextCompare) Then
                        If Not InStr(1, lst_in, LCase(Fulllist_item)) > 0 Then
                            List1.AddItem (Fulllist_item)
                            lst_in = lst_in & Fulllist_item
                        End If
                        Exit For
                    End If
                End If ' turn off if neccaserry
               Next
            End If
        End If
        
    Next i

If List1.ListCount = 0 Then
    
    'ds1 = MsgBox("No Suggestions Found ! Do you want to search again with lower filter settings ?", vbYesNo + vbMsgBoxRtlReading)
    
    'If ds1 = vbNo Then
        'Label2.Caption = "Status: No Suggestions !"
    'Else
        Call Command1_Click
    'End If
    
Else
    Label2.Caption = "Status: " & List1.ListCount & " suggestions found !"
End If
End Sub

Private Sub List1_Click() 'suggested word meanings
Dim Msg As String
Dim find As String
Dim awrds As Integer

  find = List1.List(List1.ListIndex)
  Select Case UCase(Left(find, 1))
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
  
   Msg = alphabetWords(awrds).Item(find)
  
   Label1.Caption = List1.List(List1.ListIndex) & ": " & Msg

End Sub
