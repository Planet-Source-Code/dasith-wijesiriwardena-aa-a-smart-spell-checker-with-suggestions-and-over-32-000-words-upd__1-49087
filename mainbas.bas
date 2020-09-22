Attribute VB_Name = "modMain"

'Option Explicit

Public alphabetWords(25) As New CollStrC
Public Full_list_count As Long
Public Full_list(40000) As String ' replace the 40000 with the the rough amout of words in the databse + 10,000

'Public Declare Function GetInputState Lib "user32" () As Long

Public Sub Main()
 Dim n As Integer
 
 For n = 0 To 25
   Set alphabetWords(n) = New CollStrC
 Next n
 
 LoadWords
 frmYourProgGoesHere.Show
 
End Sub
Public Sub closer()
 Dim n As Integer
 
 For n = 0 To 25
   Set alphabetWords(n) = Nothing
   frmYourProgGoesHere.ProgressBar1.Value = (n * 100) / 25
 Next n
 
For i = 0 To UBound(Full_list)
    Full_list(i) = ""
Next

 Unload Form1
 Unload frmSpelling
 Unload frmYourProgGoesHere
End Sub

Public Sub LoadWords()

On Error Resume Next
 
Dim a As Long
Dim ds As Long
Dim awrds As Integer
Dim last_awrds As Integer
Dim word As String
Dim meaning As String
Dim intFileNum  As Long
Dim strSpell As String

Form1.Show
Form1.ProgressBar1.Value = 0
DoEvents
ds = GetSetting(App.EXEName, "Words", "Count", "25000")

intFileNum = FreeFile
Open App.Path & "\word.lst" For Input As #intFileNum

  Do While Not EOF(intFileNum)
      Line Input #intFileNum, strSpell
 
      If Len(strSpell) > 0 Then
          If InStr(1, strSpell, ":") > 0 Then
             word = Left$(strSpell, InStr(1, strSpell, ":") - 1)
             meaning = Right$(strSpell, Len(strSpell) - InStr(1, strSpell, ":"))
          Else
             word = strSpell
            meaning = " N\A"
         End If
                
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
    
            If Not awrds = last_awrds Then
                Form1.Label1.Caption = "Loading" & "... " & UCase(Left$(word, 1))
                last_awrds = awrds
            End If
    
             alphabetWords(awrds).Add word, meaning
             
             If Not alphabetWords(awrds).Status = estrStatus.NameExists Or _
                        Not alphabetWords(awrds).Status = estrStatus.NameToShort Then
                Full_list(a) = word
             End If
          
          a = a + 1
          Form1.ProgressBar1.Value = (a * 100 / Val(ds))
          DoEvents
      End If
  
  Loop
  Close #intFileNum

 SaveSetting App.EXEName, "Words", "Count", a
 Full_list_count = a

  Unload Form1
  
  'MsgBox (Full_list(116))

End Sub
 
