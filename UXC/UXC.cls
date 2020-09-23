VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UXC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
  '|¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶|'
  '|¶¶              © 2002 Ariel Productions              ¶¶|'
  '|¶¶                All Rights Reserved                 ¶¶|'
  '|¶¶¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¶¶|'
  '|¶¶             Programmer - James Dougherty           ¶¶|'
  '|¶¶             Source - UXC.cls                       ¶¶|'
  '|¶¶             Object - UltimaX.dll                   ¶¶|'
  '|¶¶             Version - 1.0                          ¶¶|'
  '|¶¶             Documentation - /Tools/Docs/UXC.pdf    ¶¶|'
  '|¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶|'

Option Explicit
'
'This was made just to protect game data files.
'It was not tested on large files, so I'm not sure of its limits
'
'Benifits
'   -Small Code (for encryption)
'   -Good Protection
'
'Fallbacks
'   -Key must be 6 characters or 6 digits in length
'
'
'
'P.S. Try to crack my sample file (CrackThis.txt)
'     Post what it says if you can :)
'     Some of you hackers have tools that might be able to
'     but tell me how hard it was to crack

Public Function EncryptUXC(Data As String, Key As String) As String
 On Local Error Resume Next
 Dim i As Long
 
 If Key = "" Or Len(Key) < 6 Then
  MsgBox "Please enter a key with 6 characters", vbOKOnly Or vbInformation, "UXC"
  EncryptUXC = Data
  Exit Function
 End If
 
 For i = 1 To CDbl(Len(Key))
  EncryptUXC = EncryptUXC + Chr(Asc(Mid(Key, i, 1)) + Len(WeekdayName(5, , vbSaturday)) + Len(TimeSerial(18, 21, 1)))
 Next
 
 For i = 1 To CDbl(Len(Data))
  EncryptUXC = EncryptUXC + Chr(Asc(Mid(Data, i, 1)) + ((Len(Key) + (40 - Len(Key))) / 2) + Len(TimeSerial(3, 55, 37)) + ((Len(WeekdayName(1, , vbSunday)) / 10) + 2))
 Next
 
End Function

Public Function DecryptUXC(Data As String, Key As String) As String
 On Local Error Resume Next
 Dim i As Long
 Dim tmpData As String
 
 If Key = "" Then DecryptUXC = Data: Exit Function
 
 For i = 1 To CDbl(Len(Key))
  tmpData = tmpData + Chr(Asc(Mid(Key, i, 1)) + Len(WeekdayName(5, , vbSaturday)) + Len(TimeSerial(18, 21, 1)))
 Next
 
 If Left$(tmpData, 1) = Left$(Data, 1) And _
    Left$(tmpData, 2) = Left$(Data, 2) And _
    Left$(tmpData, 3) = Left$(Data, 3) And _
    Left$(tmpData, 4) = Left$(Data, 4) And _
    Left$(tmpData, 5) = Left$(Data, 5) And _
    Left$(tmpData, 6) = Left$(Data, 6) Then
  For i = 1 To CDbl(Len(tmpData))
   DecryptUXC = DecryptUXC + Chr(Asc(Mid(Key, i, 1)) - Len(WeekdayName(5, , vbSaturday)) - Len(TimeSerial(18, 21, 1)))
  Next
  For i = Len(tmpData) + 1 To CDbl(Len(Data))
   DecryptUXC = DecryptUXC + Chr(Asc(Mid(Data, i, 1)) - ((Len(Key) + (40 - Len(Key))) / 2) - Len(TimeSerial(3, 55, 37)) - ((Len(WeekdayName(1, , vbSunday)) / 10) + 2))
  Next
  DecryptUXC = Right$(DecryptUXC, (Len(DecryptUXC) - Len(Key)))
 Else
  DecryptUXC = Data
  Exit Function
 End If
 
End Function

Public Function ProEncryptUXC(Data As String, Key As String) As String
 Dim tmpString As String
 
 tmpString = EncryptUXC(Data, Key)
 ProEncryptUXC = EncryptUXC(tmpString, Key)
 
End Function

Public Function ProDecryptUXC(Data As String, Key As String) As String
 Dim tmpString As String
 
 tmpString = DecryptUXC(Data, Key)
 ProDecryptUXC = DecryptUXC(tmpString, Key)
 
End Function
