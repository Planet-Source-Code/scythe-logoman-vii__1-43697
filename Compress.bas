Attribute VB_Name = "Compress"
'Compression
'To get small savefiles i use compression
'Without this the save will get about 60K

'How this compression works:
'We search for the character thats most time in our Data
'Now we create a Header that holds this char and the counter (how often its in)
'After this we create a Table that holds every position of this char
'and eleminate it in the original stream

'This compression Module wasn´t from me
'But there was no name so i dont know
'who createt it
'Also i fixed some errors

Option Explicit

Public Sub CompressArray(CompData() As Byte)
 Dim OutStream() As Byte
 Dim NewStream() As Byte
 Dim DSize As Long
 Dim Chars(255) As Long
 Dim Bits(6) As Byte
 Dim ActPos As Long
 Dim Counter As Long
 Dim Best As Long
 Dim ActChar As Byte
 Dim I As Long
 Dim PosCount As Long
 Dim BitPos As Long
 Dim OutPos As Long
 Dim NewPos As Long
 Dim NoCompress As Boolean
 Dim LoopCount As Integer

 ReDim OutStream(500)
 ReDim NewStream(500)
 'Like allways an LookUp Table (for Bits)
 For I = 0 To 6
  Bits(I) = 2 ^ I
 Next

 'Here starts the compressionroutine
 Do Until NoCompress

  DSize = UBound(CompData)

  For I = 0 To 255
   Chars(I) = 0
  Next
  Best = 0

  'What byte is how often in the CompressionData
  For I = 0 To DSize
   Chars(CompData(I)) = Chars(CompData(I)) + 1
  Next

  'Search for the Best (most often byte)
  For I = 0 To 255
   If Chars(I) >= Best Then
    ActChar = I
    Best = Chars(I)
   End If
  Next

  'Setup the header
  OutStream(0) = ActChar
  OutStream(1) = Int(Best And &HFF00) / &H100
  OutStream(2) = Best And &HFF
  OutPos = 3
  NewPos = 0
  PosCount = 0
  ActPos = 0
  BitPos = 0
  Counter = 0

  'Main Loop

  'Loop until we found all Chars
  Do While Counter < Best
   'We found the Byte
   If CompData(ActPos) = ActChar Then
    'Less than 7 chars between this and last time ?
    If PosCount < 7 Then
     'Add this Position to the Table
     BitPos = BitPos Or Bits(6 - PosCount)
    Else
     'Add data to array
     AddCharToArray OutStream, OutPos, (PosCount - 7) Or 128
     BitPos = 0
     PosCount = -1
    End If
    Counter = Counter + 1
   Else
    'We found nothing so write the Char to the array
    AddCharToArray NewStream, NewPos, CompData(ActPos)
   End If

   'Move Forward
   ActPos = ActPos + 1
   'Increase difference between last found and actual
   PosCount = PosCount + 1
   'We are 7 chars away from the last position
   If PosCount = 7 Then
    'Set a new Header
    If BitPos > 0 Then
     AddCharToArray OutStream, OutPos, CInt(BitPos)
     BitPos = 0
     PosCount = 0
    End If
    'The Max our Table could write
    ElseIf PosCount = 134 Then
    AddCharToArray OutStream, OutPos, 255
    BitPos = 0
    PosCount = 0
   End If
  Loop

  'Finished search
  'Now add the data we still not added
  If BitPos > 0 Then
   AddCharToArray OutStream, OutPos, CInt(BitPos)
  End If

  For I = ActPos To DSize
   AddCharToArray NewStream, NewPos, CompData(I)
  Next


  'Can we still compress something ?
  If (OutPos + NewPos + 3) > UBound(CompData) Then
   If Best < 1100 Then
    NoCompress = True
    Exit Do
   End If
  End If

  'Copy Compressed data to old array
  ReDim CompData(OutPos + NewPos + 1)
  CompData(0) = Int(OutPos / &H100) And &HFF
  CompData(1) = OutPos And &HFF
  CopyMemory CompData(2), OutStream(0), OutPos
  CopyMemory CompData(2 + OutPos), NewStream(0), NewPos
  LoopCount = LoopCount + 1
  'Form1.txt1(0) = Form1.txt1(0) & UBound(CompData) & "  " & ActChar & "  " & Best & vbCrLf

 Loop
 ReDim Preserve CompData(UBound(CompData) + 2)
 CompData(UBound(CompData) - 1) = Int(LoopCount And &HFF00) / &H100
 CompData(UBound(CompData)) = LoopCount And &HFF

End Sub


Public Sub DeCompressArray(ByteArray() As Byte)
 Dim OutStream() As Byte
 Dim Counter As Long
 Dim Most As Long
 Dim Method As Integer
 Dim DistByte As Long
 Dim PosCount As Long
 Dim I As Long
 Dim F As Long
 Dim InpPos As Long
 Dim OutPos As Long
 Dim FilePos As Long
 Dim FileLong As Long
 Dim NewChar As Byte
 Dim LoopCount As Integer

 'How many loops do we need
 LoopCount = CInt(ByteArray(UBound(ByteArray)) + ByteArray(UBound(ByteArray) - 1) * 256)
 'Cut the Array
 ReDim Preserve ByteArray(UBound(ByteArray) - 2)

 For F = 1 To LoopCount

  FilePos = CLng(ByteArray(0)) * 256 + ByteArray(1) + 2
  NewChar = ByteArray(2)
  Most = CLng(ByteArray(3)) * 256 + ByteArray(4)
  FileLong = UBound(ByteArray) - FilePos + Most
  ReDim OutStream(FileLong)
  InpPos = 5
  Counter = 0
  OutPos = 0
  PosCount = -1

  'Main Loop
  Do While Counter < Most
   DistByte = ByteArray(InpPos)
   InpPos = InpPos + 1
   Method = (-1 * ((DistByte And 128) > 0))
   DistByte = DistByte And 127

   If Method = 1 Then
    DistByte = DistByte + 7
    For I = 1 To DistByte
     AddCharToArray OutStream, OutPos, ByteArray(FilePos)
     FilePos = FilePos + 1
    Next I
    If DistByte <> 134 Then
     AddCharToArray OutStream, OutPos, NewChar
     Counter = Counter + 1
    End If
   Else
    For I = 6 To 0 Step -1
     If Counter = Most Then Exit For
     If (DistByte And 2 ^ I) > 0 Then
      AddCharToArray OutStream, OutPos, NewChar
      Counter = Counter + 1
     Else
      AddCharToArray OutStream, OutPos, ByteArray(FilePos)
      FilePos = FilePos + 1
     End If
    Next I
   End If
  Loop

  For I = FilePos To UBound(ByteArray)
   AddCharToArray OutStream, OutPos, ByteArray(I)
   FilePos = FilePos + 1
  Next I
  ReDim ByteArray(FileLong)
  CopyMemory ByteArray(0), OutStream(0), FileLong + 1

 Next F
End Sub

'this sub will add a char into the outputstream
Private Sub AddCharToArray(ToArray() As Byte, ToPos As Long, Char As Byte)
 If ToPos > UBound(ToArray) Then
  ReDim Preserve ToArray(ToPos + 500)
 End If
 ToArray(ToPos) = Char
 ToPos = ToPos + 1
End Sub

