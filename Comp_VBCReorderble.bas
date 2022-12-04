Attribute VB_Name = "Comp_VBCReorderble"
Option Explicit

'This is a 2 run method

Private Type Bitset
    LowValue As Integer
    Needed As Integer
End Type

Private Type MinMax
    Minimum As Byte
    Maximum As Byte
End Type

Private OutPos As Long
Private OutByteBuf As Integer
Private OutBitCount As Integer
Private ReadBitPos As Integer
Private MinValToAdd(7) As Integer
Private LastChar As Byte
Private ExtraBits(12) As Bitset

'Here whe're gone try to compress the on the VBC-Reorderble method
Public Sub Compress_VBC_Reorderble(ByteArray() As Byte)
    Dim X As Long
    Dim OutStream() As Byte
    Dim NewLen As Long
    Dim Char As Byte
    Dim ExtBits As Integer
'first whe're gone try to find the best method to get the best gain/lost ratio
    NewLen = Find_Best(ByteArray)
    LastChar = 0
    ReDim OutStream(NewLen)         'worst case scenario (exact case if no followers found)
'first we're gone store the values wich belong to the lowest value of a group of 64,32,16,8 or 4 characters
    For X = 1 To 12
        'whe devide it by four, cause it's always a factor of four, to store it in six bits
        'it takes always twelve bytes and ((12*8)-(12*6))/8 = 3 bytes lost
        Call AddBitsToArray(OutStream, ExtraBits(X).LowValue / 4, 6)
    Next
    For X = 0 To UBound(ByteArray)
        Char = ByteArray(X)                             'get the next character
        ExtBits = getBitSize(Char)                      'Find number of bits to store according to char
        If ExtBits = 0 Then                             'if it is the same as the last character
            Call AddBitsToArray(OutStream, 0, 2)        'whe only need to store 2 bits
        Else
            Call AddBitsToArray(OutStream, CLng(ExtBits) + 3, 4)    'otherwise store 4 bits
        End If
        If ExtBits <> 0 Then
'extract the lowest value and store it with the minimum number of bits needed
            Call AddBitsToArray(OutStream, CLng(Char - ExtraBits(ExtBits).LowValue), ExtraBits(ExtBits).Needed)
        End If
        LastChar = Char
    Next
'maybe we have some bits leftover so lets store them
    If OutBitCount < 8 Then
        Do While OutBitCount < 8
            OutByteBuf = OutByteBuf * 2
            OutBitCount = OutBitCount + 1
        Loop
        OutStream(OutPos) = OutByteBuf: OutPos = OutPos + 1
    End If
    OutPos = OutPos - 1
    NewLen = UBound(ByteArray)
    ReDim ByteArray(OutPos + 4)
'store the original lenght of the file
    ByteArray(0) = Int(NewLen / &H1000000) And &HFF
    ByteArray(1) = Int(NewLen / &H10000) And &HFF
    ByteArray(2) = Int(NewLen / &H100) And &HFF
    ByteArray(3) = NewLen And &HFF
'and copy in the ByteArray to return it to the caller
    Call CopyMem(ByteArray(4), OutStream(0), OutPos + 1)
End Sub

'Here whe're trying to find the best methods to use
Private Function Find_Best(ByteArray() As Byte) As Long
    Dim X As Long
    Dim Y As Integer
    Dim Z As Integer
    Dim Bestway(12) As MinMax
    Dim CharCount(255) As Long
    Dim Lowest As Long
    Dim CanBeDone As Boolean
    Dim TotCount As Long
    Dim StartVal As Integer
    Dim PosCount As Integer
    Dim NewLong As Long
    Lowest = UBound(ByteArray)
'Get the frequentie of each character
    For X = 0 To UBound(ByteArray)
        CharCount(ByteArray(X)) = CharCount(ByteArray(X)) + 1
    Next
'init the coder to the standard values so you know how much bits needed for each group
    Call Init_VBC
    PosCount = 12
    Do While PosCount <> 0              'search for all groups
        For X = 0 To 255 - (2 ^ ExtraBits(PosCount).Needed - 1) Step 4 'no need to go beyond limit range
            CanBeDone = True
'try to find if the starting value isn't already occupied
            For Z = 12 To PosCount + 1 Step -1
                If X + (2 ^ ExtraBits(PosCount).Needed - 1) >= Bestway(Z).Minimum And X <= Bestway(Z).Maximum Then
                    CanBeDone = False
                    Exit For
                End If
            Next
            If CanBeDone = True Then
'if not occupied, get the total use of the particular number of bits
                TotCount = 0
                For Y = X To X + (2 ^ ExtraBits(PosCount).Needed - 1)
                    TotCount = TotCount + CharCount(Y)
                Next
                If TotCount <= Lowest Then
'if it is the lowest use, save the starting value
                    Lowest = TotCount
                    StartVal = X
                End If
            End If
        Next
'best match is found so lets store it
        NewLong = NewLong + Lowest * (ExtraBits(PosCount).Needed + 4)
        Bestway(PosCount).Minimum = StartVal
        Bestway(PosCount).Maximum = StartVal + (2 ^ ExtraBits(PosCount).Needed - 1)
        PosCount = PosCount - 1
        Lowest = UBound(ByteArray)
    Loop
'transpose them to the variable that can be used troughout the programm
    For X = 1 To 12
        ExtraBits(X).LowValue = Bestway(X).Minimum
    Next
    Find_Best = (NewLong / 8) + 9
End Function

'Here whe're gone Decompress using the VBC-Reorderble method
Public Sub DeCompress_VBC_Reorderble(ByteArray() As Byte)
    Dim X As Long
    Dim OutStream() As Byte
    Dim InpPos As Long
    Dim FileLang As Long
    Dim Char As Byte
    Dim ExtBits As Integer
'init the coder to the standard values so you know how much bits needed for each group
    Call Init_VBC
    LastChar = 0
'extract the original filelenght
    For X = 0 To 3
        FileLang = FileLang * 256 + ByteArray(X)
    Next
    InpPos = 4
'read the 12 values needed to add to the other stored values
    For X = 1 To 12
        ExtraBits(X).LowValue = ReadBitsFromArray(ByteArray, InpPos, 6) * 4
    Next
    ReDim OutStream(FileLang)
    Do While OutPos < FileLang + 1
        ExtBits = ReadBitsFromArray(ByteArray, InpPos, 2)       'read two bits
        If ExtBits = 0 Then                                     'if the two bits say 0 then the new char
            Char = LastChar                                     'is the same as the last char
        Else
'else read two bits more for the group, read the char and and the lowest value
            ExtBits = ExtBits * 4 + ReadBitsFromArray(ByteArray, InpPos, 2)
            Char = ReadBitsFromArray(ByteArray, InpPos, CInt(ExtraBits(ExtBits - 3).Needed)) + ExtraBits(ExtBits - 3).LowValue
        End If
'store the new char into the output stream and store it as the last char
        Call AddCharToArray(OutStream, OutPos, Char)
        LastChar = Char
    Loop
    OutPos = OutPos - 1
    ReDim ByteArray(OutPos)
'copy it intoe the bytearray to return it to the caller
    Call CopyMem(ByteArray(0), OutStream(0), OutPos + 1)
End Sub

'Here whe're gone initialize some variables needed troughout the program
Private Sub Init_VBC()
    OutPos = 0
    OutByteBuf = 0
    OutBitCount = 0
    ReadBitPos = 0
'                    bitsNeeded    from to char     gain/loss
    ExtraBits(0).Needed = 0     'Last Character     +6          only two bits needed to define 0
    ExtraBits(1).Needed = 2     '? - ?+3            +2
    ExtraBits(2).Needed = 2     '? - ?+3            +2
    ExtraBits(3).Needed = 2     '? - ?+3            +2
    ExtraBits(4).Needed = 2     '? - ?+3            +2
    ExtraBits(5).Needed = 2     '? - ?+3            +2
    ExtraBits(6).Needed = 2     '? - ?+3            +2
    ExtraBits(7).Needed = 2     '? - ?+3            +2
    ExtraBits(8).Needed = 2     '? - ?+3            +2
    ExtraBits(9).Needed = 5     '? - ?+31           -1
    ExtraBits(10).Needed = 6    '? - ?+63           -1
    ExtraBits(11).Needed = 6    '? - ?+63           -2
    ExtraBits(12).Needed = 6    '? - ?+63           -2
'ExtraBits().LowValue need to be defined by the program
End Sub

'Here whe're gone check the minimum amount of bits needed to store a value
Private Function getBitSize(Char As Byte) As Byte
    Dim X As Integer
    If Char = LastChar Then
        getBitSize = 0
        Exit Function
    End If
    For X = 1 To 12
        If Char >= ExtraBits(X).LowValue And Char < ExtraBits(X).LowValue + 2 ^ ExtraBits(X).Needed Then
            getBitSize = X
            Exit Function
        End If
    Next
End Function

'this sub will add an amount of bits into the outputstream
Private Sub AddBitsToArray(Toarray() As Byte, Number As Long, Numbits As Integer)
    Dim X As Long
    For X = Numbits - 1 To 0 Step -1
        OutByteBuf = OutByteBuf * 2 + (-1 * ((Number And 2 ^ X) > 0))
        OutBitCount = OutBitCount + 1
        If OutBitCount = 8 Then
            Toarray(OutPos) = OutByteBuf
            OutBitCount = 0
            OutByteBuf = 0
            OutPos = OutPos + 1
            If OutPos > UBound(Toarray) Then
                ReDim Preserve Toarray(OutPos + 500)
            End If
        End If
    Next
End Sub

'this sub will add a char into the outputstream
Private Sub AddCharToArray(Toarray() As Byte, ToPos As Long, Char As Byte)
    If ToPos > UBound(Toarray) Then
        ReDim Preserve Toarray(ToPos + 500)
    End If
    Toarray(ToPos) = Char
    ToPos = ToPos + 1
End Sub

'this sub will read an amount of bits from the inputstream
Private Function ReadBitsFromArray(FromArray() As Byte, FromPos As Long, Numbits As Integer) As Long
    Dim X As Integer
    Dim Temp As Long
    For X = 1 To Numbits
        Temp = Temp * 2 + (-1 * ((FromArray(FromPos) And 2 ^ (7 - ReadBitPos)) > 0))
        ReadBitPos = ReadBitPos + 1
        If ReadBitPos = 8 Then
            If FromPos + 1 > UBound(FromArray) Then
                Do While X < Numbits
                    Temp = Temp * 2
                    X = X + 1
                Loop
                FromPos = FromPos + 1
                Exit For
            End If
            FromPos = FromPos + 1
            ReadBitPos = 0
        End If
    Next
    ReadBitsFromArray = Temp
End Function

