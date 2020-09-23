Attribute VB_Name = "Cod_ValDownShift"
Option Explicit

Private Dictionary As String

'This coder work with a dictionary of all ascii codes
'but don't keep track of the counts
'every time a character is encountered it will be shifted with
'the character wich stand one position earlier in the dictionary

Public Sub ValueDownShifter_Coder(ByteArray() As Byte)
    Dim X As Long
    Dim Temp As Byte
    Call Init_ValueDownShifter
    For X = 0 To UBound(ByteArray)
        Temp = ByteArray(X)
        ByteArray(X) = InStr(Dictionary, Chr(Temp)) - 1
        Call Update_Model(Temp)
    Next
End Sub

Public Sub ValueDownShifter_DeCoder(ByteArray() As Byte)
    Dim X As Long
    Dim Temp As Byte
    Call Init_ValueDownShifter
    For X = 0 To UBound(ByteArray)
        Temp = ASC(Mid(Dictionary, ByteArray(X) + 1, 1))
        ByteArray(X) = Temp
        Call Update_Model(Temp)
    Next
End Sub

Private Sub Init_ValueDownShifter()
    Dim X As Integer
    Dictionary = ""
    For X = 0 To 255
        Dictionary = Dictionary & Chr(X)
    Next
End Sub

Private Sub Update_Model(Char As Byte)
    Dim Dictpos As Integer
    Dim TwistChar As String
    Dictpos = InStr(Dictionary, Chr(Char))
    If Dictpos > 1 Then
        TwistChar = Mid(Dictionary, Dictpos - 1, 1)
        Dictionary = Left(Dictionary, Dictpos - 2) & Chr(Char) & TwistChar & Mid(Dictionary, Dictpos + 1)
    End If
End Sub

