Attribute VB_Name = "Cod_MTF"
Option Explicit
'this is a Move To Front Coder which returns a lot of
'small numbers because when a value is found it will be
'placed at the start of the dictionary
'There are two methods in this module
'
'The first one uses a standard dictionary excisting of all the
'ascii characters
'The second one creates a dictionary while it is coding
'this dictionary has to be stored to get the decoder work

Public Sub MTF_CoderArray(Bytes() As Byte, Optional Dictionary As String = "")
    Dim DictString As String
    Dim NewPos As Integer
    Dim X As Long
    If Dictionary = "" Then
        For X = 0 To 255
            DictString = DictString & Chr(X)
        Next
    Else
        DictString = Dictionary
    End If
    For X = 0 To UBound(Bytes)
        NewPos = InStr(DictString, Chr(Bytes(X)))
        DictString = Chr(Bytes(X)) & Left(DictString, NewPos - 1) & Mid(DictString, NewPos + 1)
        Bytes(X) = NewPos - 1
    Next
End Sub

Public Sub MTF_DeCoderArray(Bytes() As Byte, Optional Dictionary As String = "")
    Dim DictString As String
    Dim NewString As String
    Dim NewPos As Integer
    Dim X As Long
    If Dictionary = "" Then
        For X = 0 To 255
            DictString = DictString & Chr(X)
        Next
    Else
        DictString = Dictionary
    End If
    For X = 0 To UBound(Bytes)
        NewPos = Bytes(X) + 1
        Bytes(X) = ASC(Mid(DictString, NewPos, 1))
        DictString = Mid(DictString, NewPos, 1) & Left(DictString, NewPos - 1) & Mid(DictString, NewPos + 1)
    Next
End Sub

Public Sub MTF_CoderArray2(ByteArray() As Byte)
    Dim DictString As String
    Dim OrgDict As String
    Dim NewPos As Integer
    Dim X As Long
    Dim Dictpos As Long
    For X = 0 To UBound(ByteArray)
        If InStr(DictString, Chr(ByteArray(X))) = 0 Then DictString = DictString & Chr(ByteArray(X)): OrgDict = OrgDict & Chr(ByteArray(X))
        NewPos = InStr(DictString, Chr(ByteArray(X)))
        DictString = Chr(ByteArray(X)) & Left(DictString, NewPos - 1) & Mid(DictString, NewPos + 1)
        ByteArray(X) = NewPos - 1
    Next
    Dictpos = UBound(ByteArray) + 1
    ReDim Preserve ByteArray(Len(OrgDict) + 1 + UBound(ByteArray))
    For X = 1 To Len(OrgDict)
        ByteArray(Dictpos) = ASC(Mid(OrgDict, X, 1))
        Dictpos = Dictpos + 1
    Next
    ByteArray(Dictpos) = Len(OrgDict) - 1
End Sub

Public Sub MTF_DeCoderArray2(ByteArray() As Byte)
    Dim DictString As String
    Dim DictLen As Integer
    Dim NewPos As Integer
    Dim X As Long
    DictLen = ByteArray(UBound(ByteArray)) + 1
    For X = UBound(ByteArray) - DictLen To UBound(ByteArray) - 1
        DictString = DictString & Chr(ByteArray(X))
    Next
    ReDim Preserve ByteArray(UBound(ByteArray) - DictLen - 1)
    For X = 0 To UBound(ByteArray)
        NewPos = ByteArray(X) + 1
        ByteArray(X) = ASC(Mid(DictString, NewPos, 1))
        DictString = Mid(DictString, NewPos, 1) & Left(DictString, NewPos - 1) & Mid(DictString, NewPos + 1)
    Next
End Sub

