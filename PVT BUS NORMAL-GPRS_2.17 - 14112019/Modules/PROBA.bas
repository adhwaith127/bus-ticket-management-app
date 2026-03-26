Attribute VB_Name = "PROBA"
'------------------------------------------------------------------
'
'                         PROBA Algorithm
'               (Programmable Rotating Box Algorithm)
'
'              Developed and written by Dirk Rijmenants
'
'------------------------------------------------------------------

' The PROBA algorithm uses 32 sBoxes where each byte is devided in
' two nibles (4 bit), each encryption by 16 sBoxes. There are 16
' different sBoxes to select by the first 128 bits of the key.
' These sBoxes are used as a subsitution encryption for a given
' 4 bit value. Each set of 16 sBoxes advances (rotates) as an
' odometer, where each of the 16 sBoxex has his own typical turover
' point. In addition to the normal advancing of the sBoxes, there
' is a set of sBox rotations, caused by certain output values. This
' results in a different sBox rotation process on a different given
' plaintext. The start position of each of the 32 sBoxes is set
' by the second set of 128 bits of the key.

' The key selects the sBoxes and theire start position.
' To set this key, the user has two options:
' - The PROBAsetKey sub: to enter a 32 byte string, were each
'   byte has a value from 0 to 255 (don't use this methode for
'   text keys, but for real 0-255 range values)
' - the PROBAsetExpandedKey sub: to enter a variable lenght key.
'   This methode is preferred for text keys like passwords or frazes
'   IMPORTANT REMARK: Since a RC4 style expansion is used for the
'   PROBAsetExpandedKey sub, one should use only quality keys
'   without repetitions. 'AAAA' equals 'AA', or 'ABAB' has the same
'   effect as 'ABABAB' (natural weakness of RC4). Therefore, some
'   key input quality controle is adviced to avoid these weak key,
'   or other ways must be used to transform variable keys into
'   the 256 bit key for the encryption.

' Each nible encryption full cycle is 18,446,744,073,709,551,616
' turn, however ever changing due to output interaction within
' the nible and with the other nible. This provides a hughe amount
' of different settings for the 32 stages of the encryption,
' giving a high security.

' In this demo, after encryption, the text is encoded to Ba                                                                                  se64, to
' keep the encrypted data readable. This results in an increase of
' data with about 50 percent. Be adviced that this is only a demo
' with a textbox with a limited text size. Create your own encryption
' loops to procces long strings or files, by first initializing the
' key and than using the PROBAencodeByte and PROBAdecodeByte subs.
'
' All Comments and speedup tips are most welcome.
'
' The code is freeware and can be used without restrictions, provide
' the author is given the credits and is notified.
'
' D. Rijmenants (c) 2005 mailto: dr.defcom@telenet.be

Option Explicit

Private sBoxInit(0 To 15)   As Variant
Private BoxTurnOver         As Variant
Private sBoxPos(32)         As Byte
Private sBox(32)            As Byte
Private sBoxOut(32)         As Byte
Private sBoxInvInit(32, 15) As Byte

' ------------------------------------------------------------
'                      String functions
' ------------------------------------------------------------

Public Function PROBAencodeString(aText As String, ByVal aKey As String, expandKey As Boolean) As String
'encrypt string
Dim i As Double
If expandKey = False Then
    'set 32 byte key
    Call PROBAsetKey(aKey)
    Else
    'set variable key
    Call PROBAsetExpandedKey(aKey)
    End If
For i = 1 To Len(aText)
    PROBAencodeString = PROBAencodeString & Chr(PROBAencodeByte(Asc(Mid(aText, i, 1))))
    If i Mod 100 = 0 Then
        UpdateStatus (i / Len(aText) * 100)
    End If
Next
UpdateStatus (0)
End Function

Public Function PROBAdecodeString(aText As String, ByVal aKey As String, expandKey As Boolean) As String
'decrypt string
Dim i As Double
If expandKey = False Then
    'set 32 byte key
    Call PROBAsetKey(aKey)
    Else
    'set variable key
    Call PROBAsetExpandedKey(aKey)
    End If
For i = 1 To Len(aText)
    PROBAdecodeString = PROBAdecodeString & Chr(PROBAdecodeByte(Asc(Mid(aText, i, 1))))
    If i Mod 100 = 0 Then
        UpdateStatus (i / Len(aText) * 100)
    End If
Next
UpdateStatus (0)
End Function

' ------------------------------------------------------------
'                      Encryption functions
' ------------------------------------------------------------

Private Sub PROBAinitKey()
Dim i As Byte
Dim j As Byte
'sBox configuration (encode)
sBoxInit(0) = Array(12, 8, 9, 1, 2, 4, 10, 13, 11, 3, 0, 15, 7, 6, 14, 5)
sBoxInit(1) = Array(4, 0, 10, 11, 3, 14, 9, 8, 13, 1, 2, 7, 5, 6, 12, 15)
sBoxInit(2) = Array(0, 12, 15, 2, 4, 3, 9, 13, 1, 10, 8, 11, 14, 5, 7, 6)
sBoxInit(3) = Array(7, 6, 8, 5, 0, 9, 3, 2, 1, 10, 15, 11, 14, 4, 13, 12)
sBoxInit(4) = Array(11, 13, 4, 3, 9, 10, 5, 1, 8, 12, 6, 14, 7, 15, 2, 0)
sBoxInit(5) = Array(14, 3, 4, 1, 0, 10, 5, 11, 2, 15, 6, 8, 12, 13, 9, 7)
sBoxInit(6) = Array(1, 5, 11, 12, 6, 4, 15, 0, 7, 3, 14, 9, 13, 8, 10, 2)
sBoxInit(7) = Array(5, 3, 11, 13, 2, 1, 12, 10, 0, 4, 7, 6, 14, 8, 15, 9)
sBoxInit(8) = Array(11, 3, 10, 5, 1, 14, 12, 13, 15, 2, 7, 8, 6, 0, 9, 4)
sBoxInit(9) = Array(1, 2, 6, 0, 15, 5, 13, 3, 14, 4, 10, 12, 9, 11, 8, 7)
sBoxInit(10) = Array(12, 11, 13, 3, 2, 14, 9, 4, 1, 10, 8, 7, 0, 6, 5, 15)
sBoxInit(11) = Array(3, 10, 4, 5, 0, 9, 6, 8, 7, 11, 12, 13, 2, 15, 14, 1)
sBoxInit(12) = Array(7, 0, 9, 8, 3, 10, 13, 1, 11, 4, 2, 12, 6, 14, 5, 15)
sBoxInit(13) = Array(5, 11, 4, 3, 2, 0, 12, 1, 15, 14, 6, 10, 9, 13, 7, 8)
sBoxInit(14) = Array(3, 2, 1, 10, 11, 9, 15, 4, 5, 14, 13, 0, 6, 7, 12, 8)
sBoxInit(15) = Array(2, 6, 13, 0, 15, 14, 12, 9, 8, 11, 3, 10, 5, 7, 4, 1)
'sBoxInv configuration (decode)
For i = 0 To 15
    For j = 0 To 15
    sBoxInvInit(i, sBoxInit(i)(j)) = j
    Next
Next
'turnover points per sBox (first value (0) not used!)
BoxTurnOver = Array(0, 2, 14, 6, 15, 3, 7, 11, 5, 9, 3, 14, 13, 4, 6, 8, 1)
End Sub

Private Sub PROBAsetKey(aKey As String)
' INPUT: key string containing 32 bytes
' key 256 bit (32 x 8 bit value)
' 32 x 4 (128) bits for sbox selection
' 32 x 4 (128) bits for sbox startposition selection
Dim i As Byte
Dim Ks(32) As Byte
Dim Key() As Byte
Key() = StrConv(aKey, vbFromUnicode)
'select the 32 sBoxes
For i = 0 To 15
    sBox(i + 1) = Key(i) And 15 'Hi boxes
    sBox(i + 17) = Int(Key(i) / 16)  'Lo boxes
Next
' select initial position of the 32 sboxes
For i = 0 To 15
    sBoxPos(i + 1) = Key(i + 16) And 15 'Hi boxes
    sBoxPos(i + 17) = Int(Key(i + 16) / 16) 'Lo boxes
Next
Call PROBAinitKey
End Sub

Private Sub PROBAsetExpandedKey(aKey As String)
' INPUT: variable lenght key,
' recalculated to a 256 bits key with RC4-style scramble
'
' key string containing 32 bytes
' key 256 bit (32 x 8 bits)
' 32 x 4 (128) bits for sbox selection
' 32 x 4 (128) bits for sbox startposition selection
Dim Ks(255) As Integer
Dim Ss As Integer
Dim Ps As Integer
Dim i As Integer
Dim j As Integer
Dim KeyLen As Integer
Dim Key() As Byte
Dim tmp As Byte
For i = 0 To 255
    Ks(i) = i
Next
'scramble (RC4-style)
KeyLen = Len(aKey)
Key() = StrConv(aKey, vbFromUnicode)
For i = 0 To 255
    j = (j + Ks(i) + Key(i Mod KeyLen)) Mod 255
    tmp = Ks(i)
    Ks(i) = Ks(j)
    Ks(j) = tmp
Next
'select sBoxs
For i = 1 To 16
    sBox(i) = Ks(i) And 15 'Hi boxes
    sBox(i + 16) = Int(Ks(i) / 16) 'Lo boxes
Next
'select init position sBoxs
For i = 1 To 16
    sBoxPos(i) = Ks(16 + i) And 15 'Hi boxes
    sBoxPos(i + 16) = Int(Ks(16 + i) / 16) 'Lo boxes
Next
Call PROBAinitKey
End Sub

Private Function PROBAencodeByte(aByte As Byte) As Byte
'encrypt High Nible
sBoxOut(1) = sBoxEncode((Int(aByte / 16)), 1)
sBoxOut(2) = sBoxEncode((sBoxOut(1)), 2)
sBoxOut(3) = sBoxEncode((sBoxOut(2)), 3)
sBoxOut(4) = sBoxEncode((sBoxOut(3)), 4)
sBoxOut(5) = sBoxEncode((sBoxOut(4)), 5)
sBoxOut(6) = sBoxEncode((sBoxOut(5)), 6)
sBoxOut(7) = sBoxEncode((sBoxOut(6)), 7)
sBoxOut(8) = sBoxEncode((sBoxOut(7)), 8)
sBoxOut(9) = sBoxEncode((sBoxOut(8)), 9)
sBoxOut(10) = sBoxEncode((sBoxOut(9)), 10)
sBoxOut(11) = sBoxEncode((sBoxOut(10)), 11)
sBoxOut(12) = sBoxEncode((sBoxOut(11)), 12)
sBoxOut(13) = sBoxEncode((sBoxOut(12)), 13)
sBoxOut(14) = sBoxEncode((sBoxOut(13)), 14)
sBoxOut(15) = sBoxEncode((sBoxOut(14)), 15)
sBoxOut(16) = sBoxEncode((sBoxOut(15)), 16)
'encrypt Low Nible
sBoxOut(17) = sBoxEncode((aByte And 15), 17)
sBoxOut(18) = sBoxEncode((sBoxOut(17)), 18)
sBoxOut(19) = sBoxEncode((sBoxOut(18)), 19)
sBoxOut(20) = sBoxEncode((sBoxOut(19)), 20)
sBoxOut(21) = sBoxEncode((sBoxOut(20)), 21)
sBoxOut(22) = sBoxEncode((sBoxOut(21)), 22)
sBoxOut(23) = sBoxEncode((sBoxOut(22)), 23)
sBoxOut(24) = sBoxEncode((sBoxOut(23)), 24)
sBoxOut(25) = sBoxEncode((sBoxOut(24)), 25)
sBoxOut(26) = sBoxEncode((sBoxOut(25)), 26)
sBoxOut(27) = sBoxEncode((sBoxOut(26)), 27)
sBoxOut(28) = sBoxEncode((sBoxOut(27)), 28)
sBoxOut(29) = sBoxEncode((sBoxOut(28)), 29)
sBoxOut(30) = sBoxEncode((sBoxOut(29)), 30)
sBoxOut(31) = sBoxEncode((sBoxOut(30)), 31)
sBoxOut(32) = sBoxEncode((sBoxOut(31)), 32)
'calculate encrypted byte
PROBAencodeByte = (sBoxOut(16) * 16) + sBoxOut(32)
'advance boxes
Call PROBAturnBoxes
End Function

Private Function PROBAdecodeByte(aByte As Byte) As Byte
Dim HI As Byte
Dim LO As Byte
'decrypt High Nible
sBoxOut(16) = Int(aByte / 16)
sBoxOut(15) = sBoxDecode((sBoxOut(16)), 16)
sBoxOut(14) = sBoxDecode((sBoxOut(15)), 15)
sBoxOut(13) = sBoxDecode((sBoxOut(14)), 14)
sBoxOut(12) = sBoxDecode((sBoxOut(13)), 13)
sBoxOut(11) = sBoxDecode((sBoxOut(12)), 12)
sBoxOut(10) = sBoxDecode((sBoxOut(11)), 11)
sBoxOut(9) = sBoxDecode((sBoxOut(10)), 10)
sBoxOut(8) = sBoxDecode((sBoxOut(9)), 9)
sBoxOut(7) = sBoxDecode((sBoxOut(8)), 8)
sBoxOut(6) = sBoxDecode((sBoxOut(7)), 7)
sBoxOut(5) = sBoxDecode((sBoxOut(6)), 6)
sBoxOut(4) = sBoxDecode((sBoxOut(5)), 5)
sBoxOut(3) = sBoxDecode((sBoxOut(4)), 4)
sBoxOut(2) = sBoxDecode((sBoxOut(3)), 3)
sBoxOut(1) = sBoxDecode((sBoxOut(2)), 2)
HI = sBoxDecode((sBoxOut(1)), 1)
'decrypt Low Nible
sBoxOut(32) = aByte And 15
sBoxOut(31) = sBoxDecode((sBoxOut(32)), 32)
sBoxOut(30) = sBoxDecode((sBoxOut(31)), 31)
sBoxOut(29) = sBoxDecode((sBoxOut(30)), 30)
sBoxOut(28) = sBoxDecode((sBoxOut(29)), 29)
sBoxOut(27) = sBoxDecode((sBoxOut(28)), 28)
sBoxOut(26) = sBoxDecode((sBoxOut(27)), 27)
sBoxOut(25) = sBoxDecode((sBoxOut(26)), 26)
sBoxOut(24) = sBoxDecode((sBoxOut(25)), 25)
sBoxOut(23) = sBoxDecode((sBoxOut(24)), 24)
sBoxOut(22) = sBoxDecode((sBoxOut(23)), 23)
sBoxOut(21) = sBoxDecode((sBoxOut(22)), 22)
sBoxOut(20) = sBoxDecode((sBoxOut(21)), 21)
sBoxOut(19) = sBoxDecode((sBoxOut(20)), 20)
sBoxOut(18) = sBoxDecode((sBoxOut(19)), 19)
sBoxOut(17) = sBoxDecode((sBoxOut(18)), 18)
LO = sBoxDecode((sBoxOut(17)), 17)
'calculate decrypted byte
PROBAdecodeByte = (HI * 16) + LO
'advance boxes
Call PROBAturnBoxes
End Function

Private Function sBoxEncode(aByte As Byte, aBox As Byte) As Byte
'encrypt nible with given sBox and offset
Dim pos As Byte
pos = aByte + sBoxPos(aBox)
If pos > 15 Then pos = pos - 16
sBoxEncode = sBoxInit(sBox(aBox))(pos)
End Function

Private Function sBoxDecode(aByte As Byte, aBox As Byte) As Byte
'decrypt nible with given sBoxInv and offset
Dim i As Integer
i = sBoxInvInit(sBox(aBox), aByte)
i = i - sBoxPos(aBox)
If i < 0 Then i = i + 16
sBoxDecode = i
End Function

Private Sub PROBAturnBoxes()
'advance the sBoxes by turnover or by inter-action
Dim i As Byte
'HI sBoxes, normal turns
Rotate (1)
For i = 1 To 15
If sBoxPos(i) = BoxTurnOver(sBox(i)) Then
    Rotate (i + 1)
    Else
    Exit For
    End If
Next
'LO sBoxes, normal turns
Rotate 17
For i = 17 To 31
If sBoxPos(i) = BoxTurnOver(sBox(i)) Then
    Rotate (i + 1)
    Else
    Exit For
    End If
Next
'output depended turns
If sBoxOut(1) = 0 Then Rotate (26)
If sBoxOut(1) = 0 Then Rotate (23)
If sBoxOut(17) = 0 Then Rotate (14)
If sBoxOut(17) = 0 Then Rotate (8)
If sBoxOut(3) = 0 Then Rotate (21)
If sBoxOut(18) = 0 Then Rotate (6)
If sBoxOut(2) = 0 And sBoxOut(4) = 0 Then Rotate (28)
If sBoxOut(7) = 0 And sBoxOut(12) = 0 Then Rotate (15)
If sBoxOut(20) = 0 And sBoxOut(24) = 0 Then Rotate (7)
If sBoxOut(5) = 0 And sBoxOut(6) = 0 Then Rotate (31)
If sBoxOut(18) = 0 And sBoxOut(20) = 0 Then Rotate (17)
If sBoxOut(6) + sBoxOut(27) = 8 Then Rotate (25)
If sBoxOut(10) + sBoxOut(19) = 8 Then Rotate (5)
If sBoxOut(8) + sBoxOut(21) = 8 Then Rotate (30)
If sBoxOut(7) + sBoxOut(19) = 8 Then Rotate (9)
If sBoxOut(4) + sBoxOut(7) = 8 Then Rotate (10)
If sBoxOut(2) + sBoxOut(19) = 15 Then Rotate (32)
If sBoxOut(3) + sBoxOut(22) = 15 Then Rotate (16)
If sBoxOut(6) + sBoxOut(21) = 15 Then Rotate (11)
If sBoxOut(7) + sBoxOut(19) = 15 Then Rotate (19)

''next lines are for demonstration purposes only
'' and will visualize the rotations of the Sboxes
'Dim tmp As String
'For i = 1 To 16
'tmp = Trim(Str(sBoxPos(i))) & " "
'If Len(tmp) = 2 Then tmp = "0" & tmp
'frmPatch.lblpos.Caption = frmPatch.lblpos.Caption & tmp
'Next
'frmPatch.lblpos.Caption = frmPatch.lblpos.Caption & "- - "
'For i = 17 To 32
'tmp = Trim(Str(sBoxPos(i))) & " "
'If Len(tmp) = 2 Then tmp = "0" & tmp
'frmPatch.lblpos.Caption = frmPatch.lblpos.Caption & tmp
'Next
'frmPatch.lblpos.Refresh
End Sub

Private Sub Rotate(aPos As Byte)
'advance a sBox position by 1
sBoxPos(aPos) = sBoxPos(aPos) + 1
If sBoxPos(aPos) > 15 Then sBoxPos(aPos) = sBoxPos(aPos) - 16
End Sub

