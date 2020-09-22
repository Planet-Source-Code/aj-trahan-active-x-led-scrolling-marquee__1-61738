Attribute VB_Name = "modGlobal"
Option Explicit
Type CharBMP
    Left As Long
    Width As Long
End Type
Public aCharSpace As CharBMP
Public aChars(33 To 90) As CharBMP
Public Const BULB_WIDTH = 5
Public Const CHAR_HEIGHT = 36
Public Const CHAR_THIN = 30
Public Const CHAR_WIDE = 35
Public Const PUNC1 = 10
Public Const PUNC2 = 15
Public Const PUNC3 = 25


Public Sub InitBMPStruct()
    Dim x As Long
    aCharSpace.Left = 0
    aCharSpace.Width = 5
    aChars(33).Left = 0
    aChars(33).Width = PUNC1
    aChars(34).Left = aChars(33).Left + aChars(33).Width
    aChars(34).Width = PUNC3
    For x = 35 To 38
        aChars(x).Left = aChars(x - 1).Left + aChars(x - 1).Width
        aChars(x).Width = CHAR_THIN
    Next x
    aChars(39).Left = aChars(38).Left + aChars(38).Width
    aChars(39).Width = PUNC1
    aChars(40).Left = aChars(39).Left + aChars(39).Width
    aChars(40).Width = PUNC2
    aChars(41).Left = aChars(40).Left + aChars(40).Width
    aChars(41).Width = PUNC2
    aChars(42).Left = aChars(41).Left + aChars(41).Width
    aChars(42).Width = CHAR_THIN
    aChars(43).Left = aChars(42).Left + aChars(42).Width
    aChars(43).Width = CHAR_THIN
    aChars(44).Left = aChars(43).Left + aChars(43).Width
    aChars(44).Width = PUNC2
    aChars(45).Left = aChars(44).Left + aChars(44).Width
    aChars(45).Width = PUNC3
    aChars(46).Left = aChars(45).Left + aChars(45).Width
    aChars(46).Width = PUNC1
    aChars(47).Left = aChars(46).Left + aChars(46).Width
    aChars(47).Width = CHAR_THIN
    aChars(48).Left = aChars(47).Left + aChars(47).Width
    aChars(48).Width = CHAR_THIN
    aChars(49).Left = aChars(48).Left + aChars(48).Width
    aChars(49).Width = CHAR_THIN
    aChars(50).Left = aChars(49).Left + aChars(49).Width
    aChars(50).Width = CHAR_WIDE
    aChars(51).Left = aChars(50).Left + aChars(50).Width
    aChars(51).Width = CHAR_WIDE
    aChars(52).Left = aChars(51).Left + aChars(51).Width
    aChars(52).Width = CHAR_THIN
    For x = 53 To 54
        aChars(x).Left = aChars(x - 1).Left + aChars(x - 1).Width
        aChars(x).Width = CHAR_WIDE
    Next x
    aChars(55).Left = aChars(54).Left + aChars(54).Width
    aChars(55).Width = CHAR_THIN
    For x = 56 To 57
        aChars(x).Left = aChars(x - 1).Left + aChars(x - 1).Width
        aChars(x).Width = CHAR_WIDE
    Next x
    For x = 58 To 59
        aChars(x).Left = aChars(x - 1).Left + aChars(x - 1).Width
        aChars(x).Width = PUNC1
    Next x
    For x = 60 To 62
        aChars(x).Left = aChars(x - 1).Left + aChars(x - 1).Width
        aChars(x).Width = PUNC3
    Next x
    aChars(63).Left = aChars(62).Left + aChars(62).Width
    aChars(63).Width = CHAR_THIN
    aChars(64).Left = aChars(63).Left + aChars(63).Width
    aChars(64).Width = CHAR_WIDE
    aChars(65).Left = aChars(64).Left + aChars(64).Width
    aChars(65).Width = CHAR_WIDE
    For x = 66 To 72
        aChars(x).Left = aChars(x - 1).Left + aChars(x - 1).Width
        aChars(x).Width = CHAR_WIDE
    Next x
    aChars(73).Left = aChars(72).Left + aChars(72).Width
    aChars(73).Width = CHAR_THIN
    aChars(74).Left = aChars(73).Left + aChars(73).Width
    aChars(74).Width = CHAR_WIDE
    aChars(75).Left = aChars(74).Left + aChars(74).Width
    aChars(75).Width = CHAR_THIN
    For x = 76 To 83
        aChars(x).Left = aChars(x - 1).Left + aChars(x - 1).Width
        aChars(x).Width = CHAR_WIDE
    Next x
    aChars(84).Left = aChars(83).Left + aChars(83).Width
    aChars(84).Width = CHAR_THIN
    aChars(85).Left = aChars(84).Left + aChars(84).Width
    aChars(85).Width = CHAR_WIDE
    aChars(86).Left = aChars(85).Left + aChars(85).Width
    aChars(86).Width = CHAR_THIN
    aChars(87).Left = aChars(86).Left + aChars(86).Width
    aChars(87).Width = 50
    For x = 88 To 90
        aChars(x).Left = aChars(x - 1).Left + aChars(x - 1).Width
        aChars(x).Width = CHAR_THIN
    Next x
End Sub

