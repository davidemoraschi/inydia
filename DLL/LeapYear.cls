VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LeapYear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Function to return if the specified year is a leap year
Public Function IsLeapYear(yr As Variant) As Boolean
    'If year is divisible by 4 and not divisible by 100, or
    'It is divisible by 400, it is a leap year
    If (yr Mod 4 = 0 And yr Mod 100 <> 0) Or yr Mod 400 = 0 Then
        IsLeapYear = True
    Else
        IsLeapYear = False
    End If
    
End Function
Public Function FillWellMatrix() As String()
Dim WellMatrix(96) As String

      WellMatrix(1) = "A1"
      WellMatrix(2) = "B1"
      WellMatrix(3) = "C1"
      WellMatrix(4) = "D1"
      WellMatrix(5) = "E1"
      WellMatrix(6) = "F1"
      WellMatrix(7) = "G1"
      WellMatrix(8) = "H1"
      WellMatrix(9) = "A2"
      WellMatrix(10) = "B2"
      WellMatrix(11) = "C2"
      WellMatrix(12) = "D2"
      WellMatrix(13) = "E2"
      WellMatrix(14) = "F2"
      WellMatrix(15) = "G2"
      WellMatrix(16) = "H2"
      WellMatrix(17) = "A3"
      WellMatrix(18) = "B3"
      WellMatrix(19) = "C3"
      WellMatrix(20) = "D3"
      WellMatrix(21) = "E3"
      WellMatrix(22) = "F3"
      WellMatrix(23) = "G3"
      WellMatrix(24) = "H3"
      WellMatrix(25) = "A4"
      WellMatrix(26) = "B4"
      WellMatrix(27) = "C4"
      WellMatrix(28) = "D4"
      WellMatrix(29) = "E4"
      WellMatrix(30) = "F4"
      WellMatrix(31) = "G4"
      WellMatrix(32) = "H4"
      WellMatrix(33) = "A5"
      WellMatrix(34) = "B5"
      WellMatrix(35) = "C5"
      WellMatrix(36) = "D5"
      WellMatrix(37) = "E5"
      WellMatrix(38) = "F5"
      WellMatrix(39) = "G5"
      WellMatrix(40) = "H5"
      WellMatrix(41) = "A6"
      WellMatrix(42) = "B6"
      WellMatrix(43) = "C6"
      WellMatrix(44) = "D6"
      WellMatrix(45) = "E6"
      WellMatrix(46) = "F6"
      WellMatrix(47) = "G6"
      WellMatrix(48) = "H6"
      WellMatrix(49) = "A7"
      WellMatrix(50) = "B7"
      WellMatrix(51) = "C7"
      WellMatrix(52) = "D7"
      WellMatrix(53) = "E7"
      WellMatrix(54) = "F7"
      WellMatrix(55) = "G7"
      WellMatrix(56) = "H7"
      WellMatrix(57) = "A8"
      WellMatrix(58) = "B8"
      WellMatrix(59) = "C8"
      WellMatrix(60) = "D8"
      WellMatrix(61) = "E8"
      WellMatrix(62) = "F8"
      WellMatrix(63) = "G8"
      WellMatrix(64) = "H8"
      WellMatrix(65) = "A9"
      WellMatrix(66) = "B9"
      WellMatrix(67) = "C9"
      WellMatrix(68) = "D9"
      WellMatrix(69) = "E9"
      WellMatrix(70) = "F9"
      WellMatrix(71) = "G9"
      WellMatrix(72) = "H9"
      WellMatrix(73) = "A10"
      WellMatrix(74) = "B10"
      WellMatrix(75) = "C10"
      WellMatrix(76) = "D10"
      WellMatrix(77) = "E10"
      WellMatrix(78) = "F10"
      WellMatrix(79) = "G10"
      WellMatrix(80) = "H10"
      WellMatrix(81) = "A11"
      WellMatrix(82) = "B11"
      WellMatrix(83) = "C11"
      WellMatrix(84) = "D11"
      WellMatrix(85) = "E11"
      WellMatrix(86) = "F11"
      WellMatrix(87) = "G11"
      WellMatrix(88) = "H11"
      WellMatrix(89) = "A12"
      WellMatrix(90) = "B12"
      WellMatrix(91) = "C12"
      WellMatrix(92) = "D12"
      WellMatrix(93) = "E12"
      WellMatrix(94) = "F12"
      WellMatrix(95) = "G12"
      WellMatrix(96) = "H12"

FillWellMatrix = WellMatrix
End Function
