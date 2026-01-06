' ===== Class Module: XInputState =====
Option Compare Database
Option Explicit

' Value Type
Public ValueType As XValueType

' Can be Null or Actual Value (Bool, String, etc.)
Public Value As Variant

Public IsValid As Boolean

' If IsValid is False
Public ErrorMessage As String
