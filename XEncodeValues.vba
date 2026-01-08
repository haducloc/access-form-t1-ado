Option Compare Database
Option Explicit

' Encode a list of Variant values into a pipe-delimited string, preserving empty slots
Public Function EncodeValues(ParamArray variants() As Variant) As String
    Dim vals() As Variant
    Dim i As Long
    Dim parts() As String

    ' Allow passing a single Variant array: EncodeValues(arr)
    If UBound(variants) = 0 Then
        If IsArray(variants(0)) Then
            vals = variants(0)
        Else
            vals = variants
        End If
    Else
        vals = variants
    End If

    ReDim parts(LBound(vals) To UBound(vals))
    For i = LBound(vals) To UBound(vals)
        If IsNull(vals(i)) Or IsEmpty(vals(i)) Then
            parts(i) = vbNullString
        Else
            parts(i) = CStr(vals(i))
        End If
    Next i

    EncodeValues = Join(parts, "|")
End Function

' Decode a pipe-delimited string into a String array, preserving empty elements
Public Function DecodeValues(ByVal values As String) As String()
    ' Split preserves empty tokens: "a||c" -> ["a", "", "c"]
    ' Also: Split("") -> [""] (easy to handle, no weird bounds)
    DecodeValues = Split(values, "|")
End Function
