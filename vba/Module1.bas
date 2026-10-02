Attribute VB_Name = "Module1"
Option Explicit

' No arguments: make row 1 of the active sheet bold
Public Sub BoldHeader()
    ActiveSheet.Rows(1).Font.Bold = True
End Sub

' With arguments: write text to the given cell
Public Sub WriteText(addr As String, text As String)
    ActiveSheet.Range(addr).Value = text
End Sub

' With return value: add two numbers
Public Function AddTwo(a As Double, b As Double) As Double
    AddTwo = a + b
End Function

' Array argument: sum all the values in the array
Public Function SumArray(values As Variant) As Double
    Dim v As Variant
    For Each v In values
        SumArray = SumArray + v
    Next
End Function

' Error handling inside the macro: return the result, or a text that starts with "ERROR:"
'   kind: "sqrt", "log", "acos", "exp", "div"
Public Function SafeCalc(kind As String, x As Double) As Variant
    On Error GoTo Failed
    Select Case kind
        Case "sqrt": SafeCalc = Sqr(x)                                  ' x < 0    -> error 5
        Case "log":  SafeCalc = Log(x)                                  ' x <= 0   -> error 5
        Case "acos": SafeCalc = Application.WorksheetFunction.Acos(x)   ' |x| > 1  -> error 1004
        Case "exp":  SafeCalc = Exp(x)                                  ' too big  -> error 6 (overflow)
        Case "div":  SafeCalc = 1 / x                                   ' x = 0    -> error 11
        Case Else:   Err.Raise 5, , "unknown kind: " & kind
    End Select
    Exit Function
Failed:
    SafeCalc = "ERROR: " & Err.Number & " " & Err.Description
End Function

' No error handling: an error here shows a dialog and Excel waits forever
Public Function UnsafeSqrt(x As Double) As Double
    UnsafeSqrt = Sqr(x)
End Function
