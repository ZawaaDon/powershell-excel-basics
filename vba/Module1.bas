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
