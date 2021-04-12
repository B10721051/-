Attribute VB_Name = "Module1"
Option Explicit

Sub SelectCaseDemo()

Select Case Range("A2").Value
Case "德國麻疹"
Range("B2").Value = "德國"
Case "日本腦炎"
Range("B2").Value = "日本"
Case "香港腳"
Range("B2").Value = "香港"
End Select

End Sub
Sub SelectCaseDemo2()

If (Range("B1").Value > 38) Then
Range("B2").Value = " 有症狀"

Else

Range("B2").Value = "無症狀"
End If

End Sub
