Attribute VB_Name = "Module1"
Option Explicit

Sub SelectCaseDemo()

Select Case Range("A2").Value
Case "�w��¯l"
Range("B2").Value = "�w��"
Case "�饻����"
Range("B2").Value = "�饻"
Case "����}"
Range("B2").Value = "����"
End Select

End Sub
Sub SelectCaseDemo2()

If (Range("B1").Value > 38) Then
Range("B2").Value = " ���g��"

Else

Range("B2").Value = "�L�g��"
End If

End Sub
