Attribute VB_Name = "RegEx"
Sub RegEx_Ex1()

Dim RegEx As Object, Str As String
Set RegEx = CreateObject("VBScript.RegExp")
With RegEx
    .Pattern = "[0-9]+"
End With
Str = "My Bike Number is MH-12 PP-6145"
Debug.Print RegEx.Test(Str)

End Sub
Sub RegEx_Ex2()

Dim RegEx As Object, Str As String
Set RegEx = CreateObject("VBScript.RegExp")
With RegEx
    .Pattern = "123"
    .Global = True      'If FALSE, Replaces only the first matching string'
End With
Str = "123-654-000-APY-123-XYZ-888"
Debug.Print RegEx.Replace(Str, "Replaced")

End Sub
Sub RegEx_Ex3()

Dim RegEx As Object, Str As String
Set RegEx = CreateObject("VBScript.RegExp")
With RegEx
    .Pattern = "123-XYZ"
    .Global = True
End With
Str = "123-XYZ-326-ABC-983-670-PQR-123-XYZ"
Set matches = RegEx.Execute(Str)
For Each Match In matches
    Debug.Print Match.Value
Next Match

End Sub
