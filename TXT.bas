Attribute VB_Name = "TXT"
Sub readTXT()

Dim myFile As String, text As String, textline As String, posLat As Integer, posLong As Integer
myFile = "C:\Users\mcarlini\Desktop\Origin\TEST_OUTPUT_FILE.TXT"
Open myFile For Input As #1
Do Until EOF(1)
    Line Input #1, textline
    text = text & textline
Loop
Close #1

Debug.Print text

End Sub
