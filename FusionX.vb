Sub SearchTags()
''Uses Adobe Acrobat Type Library

''Initializations
Dim fileDia As Office.FileDialog
Dim AcroXApp As Acrobat.acroApp
Dim AcroXAVDoc As Acrobat.AcroAVDoc
Dim AcroXPDDoc As Acrobat.AcroPDDoc
Dim filePath, expFilePath, folderPath, fileName, txtFile, text, textline As String
Dim regEx, jsObj As Object
Range("A2:C500").Value = "" 'Clear previous tags


''Select the PDF file
Set fileDia = Application.FileDialog(msoFileDialogFilePicker)
With fileDia
    .AllowMultiSelect = False
    .Title = "Please select the file."
    .Filters.Clear
    .Filters.Add "PDF", "*.pdf"
    .Filters.Add "All Files", "*.*"

    If .Show = True Then
        filePath = .SelectedItems(1)
        folderPath = .InitialFileName
        fileName = Right(filePath, Len(filePath) - Len(folderPath))
    Else: Exit Sub
    End If
End With


''Save the PDF as TXT
expFilePath = "C:\temp\FusionX.txt"
Set AcroXApp = CreateObject("AcroExch.App")
AcroXApp.Show
Set AcroXAVDoc = CreateObject("AcroExch.AVDoc")
AcroXAVDoc.Open filePath, "Acrobat"
Set AcroXPDDoc = AcroXAVDoc.GetPDDoc
Set jsObj = AcroXPDDoc.GetJSObject
jsObj.SaveAs expFilePath, "com.adobe.acrobat.plain-text"
AcroXAVDoc.Close False
AcroXApp.Hide
AcroXApp.Exit
txtFile = expFilePath
Open txtFile For Input As #1
Do Until EOF(1)
    Line Input #1, textline
    text = text & textline
Loop
Close #1


''Analize the TXT file with the Regular Expression
Set regEx = CreateObject("VBScript.RegExp")

i = 2
j = 2

Do While Worksheets("Pattern").Cells(i, 1).Value <> "" 'repeat for each regex formula created
    With regEx
        .Pattern = Worksheets("Pattern").Cells(i, 4).Value & Worksheets("Pattern").Cells(i, 1).Value & Worksheets("Pattern").Cells(i, 5).Value
        .Global = True
    End With
    Set matches = regEx.Execute(text)


    For Each Match In matches
        
        If InStr(1, Match.Value, "/") > 6 Then 'check if the tag need to be splitted
            
            Dim splittedMatch() As String
            splittedMatch = Split(Match.Value, "/")
                        
            k = 1

            For Each singleSplit In splittedMatch
                If k = 1 Then
                    Cells(j, 1).Value = fileName
                    Cells(j, 2).Value = singleSplit
                    'Cells(j, 3).Value = Worksheets("Pattern").Cells(i, 2).Value
                    firstSplit = Left(singleSplit, Len(singleSplit) - 1)
                    j = j + 1
                Else
                    Cells(j, 1).Value = fileName
                    Cells(j, 2).Value = firstSplit & singleSplit
                    'Cells(j, 3).Value = Worksheets("Pattern").Cells(i, 2).Value
                    j = j + 1
                End If
                k = k + 1
            Next singleSplit
        Else
            Cells(j, 1).Value = fileName
            Cells(j, 2).Value = Match.Value
            'Cells(j, 3).Value = Worksheets("Pattern").Cells(i, 2).Value
            j = j + 1
        End If
    Next Match

    i = i + 1
Loop


x = 2

Do While Cells(x, 2).Value <> ""
    lastCar = Right(Cells(x, 2).Value, 1)

    If lastCar = ")" Or lastCar = "." Or lastCar = "," Or lastCar = ";" Or lastCar = "''" Or lastCar = "”" Or lastCar = ":" Then
        Cells(x, 2).Value = Left(Cells(x, 2).Value, Len(Cells(x, 2).Value) - 1)
    End If

    x = x + 1
Loop

Range("A:C").RemoveDuplicates Columns:=2, Header:=xlYes

End Sub