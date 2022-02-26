Attribute VB_Name = "Test"
Sub Test()
    
    Dim fd As Office.FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
            
    Range("A2:E500").Value = ""
    
    Dim AcroXApp As Acrobat.acroApp
    Dim AcroXAVDoc As Acrobat.AcroAVDoc
    Dim AcroXPDDoc As Acrobat.AcroPDDoc

    Dim Filename As String, DFilename As String, jsObj As Object, Docpath As String, Docname As String
        
    Dim myFile As String, text As String, textline As String, posLat As Integer, posLong As Integer
    
    Dim RegEx As Object, Str As String
    
    With fd

      .AllowMultiSelect = False
      .Title = "Please select the file."
      
      .Filters.Clear
      .Filters.Add "PDF", "*.pdf"
      .Filters.Add "All Files", "*.*"
      
      If .Show = True Then
        Filename = .SelectedItems(1)
      Else: Exit Sub
      End If
    End With
    
    DFilename = "C:\temp\FusionX.txt"
        
    Set AcroXApp = CreateObject("AcroExch.App")
    AcroXApp.Show
    Set AcroXAVDoc = CreateObject("AcroExch.AVDoc")
    AcroXAVDoc.Open Filename, "Acrobat"
    Set AcroXPDDoc = AcroXAVDoc.GetPDDoc
    Set jsObj = AcroXPDDoc.GetJSObject
    jsObj.SaveAs DFilename, "com.adobe.acrobat.plain-text"

    AcroXAVDoc.Close False
    AcroXApp.Hide
    AcroXApp.Exit
    
    myFile = DFilename
   
    Open myFile For Input As #1

    Do Until EOF(1)
        Line Input #1, textline
        text = text & textline
    Loop
    Close #1
          
    Set RegEx = CreateObject("VBScript.RegExp")
    
    i = 2
    j = 2
    Do While Worksheets("Pattern").Cells(i, 1).Value <> ""
           
    With RegEx
        .Pattern = Worksheets("Pattern").Cells(i, 1).Value
        .Global = True
    End With

    Set matches = RegEx.Execute(text)
    For Each Match In matches
        Cells(j, 1).Value = Filename
        Cells(j, 2).Value = Match.Value
        j = j + 1
    Next Match
    
    i = i + 1

    Loop
    
    Range("A:B").RemoveDuplicates Columns:=2, Header:=xlYes
            
End Sub



