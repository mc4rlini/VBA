Attribute VB_Name = "PDF"
Sub PDFtoTXT()

    Dim AcroXApp As Acrobat.acroApp
    Dim AcroXAVDoc As Acrobat.AcroAVDoc
    Dim AcroXPDDoc As Acrobat.AcroPDDoc

    Dim Filename As String, DFilename As String, jsObj As Object

    Filename = "C:\Users\mcarlini\Desktop\Origin\079310C-000-PP-106_AFU_5_20210129.pdf"
    DFilename = "C:\Users\mcarlini\Desktop\Origin\TEST_OUTPUT_FILE.TXT"
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

End Sub



