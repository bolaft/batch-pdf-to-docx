Sub convertPDFtoDOC()
'
' convertPDFtoDOC Macro
'
    Dim docDirectory As String
    Dim pdfDirectory As String
    Dim docPath As String
    Dim doc As Document
    On Error Resume Next:

    docDirectory = "C:\Users\<USER>\DOCX\"
    pdfDirectory = "C:\Users\<USER>\PDF\"

    pdfFile = Dir(pdfDirectory & "*.*")

    Do While pdfFile <> ""
        docPath = docDirectory & pdfFile & ".docx"

        Set doc = Documents.Open(FileName:=pdfDirectory & pdfFile)
        ActiveDocument.SaveAs2 FileName:=docPath, FileFormat:=wdFormatXMLDocument

        Documents.Close
        pdfFile = Dir
    Loop
End Sub
