Option Explicit
Dim wordApp, doc, srcPath, pdfPath, totalPages, appendixPage
srcPath = "d:\study\Business Strategy\Assessment_new_v13.docx"
pdfPath = "d:\study\Business Strategy\Assessment_new_v13.pdf"
Set wordApp = CreateObject("Word.Application")
wordApp.Visible = False
wordApp.DisplayAlerts = 0
Set doc = wordApp.Documents.Open(srcPath, False, False)
totalPages = doc.ComputeStatistics(2)
appendixPage = -1
wordApp.Selection.HomeKey 6
With wordApp.Selection.Find
    .ClearFormatting
    .Text = "Appendix A. BCG Classification Evidence"
    .Forward = True
    .Wrap = 1
End With
If wordApp.Selection.Find.Execute() Then
    appendixPage = wordApp.Selection.Range.Information(3)
End If
doc.ExportAsFixedFormat pdfPath, 17
doc.Close False
wordApp.Quit
WScript.Echo "TOTAL_PAGES " & totalPages
WScript.Echo "APPENDIX_START_PAGE " & appendixPage
WScript.Echo "PDF_PATH " & pdfPath
