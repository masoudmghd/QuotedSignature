Attribute VB_Name = "QuotedSignature"
Sub sigUpdate()

url1 = Environ("AppData") + "\Microsoft\Signatures\signature.htm"
url2 = Environ("AppData") + "\Microsoft\Signatures\temp.txt"

Dim fso As Object
Set fso = CreateObject("Scripting.FileSystemObject")

If fso.fileexists(url2) = False Then
    Call fso.CopyFile(url1, url2, True)
End If

Dim txtFile As Object
Set txtFile = fso.OpenTextFile(url2, 1, False)

txt = txtFile.ReadAll
txtFile.Close
txt = Replace(txt, "„ ‰", dayQu)
Set txtFile = fso.OpenTextFile(url1, 2, False)
txtFile.Write txt
txtFile.Close
Set fso = Nothing
Set txtFile = Nothing

End Sub
Function dayQu() As String

Dim qfile As Object
Dim qworkbook As Object
Dim qsheet As Object
Dim lastrowu As Integer

quoteUrl = Environ("AppData") + "\Microsoft\Signatures\quote.xlsx"
Set qfile = CreateObject("Excel.application")
Set qworkbook = qfile.Workbooks.Open(quoteUrl)
Set qsheet = qworkbook.Worksheets(1)

randNum = Int((126 - 1 + 1) * Rnd + 1)
dquote = qsheet.Cells(randNum, 2)
dayQu = dquote
qworkbook.Close False

End Function
Sub resetSig()

Dim fso As Object
Set fso = CreateObject("Scripting.FileSystemObject")
url2 = Environ("AppData") + "\Microsoft\Signatures\temp.txt"

Call fso.deletefile(url2)

End Sub
