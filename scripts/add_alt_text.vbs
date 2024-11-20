Set objWord = CreateObject("Word.Application")

objWord.Visible = False
objWord.DisplayAlerts = False

'Use first argument passed as document path
Set doc = objWord.Documents.Open(WScript.Arguments(0))

Set alt_texts = CreateObject("System.Collections.ArrayList")

'Use third argument as alt text strings joined by +
alt_texts = Split(WScript.Arguments(2), "+")

'Iterate through tables in document and set correct alt text for each
For a = 1 to WScript.Arguments(3)
	Set tbl = doc.Tables(a)
	tbl.Descr = Replace(alt_texts(a-1), "_", " ")
Next

'Save as second argument name
Call doc.SaveAs(WScript.Arguments(1))

doc.Saved = TRUE
doc.Close
objWord.Quit
