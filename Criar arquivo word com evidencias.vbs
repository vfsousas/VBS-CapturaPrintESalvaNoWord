Set objWord = CreateObject("Word.Application")
objWord.Visible = False
objWord.Documents.Open "C:\scripts\word\modelo\modelo.doc"


username = "Luciana Giannella"

strCenario =InputBox("Entre o texto do Cenario","Input Required", "valor Padrao")
objWord.Application.ActiveDocument.Tables(1).Cell(1,2).Range = strCenario 
objWord.Application.ActiveDocument.Tables(1).Cell(2,2).Range = username
objWord.Application.ActiveDocument.Tables(1).Cell(3,2).Range = Now

intAnswer = _
MsgBox("Aprovado?", _
vbYesNo, "Resultado do teste")

If intAnswer = vbYes Then
	objWord.Application.ActiveDocument.Tables(1).Cell(4,2).Range = "[X] APROVADO [   ] REPROVADO"
Else
	objWord.Application.ActiveDocument.Tables(1).Cell(4,2).Range = "[ ] APROVADO [X] REPROVADO"
End If


strObs =InputBox("Entre observacao, caso houver","Input Required", "Nenhuma Observacao")
objWord.Application.ActiveDocument.Tables(1).Cell(5,2).Range =strObs

Const wdStory = 6
Const wdMove = 0

Set objSelection = objWord.Selection
objSelection.EndKey wdStory, wdMove

execFlag = True

Do While execFlag = True
	
	strPrint =InputBox("Entre com o texto do printscreen","Input Required", "")
	With objSelection 
		.Font.Name = "Arial"
		.Font.Size = "18"
		.TypeText strPrint
		
	End With
	
	
	objSelection.TypeParagraph()
	
	'Taking Screenshot using word object
	Set oWordBasic = CreateObject("Word.Basic")
	oWordBasic.SendKeys "{prtsc}" 
	Set oWordBasic = Nothing
	
	
	WScript.sleep 5000
	' Paste in the screen shot
	objWord.Selection.Paste
	
	objSelection.TypeParagraph()
	objSelection.TypeParagraph()
	
	Set objSelection = objWord.Selection
	objSelection.EndKey wdStory, wdMove
	
	intAnswer = _
	MsgBox("Deseja outro Print da tela?", _
	vbYesNo, "Printscreen")
	
	If intAnswer = vbNo Then
		execFlag = False
	End If
Loop

strFileName =InputBox("Entre com o nome do arquivo das evidencias","Nome do Arquivo", "")
objWord.ActiveDocument.SaveAs "C:\scripts\word\"& strFileName &".doc"
objWord.ActiveDocument.Close
objword.Quit
Set objword = Nothing


Set objWord = CreateObject("Word.Application")
objWord.Visible = True
objWord.Documents.Open "C:\scripts\word\"& strFileName &".doc"

