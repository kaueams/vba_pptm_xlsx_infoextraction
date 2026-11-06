'THIS CODE IS STILL IN DEVELOPMENT!!!

Option Explicit



Sub CRIAR_ZLIDES_FOTOS_WEEK_ATUAL()
	
	'#######vars de estrutura - inicio das declaracoes#######
	Dim WEEK_ATUAL As String
	
	Dim Excel_CUP, Excel_UFO, Excel_CROC As String
	Dim xlApp, wb, ws, targetRange, foundCell, QUINA_UPPER_LEFT, firstAddress, allResults
	Dim foundColumn, foundRow As Integer
		Set xlApp = CreateObject("Excel.Application")
		xlApp.Visible = False
	
	Dim folderPath As String
		folderPath = ActivePresentation.Path
	Dim FSO As Object
		Set FSO = CreateObject("Scripting.FileSystemObject")
	Dim targetFolder As Object
		Set targetFolder = FSO.GetFolder(folderPath)

	Dim currentFile As Object
	Dim files(1 To 100) As String
	
	Dim counter As Integer
		counter = 1
	Dim total_de_arquivos as Integer
		total_de_arquivos = 0
	'#######vars de estrutura - final das declaracoes#######
	
	For Each currentFile In targetFolder.Files
	
		files(counter) = currentFile.Name
		counter = counter + 1
		total_de_arquivos = total_de_arquivos + 1
		
	Next currentFile
	
	For counter = 1 to total_de_arquivos 
			
		If Right(LCase(files(counter)),4) = "xlsx" Then
		
			WEEK_ATUAL = Left(Right(UCase(files(counter)),8),3)
			Set wb = xlApp.Workbooks.Open(files(counter))
			Set ws = wb.Sheets(1)
			Set targetRange = ws.Columns("A:AF")
			
			Set QUINA_UPPER_LEFT = targetRange.Find(What:="ATIVADO?", LookIn:=xLValues, LookAt:=xlWhole)
			foundColumn = QUINA_UPPER_LEFT.Column
			
		End If
		
	Next counter
	
	'esse msgbox é so pra ver se QUINA_UPPER_LEFT tá achando algo
	MsgBox "OPAAAAA ACHEI O INICIO DOS DADOS EM " & QUINA_UPPER_LEFT.Address & ", BLZ?", vbInformation, "RESULTADO DAS PARADA"
	
End Sub
