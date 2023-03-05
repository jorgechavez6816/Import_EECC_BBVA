'Desarrollado por Jorge M. Chávez
'Fecha: 01/03/2023

Sub Main
	IgnoreWarning(True)
	Call ReportReaderImport()	'D:\RUC1\DATA\Archivos fuente.ILB\2022_BBVA.pdf
	Call Summarization()	'H_BBVA2022_.IMD
	Client.CloseAll
	Dim pm As Object
	Dim SourcePath As String
	Dim DestinationPath As String
	Set SourcePath = Client.WorkingDirectory
	Set DestinationPath = "D:\RUC1\DATA\_EECC"
	Client.RunAtServer False
	Set pm = Client.ProjectManagement
	pm.MoveDatabase SourcePath + "H_BBVA2022.IMD", DestinationPath
	pm.MoveDatabase SourcePath + "H.1_Resumen_BBVA.IMD", DestinationPath
	Set pm = Nothing
	Client.RefreshFileExplorer
End Sub


' Archivo - Asistente de importación: Report Reader
Function ReportReaderImport
	dbName = "H_BBVA2022.IMD"
	Client.ImportPrintReportEx "D:\RUC1\DATA\Definiciones de importación.ILB\BBVA_CTA_CTE.jpm", "D:\RUC1\DATA\Archivos fuente.ILB\2022_BBVA.pdf", dbname, FALSE, FALSE
	Client.OpenDatabase (dbName)
End Function

' Análisis: Resumen
Function Summarization
	Set db = Client.OpenDatabase("H_BBVA2022.IMD")
	Set task = db.Summarization
	task.AddFieldToSummarize "PERIODO"
	task.AddFieldToSummarize "CUENTA"
	task.AddFieldToInc "MONEDA"
	task.AddFieldToTotal "CARGOS_ABONO"
	task.AddFieldToTotal "ITF"
	dbName = "H.1_Resumen_BBVA.IMD"
	task.OutputDBName = dbName
	task.CreatePercentField = FALSE
	task.UseFieldFromFirstOccurrence = TRUE
	task.StatisticsToInclude = SM_SUM
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

