'#Reference {F5078F18-C551-11D3-89B9-0000F81FE221}#2.6#0#C:\WINXPSP2\system32\msxml2.dll#Microsoft XML, v2.6
Option Explicit

Sub Main
Dim PosInfoLine As String, strFilePath As String, StrXMLFileName As String, BCLine As String, BCNumber As String, BCRack As String, BCPos As Integer
Dim AppPath As String
Dim vExecution_ID As Long

	vExecution_ID = 1
	AppPath = "C:\Lirix"
	strFilePath = AppPath & "\log\INyDIA_Distribute_Log.MDB"
	StrXMLFileName = Create_XML_File(strFilePath, "Muestras del " & Format(Now,"dd mmmm yyyy hh_nn_ss") & ".XML", vExecution_ID)
	Call Execute_HTML_File(StrXMLFileName)

End Sub

Function Create_XML_File(strFilePath As String, StrXMLFileName As String, vExecution_ID) As String
Dim varPathCurrent As String
Dim filesys As Object
Dim xmlDoc As New DOMDocument
Dim conn As Object, rts As Object

	Set filesys = CreateObject("Scripting.FileSystemObject")
	varPathCurrent = filesys.GetParentFolderName(strFilePath)
	Set filesys = Nothing

	Call Delete_File(varPathCurrent & "\" & StrXMLFileName)
	Set conn = CreateObject("ADODB.Connection")
	conn.Provider = "Microsoft.Jet.OLEDB.4.0"
	Call conn.open(strFilePath)

	Set rts = CreateObject("ADODB.recordset")
	'Call rts.open("SELECT [SourceRack] AS Origen, [SourceTube], [TargetRack] As destino, [Position], [WellNumber], [QuotaVolume] FROM qry_ExportHTML WHERE [Execution_ID]= " & vExecution_ID, conn)
	Call rts.open("SELECT [TargetRack] AS [Destination Plate ID], [BarCode] AS [Source Sample ID], [SourceTube] AS [Posici�n origen], [Position] AS [Posici�n destino] FROM qry_ExportHTML WHERE [Execution_ID]= " & vExecution_ID, conn)
	'Save the Recordset into a DOM tree
	Call rts.Save(xmlDoc, 1)
	Call xmlDoc.insertBefore(xmlDoc.createProcessingInstruction("xml-stylesheet", "type=""text/xsl"" href=""INyDIA_Distribute_Log.XSL"""), xmlDoc.documentElement)

	'Writes the datetime of the creation
	Dim xmlFechaNode As IXMLDOMNode
	Set xmlFechaNode = xmlDoc.documentElement.appendChild(xmlDoc.createNode(NODE_ELEMENT, "fecha_hora", ""))
	xmlFechaNode.text = Format(Now, "dddd dd mmmm - hh:nn")

	Call xmlDoc.Save(varPathCurrent & "\" & StrXMLFileName)
	Set rts = Nothing
	Set conn = Nothing
	Set xmlDoc = Nothing
	Create_XML_File = varPathCurrent & "\" & StrXMLFileName

End Function

Sub Delete_File(strFilePath As String)
Dim objFSO As Object

	Set objFSO = CreateObject("Scripting.FileSystemObject")
	If (objFSO.FileExists(strFilePath)) Then
	objFSO.DeleteFile(strFilePath)
	End If
	Set objFSO = Nothing

End Sub

Sub Execute_HTML_File(strHTMFileName As String)
Dim Shl

	Set Shl = CreateObject("WScript.Shell")
	Shl.Run Chr(34) & strHTMFileName & Chr(34), 1, False
	Set Shl = Nothing

End Sub
