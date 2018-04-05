README: Data Evidence Collector
# Data-Evidence-Collector
A Microsoft Excel Application using Visual Basic for Application (VBA) programming to web scrape documents and extract supporting documents. Note: `This project won't run without a "private" application installed/accessed`. However, web scraping techniques and codes used in this project could serve as reference for other developers.

![Data-Evidence-Collector-Screenshot](https://github.com/ArielLomoctos/Visual-Basic-For-Application-Projects/blob/master/DataEvidenceCollector-Screenshot.JPG)
`Click to Download:` https://drive.google.com/file/d/1DI0cC_ip_8U7BKuoLiNZoqZ2hYzUZMTr/view

#### `OPEN INTERNET EXPLORER INSTANCE`

Note: Setup Library Reference: Ctrl+F11 > Tools > References > Find/Tick: Microsoft HTML Object Library, Microsoft Internet Controls

	'Open Internet Explorer instance and ie as visible
	
	Dim IE as InternetExplorerMedium
	Set IE = New InternetExplorerMedium
	IE.navigate "www.google.com"
	IE.Visible = True

	'Loop Until IE.Loading Readystate_Complete
	
	Do Until IE.readyState = READYSTATE_COMPLETE
		DoEvents
	Loop
	
#### `CALL, ACTIVATE, MANIPULATE HTML-TAG-ELEMENTS`

	'User Input
	IE.document.getElementById("MainContent_txtRequestNumber").Value = "Text"

	'Tick documents: Boolean (True or False)
	IE.document.getElementById("MainContent_cbxGetInbox").Checked = False

	'Click buttons:
	IE.document.getElementById("MainContent_btnSearch").Click

	'Wait / Run after some time
	Application.Wait (Now + TimeValue("00:00:05"))

	'Get value within html grid/table <td>
	IE.document.getElementById("MainContent_grdSearch").getElementsByClassName("gridDataRow")(0).getElementsByTagName("td")(4).innerText

#### `EXTRACT DOCUMENTS BASED FROM URL`

	Private Declare Function URLDownloadToFileA Lib "urlmon" (ByVal pCaller As Long, _
	ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, _
	ByVal lpfnCB As Long) As Long

	Private Function downloadfile(URL As String, LocalFilename As String) As Boolean
			Dim lngRetVal As Long
			lngRetVal = URLDownloadToFileA(0, URL, LocalFilename, 0, 0)
			If lngRetVal = 0 Then downloadfile = True
	End Function

	Sub DownloadFileWithURLlink()
		'URLlink, Filename
		downloadfile www.google.com/1fajf-fhd241d-1345f.docx, "wordfilename.docx"
	End Sub
