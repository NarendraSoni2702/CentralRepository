
Call CreateHtmlReportFile()
Function CreateHtmlReportFile()
	'Setup file system object
	Set fso =CreateObject("Scripting.FileSystemObject")

	' 1. create html file
	' 2. Open html file and paste the html content into the file

	Set fshtmlFile = fso.openTextfile("E:\TCS-canrig\QTP\ReportingImplementation\Report.html",2,True)
	Set fsCssFile = fso.openTextfile("E:\TCS-canrig\QTP\ReportingImplementation\Reporter.css",1,True)
	strData = "<Html><head>"
	strData = strData & fsCssFile.ReadAll
	strData = strData & "</head>"
	strData = strData & "<table class = ""Zebra""><caption><font size =""6""> <B> Automation Report</B></font></caption>"
	strData = strData & "<tr><th>Runby:""nsoni7""</th><th>WorkStation:""SGV09988980""</th><th>Date:""22-08-2016""</th></tr>"
	'strData = strData & similar for above whatever required we use it

	strData = strData & "<table class = ""Zebra""><thead><tr><th width =""5%"">QC_ID</th><th width =""89%"">TestCaseName</th><th width =""8%"">Result</th></tr></thead>"

	fshtmlFile.write strData
	'' create test case information & Paste into html file
	
	'' create Footer of the html file
	

End Function


Function CreateFooter()

	StrData ="</table><tale class =""bordered"" widht =80%><tr><th widht=""80%""> Total Pass </th><td>"&NoOfPassTC&"</td></tr><tr><th>Total Fail</th><td>"&NoOfFailTC&"</td></tr>"
	StrData =StrData & "<tale class =""bordered"" widht =80%><tr><td><b>Start Time : "&starttime&"</b></td><td><b>End Time : "&Endtime&"</b></td><td><b>Duration : "&Timediff&"</b></td></tr></table></body></html>"


End Function


Function GenerateHtmlReport(strTCResult)

	'step 1 open html file
	nSummaryCounter = nSummaryCounter+1
	if strTCResult=True Then
		TCResult = "Pass"
		TCResultClass="PassResult"
	else
		TCResult = "Fail"
		TCResultClass="FailResult"
	End If 
	
	'StrLogFilePath ='' Run time generated QCID log file path
	
	Set Fso= CreateObject("Scripting.FileSystemObject")
	If Fso.FileExists(StrLogFilePath)=False Then
		strDetailedLog=""
	else
		set Flag =Fso.openTextfile(StrLogFilePath,1)
		strDetailedLog =Flag.ReadAll
		Flag.close
		Set Flag =nothing
		Fso.DeleteFile StrLogFilePath , True
	End If
	
	strTestCaseName ="TestCaseName or QCID"
	SummaryReport ="11"
	
	strData ="<tr><td width=""5%""><a href =""#"" onClick = ToggleList("""&trim(QCID)&"_"&SummaryReport&""")>"&trim(QCID)&"</a></td>"
	strData =strData & "<td width =""89%"">"&trim(strTestCaseName)&"<br><div id ="""&trim(QCID)&"_"&SummaryReport&""" class =""divInfo""><br>"
	strData =strData & strDetailedLog
	strData =strData & "</div></td>"
	strData =strData & "<td width =""8%""><button class ="&TCResultClass&" type =""button"" onClick =ToggleList("""&trim(QCID)&"_"&SummaryReport&""")>"&TCResult&"&#21d5;</button></td></tr>"
	
	
End function 


Function ConvertTO_MhtFile(strHtmlUrl,strMhtmlUrl)
	set CDOObject = CreateObject("CDO.Message")
	CDOObject.CreateMHTMLBody strHtmlUrl,0,"",""
	CDOObject.Getstream.saveToFile strMhtmlUrl,2
End Function