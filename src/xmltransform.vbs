'v0.1 - 2011 Oct 16/?
'The script opens dvblink's channel storage, finds logical channel name and
'frequency and writes the values to a m3u file
'Thank you very much for your help vuego over at MediaPortal forums

strEPGStorage = "eitscanner_epg.xml"

Dim objDesc, strLang
Const ForReading = 1, ForWriting = 2
Const GMT = " +0300"
Const GMTCorrection = -10800
Set fso = CreateObject("Scripting.FileSystemObject")
Set xmlDoc = CreateObject("Microsoft.XMLDOM")
Set hash = CreateObject ("Scripting.Dictionary")
Set hash2 = CreateObject ("Scripting.Dictionary")
Set done = CreateObject ("Scripting.Dictionary")
xmlDoc.Async = "False"
maxId = 0

function epoch2date(myEpoch)
	epoch2date = DateAdd("s", myEpoch-GMTCorrection, "01/01/1970 00:00:00")
end function
function f2digitNum(num)
	if num<10 then
		f2digitNum = "0" & num
	else
		f2digitNum = num
	end if
end function

function formatDate(myDate)
	formatDate = Year(myDate)&f2digitNum(Month(myDate))&f2digitNum(Day(myDate))&f2digitNum(Hour(myDate))&f2digitNum(Minute(myDate))&"00"&GMT
end function

If fso.FileExists(strEPGStorage) Then
	Set f = fso.OpenTextFile("tvguide.xml", ForWriting, True,-1)
	If fso.FileExists("channels.xml") Then
		xmlDoc.Load("channels.xml")
		Set objNodeList = xmlDoc.documentElement.selectNodes("//channel")
		for i=0 to objNodeList.length - 1
			title = objNodeList.item(i).attributes.getNamedItem("title").value
			id = CInt(objNodeList.item(i).attributes.getNamedItem("id").value)
			hash.item(title) = id
			if id > maxId then
				maxId = id 
			end if
		next
	end if
	xmlDoc.Load(strEPGStorage)
	f.WriteLine "<?xml version=""1.0"" encoding=""utf-16"" ?>"
	f.WriteLine "<!DOCTYPE tv SYSTEM ""http://www.teleguide.info/xmltv.dtd"">"
	f.WriteLine "<tv generator-info-name=""TVH_W/2.0"">"
	Set objNodeList = xmlDoc.documentElement.selectNodes("//Channel")
	For i = 0 to objNodeList.length - 1
		name = objNodeList.Item(i).attributes.getNamedItem("Name").value
		if hash2.exists(name) then
			name = name&"+"
			' skip it
		else
			if hash.exists(name) then
				id = hash.item(name)
			else 
				maxId=  maxId+1
				id = maxId
				hash.item(name) = id
			end if
			f.WriteLine "<channel id=""" & id &""">"
			f.WriteLine "	<display-name lang=""ru"">" & name &"</display-name>"
			f.WriteLine "</channel>"
			hash2.item(name)=id
		end if
	Next
	Set objNodeList = xmlDoc.documentElement.selectNodes("//Channel")
	For i = 0 to objNodeList.length - 1
		id = 0
		name = objNodeList.Item(i).attributes.getNamedItem("Name").value
		id = hash.item(name)
		if done.exists(name) then
		' skip it
		else
			Set progNodeList = objNodeList.Item(i).selectNodes("./dvblink_epg/program")
			For j=0 to progNodeList.length - 1
				Select Case progNodeList.Item(j).selectSingleNode("./language").text
					Case "ENG"
					  strLang = "en"
					Case "RUS"
					  strLang = "ru"
					Case Else
					  strLang = "en"
				End Select
				f.WriteLine "<programme start=""" & formatDate(epoch2date(progNodeList.Item(j).selectSingleNode("./start_time").text)) &""" stop="""& formatDate(epoch2date(CLng(progNodeList.Item(j).selectSingleNode("./start_time").text) + CLng(progNodeList.Item(j).selectSingleNode("./duration").text))) &""" channel=""" & (id) &""">"
				f.WriteLine "	<title lang="""&strLang&""">" & Replace(progNodeList.Item(j).attributes.getNamedItem("name").value,"&","&amp;") &"</title>"
				Set objDesc = progNodeList.Item(j).selectSingleNode("./short_desc")
				if not (objDesc is nothing)  then
					f.WriteLine "	<desc lang="""&strLang&""">" & Replace(objDesc.text,"&","&amp;") & "</desc>"
				end if
				f.WriteLine "</programme>"
			Next
			done.item(name) = id
		end if
	Next
	f.WriteLine "</tv>"
	f.Close
	Set fHash = fso.OpenTextFile("channels.xml", ForWriting, True, -1)
	keys = hash.Keys
	items = hash.Items
	fHash.WriteLine "<?xml version=""1.0"" encoding=""utf-16"" ?>"
	fHash.WriteLine "<channels>"
	For i = 0 to hash.Count - 1
		fHash.WriteLine "<channel title=""" &Keys(i) & """ id=""" & Items(i) & """/>"
	Next
	fHash.WriteLine "</channels>"	
	fHash.Close
  Else
	Wscript.Echo strEPGStorage & " not found!"
End If