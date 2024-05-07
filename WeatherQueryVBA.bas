Public Sub getWeather()
' email: do.anh@bridgestone.com
' pw: G2j8w5?H}HRBim^
Dim xmlhttp As New MSXML2.xmlhttp, myurl As String, xmlresponse As New DOMDocument
myurl = "https://api.weatherapi.com/v1/forecast.xml?key=ef1c5c286eb442bc90782533240605&q=" & Sheets(1).Range("A2").Value & "&days=3&aqi=no&alerts=no"
xmlhttp.Open "GET", myurl, False
xmlhttp.send
xmlresponse.LoadXML (xmlhttp.responseText)
'Tham khao https://codingislove.com/list-xpath-selectors/ & https://codingislove.com/weather-app-in-excel-vba/
Sheets(1).Range("B3").Value = xmlresponse.SelectNodes("//location/country/text()")(0).Text
Sheets(1).Range("B4").Value = xmlresponse.SelectNodes("//current/last_updated/text()")(0).Text
Sheets(1).Range("B5").Value = xmlresponse.SelectNodes("//forecast/forecastday/astro/sunrise/text()")(0).Text
Sheets(1).Range("B6").Value = xmlresponse.SelectNodes("//forecast/forecastday/astro/sunset/text()")(0).Text
Sheets(1).Range("B7").Value = xmlresponse.SelectNodes("//current/feelslike_c/text()")(0).Text
Sheets(1).Range("B8").Value = xmlresponse.SelectNodes("//current/condition/text")(0).Text
Sheets(1).Range("B9").Value = xmlresponse.SelectNodes("//current/humidity")(0).Text & " %"
Sheets(1).Range("B10").Value = xmlresponse.SelectNodes("//forecast/forecastday[0]/astro/moon_phase")(0).Text
'Du bao ngay mai:
Sheets(1).Range("D3").Value = xmlresponse.SelectNodes("//forecast/forecastday[1]/date")(0).Text
Sheets(1).Range("D4").Value = xmlresponse.SelectNodes("//forecast/forecastday[1]/day/avgtemp_c/text()")(0).Text
Sheets(1).Range("D5").Value = xmlresponse.SelectNodes("//forecast/forecastday[1]/day/condition/text")(0).Text
'Du bao ngay kia:
Sheets(1).Range("D6").Value = xmlresponse.SelectNodes("//forecast/forecastday[2]/date")(0).Text
Sheets(1).Range("D7").Value = xmlresponse.SelectNodes("//forecast/forecastday[2]/day/avgtemp_c/text()")(0).Text
Sheets(1).Range("D8").Value = xmlresponse.SelectNodes("//forecast/forecastday[2]/day/condition/text")(0).Text
'MsgBox (xmlresponse.getElementsByTagName("temperature")(0).Attributes(1).Text)  Alternate method to parse XML
End Sub
