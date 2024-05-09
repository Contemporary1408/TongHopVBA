Public Sub exceljson()
Dim http As Object, JSON As Object, i As Integer
Set http = CreateObject("MSXML2.XMLHTTP")
http.Open "GET", "https://api.weatherapi.com/v1/forecast.json?key=ef1c5c286eb442bc90782533240605&q=" & Sheets(1).Range("A2").Value & "&days=3&aqi=no&alerts=no", False
http.send
Set JSON = ParseJson(http.responseText)
'For Each item In JSON
Sheets(1).Cells(3, "B").Value = JSON("location")("country")
Sheets(1).Cells(4, "B").Value = JSON("current")("last_updated")
Sheets(1).Cells(5, "B").Value = JSON("forecast")("forecastday")(1)("astro")("sunrise")
Sheets(1).Cells(6, "B").Value = JSON("forecast")("forecastday")(1)("astro")("sunset")
Sheets(1).Cells(7, "B").Value = JSON("current")("temp_c")
Sheets(1).Cells(11, "B").Value = JSON("current")("feelslike_c")
Sheets(1).Cells(2, "C").Value = "https:" & JSON("current")("condition")("icon")
Sheets(1).Range("B8").Value = JSON("current")("condition")("text")
'Next item
End Sub
