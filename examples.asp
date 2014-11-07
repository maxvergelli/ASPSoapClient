<!--#include file="/ASPSoapClient.asp" -->
<%

' Create a client object to comunicate with the webservice
Dim ws : Set ws = (New AspSoapClient)("http://www.webservicex.net/globalweather.asmx", _ 
										"http://www.webserviceX.NET", "GetWeather")

' Add parameters needed by the webservice
ws.AddParameter "CityName", "New York"
ws.AddParameter "CountryName", "United States"

' Send request to the webservice
ws.SendData

' Get response data from the webservice using an adapter to parse the XML response.
' Note: You can develop your own customized adapter to parse the XML nodes,
' 		in this example, the XmlAdapter read the parent node "CurrentWeather" and its child nodes, 
'		then fill a Dictionary object with the name and value of each child node.
Set data = ws.GetData((New XmlAdapter)("CurrentWeather"))

' You can access a single item of the response, like this...
Response.Write( "SkyConditions of New York: " & data("SkyConditions") & "<br />")

' or list all items inside the "data" Dictionary object
elements = data.Keys
For i = 0 To data.Count - 1
	Response.Write( elements(i) & ": " & data.Item(elements(i)) & "<br />" )
Next

' You can also parse a different parent node of the XML response without sending another request to the webservice!
Set data2 = ws.GetData((New XmlAdapter)("AlternativeXmlNodeName..."))

' Get the raw Soap 1.1 request sent to the webservice
Response.Write("<p>ws.GetSoapRequest(): <br /><textarea name=""textarea"" cols=""150"" rows=""10"" id=""textarea"">" & ws.GetSoapRequest() & "</textarea></p>")

' Get the raw XML response received from the webservice
Response.Write("<p>ws.GetResponseXML(): <br /><textarea name=""textarea"" cols=""150"" rows=""10"" id=""textarea"">" & ws.GetResponseXML() & "</textarea></p>")

Set ws = Nothing

%>