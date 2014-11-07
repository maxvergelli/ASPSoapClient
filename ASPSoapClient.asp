<%
'******************************************************************************** 
'
'	ASP SOAP Client v.1.3
'	Copyright (c) 2014 Max Vergelli 
'	http://www.maxvergelli.com
'	
'	Permission is hereby granted, free of charge, to any person obtaining a  copy
'	of this software and associated documentation files (the "Software"), to deal
'	in the Software without restriction, including without limitation the  rights
'	to use, copy, modify, merge, publish,  distribute,  sublicense,  and/or  sell
'	copies of the Software, and  to  permit  persons  to  whom  the  Software  is
'	furnished to do so, subject to the following conditions:
'	
'	The above copyright notice and this permission notice shall  be  included  in
'	all copies or substantial portions of the Software.
'	
'	THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY  KIND,  EXPRESS  OR
'	IMPLIED, INCLUDING BUT NOT LIMITED  TO  THE  WARRANTIES  OF  MERCHANTABILITY,
'	FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT  SHALL  THE
'	AUTHORS OR COPYRIGHT HOLDERS BE  LIABLE  FOR  ANY  CLAIM,  DAMAGES  OR  OTHER
'	LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
'	OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE  OR  OTHER  DEALINGS  IN
'	THE SOFTWARE.
'
'******************************************************************************** 

Class AspSoapClient

    Private mWebServiceURL
    Private mWebServiceNamespace
    Private mMethodName
    Private mParameters
	Private mResponseXML
	Private mSoapRequest

	Public Sub Class_Initialize()
		Set mParameters = Server.CreateObject("Scripting.Dictionary")
		mResponseXML = ""
		mSoapRequest = ""
	End Sub

    Public Default Function Constructor(WebServiceURL, WebServiceNamespace, MethodName)
		mWebServiceURL = WebServiceURL
		mWebServiceNamespace = WebServiceNamespace
		mMethodName = MethodName
        Set Constructor = Me
    End Function

	Private Sub Class_Terminate()
		Set mParameters = Nothing
	End Sub

	Public Sub SetMethod(MethodName)
		mMethodName = MethodName
	End Sub

	Public Function GetResponseXML()
		GetResponseXML = mResponseXML
	End Function

	Public Function GetSoapRequest()
		GetSoapRequest = mSoapRequest
	End Function

	Public Sub AddParameter(inputName, inputValue)
		If mParameters.Exists(inputName) Then
			mParameters.Item(inputName) = inputValue
		Else
			mParameters.Add inputName, inputValue
		End If
	End Sub

	Public Sub SendData()
		Dim objXmlHttp
        Set objXmlHttp = Server.CreateObject("Microsoft.XMLHTTP")
        mSoapRequest = "<?xml version=""1.0"" encoding=""utf-8""?>"
        mSoapRequest = mSoapRequest & "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
							"<soap:Body>" & _
							"<" & mMethodName & " xmlns=""" & mWebServiceNamespace & """>"
		Dim i,params : params = mParameters.Keys
		For i = 0 To mParameters.Count - 1
			mSoapRequest = mSoapRequest & "<" & params(i) & ">" & mParameters.Item(params(i)) & "</" & params(i) & ">"
		Next
        mSoapRequest = mSoapRequest & "</" & mMethodName & ">" & _
        					"</soap:Body></soap:Envelope>"
        With objXmlHttp
            .Open "post", mWebServiceURL, False
            .setRequestHeader "Content-Type", "text/xml; charset=utf-8"
			.setRequestHeader "SOAPAction", GetSoapActionName()
            .send mSoapRequest
        End With
        mResponseXML = objXmlHttp.responseXML.Text		
		Set objXmlHttp = Nothing
	End Sub

	Public Function GetData(XmlAdapter)
		Set GetData = XmlAdapter.ParseXml(mResponseXML)
	End Function

	Public Function GetSoapActionName()
		Dim s
		If Right(mWebServiceNamespace, 1) = "/" Then
			s = mWebServiceNamespace & mMethodName
		Else
			s = mWebServiceNamespace & "/" & mMethodName
		End If
		GetSoapActionName = s
	End Function

End Class

Class XmlAdapter
	Private mXmlNodeName
    Public Default Function Constructor(XmlNodeName)
		mXmlNodeName = XmlNodeName
        Set Constructor = Me
    End Function
	Public Function ParseXml(XmlDocument)' As Object [Dictionary]	
		Dim data : Set data = Server.CreateObject("Scripting.Dictionary")
        If XmlDocument <> "" Then
            Dim xmlDoc : Set xmlDoc = Server.CreateObject("MSXML2.DOMDocument.3.0")
            xmlDoc.loadXML XmlDocument
			Dim xmlRoot : Set xmlRoot = xmlDoc.getElementsByTagName(mXmlNodeName)(0)
			If Not (xmlRoot Is Nothing) Then
				If Not (xmlRoot.childNodes Is Nothing) Then
					Dim xmlNode
					For Each xmlNode In xmlRoot.childNodes
						If data.Exists(xmlNode.nodeName) Then
							data.Item(xmlNode.nodeName) = xmlNode.Text
						Else
							data.Add xmlNode.nodeName, xmlNode.Text
						End If
					Next  
				End If
			End If
			Set xmlRoot = Nothing
			Set xmlDoc = Nothing
        End If
		Set ParseXml = data
	End Function
End Class

%>