<%@ WebService Language="VB" Class="WebService" %>

Imports System.Web
Imports System.Web.Services
Imports System.Web.Services.Protocols

' To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line.
' <System.Web.Script.Services.ScriptService()> _
<WebService(Namespace:="http://sprintexpress.co.uk/")> _
<WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
Public Class WebService
    Inherits System.Web.Services.WebService
    
    <WebMethod()> _
    Public Function HelloWorld(ByVal sTemp As String) As String
        Return "Hello World"
    End Function

    <WebMethod()> _
       Public Function HelloInteger() As Integer
        Return 24680
    End Function

End Class
