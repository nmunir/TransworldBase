<%@ Page Language="VB" %>
<%@ Register TagPrefix="FCKeditorV2" Namespace="FredCK.FCKeditorV2" Assembly="FredCK.FCKeditorV2" %>
<script runat="server">

    Sub Page_Load
    Dim sPath As String
    Dim iIndex As Integer
    
      sPath = Request.Url.AbsolutePath
      iIndex = sPath.LastIndexOf( "_samples")
      FCKeditor1.BasePath = sPath.Remove( iIndex, sPath.Length - iIndex  )
      Label1.Text = "sPath = " & sPath & " remainder = " & FCKeditor1.BasePath
    End Sub

</script>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<%--
 * FCKeditor - The text editor for internet
 * Copyright (C) 2003-2004 Frederico Caldeira Knabben
 *
 * Licensed under the terms of the GNU Lesser General Public License:
 * 		http://www.opensource.org/licenses/lgpl-license.php
 *
 * For further information visit:
 * 		http://www.fckeditor.net/
 *
 * File Name: sample01.aspx
 * 	Sample page.
 *
 * Version:  2.1
 * Modified: 2005-02-27 19:46:20
 *
 * File Authors:
 * 		Frederico Caldeira Knabben (fredck@fckeditor.net)
--%>
<html>
<head>
    <title>FCKeditor - Sample</title> 
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta content="noindex, nofollow" name="robots" />
    <link href="../sample.css" type="text/css" rel="stylesheet" />
    <script type="text/javascript">

function FCKeditor_OnComplete( editorInstance )
{
	window.status = editorInstance.Description ;
}

		</script>
</head>
<body>
    <h1>FCKeditor - ASP.Net - Sample 1 
    </h1>
    This sample displays a normal HTML form with an FCKeditor with full features enabled. 
    <br />
    No code behind is used so you don't need to compile the ASPX pages to make it work.
    All other samples use code behind. 
    <hr />
    <form action="sampleposteddata.aspx" method="post" target="_blank">
        <FCKEDITORV2:FCKEDITOR id="FCKeditor1" runat="server" value='This is some <strong>sample text</strong>. You are using <a href="http://www.fckeditor.net/">FCKeditor</a>.'></FCKEDITORV2:FCKEDITOR>
        <asp:Label id="Label1" runat="server"></asp:Label>
        <br />
        <input type="submit" value="Submit" />
    </form>
</body>
</html>