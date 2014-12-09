<%@ Page ValidateRequest="false" Language="C#" AutoEventWireup="false" %>
<%@ Register TagPrefix="FCKeditorV2" Namespace="FredCK.FCKeditorV2" Assembly="FredCK.FCKeditorV2" %>
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
<script runat="server" language="C#">
	// This sample doesnt use a code behind file to avoid the user to have to compile
	// the page to run it.
	protected override void OnLoad(EventArgs e)
	{
		// Automatically calculates the editor base path based on the _samples directory.
		// This is usefull only for these samples. A real application should use something like this:
		// FCKeditor1.BasePath = '/FCKeditor/' ;	// '/FCKeditor/' is the default value.
		string sPath = Request.Url.AbsolutePath ;
		int iIndex = sPath.LastIndexOf( "_samples") ;
		FCKeditor1.BasePath = sPath.Remove( iIndex, sPath.Length - iIndex  ) ;
	}
</script>
<html>
	<head>
		<title>FCKeditor - Sample</title>
		<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
		<meta name="robots" content="noindex, nofollow">
		<link href="../sample.css" rel="stylesheet" type="text/css" />
		<script type="text/javascript">

function FCKeditor_OnComplete( editorInstance )
{
	window.status = editorInstance.Description ;
}

		</script>
	</head>
	<body>
		<h1>FCKeditor - ASP.Net - Sample 1</h1>
		This sample displays a normal HTML form with an FCKeditor with full features 
		enabled.
		<br>
		No code behind is used so you don't need to compile the ASPX pages to make it 
		work. All other samples use code behind.
		<hr>
		<form action="sampleposteddata.aspx" method="post" target="_blank">
			<FCKeditorV2:FCKeditor id="FCKeditor1" runat="server" value='This is some <strong>sample text</strong>. You are using <a href="http://www.fckeditor.net/">FCKeditor</a>.'></FCKeditorV2:FCKeditor>
			<br>
			<input type="submit" value="Submit">
		</form>
	</body>
</html>
