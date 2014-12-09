<%@ Page ValidateRequest="false" language="c#" Codebehind="sample02.aspx.cs" Inherits="FredCK.FCKeditorV2.Samples.Sample02" AutoEventWireup="false" %>
<%@ Register TagPrefix="fckeditorv2" Namespace="FredCK.FCKeditorV2" Assembly="FredCK.FCKeditorV2" %>
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
 * File Name: sample02.aspx
 * 	Sample page.
 * 
 * Version:  2.1
 * Modified: 2005-02-19 15:30:30
 * 
 * File Authors:
 * 		Frederico Caldeira Knabben (fredck@fckeditor.net)
--%>
<html>
	<head>
		<title>FCKeditor - Sample</title>
		<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
		<meta name="robots" content="noindex, nofollow">
		<link href="../sample.css" rel="stylesheet" type="text/css">
		<script type="text/javascript">

function FCKeditor_OnComplete( editorInstance )
{
	window.status = editorInstance.Description ;
}

		</script>
	</head>
	<body>
		<h1>FCKeditor - ASP.Net - Sample 2</h1>
		This sample displays a normal HTML form with an FCKeditor with full features 
		enabled.
		<br>
		The only difference from sample01 is that this page uses code behind and the 
		submitted data is shown in the page itself.
		<hr>
		<form method="post" runat="server">
			<FCKeditorV2:FCKeditor id="FCKeditor1" runat="server" />
			<br>
			<input id="btnSubmit" type="submit" value="Submit" runat="server">
		</form>
		<div id="eSubmittedDataBlock" runat="server">
			<hr>
			This is the submitted data:<br>
			<textarea id="txtSubmitted" runat="server" rows="10" cols="60" style="WIDTH: 100%" readonly></textarea>
		</div>
	</body>
</html>
