<%@ Page language="c#" Codebehind="BufferDataSaveResults.aspx.cs" AutoEventWireup="false" Inherits="MACROBufferBrowserWeb.BufferDataSaveResults" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" > 
<html>
  <head>
    <title>WebForm1</title>
    <meta name="GENERATOR" Content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" Content="C#">
    <meta name=vs_defaultClientScript content="JavaScript">
    <meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
	<link rel='stylesheet' href='../style/MACRO1.css' type='text/css'>
	<script language="javascript" src="../script/BufferDataBrowser.js"></script>
	<script language="javascript">
		function NavigateToPage(sURL)
		{
			window.navigate(sURL);
		}
		// screen title
		window.parent.fnSetTitle("Buffer Data Browser Save Results");
	</script>
  </head>
	<div class="clsPopMenu" id="divPopMenu" onclick="menu=this;this.tid=setTimeout('menu.style.visibility=\'hidden\'',20);" 
	onmouseout="menu=this;this.tid=setTimeout('menu.style.visibility=\'hidden\'',20);" onmouseover="clearTimeout(this.tid);">
	</div>
	<%
		// Render save results
		Response.Write( RenderSaveResultsPage() );
	%>
</html>
