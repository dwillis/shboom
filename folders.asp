<%@ LANGUAGE="VBSCRIPT" %>
<!--#include file ="adovbs.inc"-->

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual InterDev 1.0">
<META HTTP-EQUIV="Content-Type" content="text/html; charset=iso-8859-1">
<TITLE>Select or Create Folder</TITLE>
</HEAD>
<body bgColor="white">
<!--Written by Tom Torok of The New York Times
tomtorok@nytimes.com-->


<%
If Request("dr") <> "" Then
	 TheDrive = Request("dr") 
End If

ThePath = Request("p") 

dp = TheDrive & ThePath

%>
<h1 align="center">Select a directory (folder)</h1>
<blockquote>
<p align="left">Use this form to drill down to the directory in which you will place your
newly created search and query pages. (You must have write privileges.) <strong>If the
directory exists,</strong> drill down until the full path is highlighted in bold to the
left of the input box, then click the submit button.. <strong>If you want to create a new
directory</strong>, drill down to what will become the directory's parent, then enter the
directory name in the input box and click on submit.(You don't have to enter a slash.)</p>
</blockquote>

<form method="POST" action="makefolder.asp">
<strong><% = dp %></strong>
  <input type="text" name="frmPath" size="10"><input type="submit"
  value="Submit" name="B1"></p>
	<input type="hidden" name="frmDrive" value="<% = dp %>">

</form>

<blockquote>

<%
Set fs = CreateObject("Scripting.FileSystemObject")
Set f=fs.GetFolder(dp)
	For Each i In f.SubFolders
	TheNewPath =  dp & "\" & i.Name
	TheNewPath = Replace(TheNewPath,"\\","\")
	 %>
			
	
		<a href="folders.asp?p=<% = Server.URLEncode(TheNewPath) %>"><% = i.Name %><br>

<% Next %>


</blockquote>


</BODY>
</HTML>
