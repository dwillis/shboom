<%@ Language=VBScript %>
<%
strConn=Request("strConn")
TheFullPath = Request("TheFullPath")
ANewFolder = Request("ANewFolder")
SPageName = Request("SPageName")
QPageName = Request("QPageName")
TheTable = Request("TheTable")
HowMany = Request("FieldsAray").Count %>


<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>Order Fields</title>
</HEAD>
<body bgColor=#ccffff>

<!--Written by Tom Torok of The New York Times
tomtorok@nytimes.com-->


<h1 align="center">Select Messages and Field Order</h1>

<form method="POST" action="searchtype.asp">
<input type="hidden" name="strConn" value="<% = strConn %>">
<input type="hidden" value="<% = TheFullPath %>" name="TheFullPath">
<input type="hidden" value="<% = ANewFolder %>" name="ANewFolder">
<input type="hidden" value="<% = SPageName %>" name="SPageName">
<input type="hidden" value="<% = QPageName %>" name="QPageName">
<input type="hidden" value="<% = TheTable %>" name="TheTable">

<center>
<% For I = 1 to HowMany %>
	<p>Select a field for position <% = I %> <SELECT name="FieldsAray">

	<% For J = 1 to HowMany %>
		        <OPTION value=<% = Request("FieldsAray")(J) %>><% = Request("FieldsAray")(J) %></option>
	<% Next %>
    </SELECT></P>

	<p> Select message for this field (e.g. <b>Enter first name:</b>)&nbsp;<input type="text" name="BoxCaption" size="40">

	<hr width="40%" align="center">


<% Next %>
</center>
<center><p><input type="submit" value="Submit" id=submit1 name=submit1></p>
  </center>
</form>

</BODY>
</HTML>
