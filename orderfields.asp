<%@ LANGUAGE="VBSCRIPT" %>
<%

TheFullPath = Request("TheFullPath")
ANewFolder = Request("ANewFolder")
ProjectName = Request("ProjectName")
TheDB = Request("TheDB")
TheTable = Request("TheTable")
HowMany = Request("OFields").Count


%>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual InterDev 1.0">

<title>Order Fields</title>
</head>
<body bgColor="#FFCC99">

<!--Written by Tom Torok of The New York Times
tomtorok@nytimes.com-->



<h1 align="center">Order the Fields</h1>

<p align="center"><b>Select the fields in the order you want them to appear on the query
results page.</b></p>

<form method="POST" action="aliasfields.asp">
<input type="hidden" value="<% = TheFullPath %>" name="TheFullPath">
<input type="hidden" value="<% = ANewFolder %>" name="ANewFolder">
<input type="hidden" value="<% = ProjectName %>" name="ProjectName">
<input type="hidden" value="<% = TheDB %>" name="TheDB">
<input type="hidden" value="<% = TheTable %>" name="TheTable">

<table border="1" align="center">

<% For I = 1 to HowMany %>
      <tr>
	<td><% = I %>.</td>
	<td><select name="SFields" size="1">
	<% For J = 1 to HowMany %>
		        <option value="<% = Request("OFields")(J) %>"><% = Request("OFields")(J) %></option>
	<% Next %>
    </select>
	<br></td></tr>
<% Next %>

</table>
</table>
<center><p><input type="submit" value="Submit"></p>
  </center>
</form>

</body>
</html>
