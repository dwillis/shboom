<%@ LANGUAGE=VBScript%>
<!--#include file ="adovbs.inc"-->


<%
TheDB = Request("TheDB")
TheFullPath = Request("TheFullPath")
ANewFolder = Request("ANewFolder")
ProjectName = Request("ProjectName")
     
dbName= TheDB
 %>

<html>
 
<head>
<title>Select Table or View</title>
</head>
<html>
<body  bgcolor="white">
<!--#include virtual = "/thetimes/rcon.inc" -->

<%

	Set rstSchema = conn.OpenSchema(adSchemaTables)

 %>








<h2 align="center">Select a Table</h2>







<form method="post" name="tableform" action="thelink.asp">
<input type="hidden" value="<% = TheFullPath %>" name="TheFullPath">
<input type="hidden" value="<% = ANewFolder %>" name="ANewFolder">
<input type="hidden" value="<% = ProjectName %>" name="ProjectName">
<input type="hidden" value="<% = TheDB %>" name="TheDB">



	
<p><center><select id="select1" name="TheTable" size="1" ONChange='tableform.submit()'>
<option value=""></option>
			
<%	Do Until rstSchema.EOF
	
	If rstSchema("Table_Type") = "TABLE" or rstSchema("Table_Type") = "VIEW" Then %>
          <option value="<% = rstSchema("Table_Name") %>"><% = rstSchema("Table_Name") & "|" & rstSchema("Table_Type") %></option>

<%	End If
		rstSchema.MoveNext
	Loop
	rstSchema.Close
	Conn.Close
%>
</select>
</center>
<center><p><input type="submit" value="Submit"></p>
  </center>
</form>
</BODY>
</HTML>

