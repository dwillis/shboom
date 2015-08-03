<%@ LANGUAGE=VBScript%>
<!--#include file ="adovbs.inc"-->

<%
TheFullPath = Request("TheFullPath")
ANewFolder = Request("ANewFolder")
ProjectName = Request("ProjectName")
TheDB = Request("TheDB")
TheTable = Request("TheTable")

dbName= TheDB

sql="select column_name as FieldName from information_schema.columns where table_name = '" & TheTable & "' "
sql = sql & " order by ordinal_position"


 %>

<html>
<title>Select Fields</title>
</head>
<html>
<body  bgcolor="white">
<!--#include virtual = "/thetimes/rcon.inc" -->

<!--Written by Tom Torok of The New York Times
tomtorok@nytimes.com-->


<%
RS.Open sql, Conn, adOpenStatic, adLockReadOnly, adCmdText

NumRecs = RS.RecordCount

%>


<h1 align="center">Optional Linking Field</h1>

<p align="Left"><b>If you have a field that uniquely identifies records, select it below. The identifier 
will be used as a link to a page that shows all field entries for that record. If you do not have a field 
with a unique identifier simply click on Submit.</p>


<form method="post" action="projectspecs.asp">
<input type="hidden" value="<% = TheFullPath %>" name="TheFullPath">
<input type="hidden" value="<% = ANewFolder %>" name="ANewFolder">
<input type="hidden" value="<% = ProjectName %>" name="ProjectName">
<input type="hidden" value="<% = TheDB %>" name="TheDB">
<input type="hidden" value="<% = TheTable %>" name="TheTable">

<p><center><select id="select1" name="LField" size="<% = NumRecs + 1 %>">
<option selected value=""></option>
<%
Do while not rs.eof

%>
        <option value="<% = rs("FieldName") %>"><%= rs("FieldName") %></option>
<%


RS.MoveNext
Loop

 %>
</select>
</center>
<center><p><input type="submit" value="Submit"></p>
  </center>
</form>

</body>
</html>
