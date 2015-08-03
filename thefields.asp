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
<body  bgcolor="#FFCC99">
<!--#include virtual = "/thetimes/rcon.inc" -->

<!--Written by Tom Torok of The New York Times
tomtorok@nytimes.com-->


<%
RS.Open sql, Conn, adOpenStatic, adLockReadOnly, adCmdText

HowMany = RS.RecordCount

%>


<h1 align="center">Select Fields</h1>

<p align="center"><b>Choose the fields for your results pages. Order doesn't matter. You'll be
able to select the order on the next page.</p>

<p align="center">(Hold down the &quot;Ctrl&quot; key and left click with the mouse to select multiple fields.)</b></p>

<form method="post" action="orderfields.asp">
<input type="hidden" value="<% = TheFullPath %>" name="TheFullPath">
<input type="hidden" value="<% = ANewFolder %>" name="ANewFolder">
<input type="hidden" value="<% = ProjectName %>" name="ProjectName">
<input type="hidden" value="<% = TheDB %>" name="TheDB">
<input type="hidden" value="<% = TheTable %>" name="TheTable">

<p><center><select id="select1" multiple name="OFields" size="<% = HowMany %>">

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
