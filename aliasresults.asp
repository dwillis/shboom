<%@ LANGUAGE=VBScript%>
<!--#include file ="adovbs.inc"-->
<%

TheFullPath = Request("TheFullPath")
ANewFolder = Request("ANewFolder")
ProjectName = Request("ProjectName")
TheDB = Request("TheDB")
TheTable = Request("TheTable")
SFields = Request("SFields")
HowMany = Request("SFields").Count
' the following line initializes numselect, which is used to comply with code stolen from mpd
NumSelect = HowMany
DFields = Request("DFields")
TheMax = Request("TheMax")
TheRowNum = Request("TheRowNum")
LField = Request("LField")


dbName = TheDB


sql="select column_name as TheField, Data_Type as TheType, ordinal_position as ordposit from information_schema.columns where table_name = '" & TheTable & "' "
sql = sql & " and Data_type <> 'timestamp' "
sql = sql & " order by ordinal_position"



%>
<!--#include virtual = "/thetimes/rcon.inc" -->
<%


RS.Open sql, Conn, adOpenStatic, adLockReadOnly, adCmdText

%>





<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>Alias Fields</title>
</HEAD>
<body bgColor="#FFCC99">

<table align="center" border="1">
<td><b>Select</b></td>
<td><b>Data<br> Type</b></td>
<td><b>Field</b></td>
<td><b>Alias</b></td>
<td><b>Order</b></td>
</tr>

<%

Do While Not rs.EOF %>

<form method="POST" action="sup.asp">

<tr>
<td><input type="checkbox" name="C1" value="<% = rs("TheField") %>" checked></td>








