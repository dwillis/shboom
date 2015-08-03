<%@ LANGUAGE=VBScript %>
<%
'---- CursorTypeEnum Values ----
Const adOpenForwardOnly = 0
Const adOpenKeyset = 1
Const adOpenDynamic = 2
Const adOpenStatic = 3
'---- LockTypeEnum Values ----
Const adLockReadOnly = 1
Const adLockPessimistic = 2
Const adLockOptimistic = 3
Const adLockBatchOptimistic = 4
'---- CommandTypeEnum Values ----
Const adCmdUnknown = &H0008
Const adCmdText = &H0001
Const adCmdTable = &H0002
Const adCmdStoredProc = &H0004
Const adCmdFile = &H0100
Const adCmdTableDirect = &H0200


TheFullPath = Request("TheFullPath")
ANewFolder = Request("ANewFolder")
ProjectName = Request("ProjectName")
TheDB = Request("TheDB")
TheTable = Request("TheTable")
OrderBy1 = Request("frmOrderBy1")
SortOrder1 = Request("frmSortOrder1")
OrderBy2 = Request("frmOrderBy2")
SortOrder2 = Request("frmSortOrder2")
OrderBy3 = Request("frmOrderBy3")
SortOrder3 = Request("frmSortOrder3")

Set objRS = Server.CreateObject("ADODB.RecordSet")
Set objRS2 = Server.CreateObject("ADODB.RecordSet")

objRS2.Open TheFullPath & ProjectName & "_progspecs.gaga",  "Provider=MSPersist;", adOpenDynamic, adLockBatchOptimistic, adCmdFile




objRS2("OrderBy1") = OrderBy1
objRS2("SortOrder1") = SortOrder1

objRS2("OrderBy2") = OrderBy2
objRS2("SortOrder2") = SortOrder2

objRS2("OrderBy3") = OrderBy3
objRS2("SortOrder3") = SortOrder3

objRS2.Update
Set fs = CreateObject("Scripting.FileSystemObject")
Set f1 = fs.GetFile(TheFullPath & ProjectName & "_progspecs.gaga")
f1.delete


objRS2.save TheFullPath & ProjectName & "_progspecs.gaga"

set f1 = NOTHING
set fs = NOTHING

objRS.Open TheFullPath & ProjectName & "_tablespecs.gaga",  "Provider=MSPersist;", adOpenDynamic, adLockBatchOptimistic, adCmdFile


objRS.Filter = "[ShowDupe] = TRUE"


%>
<html>
<body>
<form method="POST" action="orderby.asp">
<h1 align="center">Order Your Output</h1>
<blockquote><b>&nbsp;&nbsp;&nbsp;This page lets you order the way your output will appear in the results page for filter queries .
 You can order results on up to three fields.<br>
&nbsp;&nbsp;&nbsp;The pulldowns on the left show the available fields; the pulldowns on the right let you select sort order. 
The default is ascending and if you want ascending you needn't select it. The "Ascending" option was included 
for anal compulsives who have to select something.<br>
&nbsp;&nbsp;&nbsp;Your selections will not be saved unless you first click on "Submit". Once you click "Submit" the table below 
will change to show you the first five records sorted. When you're happy with the sort order click on "Next" below the table 
to write your application.</b>
</blockquote>





<table align="center" border=1>
<tr>
<td>First:</td>
<td><SELECT name=frmOrderBy1>
<% If OrderBy1 = "" Then %>
	<OPTION selected value=""></OPTION>
<% Else %>
	<OPTION value=""></OPTION>
<% End If	
Do While Not objRS.EOF


If objRS("TheField") <> OrderBy1 Then
	Response.Write("<OPTION value=" & chr(34) & objRS("TheField") & chr(34) & ">" & objRS("TheField") & "</OPTION>" & vbCRLF)
Else
	Response.Write("<OPTION selected value=" & chr(34) & objRS("TheField") & chr(34) & ">" & objRS("TheField") & "</OPTION>" & vbCRLF)

End If

objRS.MoveNext
Loop
%>
</select>
</td>
<td><select size="1" name="frmSortOrder1">
    <option></option>
<%  Select Case SortOrder1
	Case "" %>
    <option value="ASC">Ascending</option>
    <option value="DESC">Descending</option>
<% Case "ASC" %>		
    <option selected value="ASC">Ascending</option>
    <option value="DESC">Descending</option>
<% Case "DESC" %>
    <option value="ASC">Ascending</option>
    <option selected value="DESC">Descending</option>
<% End Select %>


  </select></td>
</tr>
<tr>
<td>Second:</td>
<td><SELECT name=frmOrderBy2>
<% If OrderBy2 = "" Then %>
	<OPTION selected value=""></OPTION>
<% Else %>
	<OPTION value=""></OPTION>
<% End If
objRS.MoveFirst	
Do While Not objRS.EOF


If objRS("TheField") <> OrderBy2 Then
	Response.Write("<OPTION value=" & chr(34) & objRS("TheField") & chr(34) & ">" & objRS("TheField") & "</OPTION>" & vbCRLF)
Else
	Response.Write("<OPTION selected value=" & chr(34) & objRS("TheField") & chr(34) & ">" & objRS("TheField") & "</OPTION>" & vbCRLF)

End If

objRS.MoveNext
Loop
%>
</select>
</td>
<td><select size="1" name="frmSortOrder2">
    <option></option>
<%  Select Case SortOrder2
	Case "" %>
    <option value="ASC">Ascending</option>
    <option value="DESC">Descending</option>
<% Case "ASC" %>		
    <option selected value="ASC">Ascending</option>
    <option value="DESC">Descending</option>
<% Case "DESC" %>
    <option value="ASC">Ascending</option>
    <option selected value="DESC">Descending</option>
<% End Select %>
  </select></td>
</tr>
<tr>
<td>Third:</td>
<td><SELECT name=frmOrderBy3>
<% If OrderBy3 = "" Then %>
	<OPTION selected value=""></OPTION>
<% Else %>
	<OPTION value=""></OPTION>
<% End If	
objRS.MoveFirst
Do While Not objRS.EOF


If objRS("TheField") <> OrderBy3 Then
	Response.Write("<OPTION value=" & chr(34) & objRS("TheField") & chr(34) & ">" & objRS("TheField") & "</OPTION>" & vbCRLF)
Else
	Response.Write("<OPTION selected value=" & chr(34) & objRS("TheField") & chr(34) & ">" & objRS("TheField") & "</OPTION>" & vbCRLF)

End If

objRS.MoveNext
Loop
%>
</select>
</td>
<td><select size="1" name="frmSortOrder3">
    <option></option>
<%  Select Case SortOrder3
	Case "" %>
    <option value="ASC">Ascending</option>
    <option value="DESC">Descending</option>
<% Case "ASC" %>		
    <option selected value="ASC">Ascending</option>
    <option value="DESC">Descending</option>
<% Case "DESC" %>
    <option value="ASC">Ascending</option>
    <option selected value="DESC">Descending</option>
<% End Select %>
  </select></td>
</tr>

</table>
<INPUT TYPE="HIDDEN" NAME="RecordCount" VALUE="<% = RecordCount %>">
<INPUT TYPE="HIDDEN" NAME="TheFullPath" VALUE="<% = TheFullPath %>">
<INPUT TYPE="HIDDEN" NAME="ANewFolder" VALUE="<% = ANewFolder %>">
<INPUT TYPE="HIDDEN" NAME="ProjectName" VALUE="<% = ProjectName %>">
<INPUT TYPE="HIDDEN" NAME="TheDB" VALUE="<% = TheDB %>">
<INPUT TYPE="HIDDEN" NAME="TheTable" VALUE="<% = TheTable %>">

<center><input type="Submit" value="Submit" name="B1"></center>
</form>

<%
objRS.MoveFirst
objRS.Sort = "ShowDupeOrder"


TheOrderBy = ""

If OrderBy1 <> "" Then
	If SortOrder1 = "DESC" Then
		TheOrderBy = " order by " & orderby1 & " DESC "
	Else
		TheOrderby =  " order by " & orderby1 & " "
	End If
End If

If OrderBy2 <> "" Then
	If SortOrder2 = "DESC" Then
		TheOrderBy = TheOrderBy & ", " & orderby2 & " DESC "
	Else
		TheOrderBy = TheOrderBy & ", " & orderby2 & " "
	End If
End If


If OrderBy3 <> "" Then
	If SortOrder3 = "DESC" Then
		TheOrderBy = TheOrderBy & ", " & orderby3 & " DESC "
	Else
		TheOrderBy = TheOrderBy & ", " & orderby3 & " "
	End If
End If





sql = "Select top 5 * from " & TheTable & TheOrderBy

'Response.Write("The sql is: " & sql & "<br>")

dbName= TheDB %>
     
     
<!--#include virtual = "/thetimes/rcon.inc" -->

<% 	RS.Open sql, Conn, adOpenStatic, adLockReadOnly, adCmdText 

RSCounter = 5

TableCount = RS.RecordCount

If TableCount < 5 Then
	RSCounter = TableCount
End If

%>
<h2 align="center">This is how the results of filter queries will look sorted.</h2>

<table align="center" border="1">
<tr>

<%

Do While Not objRS.EOF %>

<td align="center" bgcolor="#0000FF"><font color="#D9FFFF" size="-1"><b><% = objRS("TheAlias") %></b></font></td>

<% 
objRS.MoveNext
Loop

%>

</tr>
<tr>
<% 
For I = 1 to RSCounter 

objRS.MoveFirst

Do While Not objRS.EOF


	FieldName = objRS("TheField")

%>

<td><font size="-1"><% = rs(FieldName) %></font></td>

<%

objRS.MoveNext
Loop
%>

</tr>
<%
rs.MoveNext
Next


%>


</table>

<form method="POST" action="done.asp">
<INPUT TYPE="HIDDEN" NAME="TheFullPath" VALUE="<% = TheFullPath %>">
<INPUT TYPE="HIDDEN" NAME="ProjectName" VALUE="<% = ProjectName %>">
<center><input type="Submit" value="Next"></center>

</form>


</body>
</html>



