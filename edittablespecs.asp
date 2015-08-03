<%@ LANGUAGE=VBScript%>
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
Response.Buffer = FALSE

TheFullPath = "e:\inetpub\wwwroot\pension\"
ProjectName = "skedb"


Set RS = Server.CreateObject("ADODB.RecordSet")

Response.Write(TheFullPath & ProjectName & "_tablespecs.gaga")

RS.Open TheFullPath & ProjectName & "_tablespecs.gaga",  "Provider=MSPersist;", adOpenDynamic, adLockBatchOptimistic, adCmdFile



%>
<html>
<body>
<P>
<center>
<TABLE BORDER=1>
<TR>
<% For i = 0 to RS.Fields.Count - 1 %>
	<TD><B><% = RS(i).Name %></B></TD>
<% Next %>
</TR>
<% Do While Not RS.EOF %>
	<TR>
	<% For i = 0 to RS.Fields.Count - 1 %>
		<TD VALIGN=TOP><% = RS(i) %></TD>
	<% Next %>
	</TR>
	<%
	RS.MoveNext
Loop
'RS.Close
rs.MoveFirst
ID = 189
RS.Find "Id = " & ID
rs("TheAlias") = "ACTRL_RPA97_OVRRIDE_CURR_AMT"

Set fs = CreateObject("Scripting.FileSystemObject")
Set f1 = fs.GetFile(TheFullPath & ProjectName & "_tablespecs.gaga")
f1.delete


RS.save TheFullPath & ProjectName & "_tablespecs.gaga"

SET fs = NOTHING
SET f1 = NOTHING

rs.close


%>
</TABLE>
</center>


</BODY>
</HTML>
