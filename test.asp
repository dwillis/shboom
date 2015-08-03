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
Set RS = Server.CreateObject("ADODB.RecordSet")
Set objRS2 = Server.CreateObject("ADODB.RecordSet")

RS.Open "c:\inetpub\wwwroot\steststaff_tablespecs.gaga",  "Provider=MSPersist;", adOpenDynamic, adLockBatchOptimistic, adCmdFile


RS.Filter="TheType like *date*"

rs.MoveFirst

Do While Not rs.Eof

	Response.Write(rs("TheType") & "<br>")

	RS.MoveNext
Loop
rs.Close
%>
