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


TheFullPath = Request("TheFullPath")
ANewFolder = Request("ANewFolder")
ProjectName = Request("ProjectName")
TheDB = Request("TheDB")
TheTable = Request("TheTable")


Set objRS = Server.CreateObject("ADODB.RecordSet")

objRS.Open TheFullPath & ProjectName & "_tablespecs.gaga",  "Provider=MSPersist;", adOpenDynamic, adLockBatchOptimistic, adCmdFile






RecordCount = Request("RecordCount")

'Response.Write("The record count is: " & RecordCount & "<br>")


For I = 1 to RecordCount



ID = Request("ID")(I)

objRS.MoveFirst
objRS.Find "Id = " & ID

'Response.Write(objRS("ShowDupeOrder") & "<br>")

TheAlias = Request("frmTheAlias")(I)
objRS("TheAlias") = TheAlias
TheOrder = Request("FrmTheOrder")(I)
objRS("TheOrder") = TheOrder
ShowPulldown = Request("frmShowPulldown(" & I & ")")
objRS("ShowPulldown") = ShowPulldown
ShowProfile = Request("frmShowProfile(" & I & ")")
objRS("ShowProfile") = ShowProfile
ShowDupe = Request("frmShowDupe(" & I & ")")
objRS("ShowDupe") = ShowDupe
ShowDupeOrder = Request("frmShowDupeOrder")(I)
objRS("ShowDupeOrder") = ShowDupeOrder
ShowCalc = Request("frmShowCalc(" & I & ")")
objRS("ShowCalc") = ShowCalc


objRS.Update
 


'Response.Write("ID = " & ID & "<br>")
'Response.Write("TheAlias = " & TheAlias & "<br>")
'Response.Write("TheOrder = " & TheOrder & "<br>")
'Response.Write("ShowPulldown = " & ShowPulldown & "<br>")
'Response.Write("ShowProfile = " & ShowProfile & "<br>")
'Response.Write("ShowDupe = " & ShowDupe & "<br>")
'Response.Write("ShowDupeOrder = " & ShowDupeOrder & "<br>")
'Response.Write("--------------------" & "<br>")





Next

Set fs = CreateObject("Scripting.FileSystemObject")
Set f1 = fs.GetFile(TheFullPath & ProjectName & "_tablespecs.gaga")
f1.delete


objRS.save TheFullPath & ProjectName & "_tablespecs.gaga"

SET fs = NOTHING
SET f1 = NOTHING


'objRS.close


objRS.Sort = "TheOrder"

%>
<html>
<body>
<h2 align="center">Fine Tune Selections</h2>
<blockquote>
<b>Use this page to customize your pages. Below the main table you will find examples of how 
your pull-down menus will look for filtering, group-by queries and calculations. An example of the table 
that will display the results of filter queries is near the bottom. When you are finished customizing your 
pages, click "Next" at the bottom of the page.</b>
</blockquote>

<form method="POST" action="custom.asp">
	<INPUT TYPE="HIDDEN" NAME="RecordCount" VALUE="<% = RecordCount %>">
<INPUT TYPE="HIDDEN" NAME="TheFullPath" VALUE="<% = TheFullPath %>">
<INPUT TYPE="HIDDEN" NAME="ANewFolder" VALUE="<% = ANewFolder %>">
<INPUT TYPE="HIDDEN" NAME="ProjectName" VALUE="<% = ProjectName %>">
<INPUT TYPE="HIDDEN" NAME="TheDB" VALUE="<% = TheDB %>">
<INPUT TYPE="HIDDEN" NAME="TheTable" VALUE="<% = TheTable %>">

    <table border="3" cellpadding="0" cellspacing="2" align="center">
      <tr>
        <td align="center" bgcolor="#F7ECCA">Field</td>
        <td align="center" bgcolor="#F7ECCA">Alias</td>
        <td align="center" bgcolor="#F7ECCA">Order</td>
        <td align="center" colspan="2" bgcolor="#F7ECCA">Show<br>
          in Filter Pulldown</td>
        <td align="center" colspan="2" bgcolor="#F7ECCA">Show in<br> Profile</td>
        <td align="center" colspan="2" bgcolor="#F7ECCA">Show in <br>Filter Results</td>
        <td align="center" bgcolor="#F7ECCA">Filter <br>Results <br>Order</td>
        <td align="center" colspan="2" bgcolor="#F7ECCA">Show in <br>Calc <br>Pulldown</td>


      </tr>


<% 
J = 0
Do While NOT objRS.EOF 
J = J+1

If TheColor = "#FFFFFF" Then
	TheColor = "#F7ECCA"
Else
	TheColor = "#FFFFFF"
End if


%>
	<INPUT TYPE="HIDDEN" NAME="ID" VALUE="<% = objRS("ID") %>">

      <tr bgcolor="<% = thecolor %>">
        <td><b><% = objRS("TheField") %></b></td>
        <td><input type="text" name="frmTheAlias" size="35" value="<% = objRS("TheAlias") %>"></td>
        <td><input type="text" name="frmTheOrder" size="5" value="<% = objRS("TheOrder") %>"></td>

<% If objRS("ShowPulldown") = TRUE Then %>
        <td bgcolor="<% = thecolor %>">Y<input type="radio" name="<% = "frmShowPullDown(" & J & ")" %>"  value="TRUE" checked></td><td bgcolor="#FDC0AC" bordercolordark="black" bordercolorlight="gray">N<input type="radio" name="<% = "frmShowPullDown(" & J & ")" %>" value="FALSE"></td>
<% Else %>
        <td bgcolor="<% = thecolor %>">Y<input type="radio" name="<% = "frmShowPullDown(" & J & ")" %>"  value="TRUE"></td><td bgcolor="#FDC0AC"  bordercolordark="black" bordercolorlight="gray">N<input type="radio" name="<% = "frmShowPullDown(" & J & ")" %>" value="FALSE" checked></td>
<% End If
If objRS("ShowProfile") = TRUE Then %>
        <td bgcolor="<% = thecolor %>">Y<input type="radio" name="<% = "frmShowProfile(" & J & ")" %>"  value="TRUE" checked></td><td bgcolor="#FDC0AC">N<input type="radio" name="<% = "frmShowProfile(" & J & ")" %>" value="FALSE"></td>
<% Else %>
        <td bgcolor="<% = thecolor %>">Y<input type="radio" name="<% = "frmShowProfile(" & J & ")" %>"  value="TRUE"></td><td bgcolor="#FDC0AC">N<input type="radio" name="<% = "frmShowProfile(" & J & ")" %>" value="FALSE" checked></td>
<% End If

If objRS("ShowDupe") = TRUE Then %>
        <td bgcolor="<% = thecolor %>">Y<input type="radio" name="<% = "frmShowDupe(" & J & ")" %>"  value="TRUE" checked></td><td bgcolor="#FDC0AC">N<input type="radio" name="<% = "frmShowDupe(" & J & ")" %>" value="FALSE"></td>
<% Else %>
        <td bgcolor="<% = thecolor %>">Y<input type="radio" name="<% = "frmShowDupe(" & J & ")" %>"  value="TRUE"></td><td bgcolor="#FDC0AC">N<input type="radio" name="<% = "frmShowDupe(" & J & ")" %>" value="FALSE" checked></td>
<% End If
%>

        <td><input type="text" name="frmShowDupeOrder" size="5" value="<% = objRS("ShowDupeOrder") %>"></td>

<%
If (objRS("numeric_precision") > 5) and (objRS("TheType") <> "timestamp") Then
	If objRS("ShowCalc") = TRUE Then %>
        	<td bgcolor="<% = thecolor %>">Y<input type="radio" name="<% = "frmShowCalc(" & J & ")" %>"  value="TRUE" checked></td><td bgcolor="#FDC0AC">N<input type="radio" name="<% = "frmShowCalc(" & J & ")" %>" value="FALSE"></td>
	<% Else %>
        	<td bgcolor="<% = thecolor %>">Y<input type="radio" name="<% = "frmShowCalc(" & J & ")" %>"  value="TRUE"></td><td bgcolor="#FDC0AC">N<input type="radio" name="<% = "frmShowCalc(" & J & ")" %>" value="FALSE" checked></td>
	<% End If
Else %>
	<td bgcolor="<% = thecolor %>">&nbsp;</td><td bgcolor="#FDC0AC">&nbsp;</td>
	<INPUT TYPE="HIDDEN" NAME="<% = "frmShowCalc(" & J & ")" %>" VALUE="FALSE">

      </tr>

<% End If 


	objRS.MoveNext
Loop 

'objRS.Close

%>
       <tr>
        <td align="center" bgcolor="#F7ECCA">Field</td>
        <td align="center" bgcolor="#F7ECCA">Alias</td>
        <td align="center" bgcolor="#F7ECCA">Order</td>
        <td align="center" colspan="2" bgcolor="#F7ECCA">Show<br>
          in Filter Pulldown</td>
        <td align="center" colspan="2" bgcolor="#F7ECCA">Show in<br> Profile</td>
        <td align="center" colspan="2" bgcolor="#F7ECCA">Show in <br>Filter Results</td>
        <td align="center" bgcolor="#F7ECCA">Filter <br>Results <br>Order</td>
        <td align="center" colspan="2" bgcolor="#F7ECCA">Show in <br>Calc <br>Pulldown</td>
</tr>
</table>
  <p>
<center><input type="Submit" value="Submit" name="B1"><input type="reset" value="Reset" name="B2"></center></p>
</form>

<hr>
<%

'--------------------------- make a sample pulldown menu --------------------------------------

objRS.MoveFirst

objRS.Filter = "[ShowPulldown] = TRUE"




'PDS is Pulldown String
PDS = "<select size =""1"" name=""SamplePulldown"">"&vbCrLf




Do While NOT objRS.EOF

	PDS=PDS & "<option value=" & chr(34) & objRS("TheField") & chr(34) & ">" & objRS("TheAlias") & "</option>" & vbCrLf

	objRS.MoveNext

Loop

PDS=PDS & "</select>"

%>

<h2 align="center">This is Your Drop-down Menu for Filtering and Group-by Queries</h2>
<center>
<% = PDS %></center>
<hr>
<%

'---------------------------------------------------------------------

objRS.Filter=""
objRS.MoveFirst

objRS.Filter = "[ShowCalc] = TRUE"

Countem=objRS.RecordCount


If Countem >= 1 Then 
'PDS is Pulldown String
PDS = ""
PDS = "<select size =""1"" name=""SampleCalcPulldown"">"&vbCrLf




Do While NOT objRS.EOF

	PDS=PDS & "<option value=" & chr(34) & objRS("TheField") & chr(34) & ">" & objRS("TheAlias") & "</option>" & vbCrLf

	objRS.MoveNext

Loop

PDS=PDS & "</select>"

%>

<h2 align="center">This is Your Drop-down Menu for Calculations</h2>
<center>
<% = PDS %></center>

<% Else %>

<h2 align="center">There are no fields available for Calcuations.<br> The calculations page will not be built.</h2>



<% End If %>

<hr>

<%

objRS.Filter=""

objRS.MoveFirst


objRS.Filter = "[ShowDupe] = TRUE"

objRS.Sort = "ShowDupeOrder"

DupeCount = objRS.RecordCount

If DupeCount <= 0 Then %>

<h2 align="center">You did not select any fields to show in filter query results.<br>
You must select at least one field to continue.</h2>

<% Else

sql = "Select top 5 * from " & TheTable

dbName= TheDB %>
     
     
<!--#include virtual = "/thetimes/rcon.inc" -->

<% 	RS.Open sql, Conn, adOpenStatic, adLockReadOnly, adCmdText 

RSCounter = 5

TableCount = RS.RecordCount

If TableCount < 5 Then
	RSCounter = TableCount
End If

%>
<h2 align="center">This is how the results of filter queries will look.<br>
You will be able to order the results in the next step.</h2>

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


<% End If %>

<form method="POST" action="orderby.asp">
<INPUT TYPE="HIDDEN" NAME="RecordCount" VALUE="<% = RecordCount %>">
<INPUT TYPE="HIDDEN" NAME="TheFullPath" VALUE="<% = TheFullPath %>">
<INPUT TYPE="HIDDEN" NAME="ANewFolder" VALUE="<% = ANewFolder %>">
<INPUT TYPE="HIDDEN" NAME="ProjectName" VALUE="<% = ProjectName %>">
<INPUT TYPE="HIDDEN" NAME="TheDB" VALUE="<% = TheDB %>">
<INPUT TYPE="HIDDEN" NAME="TheTable" VALUE="<% = TheTable %>">

<center><input type="Submit" value="Next"></center></p>
</form>



</body>
</html>





<%
objRS.Close
%>








