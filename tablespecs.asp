<% Response.Buffer = FALSE %>
<!--#include virtual ="/thetimes/adovbs.inc"-->
<%  

TheFullPath = Request("TheFullPath")
ANewFolder = Request("ANewFolder")
ProjectName = Request("ProjectName")
TheDB = Request("TheDB")
TheTable = Request("TheTable")
LField = Request("LField")

'Response.Write("The Full Path is: " & TheFullPath & "<br>")


  ' Declare all ADO Objects to be Used
  Dim objRS                 ' ADO Recordset Obejct

  Set objRS = Server.CreateObject("ADODB.Recordset")    ' Create the Recordset
  objRS.CursorType = adOpenStatic       ' Specify the cursor location and type
  objRS.CursorLocation = adUseClient

  ' Attempt to opend the file
  ' If the File does not exist an error will be raised


' rename the file and kill the old on
Const ForReading = 1, ForWriting = 2, ForAppending = 3
Set fs = CreateObject("Scripting.FileSystemObject")
' open the source file for reading
If fs.FileExists(TheFullPath & ProjectName & "_tablespecs.gaga") Then


	Set f1 = fs.GetFile(TheFullPath & ProjectName & "_tablespecs.gaga")
	f1.Copy (TheFullPath & ProjectName & "_tablespecs.gaga.old")
	f1.delete

End If



On Error Resume Next 

  objRS.Open TheFullPath & ProjectName & "_tablepecs.gaga"
 ' If the file does not exist we will get error no 3709
  If err.number = 3709 Then
    err.number = 0                                 ' Clear the error number

    objRS.Fields.Append "ID", adInteger
    objRS.Fields.Append "TheField", adVarChar, 150    ' Create the Recordset
    objRS.Fields.Append "TheAlias", adVarChar, 150
    objRS.Fields.Append "TheOrder", adInteger
    objRS.Fields.Append "TheType", adVarchar, 50
    objRS.Fields.Append "Numeric_Precision", adInteger
    objRS.Fields.Append "ShowPulldown", adBoolean
    objRS.Fields.Append "ShowProfile", adBoolean
    objRS.Fields.Append "ShowDupe", adBoolean
    objRS.Fields.Append "ShowDupeOrder", adInteger
    objRS.Fields.Append "ShowCalc", adBoolean


    objRS.Open
  End If




' ------------- PUT SQL STATEMENT HERE --------------------

sql = "select ordinal_position as ID, column_name as TheField, column_name as TheAlias, ordinal_position * 100 as TheOrder, "
sql=sql & " Data_Type as TheType, Numeric_Precision "
sql=sql & " from information_schema.columns where table_name = '" & TheTable & "' order by ordinal_Position"

' Response.Write("The SQL is: " & sql & "<br>")

'------------ PUT DATABASE NAME HERE  ---------------------
dbName = TheDB %>



<!--#include virtual = "/thetimes/rcon.inc" -->

<%


RS.Open sql, Conn, adOpenStatic, adLockReadOnly, adCmdText



Do While NOT rs.EOF

		objRS.AddNew
		objRS("ID") = rs("ID")
		objRS("TheField") = rs("TheField")
		objRS("TheAlias") = rs("TheAlias")
		objRS("TheOrder") = rs("TheOrder")
		objRS("TheType") = rs("TheType")
		objRS("Numeric_Precision") = rs("Numeric_Precision")
		If objRS("TheType") = "timestamp" Then
			objRS("ShowPulldown") = FALSE
			objRS("ShowProfile") = FALSE
		Else
			objRS("ShowPulldown") = TRUE
			objRS("ShowProfile") = TRUE
		End If
		objRS("ShowDupe") = FALSE
		objRS("ShowDupeOrder") = rs("TheOrder")
		If objRS("numeric_precision") > 5 and objRS("TheType") <> "timestamp" Then
			objRS("ShowCalc") = TRUE
		Else
			objRS("ShowCalc") = FALSE
		End If
		




rs.MoveNext
Loop


objRS.UpdateBatch

  objRS.Save TheFullPath & ProjectName & "_tablespecs.gaga"
  ' If the file exist we will get an error no 58
  If err.number = 58 Then
    ' If it exist we can just issue the save method 
    ' to save it the currently opened path	
Set fs = CreateObject("Scripting.FileSystemObject")
Set f1 = fs.GetFile(TheFullPath & ProjectName & "_progspecs.gaga")
f1.delete


objRS.save TheFullPath & ProjectName & "_progspecs.gaga"

set f1 = NOTHING
set fs = NOTHING
  End If




'objRS.MoveFirst


objRS.Sort = ("TheOrder")    ' Sort the Recorset





















RecordCount = objRS.RecordCount

%>

<html>
<body bgcolor="#FFCC99">
<h1 align="center">Table Specs</h1>
<p align="left"><b>This page lets you customize a number of features relative to
the table or view you selected. The options for each column are explained below
the table. When you click Submit, you will see how the options you selected will
appear on input and results pages. You will also see another copy of this form
so you can fine tune your selections.</b></p>



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

objRS.Close

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
<center><input type="submit" value="Submit" name="B1"><input type="reset" value="Reset" name="B2"></center></p>
</form>

  <table align="center" border="1">
    <tr>
      <td><b>Field</b></td>
      <td>The name of the field. Unchangeable.</td>
    </tr>
    <tr>
      <td><b>Alias</b></td>
      <td>You can assign a plain-English Name to a field. Spaces are OK. 150
        Chars max.</td>
    </tr>
    <tr>
      <td><b>Order</b></td>
      <td>The order that fields will appear in filter and group-by pulldowns and
        in the profile. Lower numbers will be shown first.</td>
    </tr>
    <tr>
      <td><b>Show in Filter Pulldown</b></td>
      <td>Determine whether a field should appear in the filter and group by
        pulldowns</td>
    </tr>
    <tr>
      <td><b>Show in Profile</b></td>
      <td>Determine whether a field should be shown on the profile page.</td>
    </tr>
    <tr>
      <td><b>Show in Filter Results</b></td>
      <td>Filter queries have fixed results sets. Select the fields you want to
        show in filter query results. The same fields will be displayed on
        drilldowns from the group-by and calculations results pages.</td>
    </tr>
    <tr>
      <td><b>Filter Results Order</b></td>
      <td>This column lets you order the columns for the filter results set. You
        can set the sort order on a later page.</td>
    </tr>
    <tr>
      <td><b>Show in Calc Pulldown</b></td>
      <td>You will see options only for fields amenable to calculations. If you
        deselect all the available fields no calculations page will be created. Data types "tinyint" and "smallint" are not 
	considered amenable to calculations.</td>
    </tr>
  </table>



</body>
</html>


