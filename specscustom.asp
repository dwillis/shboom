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


Set RS = Server.CreateObject("ADODB.RecordSet")

RS.Open TheFullPath & ProjectName & "_progspecs.gaga",  "Provider=MSPersist;", adOpenDynamic, adLockBatchOptimistic, adCmdFile



rs("ProjectAlias")=Trim(Request("frmProjectAlias"))
rs("MaxFields")=Trim(Request("frmMaxFields"))
rs("FilterTitle")=Trim(Request("frmFilterTitle"))
rs("GroupByTitle")=Trim(Request("frmGroupByTitle"))
rs("CalcTitle")=Trim(Request("frmCalcTitle"))
rs("MaxRows")=Trim(Request("frmMaxRows"))
rs("MaxRecords")=Trim(Request("frmMaxRecords"))
rs("AllowUserMod")=Trim(Request("frmAllowUserMod"))
rs("DisplaySQL")=Trim(Request("frmDisplaySQL"))
rs.Update
 
Set fs = CreateObject("Scripting.FileSystemObject")
Set f1 = fs.GetFile(TheFullPath & ProjectName & "_progspecs.gaga")
f1.delete


RS.save TheFullPath & ProjectName & "_progspecs.gaga"

set f1 = NOTHING
set fs = NOTHING


'rs.save
'rs.close

DaLinks = "<center><small><font color=""blue""><u>" & rs("FilterTitle") & " </u> | "
DaLinks=DaLinks & " <u>" & rs("GroupbyTitle") &  " </u> | "
DaLinks=DaLinks & " <u>" & rs("CalcTitle") &  " </u> | <u>Search Tips</u>"

'response.write(rs("AllowUserMod") & "<br>")
If rs("AllowUserMod") = "YES" Then
	DaLinks = DaLinks & " | <u>Modify</u></font></small></center>"
Else
	DaLinks = DaLinks & "</font></small></center>"
End If





%>

<html>
<head>
<title>Specs Page</title>
<style>
<!--
select       { font-size: 7pt }
-->
</style>
</head>
<body>
<table width="100%"  bgcolor="white">
<tr>
<td>
<h1 align="center">Modify Input Pages</h1>
<blockquote><b>The three input pages for your project will look like the sample pages below. If you want 
to change specifications enter the changes in the appropriate input box(es) and click on "Submit Changes." (You're
stuck with the ugly colors.) If the pages are OK, click on "Next."</b>
</blockquote>


<P>
<center>
<form method="POST" action="specscustom.asp">
<INPUT TYPE="HIDDEN" NAME="TheFullPath" VALUE="<% = TheFullPath %>">
<INPUT TYPE="HIDDEN" NAME="ANewFolder" VALUE="<% = ANewFolder %>">
<INPUT TYPE="HIDDEN" NAME="ProjectName" VALUE="<% = ProjectName %>">
<INPUT TYPE="HIDDEN" NAME="TheDB" VALUE="<% = TheDB %>">
<INPUT TYPE="HIDDEN" NAME="TheTable" VALUE="<% = TheTable %>">





  <div align="center">
    <center>
    <table border="1">
      <tr>
        <td align="right"><b>Project Name to Appear on Each Page</td></b>
        <td><b><input type="text" name="frmProjectAlias" size="75" value="<% = Trim(rs("ProjectAlias")) %>"></td></b>
      </tr>
      <tr>
        <td align="right"><b>Max number of fields for filtering or group by queries</td></b>
        <td><b><input type="text" name="frmMaxFields" size="5"  value="<% = Trim(rs("MaxFields")) %>"></td></b>
      </tr>
      <tr>
        <td align="right"><b>Title for Filter Page</td></b>
        <td><b><input type="text" name="frmFilterTitle" size="75" value="<% = Trim(rs("FilterTitle")) %>"></td></b>
      </tr>
      <tr>
        <td align="right"><b>Title for Group-by Page</td></b>
        <td><b><input type="text" name="frmGroupByTitle" size="75" value="<% = Trim(rs("GroupByTitle")) %>"></td></b>
      </tr>
      <tr>
        <td align="right"><b>Title for Calculations Page</td></b>
        <td><b><input type="text" name="frmCalcTitle" size="75" value="<% = Trim(rs("CalcTitle")) %>"></td></b>
      </tr>
      <tr>
        <td align="right"><b>Number of rows to display for filter query results</td></b>
        <td><b><input type="text" name="frmMaxRows" size="8" value="<% = Trim(rs("MaxRows")) %>"></td></b>
      </tr>
      <tr>
        <td align="right"><b>Max number of records to return for filter query results</td></b>
        <td><b><input type="text" name="frmMaxRecords" size="10" value="<% = Trim(rs("MaxRecords")) %>"></td></b>
      </tr>
      <tr>
        <td align="right"><b>Allow User to Modify Application</td></b>
        <td><b><select size="1" name="frmAllowUserMod">

<% If rs("AllowUserMod") = "YES" Then %>

            <option value="NO">No</option>
            <option selected value="YES">Yes</option>

<% Else %>

            <option selected value="NO">No</option>
            <option value="YES">Yes</option>

<% End If %>

          </select></td></b>
      </tr>
      <tr>
        <td align="right"><b>Display SQL Statement for Group By and Calculations results</td></b>
        <td><b><select size="1" name="frmDisplaySQL">
<% If rs("DisplaySQL") = "YES" Then %>
            <option value="NO">No</option>
            <option selected value="YES">Yes</option>

<% Else %>

            <option selected value="NO">No</option>
            <option value="YES">Yes</option>

<% End If %>

          </select></td></b>
      </tr>
    </table>
    </center>
  </div>
  <p>
<center><input type="submit" value="Submit Changes" name="B1"><input type="reset" value="Reset" name="B2"></center></p>
</form>

<form method="POST" action="tablespecs.asp">
<INPUT TYPE="HIDDEN" NAME="TheFullPath" VALUE="<% = TheFullPath %>">
<INPUT TYPE="HIDDEN" NAME="ANewFolder" VALUE="<% = ANewFolder %>">
<INPUT TYPE="HIDDEN" NAME="ProjectName" VALUE="<% = ProjectName %>">
<INPUT TYPE="HIDDEN" NAME="TheDB" VALUE="<% = TheDB %>">
<INPUT TYPE="HIDDEN" NAME="TheTable" VALUE="<% = TheTable %>">


<center><input type="submit" value="Next" name="B1"></center>
</form>
</td>
</tr>
</table>


<hr width="100%">
<table width="100%" bgcolor="#ECF5FF">
<tr><td>
<% = DaLinks %>


<%
ProjAlias = "<h1 align=""center"">" & rs("ProjectAlias") & "</h1>"
FiltTitle = "<h2 align=""center"">" & rs("FilterTitle") & "</h2>"
FiltMess = "<h4 align=""center"">Select records that meet the following conditions:</h4>"

FiltTop = ProjAlias & vbCrLf & FiltTitle & vbCRLF & FiltMess & vbCrLf
FiltTable = "<center>" & vbCrLf
FiltTable = FiltTable & "<table>" & vbCrLf
FiltTable = FiltTable & "<tr>" & vbCrLf
FiltTable = FiltTable & "  <td align=""center""><b>Field</b></td>" & vbCrLf
FiltTable = FiltTable & "  <td align=""center""><b>Show Rows Where</b></td>" & vbCrLf
FiltTable = FiltTable & "  <td align=""center""><b>Value 1</td></b>" & vbCrLf
FiltTable = FiltTable & "  <td align=""center""><b>And / or</td></b>" & vbCrLf
FiltTable = FiltTable & "  <td align=""center""><b>Show Rows Where</b></td>" & vbCrLf
FiltTable = FiltTable & "  <td align=""center""><b>Value 2</td></b>" & vbCrLf
FiltTable = FiltTable & "" & vbCrLf
FiltTable = FiltTable & "</tr>" & vbCrLf


' begin building the row with field and option pulldowns

FieldsSelect = "<tr>" & vbCrLf
FieldsSelect = FieldsSelect & "<td align=""center""><select size=""1"" name=""frmSource"">" & vbCrLf
FieldsSelect = FieldsSelect & "<option value="""""" selected></option>" & vbCrLf
FieldsSelect = FieldsSelect & "<option value=""SampleField1"">Sample Field 1</option>" & vbCrLf
FieldsSelect = FieldsSelect & "<option value=""SampleField2"">Sample Field 2</option>" & vbCrLf
FieldsSelect = FieldsSelect & "<option value=""SampleField3"">Sample Field 3</option>" & vbCrLf
FieldsSelect = FieldsSelect & "</select>" & vbCrLf
FieldsSelect = FieldsSelect & "</td>" & vbCrLf

TheOptions = TheOptions & "<td align=""center""><select size=""1"" name=""frmFilter1"">" & vbCrLf
TheOptions = TheOptions & "<option value="""" selected></option>" & vbCrLf
TheOptions = TheOptions & "<option value=""eq"">equals</option>" & vbCrLf
TheOptions = TheOptions & "<option value=""noteq"">is not equal to</option>" & vbCrLf
TheOptions = TheOptions & "<option value=""gt"">is greater than</option>" & vbCrLf
TheOptions = TheOptions & "<option value=""gtet"">is greater than or equal to</option>" & vbCrLf
TheOptions = TheOptions & "<option value=""lt"">is less than</option>" & vbCrLf
TheOptions = TheOptions & "<option value=""ltet"">is less than or equal to</option>" & vbCrLf
TheOptions = TheOptions & "<option value=""bw"">begins with</option>" & vbCrLf
TheOptions = TheOptions & "<option value=""dnbw"">does not begin with</option>" & vbCrLf
TheOptions = TheOptions & "<option value=""ew"">ends with</option>" & vbCrLf
TheOptions = TheOptions & "<option value=""dnew"">does not end with</option>" & vbCrLf
TheOptions = TheOptions & "<option value=""con"">contains</option>" & vbCrLf
TheOptions = TheOptions & "<option value=""dncon"">does not contain</option>" & vbCrLf
TheOptions = TheOptions & "<option value=""blank"">Is Blank</option>" & vbCrLf
TheOptions = TheOptions & "<option value=""notblank"">Is Not Blank</option>" & vbCrLf
TheOptions = TheOptions & "</select>" & vbCrLf
TheOptions = TheOptions & "</td>" & vbCrLf


ForVal1 = "</td>" & vbCrLf
ForVal1 = ForVal1 & "<td align=""center""><input type=""text"" name=""frmValue1""  size=""10""></td>" & vbCrLf


ForVal2 = "</td>" & vbCrLf
ForVal2 = ForVal2 & "<td align=""center""><input type=""text"" name=""frmValue2""  size=""10""></td>" & vbCrLf


TheAndOr = TheAndOr & "<td align=""center""><select size=""1"" name=""frmAndOr"">" & vbCrLf
TheAndOr = TheAndOr & "<option value="""" selected></option>" & vbCrLf
TheAndOr = TheAndOr & "<option value=""and"">And</option>" & vbCrLf
TheAndOr = TheAndOr & "<option value=""or"">Or</option>" & vbCrLf
TheAndOr = TheAndOr & "</select>" & vbCrLf
TheAndOr = TheAndOr & "</td>" & vbcrLF


FilterRow = FieldsSelect & TheOptions & ForVal1 & TheAndOr & TheOptions & ForVal2 & "</tr>" & vbCrLF

TheButtons = "<center><p><input type=""submit"" value=""Submit""><input " & vbCRFL
TheButtons=TheButtons & " type=""reset"" value=""Reset"" name=""B2""></p></center>"


TheFilter = FiltTop & TheButtons & FiltTable



%>

<% = TheFilter  %>

<% For I = 1 to rs("MaxFields") %>

<% = FilterRow %>

<% Next %>



</table>
<% = TheButtons %>
</td></tr></table>
<hr width="100%">
<table width="100%" bgcolor="#FFCC99">
<tr>
<td>
<% = DaLinks %>


<%

GBTitle = "<h2 align=""center"">" & rs("GroupByTitle") & "</h2>"
GBMess = "<h4 align=""center"">Fields will be grouped, displayed and counted in the order you select</h4>"

TheGBTop="<center>"&vbCRLF
TheGBTop=TheGBTop & "<table>"&vbCRLF
TheGBTop=TheGBTop & "<tr>"&vbCRLF
TheGBTop=TheGBTop & "  <td align=""center""><b>Select one or more fields to group by</b></td>"&vbCRLF
TheGBTop=TheGBTop & "</tr>"&vbCRLF
TheGBTop=TheGBTop & ""&vbCRLF
TheGBTop=TheGBTop & ""&vbCRLF
TheGBTop=TheGBTop & ""&vbCRLF








GBTOP = ProjAlias & vbCrLf & GBTitle & vbCRLF & GBMess & vbCrLf & TheButtons & TheGBTOP



GBPullDown="<tr>"&vbCRLF
GBPullDown=GBPullDown & "<td align=""center""><select size=""1"" name=""frmSource"">"&vbCRLF
GBPullDown=GBPullDown & "<option value="""""" selected></option>"&vbCRLF
GBPullDown=GBPullDown & "<option value=""SampleField1"">Sample Field 1</option>"&vbCRLF
GBPullDown=GBPullDown & "<option value=""SampleField2"">Sample Field 2</option>"&vbCRLF
GBPullDown=GBPullDown & "<option value=""SampleField3"">Sample Field 3</option>"&vbCRLF
GBPullDown=GBPullDown & "</select>" & vbCRLF
GBPullDown=GBPullDown & "</tr>" & vbCRLF



%>

<% = GBTOP %>


<% For I = 1 to rs("MaxFields") %>

<% = GBPullDown %>

<% Next %>
</table>
<p>


<% = FiltMess & FiltTable %>







<% For I = 1 to rs("MaxFields") %>

<% = FilterRow %>

<% Next %>



</table>
<% = TheButtons %>
</td>
</tr>
</table>


<hr width="100%">
<table width="100%" bgcolor="#CCFFCC">
<tr>
<td>
<% = DaLinks %>


<% 

CTitle = "<h2 align=""center"">" & rs("CalcTitle") & "</h2>"



CTopper="<center>" & vbCrLF
CTopper=CTopper & "<table align=""center"" border=""1"">" & vbCrLF
CTopper=CTopper & "<tr>" & vbCrLF
CTopper=CTopper & "<td align=""center""><b>Do Calculations on:</b></td>" & vbCrLF
CTopper=CTopper & "<td align=""center""><b>Calculate:</b></td>" & vbCrLF
CTopper=CTopper & "</tr>" & vbCrLF
CTopper=CTopper & "" & vbCrLF
CTopper=CTopper & "<td align=""center""><select size=""1"" name=""frmExp"">" & vbCrLF
CTopper=CTopper & "<option value="""""" selected></option>" & vbCrLF
CTopper=CTopper & "<option value=""NumericOrDateField1"">Numeric or Date Field 1</option>" & vbCrLF
CTopper=CTopper & "<option value=""NumericOrDateField2"">Numeric or Date Field 2</option>" & vbCrLF
CTopper=CTopper & "<option value=""NumericOrDateField3"">Numeric or Date Field 3</option>" & vbCrLF
CTopper=CTopper & "</select>" & vbCrLF
CTopper=CTopper & "" & vbCrLF
CTopper=CTopper & "</td>" & vbCrLF
CTopper=CTopper & "<td align=""center"">" & vbCrLF
CTopper=CTopper & "<select size=""1"" name=""frmCalc"">" & vbCrLF
CTopper=CTopper & "<option value=""AVG"">Average</option>" & vbCrLF
CTopper=CTopper & "<option value=""COUNT"">Count</option>" & vbCrLF
CTopper=CTopper & "<option value=""MIN"">Minimum</option>" & vbCrLF
CTopper=CTopper & "<option value=""Max"">Maximum</option>" & vbCrLF
CTopper=CTopper & "<option value=""SUM"">Sum</option>" & vbCrLF
CTopper=CTopper & "</select>" & vbCrLF
CTopper=CTopper & "</td>" & vbCrLF
CTopper=CTopper & "</tr>" & vbCrLF
CTopper=CTopper & "</table>" & vbCrLF

CalcTop = ProjAlias & vbCrLf & CTitle & vbCRLF & TheButtons & CTopper


%>

<% = CalcTop %>

<%

GBTOP = GBMess & vbCrLf & TheGBTOP


%>

<% = GBTOP %>


<% For I = 1 to rs("MaxFields") %>

<% = GBPullDown %>

<% Next %>
</table>
<p>


<% = FiltMess & FiltTable %>







<% For I = 1 to rs("MaxFields") %>

<% = FilterRow %>

<% Next %>



</table>
<% = TheButtons %>
</td>
</tr>
</table>

<hr width="100%">


</body>
</html>






