<% Response.Buffer = FALSE %>
<!--#include virtual ="/thetimes/adovbs.inc"-->
<%  

TheFullPath = Request("TheFullPath")
ANewFolder = Request("ANewFolder")
ProjectName = Request("ProjectName")
TheDB = Request("TheDB")
TheTable = Request("TheTable")
LField = Request("LField")

If right(TheFullPath, 1) <> "\" Then
	TheFullPath = TheFullPath & "\"
End If



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
If fs.FileExists(TheFullPath & ProjectName & "_progspecs.gaga") Then


	Set f1 = fs.GetFile(TheFullPath & ProjectName & "_progspecs.gaga")
	f1.Copy (TheFullPath & ProjectName & "_progspecs.gaga.old")
	f1.delete

End If



On Error Resume Next 

  objRS.Open TheFullPath & ProjectName & "_progspecs.gaga"
 ' If the file does not exist we will get error no 3709
  If err.number = 3709 Then
    err.number = 0                                 ' Clear the error number



    objRS.Fields.Append "TheFullPath", adVarChar, 150
    objRS.Fields.Append "ProjectName", adVarChar, 150    ' Create the Recordset
    objRS.Fields.Append "TheDB", adVarChar, 150
    objRS.Fields.Append "TheTable", adVarChar, 150
    objRS.Fields.Append "LinkingField", adVarChar, 150
    objRS.Fields.Append "ProjectAlias", adVarChar, 150
    objRS.Fields.Append "MaxFields", adInteger
    objRS.Fields.Append "FilterTitle", adVarChar, 150
    objRS.Fields.Append "GroupByTitle", adVarChar, 150
    objRS.Fields.Append "CalcTitle", adVarChar, 150
    objRS.Fields.Append "MaxRows", adInteger
    objRS.Fields.Append "MaxRecords", adInteger
    objRS.Fields.Append "AllowUserMod", adVarChar, 3
    objRS.Fields.Append "DisplaySQL", adVarChar, 3
    objRS.Fields.Append "OrderBy1", adVarchar, 150
    objRS.Fields.Append "SortOrder1", adVarchar, 4
    objRS.Fields.Append "OrderBy2", adVarchar, 150
    objRS.Fields.Append "SortOrder2", adVarchar, 4
    objRS.Fields.Append "OrderBy3", adVarchar, 150
    objRS.Fields.Append "SortOrder3", adVarchar, 4




    objRS.Open
  End If




		objRS.AddNew
		objRS("TheFullPath") = TheFullPath
		objRS("ProjectName") = ProjectName
		objRS("TheDB") = TheDB
		objRS("TheTable") = TheTable
		objRS("LinkingField") = LField

objRS.Update

  objRS.Save TheFullPath & ProjectName & "_progspecs.gaga"
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









%>


<html>

<head>
<title>Specs Page</title>
</head>

<body bgcolor="white">
<h1 align="center">Project Specs</h1>
<blockquote>
<p align="left"><b>Use this page to specify a project alias that will appear on the
top of each input page, the maximum number of fields to be used for filter or
group-by queries, and the maximum number of rows to display and records to
return for filter queries. You can also specify whether users will be allowed to
modify the application and whether to show sql queries with the results of group
by and calculation queries.</p>
<p align="left">Click on the link for items marked with (H) to see help on that item.</b></p>
</blockquote>

<form method="POST" action="specscustom.asp">
<INPUT TYPE="HIDDEN" NAME="TheFullPath" VALUE="<% = TheFullPath %>">
<INPUT TYPE="HIDDEN" NAME="ANewFolder" VALUE="<% = ANewFolder %>">
<INPUT TYPE="HIDDEN" NAME="ProjectName" VALUE="<% = ProjectName %>">
<INPUT TYPE="HIDDEN" NAME="TheDB" VALUE="<% = TheDB %>">
<INPUT TYPE="HIDDEN" NAME="TheTable" VALUE="<% = TheTable %>">



  <div align="center">
    <center>
    <table border="0">
      <tr>
        <td align="right"><b>Project Name to Appear on Each Page</td></b>
        <td><b><input type="text" name="frmProjectAlias" size="75"></td></b>
      </tr>
      <tr>
        <td align="right"><b>Max number of fields for filtering or group by queries <a href="#Maxnumber">(H)</a>
</td></b>
        <td><b><input type="text" name="frmMaxFields" size="5" value="5"></td></b>
      </tr>
      <tr>
        <td align="right"><b>Title for Filter Page</td></b>
        <td><b><input type="text" name="frmFilterTitle" size="75" value="Filter"></td></b>
      </tr>
      <tr>
        <td align="right"><b>Title for Group-by Page</td></b>
        <td><b><input type="text" name="frmGroupByTitle" size="75" value="Group By"></td></b>
      </tr>
      <tr>
        <td align="right"><b>Title for Calculations Page</td></b>
        <td><b><input type="text" name="frmCalcTitle" size="75" value="Calculations"></td></b>
      </tr>
      <tr>
        <td align="right"><b>Number of rows to display for filter query results <a href="#Maxfields">(H)</a>
</td></b>
        <td><b><input type="text" name="frmMaxRows" size="8" value="50"></td></b>
      </tr>
      <tr>
        <td align="right"><b>Max number of records to return for filter query results <a href="#Numberofrows">(H)</a>
</td></b>
        <td><b><input type="text" name="frmMaxRecords" size="10" value="1000"></td></b>
      </tr>
      <tr>
        <td align="right"><b>Allow User to Modify Application <a href="#Allow">(H)</a>
</td></b>
        <td><b><select size="1" name="frmAllowUserMod">
            <option selected value="NO">No</option>
            <option value="YES">Yes</option>
          </select></td></b>
      </tr>
      <tr>
        <td align="right"><b>Display SQL Statement for Group By and Calculations results <a href="#Display">(H)</a>
</td></b>
        <td><b><select size="1" name="frmDisplaySQL">
            <option selected value="NO">No</option>
            <option value="YES">Yes</option>
          </select></td></b>
      </tr>
    </table>
    </center>
  </div>
  <p>
<center><input type="submit" value="Submit" name="B1"><input type="reset" value="Reset" name="B2"></center></p>
</form>
<p align="left"><b><a name="Maxfields">Max</a> number of fields for filtering or
group by queries:</b></p>
<p align="left">This option lets you determine how many fields can be included
in filter or group by queries. It's rare where you'll want more than three or
four but you can enter as many as you like.</p>
<p align="left"><b><a name="Numberofrows">Number</a> of rows to display for
filter query results:</b></p>
<p align="left">Here you can specify how many rows will be returned for each
page of your filter query results. If more rows are returned than specified,
buttons will be included to fetch spillover pages.</p>
<p align="left"><b><a name="Maxnumber">Max</a> number of records to return for
filter query results:</b></p>
<p align="left">When large data sets are returned, the entire dataset is fetched
before results are parceled to pages. You can cut down server load by imposing a
low limit of max records to return. If there are more records that were not
fetched the user will be notified.</p>
<p align="left"><b><a name="Allow">Allow </a>User to Modify Application:</b></p>
<p align="left">This option allows users to change just about everything in the
app but the database, table and unique id. You should not select this option if
you are moving the app to a secure server.</p>
<p align="left"><b><a name="Display">Display </a>SQL Statement for Group-by and
Calculations results:</b></p>
<p align="left">Sometimes it's convenient to store results of group-by and
calculations queries in Excel, where the results can be sorted and filtered. If
you select the option to display the SQL statement, the sql used to generate the
results will be shown on top of the results page. That way, you'll have a record
of what produced the results and you can verify your analysis later.</p>

</body>

</html>

