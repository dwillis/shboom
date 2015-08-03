<%@ LANGUAGE=VBScript%>
<!--#include virtual = "/thetimes/ADOVBS.inc" -->
<%
     
    
     
' ------------------------- Open a Database Connection -------------
     
     
dbName="master" 

dim lastpart(11)
LastPart(1)="_efilter.inc"
LastPart(2)="_expressionresults.asp"
LastPart(3)="_expressions.asp"
LastPart(4)="_filter.asp"
LastPart(5)="_filter.inc"
LastPart(6)="_filterresults.asp"
LastPart(7)="_groupby.asp"
LastPart(8)="_groupbyresults.asp"
LastPart(9)="_inputstuff.inc"
LastPart(10)="_mathoptions.inc"


TheFullPath = Request("TheFullPath")
ANewFolder = Request("ANewFolder")
ProjectName = Trim(Request("ProjectName"))


' flags for no spaces or periods
ProjPageFlag = True


' flags for  no entry
ProjPageOK = True


' flags for exists
ProjPageEFlag = True


Set fs = CreateObject("Scripting.FileSystemObject")


' test if seach page is OK

If ProjectName <> "" Then

	ProjectName = Replace(ProjectName,".asp","")

	If InStr(ProjectName,".") > 0 or InStr(ProjectName," ") > 0 Then
		ProjPageFlag = False
	Else
		

	For I = 1 to 10	

		TheFileName = ProjectName & LastPart(I)
		If fs.FileExists(TheFullPath & TheFileName) Then
			ProjPageEFlag = False
		End If

	Next

	End If

	


Else

	ProjPageOK = False

End If


%>

<html>
<head>

<title>Check Files</title>
</head>
<body bgColor="white">
<!--Written by Tom Torok of The New York Times
tomtorok@nytimes.com-->
<!--#include virtual = "/thetimes/rcon.inc" -->
 
<%

	SQL="SELECT name from sysdatabases order by name"
	Rs.Open sql, Conn



%>


<% If ProjPageFlag and ProjPageOK and ProjPageEFlag Then %>

<h1 align="center">Project Name Accepted</h1>

<hr width="55%">

<h1 align="center">Select a Database</h1>


<form method="post" name="dbform" action="thetable.asp">
	<input type="hidden" value="<% = TheFullPath %>" name="TheFullPath">
	<input type="hidden" value="<% = ANewFolder %>" name="ANewFolder">
	<input type="hidden" value="<% = ProjectName %>" name="ProjectName">

 <center><select id="select1" name="TheDB" size="1" ONChange='dbform.submit()'></p>
<option value=""></option>


<%
Do While NOT RS.EOF

%>
        <option value="<% = rs("name") %>"><% = rs("name") %></option>
<%

rs.MoveNext

Loop
rs.Close
%>
</select>
  </center>
</form>

<% Else %>

<h1 align="center"><font color="#FF0000">Project Name Rejected</font></h1>


<h3 align="center">Either the name is invalid or you already have one or more files with the names that would be created.</h3>


<p align="center"><b>Click on the back button and try again.</b></p>

<% End If %>

</body>
</html>
