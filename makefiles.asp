<%@ LANGUAGE="VBSCRIPT" %>
<% 
TheFullPath = Request("TheFullPath")
ANewFolder = Request("ANewFolder")

%>


<html>
<head>

<title>Name Project</title>
</head>
<body bgColor="white">
<!--Written by Tom Torok of The New York Times
tomtorok@nytimes.com-->

<h1 align="center">Name Your Project</h1>

<p align="left">Name your project. Project names must be web compliant, i.e.,
they cannot contain spaces or start with numbers. <br>
The following pages will be created with your project name as the root:
<blockquote>
<b>projectname_drilldown.asp<br>
projectname_efilter.inc<br>
projectname_expressionresults.asp<br>
projectname_expressions.asp<br>
projectname_filter.asp<br>
projectname_filter.inc<br>
projectname_filterresults.asp<br>
projectname_groupby.asp<br>
projectname_groupbyresults.asp<br>
projectname_inputstuff.inc<br>
projectname_links.inc<br>
projectname_mathoptions.inc<br>
projectname_profile.asp</b><br>
</blockquote>

</p>

<form method="POST" action="checkfiles.asp">
	<input type="hidden" value="<% = TheFullPath %>" name="TheFullPath">
	<input type="hidden" value="<% = ANewFolder %>" name="ANewFolder">
  <div align="center"><p><strong>Enter Project Name<br>
  </strong><input type="text" name="ProjectName" size="20"></p>
  </div></p>
  <div align="center"><center><p><input type="submit" value="Submit" name="B1"><input type="reset" value="Reset" name="B2"></p>
  </center>
</form>

<% If ANewFolder = "N" Then %>
<h2 align="center">The following files exist in the directory.</h2>
<b>
<%  Set fso = CreateObject("Scripting.FileSystemObject")
  Set f = fso.GetFolder(TheFullPath)
  Set fc = f.Files
  For Each f1 in fc
	Response.Write(f1.name & "<br>")
  Next
End If
%>
</b>
</body>
</html>
