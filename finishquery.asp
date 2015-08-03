<%@ LANGUAGE="VBSCRIPT" %>
<%
TheFullPath = Request("TheFullPath")
ANewFolder = Request("ANewFolder")
ProjectName = Request("ProjectName")
TheDB = Request("TheDB")
TheTable = Request("TheTable")
SFields = Request("SFields")
HowMany = Request("SFields").Count
DFields = Request("DFields")


%>

<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual InterDev 1.0">

<title>Paging Information</title>
</head>
<body bgColor="#ffcc99">


<!--Written by Tom Torok of The New York Times
tomtorok@nytimes.com-->



<h1 align="center">Almost Done!</h1>

<p align="center"><b>Finally, enter the following information for paging:</b></p>

<form method="POST" action="aliasresults.asp">

<input type="hidden" value="<% = TheFullPath %>" name="TheFullPath">
<input type="hidden" value="<% = ANewFolder %>" name="ANewFolder">
<input type="hidden" value="<% = ProjectName %>" name="ProjectName">
<input type="hidden" value="<% = TheDB %>" name="TheDB">
<input type="hidden" value="<% = TheTable %>" name="TheTable">
<% For I = 1 to HowMany %>
<input type="hidden" value="<% = Request("SFields")(I) %>" name="SFields">
<input type="hidden" value="<% = Request("DFields")(I) %>" name="DFields">
<% Next %>
<input type="hidden" value="<% = Request("LField") %>" name="LField">
<center><p><strong>Enter the number of rows
  to be returned on a query page</strong></p>
  </center></div><div align="center"><center><p><input type="text" name="TheRowNum" size="5"></p>
  </center></div><div align="center"><center><p><strong>Enter the maximum number of records
  a query should return</strong></p>
  </center></div><div align="center"><center><p><input type="text" name="TheMax" size="5"></p>
  </center></div><div align="center"><center><p><input type="submit" value="Submit" name="B1"></p>
  </center></div>
</form>

</body>
</html>
