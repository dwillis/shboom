<%@ LANGUAGE="VBSCRIPT" %>
<%

TheFullPath = Request("TheFullPath")
ANewFolder = Request("ANewFolder")
ProjectName = Request("ProjectName")
TheDB = Request("TheDB")
TheTable = Request("TheTable")
SFields = Request("SFields")
HowMany = Request("SFields").Count

DupeFlag = "N"
' Test that all selections from SFields were used
Dim Un
Redim un(HowMany+1)
For I = 1 to HowMany
	Un(I) = Request("SFields")(I)
		For J = 1 to HowMany
			If J <> I Then
				If Un(I) = Request("SFields")(J) Then
					DupeFlag = "Y"
				End If
			End If
		Next
Next






%>

<html>
<head>
<title>Name Columns</title>
</head>
<body bgColor="#ffcc99">


<!--Written by Tom Torok of The New York Times
tomtorok@nytimes.com-->



<h1 align="center">Select Column Headings</h1>

<p align="left">Enter the column headings of your choice for each field. To use the field
name as a column heading, leave the field alone. To replace the field name with a
different heading, delete the field name and enter a heading.</p>
<% If DupeFlag = "Y" Then ' print a caution for dupe field %>
<p align="left"><font color="#FF0000"><strong>Note: There are duplicate field names, meaning you failed to select at least 
one field on the last page. If that's what you wanted to do, fine. If not, 
hit the "back" key and try again.</strong></font></p>
<% End If ' end caution for dupe field %>
<form method="POST" action="thelink.asp">

<input type="hidden" value="<% = TheFullPath %>" name="TheFullPath">
<input type="hidden" value="<% = ANewFolder %>" name="ANewFolder">
<input type="hidden" value="<% = ProjectName %>" name="ProjectName">
<input type="hidden" value="<% = TheDB %>" name="TheDB">
<input type="hidden" value="<% = TheTable %>" name="TheTable">
<% For I = 1 to HowMany %>
<input type="hidden" value="<% = Request("SFields")(I) %>" name="SFields">
<% Next %>
<table border="1" align="center">
<tr>
<% For I = 1 to HowMany %>
	<td><input type="text" name="Dfields" size="18" value="<% = Request("SFields")(I) %>"></td>
<% Next %>

</tr>
</table>
<center><p><input type="submit" value="Submit"></p>
  </center>
</form>



</body>
</html>
