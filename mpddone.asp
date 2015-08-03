<%@ Language=VBScript %>
<%
strConn=Request("strConn")
TheFullPath = Request("TheFullPath")
ANewFolder = Request("ANewFolder")
SPageName = Request("SPageName")
QPageName = Request("QPageName")
TheTable = Request("TheTable")
BoxCaption = Request("BoxCaption")
FieldsAray = Request("FieldsAray")
BoxSize = Request("BoxSize")
SearchType = Request("SearchType")
HTMLTitle = Request("HTMLTitle")
TheHeading = Request("TheHeading")
Instruct = Request("Instruct")
SFields = Request("SFields")
DFields = Request("DFields")
TheMax = Request("TheMax")
TheRowNum = Request("TheRowNum")
NumFields = Request("FieldsAray").Count
NumSelect = Request("SFields").Count

'Response.Write(TheDSN & "<br>")
'Response.Write(TheLogin & "<br>")
'Response.Write(ThePword & "<br>")
'Response.Write(TheFullPath & "<br>")
'Response.Write(ANewFolder & "<br>")
'Response.Write(SPageName & "<br>")
'Response.Write(QPageName & "<br>")
'Response.Write(TheTable & "<br>")
'Response.Write(BoxCaption & "<br>")
'Response.Write(FieldsAray & "<br>")
'Response.Write(BoxSize & "<br>")
'Response.Write(SearchType & "<br>")
'Response.Write(HTMLTitle & "<br>")
'Response.Write(TheHeading & "<br>")
'Response.Write(Instruct & "<br>")
'Response.Write(SFields & "<br>")
'Response.Write(DFields & "<br>")
'Response.Write(TheMax & "<br>")
'Response.Write(TheRowNum & "<br>")
'Response.Write(NumFields)
' Start writing pages

Const ForReading = 1, ForWriting = 2, ForAppending = 3
Set fs = CreateObject("Scripting.FileSystemObject")

' write the search page 
Set sp = fs.CreateTextFile(TheFullPath & "\" & SPageName, True)

sp.WriteLine(Chr(60) + Chr(37) + "@ LANGUAGE=VBScript" + chr(37) + ">")
sp.WriteLine("<html>")
sp.WriteLine("<head>")
sp.WriteLine("<title>" + HTMLTitle + "</title>")
sp.WriteLine("</head>")
sp.WriteLine("<body bgcolor=" + chr(34) + "#FFFDEC" + chr(34) + ">")
sp.WriteLine("<h1 align=" + chr(34) + "center" + chr(34) + ">" + TheHeading + "</h1>")
sp.WriteLine("<h4 align=" + chr(34) + "center" + chr(34) + "><font color=" + chr(34) + "#CDA21F" + chr(34) + ">" + Instruct + "</font></h4>")
sp.WriteLine("<form method=" + chr(34) + "POST" + chr(34) + " action=" + chr(34) + QPageName + chr(34) + ">")
sp.WriteLine("<center>")

For I = 1 to NumFields
	sp.WriteLine("  <p><strong>" + Request("BoxCaption")(I) + "<br>")
	sp.WriteLine("  </strong><input type=" + chr(34) + "text" + chr(34) + " name=" + chr(34) + "frm" + Request("FieldsAray")(I) + chr(34) + " size=" + chr(34) + Request("BoxSize")(I) + chr(34) + "></p>")
Next
sp.WriteLine("  <div align=" + chr(34) + "center" + chr(34) + "><center><p><input type=" + chr(34) + "submit" + chr(34) + " value=" + chr(34) + "Submit" + chr(34) + "><input")
sp.WriteLine("  type=" + chr(34) + "reset" + chr(34) + " value=" + chr(34) + "Reset" + chr(34) + " name=" + chr(34) + "B2" + chr(34) + "></p>")
sp.WriteLine("  </center></div>")
sp.WriteLine("</form>")
sp.WriteLine("<p align=" + chr(34) + "center" + chr(34) + "><em>This page was generated from a a web application written by Tom Torok of The New York Times.<br>")
sp.WriteLine("<a href=" + chr(34) + "mailto:tomtorok@nytimes.com" + chr(34) + ">Comments</a></em></p>")
sp.WriteLine("</body>")
sp.WriteLine("</html>")
sp.Close


' write the querypage qp

Set qp = fs.CreateTextFile(TheFullPath & "\" & QPageName, True)

qp.WriteLine(Chr(60) + Chr(37) + "@ LANGUAGE=VBScript" + chr(37) + ">")
qp.WriteLine(Chr(60) + Chr(37))
qp.WriteLine("'---- CursorTypeEnum Values ----")
qp.WriteLine("Const adOpenForwardOnly = 0")
qp.WriteLine("Const adOpenKeyset = 1")
qp.WriteLine("Const adOpenDynamic = 2")
qp.WriteLine("Const adOpenStatic = 3")
qp.WriteLine("'---- LockTypeEnum Values ----")
qp.WriteLine("Const adLockReadOnly = 1")
qp.WriteLine("Const adLockPessimistic = 2")
qp.WriteLine("Const adLockOptimistic = 3")
qp.WriteLine("Const adLockBatchOptimistic = 4")
qp.WriteLine("'---- CommandTypeEnum Values ----")
qp.WriteLine("Const adCmdUnknown = &H0008")
qp.WriteLine("Const adCmdText = &H0001")
qp.WriteLine("Const adCmdTable = &H0002")
qp.WriteLine("Const adCmdStoredProc = &H0004")
qp.WriteLine("Const adCmdFile = &H0100")
qp.WriteLine("Const adCmdTableDirect = &H0200")

qp.WriteLine("'---Set Max Records and Number of Rows Displayed")
qp.WriteLine("maxRecs = " & TheMax)
qp.WriteLine("DisplayRows = " & TheRowNum)
qp.WriteLine("'Test to see if the " + chr(34) + "New Query" + chr(34) + " button was clicked")
qp.WriteLine("	If Request(" + chr(34) + "GoBack" + chr(34) + ") = " + chr(34) + "New Query" + chr(34) + " Then")
qp.WriteLine("			Response.Redirect(" + chr(34) + SPageName + chr(34) + ")")
qp.WriteLine("	End If")
qp.WriteLine("' If there is no sql statement from asp form")
qp.WriteLine("' test for input from " + SPageName)
qp.WriteLine("If Request(" + chr(34) + "sql" + chr(34) + ") = " + chr(34) + "" + chr(34) + " Then ' start building sql")
qp.WriteLine("' test for three characters")

' Build ValCheck to see if at least 3 characters were entered

ValCheck = "If Len(Request(" + chr(34) + "frm" + Request("FieldsAray")(1) + chr(34) + "))"
If NumFields > 1 Then
	For I = 2 to NumFields
		ValCheck = ValCheck + " + Len(Request(" + chr(34) + "frm" + Request("FieldsAray")(I) + chr(34) + "))"
	Next
End IF
ValCheck = ValCheck + " < 3 Then" 

qp.WriteLine(ValCheck)
qp.WriteLine("	Response.Redirect(" + chr(34) + SPageName + chr(34) + ")")
qp.WriteLine("Else")
qp.WriteLine("' build a Select statement")

' build the select

TheSelect = " SQLSelect = " + chr(34) + "Select " + Request("SFields")(1)
If NumSelect > 1 Then
	For I = 2 to NumSelect
		TheSelect = TheSelect + ", " + Request("SFields")(I)
	Next
End If

TheSelect = TheSelect + " from " + TheTable + chr(34)

qp.WriteLine(TheSelect)
qp.WriteLine("'build SQLWhere")
qp.WriteLine("'Test to see if anything was in the frm" + Request("FieldsAray")(1) + " field.") 
qp.WriteLine("	If Request(" + chr(34) + "frm" + Request("FieldsAray")(1) + chr(34) + ") <> " + chr(34) + "" + chr(34) + " Then")
qp.WriteLine("		str" + Request("FieldsAray")(1) + " = Trim(Request(" + chr(34) + "frm" + Request("FieldsAray")(1) + chr(34) + "))")
Select Case Request("SearchType")(1)
	Case "A"
		WherePart = Request("FieldsAray")(1) + " Like '%" + chr(34) + " & str" + Request("FieldsAray")(1) + " & " + chr(34) + "%'"
	Case "E"
		WherePart = Request("FieldsAray")(1) + " = '" + chr(34) + " & str" + Request("FieldsAray")(1) + " & " + chr(34) + "'"
	Case "S"
		WherePart = Request("FieldsAray")(1) + " Like '" + chr(34) + " & str" + Request("FieldsAray")(1) + " & " + chr(34) + "%'"
	Case "N"
		WherePart = Request("FieldsAray")(1) + " =" + chr(34) + " & str" + Request("FieldsAray")(1)
End Select
If Request("SearchType")(1) = "N" Then
qp.WriteLine("		SQLWhere = " + chr(34) + " Where " + WherePart)
qp.WriteLine("	Else")
qp.WriteLine("		SQLWhere = " + chr(34) + "" + chr(34))
qp.WriteLine("	End If")
Else
qp.WriteLine("		SQLWhere = " + chr(34) + " Where " + WherePart + chr(34))
qp.WriteLine("	Else")
qp.WriteLine("		SQLWhere = " + chr(34) + "" + chr(34))
qp.WriteLine("	End If")
End If



If NumFields > 1 Then
	For I = 2 to NumFields
		qp.WriteLine("'Test to see if anything was in the frm" + Request("FieldsAray")(I) + " field.") 
		qp.WriteLine("	If Request(" + chr(34) + "frm" + Request("FieldsAray")(I) + chr(34) + ") <> " + chr(34) + "" + chr(34) + " Then")
		qp.WriteLine("		str" + Request("FieldsAray")(I) + " = Trim(Request(" + chr(34) + "frm" + Request("FieldsAray")(I) + chr(34) + "))")
		Select Case Request("SearchType")(I)
			Case "A"
				WherePart = Request("FieldsAray")(I) + " Like '%" + chr(34) + " & str" + Request("FieldsAray")(I) + " & " + chr(34) + "%'"
			Case "E"
				WherePart = Request("FieldsAray")(I) + " = '" + chr(34) + " & str" + Request("FieldsAray")(I) + " & " + chr(34) + "'"
			Case "S"
				WherePart = Request("FieldsAray")(I) + " Like '" + chr(34) + " & str" + Request("FieldsAray")(I) + " & " + chr(34) + "%'"
			Case "N"
				WherePart = Request("FieldsAray")(I) + " = " + chr(34) + "& str" + Request("FieldsAray")(I)

		End Select
If Request("SearchType")(I)="N" Then
		qp.WriteLine("		If sqlWhere <> " + chr(34) + "" + chr(34) + " Then")
		qp.WriteLine("			SQLWhere = SQLWhere & " + chr(34) + " and " + WherePart)
		qp.WriteLine("		Else")
		qp.WriteLine("			SQLWhere = " + chr(34) + " Where " + WherePart)
		qp.WriteLine("		End If")
		qp.WriteLine("	End If")
Else
		qp.WriteLine("		If sqlWhere <> " + chr(34) + "" + chr(34) + " Then")
		qp.WriteLine("			SQLWhere = SQLWhere & " + chr(34) + " and " + WherePart + chr(34))
		qp.WriteLine("		Else")
		qp.WriteLine("			SQLWhere = " + chr(34) + " Where " + WherePart + chr(34))
		qp.WriteLine("		End If")
		qp.WriteLine("	End If")
End If

		
	Next
End If
qp.WriteLine("'build a sql statement")
qp.WriteLine("sql = SQLSelect & SQLWhere")
qp.WriteLine("End If")
qp.WriteLine("' if there is no sql passed from asp")
qp.WriteLine("Else")
qp.WriteLine("' create sql")
qp.WriteLine("	sql = Request(" + chr(34) + "sql" + chr(34) + ")")
qp.WriteLine("End If")
qp.WriteLine("' open a database connection")
qp.WriteLine("Set Conn = Server.CreateObject(" + chr(34) + "ADODB.Connection" + chr(34) + ")")
qp.WriteLine("Set RS = Server.CreateObject(" + chr(34) + "ADODB.RecordSet" + chr(34) + ")")
qp.WriteLine("StrConn=" + chr(34) + strConn + chr(34))
qp.WriteLine("Conn.Open strConn")
qp.WriteLine("'---Limit Records and Page Size")
qp.WriteLine("RS.MaxRecords = maxRecs")
qp.WriteLine("RS.PageSize = DisplayRows 'Number of rows per page")
qp.WriteLine("RS.Open sql, Conn, adOpenStatic, adLockReadOnly, adCmdText")
qp.WriteLine("' get the value of the page and record counts")
qp.WriteLine("' and load them into variables")
qp.WriteLine("TotalPages = RS.PageCount")
qp.WriteLine("TotalRecs = RS.RecordCount")
qp.WriteLine("' The test sees if there is a value to pick")
qp.WriteLine("' up from the buttons for prior and next page")
qp.WriteLine("' If there is no value, it sets the PageNumber to 1")
qp.WriteLine("ScrollAction = Request(" + chr(34) + "ScrollAction" + chr(34) + ")")
qp.WriteLine("if ScrollAction <> " + chr(34) + "" + chr(34) + " Then")
qp.WriteLine("	PageNo = ScrollAction")
qp.WriteLine("	if PageNo < 1 Then ")
qp.WriteLine("		PageNo = 1")
qp.WriteLine("	end if")
qp.WriteLine("else")
qp.WriteLine("	PageNo = 1")
qp.WriteLine("end if")
qp.WriteLine("' if any records returned tell the recordset")
qp.WriteLine("' which page you want displayed")
qp.WriteLine("If Not RS.EOF Then")
qp.WriteLine("	RS.AbsolutePage = PageNo")
qp.WriteLine("End If")
qp.WriteLine(Chr(37) + Chr(62))
qp.WriteLine("<HTML>")
qp.WriteLine("<HEAD>")
qp.WriteLine("<TITLE>Query Results Page</TITLE>")
qp.WriteLine("</HEAD>")
qp.WriteLine("<body bgcolor=" + chr(34) + "#FFFFFF" + chr(34) + ">")
qp.WriteLine(Chr(60) + Chr(37))
qp.WriteLine("' creates a form and ")
qp.WriteLine("' loads the value of sql into a hidden ")
qp.WriteLine("' input type called SQL")
qp.WriteLine(Chr(37) + Chr(62))
qp.WriteLine("	<FORM METHOD=POST ACTION=" + chr(34) + QPageName + chr(34) + ">")
qp.WriteLine("	<INPUT TYPE=" + chr(34) + "HIDDEN" + chr(34) + " NAME=" + chr(34) + "sql" + chr(34) + " VALUE=" + chr(34) + "<" + Chr(37) + " = sql " + Chr(37) + ">" + chr(34) + ">")
qp.WriteLine(Chr(60) + Chr(37))
qp.WriteLine("' Display the page you are on")
qp.WriteLine("' display the total # of records")
qp.WriteLine("' and if you are maxed out")
qp.WriteLine(" If Not RS.EOF Then")
qp.WriteLine(Chr(37) + Chr(62))
qp.WriteLine("<H3>Search Results</H3>")
qp.WriteLine("<B>Page <" + Chr(37) + "=PageNo" + Chr(37) + "> of <" + Chr(37) + "=TotalPages" + Chr(37) + "></B> [Total records: <" + Chr(37) + "=TotalRecs" + Chr(37) + ">")
qp.WriteLine("<" + Chr(37) + "	If TotalRecs = MaxRecs Then")
qp.WriteLine("		Response.Write(" + chr(34) + " (the max allowed.)]" + chr(34) + ")")
qp.WriteLine("	else")
qp.WriteLine("		Response.Write(" + chr(34) + "]" + chr(34) + ")")
qp.WriteLine("	End If " + Chr(37) + ">")
qp.WriteLine("<P>")
qp.WriteLine("<div align=" + chr(34) + "center" + chr(34) + "><center>")
qp.WriteLine("<table border=" + chr(34) + "1" + chr(34) + ">")
qp.WriteLine("  <tr>")
For I = 1 to NumSelect
	qp.WriteLine("    <td bgcolor=" + chr(34) + "#F7ECCA" + chr(34) + "><strong>" + Request("DFields")(I) + "</strong></td>")
Next
qp.WriteLine("  </tr>")
qp.WriteLine(Chr(60) + Chr(37))
qp.WriteLine("' initialize the row count and start the")
qp.WriteLine("' do while loop to print this page of records")
qp.WriteLine("	RowCount = rs.PageSize")
qp.WriteLine(" Do While Not RS.EOF and RowCount > 0" + Chr(37) + ">")
qp.WriteLine("  <tr>")
For I = 1 to NumSelect
	qp.WriteLine("    <td bgcolor=" + chr(34) + "#fffdec" + chr(34) + "><" + Chr(37) + " = rs(" + chr(34) + Request("SFields")(I) + chr(34) + ") " + Chr(37) + "></td>")
Next
qp.WriteLine("  </tr>")
qp.WriteLine(Chr(60) + Chr(37))
qp.WriteLine("' decrease the value of RowCount by 1")
qp.WriteLine("	RowCount = RowCount - 1")
qp.WriteLine("	RS.MoveNext")
qp.WriteLine("Loop")
qp.WriteLine(Chr(37) + Chr(62))
qp.WriteLine("</TABLE>")
qp.WriteLine("</div></center>")
qp.WriteLine("</p>")
qp.WriteLine(Chr(60) + Chr(37))
qp.WriteLine("' create prior next buttons")
qp.WriteLine("if TotalPages > 1 then " + Chr(37) + ">")
qp.WriteLine("<Strong>Pages: </Strong>")
qp.WriteLine("	<" + Chr(37) + " For I = 1 to TotalPages")
qp.WriteLine("	   if CInt(PageNo) <> I Then " + Chr(37) + ">")
qp.WriteLine("	<INPUT TYPE=SUBMIT NAME=ScrollAction VALUE=" + Chr(34) + "<" + Chr(37) + " = I " + Chr(37) + ">" + chr(34) + ">")
qp.WriteLine("	<" + Chr(37) + " end if ")
qp.WriteLine("  Next ")
qp.WriteLine(" end if " + Chr(37) + ">") 
qp.WriteLine("<" + Chr(37) + " ")
qp.WriteLine("'----------------------------------------------")
qp.WriteLine("'		If No results")
qp.WriteLine("'----------------------------------------------")
qp.WriteLine("Else " + Chr(37) + ">")
qp.WriteLine("	<h3>Your query returned no data.<br></h3>")
qp.WriteLine(" ")
qp.WriteLine("<" + Chr(37) + " End If " + Chr(37) + ">")
qp.WriteLine("<INPUT TYPE=" + chr(34) + "SUBMIT" + chr(34) + " NAME = " + chr(34) + "GoBack" + chr(34) + " VALUE = " + chr(34) + "New Query" + chr(34) + ">")
qp.WriteLine(Chr(60) + Chr(37))
qp.WriteLine("rs.Close")
qp.WriteLine("conn.Close")
qp.WriteLine("set rs = nothing")
qp.WriteLine("set Conn = nothing")
qp.WriteLine(Chr(37) + Chr(62))
qp.WriteLine("<p align=" + chr(34) + "center" + chr(34) + "><em>This page was generated by a web application developed by Tom Torok of The New York Times.<br>")
qp.WriteLine("<a href=" + chr(34) + "mailto:tomtorok@nytimes.com" + chr(34) + ">Comments</a></em></p>")
qp.WriteLine("</FORM>")
qp.WriteLine("</BODY>")
qp.WriteLine("</HTML>")

qp.Close


%>

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>Pages Completed</title>
</HEAD>
<body bgColor=#ccffff>

<!--Written by Tom Torok of The New York Times
tomtorok@nytimes.com-->


<h1 align="center">Pages Completed</h1>
<P align=center>The Search Page: <STRONG><% = SPageName %></STRONG></P>
<P align=center>and the Query Page: <STRONG><% = QPageName %></STRONG></P>
<P align=center>were written to the following directory:</P>
<P align=center><STRONG><% = TheFullPath %></STRONG></P>
<% If InStr(TheFullPath,"wwwroot")>0 Then
	WhereStart = InStr(TheFullPath,"wwwroot")
	TheLink = Right(TheFullPath,Len(TheFullPath)-(WhereStart+7))
	HowLong = Len(TheLink)
	If Right(TheLink,1) = "/" Then
		TheLink = Left(TheLink,HowLong-1)
	End If
	TheLink = TheLink & "/" & SPageName
	TheLink = "../" & TheLink
	TheLink=Replace(TheLink,"\","/")
	TheLink=Replace(TheLink,"//","/")
%>

<P align=center>To check them out <A href="<% = TheLink %>">click 
here.</A></P>
<% End If %>

</BODY>
</HTML>
