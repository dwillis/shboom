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

TheFullPath = Request("TheFullPath")
'ANewFolder = Request("ANewFolder")
ProjectName = Request("ProjectName")
'TheDB = Request("TheDB")
'TheTable = Request("TheTable")
'SFields = Request("SFields")
'HowMany = Request("SFields").Count
' the following line initializes numselect, which is used to comply with code stolen from mpd
'NumSelect = HowMany
'DFields = Request("DFields")
'TheMax = Request("TheMax")
'TheRowNum = Request("TheRowNum")
'LField = Request("LField")

'the linkfieldflag variable will be set to true if the link field is selected as part of the display
LinkFieldFlag = FALSE


'Response.Write("The Linking field is: " & LField & "<br>")

'ProjectName = UCase(Left(ProjectName, 1)) & Right(ProjectName, len(ProjectName) - 1)

Set RS = Server.CreateObject("ADODB.RecordSet")
Set objRS2 = Server.CreateObject("ADODB.RecordSet")

RS.Open TheFullPath & ProjectName & "_tablespecs.gaga",  "Provider=MSPersist;", adOpenDynamic, adLockBatchOptimistic, adCmdFile
objRS2.Open TheFullPath & ProjectName & "_progspecs.gaga",  "Provider=MSPersist;", adOpenDynamic, adLockBatchOptimistic, adCmdFile




Const ForReading = 1, ForWriting = 2, ForAppending = 3
Set fs = CreateObject("Scripting.FileSystemObject")
' create blank header and footer include files if there are none
If NOT fs.FileExists(TheFullPath & ProjectName & "_header.inc") Then

        Set sp = fs.CreateTextFile(TheFullPath & ProjectName & "_header.inc", True)
        sp.WriteLine(" ")
        sp.close
        set sp = nothing


        Set sp = fs.CreateTextFile(TheFullPath & ProjectName & "_footer.inc", True)
        sp.WriteLine("")
        sp.WriteLine("")
        sp.WriteLine("<p align=""center""><em>Developed by Tom Torok of The New York Times.<br>")
        sp.WriteLine("<a href=""mailto:tomtorok@nytimes.com"">Comments</a></em></p>")
        sp.WriteLine(" ")
        sp.close
        set sp = nothing

End If ' end creation of header and footer


dbName = objRS2("TheDB")

' Start writing pages
' start with include input files

rs.Filter =  ("ShowPulldown = TRUE")
rs.sort = ("TheOrder")    ' Sort the Recorset



'sql="select column_name as TheField, Data_Type as TheType from information_schema.columns where table_name = '" & TheTable & "' "
'sql = sql & " order by ordinal_position"


' write the _filter page
Set sp = fs.CreateTextFile(TheFullPath & ProjectName & "_filter.inc", True)

sp.WriteLine("<td align=""center""><select size=""1"" name=""frmSource"">")
sp.WriteLine("<option value="""""" selected></option>")



Do While Not rs.EOf

Select Case rs("TheType")


        Case "bigint", "bit", "decimal", "float", "int", "money", "numeric", "real", "smallint", "smallmoney", "tinyint"
                sp.WriteLine("<option value=""" & rs("TheField") & "*N" & """>" & rs("TheAlias") & "</option>")
        Case "datetime", "smalldatetime"
                sp.WriteLine("<option value=""" & rs("TheField") & "*D" & """>" & rs("TheAlias") & "</option>")
        Case "char", "nchar", "ntext", "nvarchar", "text", "varchar"
                sp.WriteLine("<option value=""" & rs("TheField") & "*T" & """>" & rs("TheAlias") & "</option>")
End Select

rs.MoveNext

Loop
sp.WriteLine("</select>")
sp.WriteLine(" ")
sp.WriteLine(" ")

sp.WriteLine("<td align=""center""><select size=""1"" name=""frmFilter1"">")
sp.WriteLine("<option value="""" selected></option>")
sp.WriteLine("<option value=""eq"">equals</option>")
sp.WriteLine("<option value=""noteq"">is not equal to</option>")
sp.WriteLine("<option value=""gt"">is greater than</option>")
sp.WriteLine("<option value=""gtet"">is greater than or equal to</option>")
sp.WriteLine("<option value=""lt"">is less than</option>")
sp.WriteLine("<option value=""ltet"">is less than or equal to</option>")
sp.WriteLine("<option value=""bw"">begins with</option>")
sp.WriteLine("<option value=""dnbw"">does not begin with</option>")
sp.WriteLine("<option value=""ew"">ends with</option>")
sp.WriteLine("<option value=""dnew"">does not end with</option>")
sp.WriteLine("<option value=""con"">contains</option>")
sp.WriteLine("<option value=""dncon"">does not contain</option>")
sp.WriteLine("<option value=""blank"">Is Blank</option>")
sp.WriteLine("<option value=""notblank"">Is Not Blank</option>")
sp.WriteLine("</select>")


sp.WriteLine("</td>")
sp.WriteLine("  <td align=""center""><input type=""text"" name=""frmValue1""  size=""10""></td>")
sp.WriteLine("<td align=""center""><select size=""1"" name=""frmAndOr"">")
sp.WriteLine("<option value="""" selected></option>")
sp.WriteLine("<option value=""and"">And</option>")
sp.WriteLine("<option value=""or"">Or</option>")
sp.WriteLine("</select>")
sp.WriteLine("</td>")
sp.WriteLine("<td align=""center""><select size=""1"" name=""frmFilter2"">")
sp.WriteLine("<option value="""" selected></option>")
sp.WriteLine("<option value=""eq"">equals</option>")
sp.WriteLine("<option value=""noteq"">is not equal to</option>")
sp.WriteLine("<option value=""gt"">is greater than</option>")
sp.WriteLine("<option value=""gtet"">is greater than or equal to</option>")
sp.WriteLine("<option value=""lt"">is less than</option>")
sp.WriteLine("<option value=""ltet"">is less than or equal to</option>")
sp.WriteLine("<option value=""bw"">begins with</option>")
sp.WriteLine("<option value=""dnbw"">does not begin with</option>")
sp.WriteLine("<option value=""ew"">ends with</option>")
sp.WriteLine("<option value=""dnew"">does not end with</option>")
sp.WriteLine("<option value=""con"">contains</option>")
sp.WriteLine("<option value=""dncon"">does not contain</option>")
sp.WriteLine("<option value=""blank"">Is Blank</option>")
sp.WriteLine("<option value=""notblank"">Is Not Blank</option>")
sp.WriteLine("</select>")
sp.WriteLine("</td>")
sp.WriteLine("  <td align=""center""><input type=""text"" name=""frmValue2""  size=""10""></td>")



sp.close
set sp = nothing

'--------------------------------------------------------------

' write the _efilter page
Set sp = fs.CreateTextFile(TheFullPath  & ProjectName & "_efilter.inc", True)

sp.WriteLine("<td align=""center""><select size=""1"" name=""frmESource"">")
sp.WriteLine("<option value="""""" selected></option>")

rs.MoveFirst

Do While Not rs.EOf

Select Case rs("TheType")


        Case "bigint", "bit", "decimal", "float", "int", "money", "numeric", "real", "smallint", "smallmoney", "tinyint"
                sp.WriteLine("<option value=""" & rs("TheField") & "*N" & """>" & rs("TheAlias") & "</option>")
        Case "datetime", "smalldatetime"
                sp.WriteLine("<option value=""" & rs("TheField") & "*D" & """>" & rs("TheAlias") & "</option>")
        Case "char", "nchar", "ntext", "nvarchar", "text", "varchar"
                sp.WriteLine("<option value=""" & rs("TheField") & "*T" & """>" & rs("TheAlias") & "</option>")
End Select

rs.MoveNext

Loop

sp.WriteLine("</select>")
sp.WriteLine(" ")
sp.WriteLine(" ")
sp.WriteLine(" ")
sp.WriteLine(" ")

sp.WriteLine("<td align=""center""><select size=""1"" name=""frmFilter1"">")
sp.WriteLine("<option value="""" selected></option>")
sp.WriteLine("<option value=""eq"">equals</option>")
sp.WriteLine("<option value=""noteq"">is not equal to</option>")
sp.WriteLine("<option value=""gt"">is greater than</option>")
sp.WriteLine("<option value=""gtet"">is greater than or equal to</option>")
sp.WriteLine("<option value=""lt"">is less than</option>")
sp.WriteLine("<option value=""ltet"">is less than or equal to</option>")
sp.WriteLine("<option value=""bw"">begins with</option>")
sp.WriteLine("<option value=""dnbw"">does not begin with</option>")
sp.WriteLine("<option value=""ew"">ends with</option>")
sp.WriteLine("<option value=""dnew"">does not end with</option>")
sp.WriteLine("<option value=""con"">contains</option>")
sp.WriteLine("<option value=""dncon"">does not contain</option>")
sp.WriteLine("<option value=""blank"">Is Blank</option>")
sp.WriteLine("<option value=""notblank"">Is Not Blank</option>")

sp.WriteLine("</select>")
sp.WriteLine("</td>")
sp.WriteLine("  <td align=""center""><input type=""text"" name=""frmValue1""  size=""10""></td>")
sp.WriteLine("<td align=""center""><select size=""1"" name=""frmAndOr"">")
sp.WriteLine("<option value="""" selected></option>")
sp.WriteLine("<option value=""and"">And</option>")
sp.WriteLine("<option value=""or"">Or</option>")
sp.WriteLine("</select>")
sp.WriteLine("</td>")
sp.WriteLine("<td align=""center""><select size=""1"" name=""frmFilter2"">")
sp.WriteLine("<option value="""" selected></option>")
sp.WriteLine("<option value=""eq"">equals</option>")
sp.WriteLine("<option value=""noteq"">is not equal to</option>")
sp.WriteLine("<option value=""gt"">is greater than</option>")
sp.WriteLine("<option value=""gtet"">is greater than or equal to</option>")
sp.WriteLine("<option value=""lt"">is less than</option>")
sp.WriteLine("<option value=""ltet"">is less than or equal to</option>")
sp.WriteLine("<option value=""bw"">begins with</option>")
sp.WriteLine("<option value=""dnbw"">does not begin with</option>")
sp.WriteLine("<option value=""ew"">ends with</option>")
sp.WriteLine("<option value=""dnew"">does not end with</option>")
sp.WriteLine("<option value=""con"">contains</option>")
sp.WriteLine("<option value=""dncon"">does not contain</option>")
sp.WriteLine("<option value=""blank"">Is Blank</option>")
sp.WriteLine("<option value=""notblank"">Is Not Blank</option>")
sp.WriteLine("</select>")
sp.WriteLine("</td>")
sp.WriteLine("  <td align=""center""><input type=""text"" name=""frmValue2""  size=""10""></td>")



sp.close
set sp = nothing

' ------------------------- ended with creation of _efilter.inc


'--------------------------------------------------------------

' write the _inputstuff page
Set sp = fs.CreateTextFile(TheFullPath  & ProjectName & "_inputstuff.inc", True)

sp.WriteLine("<td align=""center""><select size=""1"" name=""frmSource"">")
sp.WriteLine("<option value="""""" selected></option>")

rs.MoveFirst

Do While Not rs.EOf

        If InStr(rs("TheType"), "datetime") Then
                sp.WriteLine("<option value=""" & rs("TheField") & """>" & rs("TheAlias") & "</option>")
' --------------- added Dec 20
                sp.WriteLine("<option value=""" & rs("TheField") & "*M" & """>" & rs("TheAlias") & " (yr-month)</option>")
' --------------- added Dec 20
                sp.WriteLine("<option value=""" & rs("TheField") & "*Y" & """>" & rs("TheAlias") & " (year)</option>")
        Else
                sp.WriteLine("<option value=""" & rs("TheField") & """>" & rs("TheAlias") & "</option>")
        END IF

rs.MoveNext
Loop

        sp.WriteLine("</select>")

sp.close
set sp = nothing


' ------------------------- ended with creation of _inputstuff.inc







' write the _filter.asp page
Set sp = fs.CreateTextFile(TheFullPath  & ProjectName & "_filter.asp", True)



sp.WriteLine(Chr(60) + Chr(37) + "@ LANGUAGE=VBScript" + chr(37) + ">")
sp.WriteLine("")
sp.WriteLine("<head>")
sp.WriteLine("<title>" & objRS2("ProjectAlias") & "</title>")
sp.WriteLine("<LINK REL=""stylesheet"" TYPE=""text/css"" HREF=""/thetimes/shboom.css"">")
sp.WriteLine("<style>")
sp.WriteLine("<!--")
sp.WriteLine("select       { font-size: 7pt }")
sp.WriteLine("-->")
sp.WriteLine("</style>")
sp.WriteLine("</head>")
sp.WriteLine("<html>")
sp.WriteLine("")
sp.WriteLine("<body  bgcolor=""white"">")
sp.WriteLine("<!" & "--#include file = """ & ProjectName & "_links.inc"" -->")
sp.WriteLine("<!" & "--#include file = """ & ProjectName & "_header.inc"" -->")
sp.WriteLine("")
sp.WriteLine("<form method=""POST"" action=""" & ProjectName & "_filterresults.asp"">")

sp.WriteLine("<p><table>")
sp.WriteLine("<tr><td colspan=6 class=pageTitle>" & objRS2("FilterTitle") & "</td></tr>")
sp.WriteLine("<tr><td colspan=6 class=instructions>Select records that meet the following conditions:</td></tr>")

sp.WriteLine("<tr>")
sp.WriteLine("  <td class=""fieldHeading"">Field</td>")
sp.WriteLine("  <td class=""fieldHeading"">Show Rows Where</td>")
sp.WriteLine("  <td class=""fieldHeading"">Value 1</td>")
sp.WriteLine("  <td class=""fieldHeading"">And / or</td>")
sp.WriteLine("  <td class=""fieldHeading"">Show Rows Where</td>")
sp.WriteLine("  <td class=""fieldHeading"">Value 2</td>")
sp.WriteLine("")
sp.WriteLine("</tr>")
sp.WriteLine("")
sp.WriteLine("<% For I = 1 to " & objRS2("MaxFields") & " " &  chr(37) & ">")
sp.WriteLine("")
sp.WriteLine("<tr>")
sp.WriteLine("<!" & "--#include file = """ & ProjectName & "_filter.inc"" -->")
sp.WriteLine("</tr>")
sp.WriteLine("")
sp.WriteLine("<% Next " & chr(37) & ">")
sp.WriteLine("</table>")
sp.WriteLine("  <center><p><input type=""submit"" value=""Submit""><input")
sp.WriteLine("  type=""reset"" value=""Reset"" name=""B2""></p></center>")
sp.WriteLine("")
sp.WriteLine("</form>")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("<!" & "--#include file = """ & ProjectName & "_footer.inc"" -->")
sp.WriteLine("</body>")
sp.WriteLine("</html>")

sp.Close


' -----------------------------



' write the _filterresults.asp page
Set sp = fs.CreateTextFile(TheFullPath  & ProjectName & "_filterresults.asp", True)

sp.WriteLine(Chr(60) + Chr(37) + "@ LANGUAGE=VBScript" + chr(37) + ">")
sp.WriteLine("")
sp.WriteLine(Chr(60) + Chr(37))
sp.WriteLine("'---- CursorTypeEnum Values ----")
sp.WriteLine("Const adOpenForwardOnly = 0")
sp.WriteLine("Const adOpenKeyset = 1")
sp.WriteLine("Const adOpenDynamic = 2")
sp.WriteLine("Const adOpenStatic = 3")
sp.WriteLine("'---- LockTypeEnum Values ----")
sp.WriteLine("Const adLockReadOnly = 1")
sp.WriteLine("Const adLockPessimistic = 2")
sp.WriteLine("Const adLockOptimistic = 3")
sp.WriteLine("Const adLockBatchOptimistic = 4")
sp.WriteLine("'---- CommandTypeEnum Values ----")
sp.WriteLine("Const adCmdUnknown = &H0008")
sp.WriteLine("Const adCmdText = &H0001")
sp.WriteLine("Const adCmdTable = &H0002")
sp.WriteLine("Const adCmdStoredProc = &H0004")
sp.WriteLine("Const adCmdFile = &H0100")
sp.WriteLine("Const adCmdTableDirect = &H0200")
sp.WriteLine("Server.ScriptTimeout = 600")

sp.WriteLine("'---Set Max Records and Number of Rows Displayed")
sp.WriteLine("maxRecs = " & objRS2("MaxRecords"))
sp.WriteLine("DisplayRows = " & objRS2("MaxRows"))
sp.WriteLine("'Test to see if the " + chr(34) + "New Query" + chr(34) + " button was clicked")
sp.WriteLine("  If Request(" + chr(34) + "GoBack" + chr(34) + ") = " + chr(34) + "New Query" + chr(34) + " Then")
sp.WriteLine("                  Response.Redirect(" + chr(34) + ProjectName + "_filter.asp" + chr(34) + ")")
sp.WriteLine("  End If")
sp.WriteLine("' If there is no sql statement from asp form")
sp.WriteLine("' test for input from same page")
sp.WriteLine("If Request(" + chr(34) + "sql" + chr(34) + ") = " + chr(34) + "" + chr(34) + " Then ' start building sql")
sp.WriteLine("")

' NP is number of pulldowns
NP = objRS2("MaxFields") + 1

sp.WriteLine("")
sp.WriteLine("Dim TheSource(" & NP & "), TheType(" & NP & "), Filter1(" & NP & "), Filter2(" & NP & "), Value1(" & NP & "), Value2(" & NP & "), AndOr(" & NP & ")")
sp.WriteLine("")
sp.WriteLine("TheCount=0")
sp.WriteLine("For I = 1 to " & objRS2("MaxFields"))
sp.WriteLine("")
sp.WriteLine("  TestSource = Request(""frmSource"")(I)")
sp.WriteLine("  If TestSource = """" Then")
sp.WriteLine("          Exit For")
sp.WriteLine("  Else")
sp.WriteLine("          SplitSource=Split(TestSource, ""*"")")
sp.WriteLine("          TheSource(I) = SplitSource(0)")
sp.WriteLine("          TheType(I) = UCASE(SplitSource(1))")
sp.WriteLine("          Filter1(I) = Request(""frmFilter1"")(I)")
sp.WriteLine("          Filter2(I) = Request(""frmFilter2"")(I)")
sp.WriteLine("          Value1(I) = Request(""frmValue1"")(I)")
sp.WriteLine("          Value2(I) = Request(""frmValue2"")(I)")
sp.WriteLine("          AndOr(I) = Request(""frmAndOr"")(I)")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("          TheCount = I")
sp.WriteLine("  End If")
sp.WriteLine("'Response.Write(I)")
sp.WriteLine("Next")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("' make the sql where clause")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("sqlWhere = ""WHERE """)
sp.WriteLine("For I = 1 to TheCount")
sp.WriteLine("WhereHolder1 = """"")
sp.WriteLine("WhereHolder2 = """"")
sp.WriteLine("")
sp.WriteLine("'Response.Write(""The Source is: "" & TheSource(I) & ""<br>"")")
sp.WriteLine("")
sp.WriteLine("' see if there's an and/or clause")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("  If AndOr(I) = """" Then")
sp.WriteLine("          Select Case TheType(I)")
sp.WriteLine("                  Case ""N""")
sp.WriteLine("                          WhereHolder1 = MakeWhereNum(TheSource(I), Filter1(I), Value1(I))")
sp.WriteLine("                  Case ""T""")
sp.WriteLine("                          WhereHolder1 = MakeWhereText(TheSource(I), Filter1(I), Value1(I))")
sp.WriteLine("                  Case ""D""")
sp.WriteLine("                          WhereHolder1 = MakeWhereDate(TheSource(I), Filter1(I), Value1(I))")
sp.WriteLine("          End Select")
sp.WriteLine("")
sp.WriteLine("          If I <> TheCount Then")
sp.WriteLine("                  sqlWhere = sqlWhere & "" ("" & WhereHolder1 & "") AND """)
sp.WriteLine("          Else")
sp.WriteLine("                  sqlWhere = sqlWhere & "" ("" & WhereHolder1 & "") """)
sp.WriteLine("          End If")
sp.WriteLine("                  ")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("  Else ' there's an and or or     ")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("          Select Case TheType(I)")
sp.WriteLine("                  Case ""N""")
sp.WriteLine("                          WhereHolder1 = MakeWhereNum(TheSource(I), Filter1(I), Value1(I))")
sp.WriteLine("                          WhereHolder2 = MakeWhereNum(TheSource(I), Filter2(I), Value2(I))")
sp.WriteLine("                  Case ""T""")
sp.WriteLine("                          WhereHolder1 = MakeWhereText(TheSource(I), Filter1(I), Value1(I))")
sp.WriteLine("                          WhereHolder2 = MakeWhereText(TheSource(I), Filter2(I), Value2(I))")
sp.WriteLine("                  Case ""D""")
sp.WriteLine("                          WhereHolder1 = MakeWhereDate(TheSource(I), Filter1(I), Value1(I))")
sp.WriteLine("                          WhereHolder2 = MakeWhereDate(TheSource(I), Filter2(I), Value2(I))")
sp.WriteLine("          End Select")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("          If I <> TheCount Then")
sp.WriteLine("                  sqlWhere = sqlWhere & "" (("" & WhereHolder1 & "") "" & AndOr(I) & "" ("" & WhereHolder2 & "")) AND """)
sp.WriteLine("          Else")
sp.WriteLine("                  sqlWhere = sqlWhere & "" (("" & WhereHolder1 & "") "" & AndOr(I) & "" ("" & WhereHolder2& "")) """)
sp.WriteLine("          End If")
sp.WriteLine("                  ")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("  End If")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("Next")
sp.WriteLine("")

sp.WriteLine("' build a Fields List")

' build the select


rs.Filter = ""
rs.Filter = ("ShowDupe = TRUE")
rs.Sort  = ("ShowDupeOrder")

CountTheFields = rs.RecordCount
TheFields = ""

For I = 1 to CountTheFields

        If objRS2("LinkingField") = rs("TheField") Then
                LinkFieldFlag = TRUE
        End If

        TheFields = TheFields & rs("TheField")

        If I <> CountTheFields Then
                TheFields = TheFields & ", "
        ElseIf LinkFieldFlag = FALSE and objRS2("LinkingField") <> "" Then
                TheFields = TheFields & ", " & objRS2("LinkingField")
        End If

rs.MoveNext
Next

OrderBy1 = objRS2("OrderBy1")
SortOrder1 = objRS2("SortOrder1")
OrderBy2 = objRS2("OrderBy2")
SortOrder2 = objRS2("SortOrder2")
OrderBy3 = objRS2("OrderBy3")
SortOrder3 = objRS2("SortOrder3")






TheOrderBy = ""

If OrderBy1 <> "" Then
        If SortOrder1 = "DESC" Then
                TheOrderBy = " order by " & orderby1 & " DESC "
        Else
                TheOrderby =  " order by " & orderby1 & " "
        End If
End If

If OrderBy2 <> "" Then
        If SortOrder2 = "DESC" Then
                TheOrderBy = TheOrderBy & ", " & orderby2 & " DESC "
        Else
                TheOrderBy = TheOrderBy & ", " & orderby2 & " "
        End If
End If


If OrderBy3 <> "" Then
        If SortOrder3 = "DESC" Then
                TheOrderBy = TheOrderBy & ", " & orderby3 & " DESC "
        Else
                TheOrderBy = TheOrderBy & ", " & orderby3 & " "
        End If
End If




sp.WriteLine("TheOrderBy = " & chr(34) & TheOrderBy & chr(34))

sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
TheSQL = "Select " & TheFields & " from " & objRS2("TheTable") & " "
sp.WriteLine(" SQL =" & chr(34) & TheSQL & chr(34) & " & sqlwhere & TheOrderBy")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("Else")
sp.WriteLine("' create sql")
sp.WriteLine("  sql = Request(""sql"")")
sp.WriteLine("End If")
sp.WriteLine("")

sp.WriteLine("")
sp.WriteLine("dbName = " & chr(34)  & objRS2("TheDB") & chr(34) & " "  & chr(37) + ">")
sp.WriteLine("")
sp.WriteLine("")

sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("<!" & "--#include virtual = ""/thetimes/longrcon.inc"" -->")
sp.WriteLine("")
sp.WriteLine(Chr(60) + Chr(37))


sp.WriteLine("")
sp.WriteLine("'---Limit Records and Page Size")
sp.WriteLine("RS.MaxRecords = maxRecs")
sp.WriteLine("RS.PageSize = DisplayRows 'Number of rows per page")
sp.WriteLine("RS.Open sql, Conn, adOpenStatic, adLockReadOnly, adCmdText")
sp.WriteLine("' get the value of the page and record counts")
sp.WriteLine("' and load them into variables")
sp.WriteLine("TotalPages = RS.PageCount")
sp.WriteLine("TotalRecs = RS.RecordCount")
sp.WriteLine("' The test sees if there is a value to pick")
sp.WriteLine("' up from the buttons for prior and next page")
sp.WriteLine("' If there is no value, it sets the PageNumber to 1")
sp.WriteLine("ScrollAction = Request(""ScrollAction"")")
sp.WriteLine("if ScrollAction <> """" Then")
sp.WriteLine("  PageNo = ScrollAction")
sp.WriteLine("  if PageNo < 1 Then ")
sp.WriteLine("          PageNo = 1")
sp.WriteLine("  end if")
sp.WriteLine("else")
sp.WriteLine("  PageNo = 1")
sp.WriteLine("end if")
sp.WriteLine("' if any records returned tell the recordset")
sp.WriteLine("' which page you want displayed")
sp.WriteLine("If Not RS.EOF Then")
sp.WriteLine("  RS.AbsolutePage = PageNo")
sp.WriteLine("End If")
sp.WriteLine(chr(37) & ">")
sp.WriteLine("<HTML>")
sp.WriteLine("<HEAD>")
sp.WriteLine("<TITLE>Query Results Page</TITLE>")
sp.WriteLine("<LINK REL=""stylesheet"" TYPE=""text/css"" HREF=""/thetimes/shboom.css"">")
sp.WriteLine("</HEAD>")
sp.WriteLine("<body bgcolor=""white"">")
sp.WriteLine("<!" & "--#include file = """ & ProjectName & "_links.inc"" -->")
sp.WriteLine("<!" & "--#include file = """ & ProjectName & "_header.inc"" -->")
sp.WriteLine(Chr(60) + Chr(37))

sp.WriteLine("' creates a form and ")
sp.WriteLine("' loads the value of sql into a hidden ")
sp.WriteLine("' input type called SQL")

sp.WriteLine(chr(37) & ">")

sp.WriteLine("  <FORM METHOD=POST ACTION=""" & ProjectName & "_filterresults.asp"">")

' The following line is screwed up steal from mpd

sp.WriteLine("  <INPUT TYPE=" + chr(34) + "HIDDEN" + chr(34) + " NAME=" + chr(34) + "sql" + chr(34) + " VALUE=" + chr(34) + "<" + Chr(37) + " = sql " + Chr(37) + ">" + chr(34) + ">")

sp.WriteLine(Chr(60) + Chr(37))
sp.WriteLine("' Display the page you are on")
sp.WriteLine("' display the total # of records")
sp.WriteLine("' and if you are maxed out")
sp.WriteLine(" If Not RS.EOF Then")
sp.WriteLine(Chr(37) + Chr(62))
sp.WriteLine("<H1 class=pageTitle>" & objRS2("ProjectAlias") & " Filter Search Results</H1>")
sp.WriteLine("<B>Page <" + Chr(37) + "=PageNo" + Chr(37) + "> of <" + Chr(37) + "=TotalPages" + Chr(37) + "></B> [Total records: <" + Chr(37) + "=TotalRecs" + Chr(37) + ">")
sp.WriteLine("<" + Chr(37) + "  If TotalRecs = MaxRecs Then")
sp.WriteLine("          Response.Write(" + chr(34) + " (the max allowed.)]" + chr(34) + ")")
sp.WriteLine("  else")
sp.WriteLine("          Response.Write(" + chr(34) + "]" + chr(34) + ")")
sp.WriteLine("  End If " + Chr(37) + ">")
sp.WriteLine("<P>")
LField = objRS2("LinkingField") & ""
If Lfield <> "" Then
        sp.WriteLine("<center><small><b>Click on (p) to see a printable version of the complete record.</b></small></center>")
End If
sp.WriteLine("<table class=resultsTable cellpadding=0 cellspacing=0>")
sp.WriteLine("  <tr>")
rs.MoveFirst
For I = 1 to CountTheFields
        sp.WriteLine("    <td class=resultsHeading>" + rs("TheAlias") + "</td>")
        rs.MoveNext
Next
sp.WriteLine("  </tr>")
sp.WriteLine(Chr(60) + Chr(37))
sp.WriteLine("' initialize the row count and start the")
sp.WriteLine("' do while loop to print this page of records")
sp.WriteLine("  RowCount = rs.PageSize")
sp.WriteLine(" Do While Not RS.EOF and RowCount > 0" + Chr(37) + ">")
sp.WriteLine("  <tr>")
rs.MoveFirst
' subbed ------------------ Dec 16, 2005
For I = 1 to CountTheFields
        If objRS2("LinkingField") <> "" and I = 1 Then
                sp.WriteLine("    <td class=resultsData><" + Chr(37) + " = rs(" + chr(34) + rs("TheField") + chr(34) + ") " + Chr(37) + ">")
                sp.WriteLine("    <a href=""" & ProjectName & "_profile.asp?" & objRS2("LinkingField") & "=<" & chr(37) + " = rs(" + chr(34) + objRS2("LinkingField") + chr(34) + ") " + Chr(37) + ">" + chr(34) & " class=resultsDrillDown> (p)</a></td>")


        Else

	Select Case rs("TheType")

		Case "decimal", "float", "money", "numeric", "real", "smallmoney"

			sp.WriteLine(Chr(60) + Chr(37) + " If IsNull(rs(" + chr(34) + RS("TheField") + chr(34) + ")) Then " + Chr(37) + ">")
                	sp.WriteLine("    <td class=resultsDataRight><" + Chr(37) + " = rs(" + chr(34) + RS("TheField") + chr(34) + ") " + Chr(37) + ">&nbsp;</td>")
			sp.WriteLine(Chr(60) + Chr(37) + " Else " + Chr(37) + ">")
 	              	sp.WriteLine("    <td class=resultsDataRight><" + Chr(37) + " = FormatNumber(rs(" + chr(34) + RS("TheField") + chr(34) + "), 2) " + Chr(37) + "></td>")
			sp.WriteLine(Chr(60) + Chr(37) + " End If " + Chr(37) + ">")

		Case "bigint", "int"

			sp.WriteLine(Chr(60) + Chr(37) + " If IsNull(rs(" + chr(34) + RS("TheField") + chr(34) + ")) Then " + Chr(37) + ">")
                	sp.WriteLine("    <td class=resultsDataRight><" + Chr(37) + " = rs(" + chr(34) + RS("TheField") + chr(34) + ") " + Chr(37) + ">&nbsp;</td>")
			sp.WriteLine(Chr(60) + Chr(37) + " Else " + Chr(37) + ">")
 	              	sp.WriteLine("    <td class=resultsDataRight><" + Chr(37) + " = FormatNumber(rs(" + chr(34) + RS("TheField") + chr(34) + "), 0) " + Chr(37) + "></td>")
			sp.WriteLine(Chr(60) + Chr(37) + " End If " + Chr(37) + ">")



		Case else
                	sp.WriteLine("    <td class=resultsData><" + Chr(37) + " = rs(" + chr(34) + RS("TheField") + chr(34) + ") " + Chr(37) + ">&nbsp;</td>")

	End Select


        End If
rs.MoveNext
' subbed ---------- Dec. 16, 2005
Next
sp.WriteLine("  </tr>")
sp.WriteLine(Chr(60) + Chr(37))
sp.WriteLine("' decrease the value of RowCount by 1")
sp.WriteLine("  RowCount = RowCount - 1")
sp.WriteLine("  RS.MoveNext")
sp.WriteLine("Loop")
sp.WriteLine(Chr(37) + Chr(62))
sp.WriteLine("</TABLE>")
sp.WriteLine("</p>")
sp.WriteLine(Chr(60) + Chr(37))
sp.WriteLine("' create prior next buttons")
sp.WriteLine("if TotalPages > 1 then " + Chr(37) + ">")
sp.WriteLine("<Strong>Pages: </Strong>")
sp.WriteLine("  <" + Chr(37) + " For I = 1 to TotalPages")
sp.WriteLine("     if CInt(PageNo) <> I Then " + Chr(37) + ">")
sp.WriteLine("  <INPUT TYPE=SUBMIT NAME=ScrollAction VALUE=" + Chr(34) + "<" + Chr(37) + " = I " + Chr(37) + ">" + chr(34) + ">")
sp.WriteLine("  <" & chr(37) & " else " & chr(37) & ">")
sp.WriteLine("  <b><" & chr(37) & " = I " & chr(37) & "></b>")
sp.WriteLine("  <" + Chr(37) + " end if ")
sp.WriteLine("  Next ")
sp.WriteLine(" end if " + Chr(37) + ">") 
sp.WriteLine("<" + Chr(37) + " ")
sp.WriteLine("'----------------------------------------------")
sp.WriteLine("'         If No results")
sp.WriteLine("'----------------------------------------------")
sp.WriteLine("Else " + Chr(37) + ">")
sp.WriteLine("  <h3>Your query returned no data.<br></h3>")
sp.WriteLine(" ")
sp.WriteLine("<" + Chr(37) + " End If " + Chr(37) + ">")
sp.WriteLine("<INPUT TYPE=" + chr(34) + "SUBMIT" + chr(34) + " NAME = " + chr(34) + "GoBack" + chr(34) + " VALUE = " + chr(34) + "New Query" + chr(34) + ">")
sp.WriteLine(Chr(60) + Chr(37))
sp.WriteLine("rs.Close")
sp.WriteLine("conn.Close")
sp.WriteLine("set rs = nothing")
sp.WriteLine("set Conn = nothing")
sp.WriteLine(Chr(37) + Chr(62))
sp.WriteLine("</FORM>")
sp.WriteLine("<!" & "--#include file = """ & ProjectName & "_footer.inc"" -->")
sp.WriteLine("</BODY>")
sp.WriteLine("</HTML>")

sp.WriteLine("<" & chr(37))

sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("'Response.Write(""The where clause is: "" & sqlWhere & ""<br>"")")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("Function MakeWhereNum(TheFieldName, TheOperator, TheValue)")
sp.WriteLine("")
sp.WriteLine("Select Case TheOperator")
sp.WriteLine("  Case ""eq""")
sp.WriteLine("          MakeWhereNum = "" "" & TheFieldName & "" = "" & TheValue")
sp.WriteLine("  Case ""noteq""")
sp.WriteLine("          MakeWhereNum = "" "" & TheFieldName & "" <> "" & TheValue")
sp.WriteLine("  Case ""gt""")
sp.WriteLine("          MakeWhereNum = "" "" & TheFieldName & "" > "" & TheValue")
sp.WriteLine("  Case ""gtet""")
sp.WriteLine("          MakeWhereNum = "" "" & TheFieldName & "" >= "" & TheValue")
sp.WriteLine("  Case ""lt""")
sp.WriteLine("          MakeWhereNum = "" "" & TheFieldName & "" < "" & TheValue")
sp.WriteLine("  Case ""ltet""")
sp.WriteLine("          MakeWhereNum = "" "" & TheFieldName & "" <= "" & TheValue")
sp.WriteLine("  Case ""bw""")
sp.WriteLine("          MakeWhereNum = "" Cast("" & TheFieldName & "" as varchar(25)) Like '"" & TheValue & ""%'""")
sp.WriteLine("  Case ""dnbw""")
sp.WriteLine("          MakeWhereNum = "" Cast("" & TheFieldName & "" as varchar(25)) NOT Like '"" & TheValue & ""%'""")
sp.WriteLine("  Case ""ew""")
sp.WriteLine("          MakeWhereNum = "" Cast("" & TheFieldName & "" as varchar(25)) Like '%"" & TheValue & ""'""")
sp.WriteLine("  Case ""dnew""")
sp.WriteLine("          MakeWhereNum = "" Cast("" & TheFieldName & "" as varchar(25)) NOT Like '%"" & TheValue & ""'""")
sp.WriteLine("  Case ""con""")
sp.WriteLine("          MakeWhereNum = "" Cast("" & TheFieldName & "" as varchar(25)) Like '%"" & TheValue & ""%'""")
sp.WriteLine("  Case ""dncon""")
sp.WriteLine("          MakeWhereNum = "" Cast("" & TheFieldName & "" as varchar(25)) NOT Like '%"" & TheValue & ""%'""")
sp.WriteLine("  Case ""blank""")
sp.WriteLine("          MakeWhereNum = "" "" & TheFieldName & "" is null """)
sp.WriteLine("  Case ""notblank""")
sp.WriteLine("          MakeWhereNum = "" "" & TheFieldName & "" is NOT null  "" ")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("End Select")
sp.WriteLine("")
sp.WriteLine("End Function")
sp.WriteLine("")
sp.WriteLine("Function MakeWhereText(TheFieldName, TheOperator, TheValue)")
sp.WriteLine("If InStr(TheValue, ""'"") Then")
sp.WriteLine("  TheValue = Replace(TheValue, ""'"", ""''"")")
sp.WriteLine("End If")
sp.WriteLine("Select Case TheOperator")
sp.WriteLine("  Case ""eq""")
sp.WriteLine("          MakeWhereText = "" "" & TheFieldName & "" = '"" & TheValue & ""'""")
sp.WriteLine("          If InStr(MakeWhereText, ""]"") and InStr(MakeWhereText, ""["") Then")
sp.WriteLine("                  MakeWhereText = Replace(MakeWhereText, ""'"", """")")
sp.WriteLine("          End If")
sp.WriteLine("  Case ""noteq""")
sp.WriteLine("          MakeWhereText = "" "" & TheFieldName & "" <> '"" & TheValue & ""'""")
sp.WriteLine("          If InStr(MakeWhereText, ""]"") and InStr(MakeWhereText, ""["") Then")
sp.WriteLine("                  MakeWhereText = Replace(MakeWhereText, ""'"", """")")
sp.WriteLine("          End If")
sp.WriteLine("  Case ""gt""")
sp.WriteLine("          MakeWhereText = "" "" & TheFieldName & "" > '"" & TheValue & ""'""")
sp.WriteLine("          If InStr(MakeWhereText, ""]"") and InStr(MakeWhereText, ""["") Then")
sp.WriteLine("                  MakeWhereText = Replace(MakeWhereText, ""'"", """")")
sp.WriteLine("          End If")
sp.WriteLine("  Case ""gtet""")
sp.WriteLine("          MakeWhereText = "" "" & TheFieldName & "" >=  '"" & TheValue & ""'""")
sp.WriteLine("          If InStr(MakeWhereText, ""]"") and InStr(MakeWhereText, ""["") Then")
sp.WriteLine("                  MakeWhereText = Replace(MakeWhereText, ""'"", """")")
sp.WriteLine("          End If")
sp.WriteLine("  Case ""lt""")
sp.WriteLine("          MakeWhereText = "" "" & TheFieldName & "" <  '"" & TheValue & ""'""")
sp.WriteLine("          If InStr(MakeWhereText, ""]"") and InStr(MakeWhereText, ""["") Then")
sp.WriteLine("                  MakeWhereText = Replace(MakeWhereText, ""'"", """")")
sp.WriteLine("          End If")
sp.WriteLine("  Case ""ltet""")
sp.WriteLine("          MakeWhereText = "" "" & TheFieldName & "" <= '"" & TheValue & ""'""")
sp.WriteLine("          If InStr(MakeWhereText, ""]"") and InStr(MakeWhereText, ""["") Then")
sp.WriteLine("                  MakeWhereText = Replace(MakeWhereText, ""'"", """")")
sp.WriteLine("          End If")
sp.WriteLine("  Case ""bw""")
sp.WriteLine("          MakeWhereText = "" "" & TheFieldName & "" Like '"" & TheValue & ""%'""")
sp.WriteLine("  Case ""dnbw""")
sp.WriteLine("          MakeWhereText = "" "" & TheFieldName & "" NOT Like '"" & TheValue & ""%'""")
sp.WriteLine("  Case ""ew""")
sp.WriteLine("          MakeWhereText = "" "" & TheFieldName & "" Like '%"" & TheValue & ""'""")
sp.WriteLine("  Case ""dnew""")
sp.WriteLine("          MakeWhereText = "" "" & TheFieldName & "" NOT Like '%"" & TheValue & ""'""")
sp.WriteLine("  Case ""con""")
sp.WriteLine("          MakeWhereText = "" "" & TheFieldName & "" Like '%"" & TheValue & ""%'""")
sp.WriteLine("  Case ""dncon""")
sp.WriteLine("          MakeWhereText = "" "" & TheFieldName & "" NOT Like '%"" & TheValue & ""%'""")
sp.WriteLine("  Case ""blank""")
sp.WriteLine("          MakeWhereText = "" ("" & TheFieldName & "" is Null or "" & TheFieldName & "" = '') """)
sp.WriteLine("  Case ""notblank""")
sp.WriteLine("          MakeWhereText = "" ("" & TheFieldName & "" is NOT Null and "" & TheFieldName & "" <> '') """)
sp.WriteLine("End Select")
sp.WriteLine("")
sp.WriteLine("End Function")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("Function MakeWhereDate(TheFieldName, TheOperator, TheValue)")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("Select Case TheOperator")
sp.WriteLine("  Case ""eq""")
sp.WriteLine("          MakeWhereDate = "" "" & TheFieldName & "" = '"" & TheValue & ""'""")
sp.WriteLine("          If InStr(MakeWhereDate, ""]"") and InStr(MakeWhereDate, ""["") Then")
sp.WriteLine("                  MakeWhereDate = Replace(MakeWhereDate, ""'"", """")")
sp.WriteLine("          End If")
sp.WriteLine("  Case ""noteq""")
sp.WriteLine("          MakeWhereDate = "" "" & TheFieldName & "" <> '"" & TheValue & ""'""")
sp.WriteLine("          If InStr(MakeWhereDate, ""]"") and InStr(MakeWhereDate, ""["") Then")
sp.WriteLine("                  MakeWhereDate = Replace(MakeWhereDate, ""'"", """")")
sp.WriteLine("          End If")
sp.WriteLine("  Case ""gt""")
sp.WriteLine("          MakeWhereDate = "" "" & TheFieldName & "" > '"" & TheValue & ""'""")
sp.WriteLine("          If InStr(MakeWhereDate, ""]"") and InStr(MakeWhereDate, ""["") Then")
sp.WriteLine("                  MakeWhereDate = Replace(MakeWhereDate, ""'"", """")")
sp.WriteLine("          End If")
sp.WriteLine("  Case ""gtet""")
sp.WriteLine("          MakeWhereDate = "" "" & TheFieldName & "" >=  '"" & TheValue & ""'""")
sp.WriteLine("          If InStr(MakeWhereDate, ""]"") and InStr(MakeWhereDate, ""["") Then")
sp.WriteLine("                  MakeWhereDate = Replace(MakeWhereDate, ""'"", """")")
sp.WriteLine("          End If")
sp.WriteLine("  Case ""lt""")
sp.WriteLine("          MakeWhereDate = "" "" & TheFieldName & "" <  '"" & TheValue & ""'""")
sp.WriteLine("          If InStr(MakeWhereDate, ""]"") and InStr(MakeWhereDate, ""["") Then")
sp.WriteLine("                  MakeWhereDate = Replace(MakeWhereDate, ""'"", """")")
sp.WriteLine("          End If")
sp.WriteLine("  Case ""ltet""")
sp.WriteLine("          MakeWhereDate = "" "" & TheFieldName & "" <= '"" & TheValue & ""'""")
sp.WriteLine("          If InStr(MakeWhereDate, ""]"") and InStr(MakeWhereDate, ""["") Then")
sp.WriteLine("                  MakeWhereDate = Replace(MakeWhereDate, ""'"", """")")
sp.WriteLine("          End If")
sp.WriteLine("  Case ""bw""")
sp.WriteLine("          MakeWhereDate = "" Cast("" & TheFieldName & "" as varchar(25)) Like '"" & TheValue & ""%'""")
sp.WriteLine("  Case ""dnbw""")
sp.WriteLine("          MakeWhereDate = "" Cast("" & TheFieldName & "" as varchar(25)) NOT Like '"" & TheValue & ""%'""")
sp.WriteLine("  Case ""ew""")
sp.WriteLine("          MakeWhereDate = "" Cast("" & TheFieldName & "" as varchar(25)) Like '%"" & TheValue & ""'""")
sp.WriteLine("  Case ""dnew""")
sp.WriteLine("          MakeWhereDate = "" Cast("" & TheFieldName & "" as varchar(25)) NOT Like '%"" & TheValue & ""'""")
sp.WriteLine("  Case ""con""")
sp.WriteLine("          MakeWhereDate = "" Cast("" & TheFieldName & "" as varchar(25)) Like '%"" & TheValue & ""%'""")
sp.WriteLine("  Case ""dncon""")
sp.WriteLine("          MakeWhereDate = "" Cast("" & TheFieldName & "" as varchar(25)) NOT Like '%"" & TheValue & ""%'""")
sp.WriteLine("  Case ""blank""")
sp.WriteLine("          MakeWhereDate = "" "" & TheFieldName & "" is NULL "" ")
sp.WriteLine("  Case ""notblank""")
sp.WriteLine("          MakeWhereDate = "" "" & TheFieldName & "" is NOT NULL "" ")
sp.WriteLine("")
sp.WriteLine("End Select")
sp.WriteLine("")
sp.WriteLine("End Function")
sp.WriteLine("")

sp.WriteLine(chr(37) & ">")



sp.Close


' -----------------------------  drill down page



' write the _drilldown.asp page
Set sp = fs.CreateTextFile(TheFullPath  & ProjectName & "_drilldown.asp", True)

sp.WriteLine(Chr(60) + Chr(37) + "@ LANGUAGE=VBScript" + chr(37) + ">")
sp.WriteLine("")
sp.WriteLine(Chr(60) + Chr(37))
sp.WriteLine("'---- CursorTypeEnum Values ----")
sp.WriteLine("Const adOpenForwardOnly = 0")
sp.WriteLine("Const adOpenKeyset = 1")
sp.WriteLine("Const adOpenDynamic = 2")
sp.WriteLine("Const adOpenStatic = 3")
sp.WriteLine("'---- LockTypeEnum Values ----")
sp.WriteLine("Const adLockReadOnly = 1")
sp.WriteLine("Const adLockPessimistic = 2")
sp.WriteLine("Const adLockOptimistic = 3")
sp.WriteLine("Const adLockBatchOptimistic = 4")
sp.WriteLine("'---- CommandTypeEnum Values ----")
sp.WriteLine("Const adCmdUnknown = &H0008")
sp.WriteLine("Const adCmdText = &H0001")
sp.WriteLine("Const adCmdTable = &H0002")
sp.WriteLine("Const adCmdStoredProc = &H0004")
sp.WriteLine("Const adCmdFile = &H0100")
sp.WriteLine("Const adCmdTableDirect = &H0200")
sp.WriteLine("Server.ScriptTimeout = 600")

sp.WriteLine("'---Set Max Records and Number of Rows Displayed")
sp.WriteLine("maxRecs = " & objRS2("MaxRecords"))
sp.WriteLine("DisplayRows = " & objRS2("MaxRows"))
sp.WriteLine("'Test to see if the " + chr(34) + "New Query" + chr(34) + " button was clicked")
sp.WriteLine("  If Request(" + chr(34) + "GoBack" + chr(34) + ") = " + chr(34) + "New Query" + chr(34) + " Then")
sp.WriteLine("                  Response.Redirect(" + chr(34) + ProjectName + "_filter.asp" + chr(34) + ")")
sp.WriteLine("  End If")
sp.WriteLine("' If there is no sql statement from asp form")
sp.WriteLine("' test for input from same page")
sp.WriteLine("If Request(" + chr(34) + "sql" + chr(34) + ") = " + chr(34) + "" + chr(34) + " Then ' start building sql")
sp.WriteLine("")
sp.WriteLine("")

TheSQL = "Select " & TheFields & " from " & objRS2("TheTable") & " "

rs.Filter=""
rs.Filter=("ShowPulldown = TRUE")
rs.MoveFirst

Do While Not RS.EOF

Select Case rs("TheType")




             Case "char", "nchar", "varchar", "nvarchar", "text", "ntext"

                sp.WriteLine("'Test to see if anything was in the frm" + rs("TheField") + " field.") 
                sp.WriteLine("  If Request(" + chr(34) + "frm" + rs("TheField") + chr(34) + ") <> " + chr(34) + "" + chr(34) + " Then")
                sp.WriteLine("          str" + rs("TheField") + " = Request(" + chr(34) + "frm" + rs("TheField") + chr(34) + ")")
                sp.WriteLine("          q" + rs("TheField") + " = Replace(str" + rs("TheField") + ",""'"",""''"")")
                WherePart = rs("TheField") + " = '" + chr(34) + " & q" + rs("TheField") + " & " + chr(34) + "'" + chr(34)
                NullWherePart = "(" & rs("TheField") & " IS NULL or " & rs("TheField") & " = '')" & chr(34)

                sp.WriteLine("     If q" + rs("TheField") + " = ""NULL"" Then ")
                sp.WriteLine("          If sqlWhere <> " + chr(34) + "" + chr(34) + " Then")
                sp.WriteLine("                  SQLWhere = SQLWhere & " + chr(34) + " and " + NullWherePart)
                sp.WriteLine("          Else")
                sp.WriteLine("                  SQLWhere = " + chr(34) + " Where " + NullWherePart)
                sp.WriteLine("          End If")
                sp.WriteLine("     Else")
                sp.WriteLine("          If sqlWhere <> " + chr(34) + "" + chr(34) + " Then")
                sp.WriteLine("                  SQLWhere = SQLWhere & " + chr(34) + " and " + WherePart)
                sp.WriteLine("          Else")
                sp.WriteLine("                  SQLWhere = " + chr(34) + " Where " + WherePart)
                sp.WriteLine("          End If")
                sp.WriteLine("     End If")
                sp.WriteLine("  End If")

             Case "datetime", "smalldatetime"

                sp.WriteLine("'Test to see if anything was in the frm" + rs("TheField") + " field.") 
                sp.WriteLine("  If Request(" + chr(34) + "frm" + rs("TheField") + chr(34) + ") <> " + chr(34) + "" + chr(34) + " Then")
                sp.WriteLine("          str" + rs("TheField") + " = Request(" + chr(34) + "frm" + rs("TheField") + chr(34) + ")")
                sp.WriteLine("          q" + rs("TheField") + " = Replace(str" + rs("TheField") + ",""'"",""''"")")
                WherePart = rs("TheField") + " = '" + chr(34) + " & q" + rs("TheField") + " & " + chr(34) + "'" + chr(34)
                NullWherePart = "(" & rs("TheField") & " IS NULL or " & rs("TheField") & " = '')" & chr(34)
                YearWherePart = "DatePart(yyyy, " & rs("TheField") & ")= " & chr(34) & " & q" & rs("TheField") 

                sp.WriteLine("     If q" + rs("TheField") + " = ""NULL"" Then ")
                sp.WriteLine("          If sqlWhere <> " + chr(34) + "" + chr(34) + " Then")
                sp.WriteLine("                  SQLWhere = SQLWhere & " + chr(34) + " and " + NullWherePart)
                sp.WriteLine("          Else")
                sp.WriteLine("                  SQLWhere = " + chr(34) + " Where " + NullWherePart)
                sp.WriteLine("          End If")


                sp.WriteLine("     ElseIf Len(q" & rs("TheField") & ")=4 Then")
                sp.WriteLine("          If sqlWhere <> " + chr(34) + "" + chr(34) + " Then")
                sp.WriteLine("                  SQLWhere = SQLWhere & " + chr(34) + " and " + YearWherePart)
                sp.WriteLine("          Else")
                sp.WriteLine("                  SQLWhere = " + chr(34) + " Where " + YearWherePart)
                sp.WriteLine("          End If")


' -------------------- Changed Dec. 21


                sp.WriteLine("     ElseIf Len(q" & rs("TheField") & ")=7 Then")
		sp.WriteLine("		SplitIt = Split(q" & rs("TheField") & ", " & chr(34) & "-" & chr(34) & ")")
                sp.WriteLine("          If sqlWhere <> " + chr(34) + "" + chr(34) + " Then")
'                  SQLWhere = SQLWhere & " and DatePart(yyyy, REVISEDENDDATE)= " & SplitIt(0) & " and DatePart(mm, RevisedendDate) = " & SplitIt(1)

                sp.WriteLine("                  SQLWhere = SQLWhere & "" and DatePart(yyyy, " & rs("TheField") & ")= "" & SplitIt(0) & "" and DatePart(mm, " & RS("TheField") & ") = "" & SplitIt(1)")
                sp.WriteLine("          Else")
                sp.WriteLine("                  SQLWhere = " + chr(34) + " Where DatePart(yyyy, " & rs("TheField") & ")= "" & SplitIt(0) & "" and DatePart(mm, " & RS("TheField") & ") = "" & SplitIt(1)")
                sp.WriteLine("          End If")

'---------------------------- Changed Dec. 21




                sp.WriteLine("     Else")
                sp.WriteLine("          If sqlWhere <> " + chr(34) + "" + chr(34) + " Then")
                sp.WriteLine("                  SQLWhere = SQLWhere & " + chr(34) + " and " + WherePart)
                sp.WriteLine("          Else")
                sp.WriteLine("                  SQLWhere = " + chr(34) + " Where " + WherePart)
                sp.WriteLine("          End If")
                sp.WriteLine("     End If")
                sp.WriteLine("  End If")

             Case else ' these would be numbers

                sp.WriteLine("'Test to see if anything was in the frm" + rs("TheField") + " field.") 
                sp.WriteLine("  If Request(" + chr(34) + "frm" + rs("TheField") + chr(34) + ") <> " + chr(34) + "" + chr(34) + " Then")
                sp.WriteLine("          str" + rs("TheField") + " = Request(" + chr(34) + "frm" + rs("TheField") + chr(34) + ")")
                WherePart = rs("TheField") + " = "" & str" + rs("TheField")
                NullWherePart =  rs("TheField") + " IS NULL " & chr(34)
                sp.WriteLine("     If q" + rs("TheField") + " = ""NULL"" Then ")
                sp.WriteLine("          If sqlWhere <> " + chr(34) + "" + chr(34) + " Then")
                sp.WriteLine("                  SQLWhere = SQLWhere & " + chr(34) + " and " + NullWherePart)
                sp.WriteLine("          Else")
                sp.WriteLine("                  SQLWhere = " + chr(34) + " Where " + NullWherePart)
                sp.WriteLine("          End If")
                sp.WriteLine("     Else")
                sp.WriteLine("          If sqlWhere <> " + chr(34) + "" + chr(34) + " Then")
                sp.WriteLine("                  SQLWhere = SQLWhere & " + chr(34) + " and " + WherePart)
                sp.WriteLine("          Else")
                sp.WriteLine("                  SQLWhere = " + chr(34) + " Where " + WherePart)
                sp.WriteLine("          End If")
                sp.WriteLine("  End If")
                sp.WriteLine("     End If")




         End Select


rs.MoveNext

Loop

'sp.WriteLine("")
'sp.WriteLine("")
'sp.WriteLine("PassedSQL = Trim(Request(""frmsqlWhere""))")
'sp.WriteLine("If PassedSQL <> """" Then")
'sp.WriteLine("  WhereMarker = InStr(PassedSQL, ""WHERE"")")
'sp.WriteLine("  PassedSQL = "" AND "" & Right(PassedSQL, len(passedSQL) -(WhereMarker + 5))")
'sp.WriteLine("  sqlWhere = sqlWhere & PassedSQL")
'sp.WriteLine("End If")
'sp.WriteLine("")


sp.WriteLine("")
sp.WriteLine("' Test to see if there's only a where clause passed from the expressions page")
sp.WriteLine("")
sp.WriteLine("If Not(IsEmpty(Request(""&frmsqlWhere""))) Then")
sp.WriteLine("")
sp.WriteLine("	SQLWhere = "" "" & Request(""&frmsqlWhere"")")
sp.WriteLine("Else")
sp.WriteLine("")
sp.WriteLine("	PassedSQL = Trim(Request(""frmsqlWhere""))")
sp.WriteLine("	'Response.Write(PassedSQL & ""<br>"")")
sp.WriteLine("")
sp.WriteLine("	If PassedSQL <> """" Then")
sp.WriteLine("		WhereMarker = InStr(PassedSQL, ""WHERE"")")
sp.WriteLine("		PassedSQL = "" AND "" & Right(PassedSQL, len(passedSQL) -(WhereMarker + 5))")
sp.WriteLine("		sqlWhere = sqlWhere & PassedSQL")
sp.WriteLine("	End If")
sp.WriteLine("")
sp.WriteLine("End If")




sp.WriteLine("TheOrderBy = " & chr(34) & TheOrderBy & chr(34))


sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("'Combine the sql statement")
sp.WriteLine(" SQL =" & chr(34) & TheSQL & chr(34) & " & sqlwhere & TheOrderBy")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("Else")
sp.WriteLine("' create sql")
sp.WriteLine("  sql = Request(""sql"")")
sp.WriteLine("End If")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("dbName = " & chr(34)  & objRS2("TheDB") & chr(34) & " "  & chr(37) + ">")
sp.WriteLine("")
sp.WriteLine("")

sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("<!" & "--#include virtual = ""/thetimes/longrcon.inc"" -->")
sp.WriteLine("")
sp.WriteLine(Chr(60) + Chr(37))


sp.WriteLine("")
sp.WriteLine("'---Limit Records and Page Size")
sp.WriteLine("RS.MaxRecords = maxRecs")
sp.WriteLine("RS.PageSize = DisplayRows 'Number of rows per page")
sp.WriteLine("RS.Open sql, Conn, adOpenStatic, adLockReadOnly, adCmdText")
sp.WriteLine("' get the value of the page and record counts")
sp.WriteLine("' and load them into variables")
sp.WriteLine("TotalPages = RS.PageCount")
sp.WriteLine("TotalRecs = RS.RecordCount")
sp.WriteLine("' The test sees if there is a value to pick")
sp.WriteLine("' up from the buttons for prior and next page")
sp.WriteLine("' If there is no value, it sets the PageNumber to 1")
sp.WriteLine("ScrollAction = Request(""ScrollAction"")")
sp.WriteLine("if ScrollAction <> """" Then")
sp.WriteLine("  PageNo = ScrollAction")
sp.WriteLine("  if PageNo < 1 Then ")
sp.WriteLine("          PageNo = 1")
sp.WriteLine("  end if")
sp.WriteLine("else")
sp.WriteLine("  PageNo = 1")
sp.WriteLine("end if")
sp.WriteLine("' if any records returned tell the recordset")
sp.WriteLine("' which page you want displayed")
sp.WriteLine("If Not RS.EOF Then")
sp.WriteLine("  RS.AbsolutePage = PageNo")
sp.WriteLine("End If")
sp.WriteLine(chr(37) & ">")
sp.WriteLine("<HTML>")
sp.WriteLine("<HEAD>")
sp.WriteLine("<TITLE>Drill-down Results Page</TITLE>")
sp.WriteLine("<LINK REL=""stylesheet"" TYPE=""text/css"" HREF=""/thetimes/shboom.css"">")
sp.WriteLine("<style>")
sp.WriteLine("<!--")
sp.WriteLine("th           { background-color: #0000FF; color: #D9FFFF; font-size: 8pt }")
sp.WriteLine("td           { background-color: #FFFFFF; font-size: 8pt }")
sp.WriteLine("-->")
sp.WriteLine("</style>")
sp.WriteLine("</HEAD>")
sp.WriteLine("<body bgcolor=""white"">")
sp.WriteLine("<!" & "--#include file = """ & ProjectName & "_links.inc"" -->")
sp.WriteLine("<!" & "--#include file = """ & ProjectName & "_header.inc"" -->")
sp.WriteLine(Chr(60) + Chr(37))

sp.WriteLine("' creates a form and ")
sp.WriteLine("' loads the value of sql into a hidden ")
sp.WriteLine("' input type called SQL")

sp.WriteLine(chr(37) & ">")

sp.WriteLine("  <FORM METHOD=POST ACTION=""" & ProjectName & "_drilldown.asp"">")

' The following line is screwed up steal from mpd

sp.WriteLine("  <INPUT TYPE=" + chr(34) + "HIDDEN" + chr(34) + " NAME=" + chr(34) + "sql" + chr(34) + " VALUE=" + chr(34) + "<" + Chr(37) + " = sql " + Chr(37) + ">" + chr(34) + ">")

sp.WriteLine(Chr(60) + Chr(37))
sp.WriteLine("' Display the page you are on")
sp.WriteLine("' display the total # of records")
sp.WriteLine("' and if you are maxed out")
sp.WriteLine(" If Not RS.EOF Then")
sp.WriteLine(Chr(37) + Chr(62))
sp.WriteLine("<H1 class=pageTitle>" & objRS2("ProjectAlias") & " Details</H1>")
sp.WriteLine("<B>Page <" + Chr(37) + "=PageNo" + Chr(37) + "> of <" + Chr(37) + "=TotalPages" + Chr(37) + "></B> [Total records: <" + Chr(37) + "=TotalRecs" + Chr(37) + ">")
sp.WriteLine("<" + Chr(37) + "  If TotalRecs = " & objRS2("MaxRecords") & " Then")
sp.WriteLine("          Response.Write(" + chr(34) + " (the max allowed.)]" + chr(34) + ")")
sp.WriteLine("  else")
sp.WriteLine("          Response.Write(" + chr(34) + "]" + chr(34) + ")")
sp.WriteLine("  End If " + Chr(37) + ">")
sp.WriteLine("<P>")
If Lfield <> "" Then
        sp.WriteLine("<center><small><b>Click on (p) to see a printable version of the complete record.</b></small></center>")
End If
sp.WriteLine("<table class=resultsTable")
sp.WriteLine("  <tr>")
rs.Filter = ""
rs.Filter = ("ShowDupe = TRUE")
rs.Sort  = ("ShowDupeOrder")
rs.MoveFirst
For I = 1 to rs.RecordCount
        sp.WriteLine("    <td class=resultsHeading>" + rs("TheAlias") + "</td>")
        rs.MoveNext
Next
sp.WriteLine("  </tr>")
sp.WriteLine(Chr(60) + Chr(37))
sp.WriteLine("' initialize the row count and start the")
sp.WriteLine("' do while loop to print this page of records")
sp.WriteLine("  RowCount = rs.PageSize")
sp.WriteLine(" Do While Not RS.EOF and RowCount > 0" + Chr(37) + ">")
sp.WriteLine("  <tr>")
rs.MoveFirst
For I = 1 to rs.RecordCount
' subbed ------------------ Dec 16, 2005
        If objRS2("LinkingField") <> "" and I = 1 Then
                sp.WriteLine("    <td class=resultsData><" + Chr(37) + " = rs(" + chr(34) + rs("TheField") + chr(34) + ") " + Chr(37) + ">")
                sp.WriteLine("    <a href=""" & ProjectName & "_profile.asp?" & objRS2("LinkingField") & "=<" & chr(37) + " = rs(" + chr(34) + objRS2("LinkingField") + chr(34) + ") " + Chr(37) + ">" + chr(34) & " class=resultsDrillDown> (p)</a></td>")


        Else

	Select Case rs("TheType")

		Case "decimal", "float", "money", "numeric", "real", "smallmoney"

			sp.WriteLine(Chr(60) + Chr(37) + " If IsNull(rs(" + chr(34) + RS("TheField") + chr(34) + ")) Then " + Chr(37) + ">")
                	sp.WriteLine("    <td class=resultsDataRight><" + Chr(37) + " = rs(" + chr(34) + RS("TheField") + chr(34) + ") " + Chr(37) + ">&nbsp;</td>")
			sp.WriteLine(Chr(60) + Chr(37) + " Else " + Chr(37) + ">")
 	              	sp.WriteLine("    <td class=resultsDataRight><" + Chr(37) + " = FormatNumber(rs(" + chr(34) + RS("TheField") + chr(34) + "), 2) " + Chr(37) + "></td>")
			sp.WriteLine(Chr(60) + Chr(37) + " End If " + Chr(37) + ">")

		Case "bigint", "int"

			sp.WriteLine(Chr(60) + Chr(37) + " If IsNull(rs(" + chr(34) + RS("TheField") + chr(34) + ")) Then " + Chr(37) + ">")
                	sp.WriteLine("    <td class=resultsDataRight><" + Chr(37) + " = rs(" + chr(34) + RS("TheField") + chr(34) + ") " + Chr(37) + ">&nbsp;</td>")
			sp.WriteLine(Chr(60) + Chr(37) + " Else " + Chr(37) + ">")
 	              	sp.WriteLine("    <td class=resultsDataRight><" + Chr(37) + " = FormatNumber(rs(" + chr(34) + RS("TheField") + chr(34) + "), 0) " + Chr(37) + "></td>")
			sp.WriteLine(Chr(60) + Chr(37) + " End If " + Chr(37) + ">")



		Case else
                	sp.WriteLine("    <td class=resultsData><" + Chr(37) + " = rs(" + chr(34) + RS("TheField") + chr(34) + ") " + Chr(37) + ">&nbsp;</td>")

	End Select


        End If
rs.MoveNext
' subbed ---------- Dec. 16, 2005
Next
sp.WriteLine("  </tr>")
sp.WriteLine(Chr(60) + Chr(37))
sp.WriteLine("' decrease the value of RowCount by 1")
sp.WriteLine("  RowCount = RowCount - 1")
sp.WriteLine("  RS.MoveNext")
sp.WriteLine("Loop")
sp.WriteLine(Chr(37) + Chr(62))
sp.WriteLine("</TABLE>")
sp.WriteLine("</div></center>")
sp.WriteLine("</p>")
sp.WriteLine(Chr(60) + Chr(37))
sp.WriteLine("' create prior next buttons")
sp.WriteLine("if TotalPages > 1 then " + Chr(37) + ">")
sp.WriteLine("<Strong>Pages: </Strong>")
sp.WriteLine("  <" + Chr(37) + " For I = 1 to TotalPages")
sp.WriteLine("     if CInt(PageNo) <> I Then " + Chr(37) + ">")
sp.WriteLine("  <INPUT TYPE=SUBMIT NAME=ScrollAction VALUE=" + Chr(34) + "<" + Chr(37) + " = I " + Chr(37) + ">" + chr(34) + ">")
sp.WriteLine("  <" & chr(37) & " else " & chr(37) & ">")
sp.WriteLine("  <b><" & chr(37) & " = I " & chr(37) & "></b>")
sp.WriteLine("  <" + Chr(37) + " end if ")
sp.WriteLine("  Next ")
sp.WriteLine(" end if " + Chr(37) + ">") 
sp.WriteLine("<" + Chr(37) + " ")
sp.WriteLine("'----------------------------------------------")
sp.WriteLine("'         If No results")
sp.WriteLine("'----------------------------------------------")
sp.WriteLine("Else " + Chr(37) + ">")
sp.WriteLine("  <h3>Your query returned no data.<br></h3>")
sp.WriteLine(" ")
sp.WriteLine("<" + Chr(37) + " End If " + Chr(37) + ">")
sp.WriteLine("<INPUT TYPE=" + chr(34) + "SUBMIT" + chr(34) + " NAME = " + chr(34) + "GoBack" + chr(34) + " VALUE = " + chr(34) + "New Query" + chr(34) + ">")
sp.WriteLine(Chr(60) + Chr(37))
sp.WriteLine("rs.Close")
sp.WriteLine("conn.Close")
sp.WriteLine("set rs = nothing")
sp.WriteLine("set Conn = nothing")
sp.WriteLine(Chr(37) + Chr(62))
sp.WriteLine("</FORM>")
sp.WriteLine("<!" & "--#include file = """ & ProjectName & "_footer.inc"" -->")
sp.WriteLine("</BODY>")
sp.WriteLine("</HTML>")





sp.Close

' ------------------------------ End Drill Down


' -----------------------------

' ------------------------------ Start _groupby

' write the _groupby.asp page
Set sp = fs.CreateTextFile(TheFullPath  & ProjectName & "_groupby.asp", True)



sp.WriteLine(Chr(60) + Chr(37) + "@ LANGUAGE=VBScript" + chr(37) + ">")
sp.WriteLine("")
sp.WriteLine("<head>")
sp.WriteLine("<title>Group by Page</title>")
sp.WriteLine("<LINK REL=""stylesheet"" TYPE=""text/css"" HREF=""/thetimes/shboom.css"">")
sp.WriteLine("</style>")
sp.WriteLine("</head>")
sp.WriteLine("<html>")
sp.WriteLine("")
sp.WriteLine("<body  bgcolor=""white"">")
sp.WriteLine("<!" & "--#include file = """ & ProjectName & "_links.inc"" -->")
sp.WriteLine("<!" & "--#include file = """ & ProjectName & "_header.inc"" -->")
sp.WriteLine("")


sp.WriteLine("<form method=""POST"" action=""" & ProjectName & "_groupbyresults.asp"">")
sp.WriteLine("<p>")
'new
sp.WriteLine("<table border=""0"">")
sp.WriteLine("<tr><td colspan=3 class=pageTitle>" & objRS2("GroupByTitle") & "</td></tr>")
sp.WriteLine("<tr><td colspan=3 class=instructions>Fields will be grouped, displayed and counted in the order you select.</td></tr>")

sp.WriteLine("<tr>")
'end new



'new
sp.WriteLine("  <td class=fieldInstructions>Select one or more 	fields to group by:</td>")
'end new
'sp.WriteLine("</tr>")


sp.WriteLine("")
sp.WriteLine("<td><table cellspacing=0 cellpadding=0 border=0>")
sp.WriteLine("<tr>")
sp.WriteLine("<" & chr(37) & " For I = 1 to " & objRS2("MaxFields") & " " & chr(37) & ">")
sp.WriteLine("")
sp.WriteLine("<tr>")
sp.WriteLine("<!" & "--#include file = """ & ProjectName & "_inputstuff.inc"" -->")
sp.WriteLine("</tr>")
sp.WriteLine("")
sp.WriteLine("<" & chr(37) & " Next " & chr(37) & ">")

sp.WriteLine("</tr>")
sp.WriteLine("</table>")

sp.WriteLine( "</td>")
sp.WriteLine("<td align=right valign=top><input type=submit value=Submit></td>")
sp.WriteLine("</tr>")

' new
sp.WriteLine("<tr><td class=fieldInstructions>Show results in:</td>")
sp.WriteLine("<td><select name=""frmShowOrder"">")
sp.WriteLine("<option value=""default"" selected>The Same Order as Selected Above</option>")
sp.WriteLine("<option value="" order by count(*)"" >From the Lowest Count to the Highest</option>")
sp.WriteLine("<option value="" order by count(*) desc"">From the Highest Count to the Lowest</option>")
sp.WriteLine("</select>")
sp.WriteLine("</td></tr>")

sp.WriteLine("<tr><td class=fieldInstructions>Maximum number of rows to show:</td>")

sp.WriteLine("<td><input type=""text"" name=""frmMaxRet""  size=""10""></td></tr>")
sp.WriteLine("<tr><td class=fieldDetails colspan=2>If you leave this blank or enter a value greater than 5,000, a maximum of 5,000 records will be returned.")
sp.WriteLine("</td>")
sp.WriteLine("</tr>")
sp.WriteLine("</table>")

' end new
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("<h4>Select records that meet the following conditions:</h4>")
sp.WriteLine("<table>")
sp.WriteLine("<tr>")
sp.WriteLine("  <td class=fieldHeading>Field</td>")
sp.WriteLine("  <td class=fieldHeading>Show Rows Where</td>")
sp.WriteLine("  <td class=fieldHeading>Value 1</td>")
sp.WriteLine("  <td class=fieldHeading>And / or</td>")
sp.WriteLine("  <td class=fieldHeading>Show Rows Where</td>")
sp.WriteLine("  <td class=fieldHeading>Value 2</td>")
sp.WriteLine("")
sp.WriteLine("</tr>")
sp.WriteLine("")
sp.WriteLine("<% For I = 1 to " & objRS2("MaxFields") & " " & chr(37) & ">")
sp.WriteLine("")
sp.WriteLine("<tr>")
sp.WriteLine("<!" & "--#include file = """ & ProjectName & "_efilter.inc"" -->")
sp.WriteLine("</tr>")
sp.WriteLine("")
sp.WriteLine("<% Next " & chr(37) & ">")
sp.WriteLine("</table>")
sp.WriteLine("  <center><p><input type=""submit"" value=""Submit""><input")
sp.WriteLine("  type=""reset"" value=""Reset"" name=""B2""></p></center>")
sp.WriteLine("")
sp.WriteLine("</form>")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("<!" & "--#include file = """ & ProjectName & "_footer.inc"" -->")
sp.WriteLine("</body>")
sp.WriteLine("</html>")

sp.Close




' write the _groupbyresults.asp page
Set sp = fs.CreateTextFile(TheFullPath  & ProjectName & "_groupbyresults.asp", True)

sp.WriteLine(Chr(60) + Chr(37) + "@ LANGUAGE=VBScript" + chr(37) + ">")
sp.WriteLine("")
sp.WriteLine(Chr(60) + Chr(37))
sp.WriteLine("'---- CursorTypeEnum Values ----")
sp.WriteLine("Const adOpenForwardOnly = 0")
sp.WriteLine("Const adOpenKeyset = 1")
sp.WriteLine("Const adOpenDynamic = 2")
sp.WriteLine("Const adOpenStatic = 3")
sp.WriteLine("'---- LockTypeEnum Values ----")
sp.WriteLine("Const adLockReadOnly = 1")
sp.WriteLine("Const adLockPessimistic = 2")
sp.WriteLine("Const adLockOptimistic = 3")
sp.WriteLine("Const adLockBatchOptimistic = 4")
sp.WriteLine("'---- CommandTypeEnum Values ----")
sp.WriteLine("Const adCmdUnknown = &H0008")
sp.WriteLine("Const adCmdText = &H0001")
sp.WriteLine("Const adCmdTable = &H0002")
sp.WriteLine("Const adCmdStoredProc = &H0004")
sp.WriteLine("Const adCmdFile = &H0100")
sp.WriteLine("Const adCmdTableDirect = &H0200")
sp.WriteLine("Server.ScriptTimeout = 600")
sp.WriteLine("")
sp.WriteLine("Dim TheSource(" & NP & "), TheESource(" & NP & "), TheType(" & NP & "), Filter1(" & NP & "), Filter2(" & NP & "), Value1(" & NP & "), Value2(" & NP & "), AndOr(" & NP & ")")
sp.WriteLine("")
sp.WriteLine("TheCount=0")
sp.WriteLine("For I = 1 to " & objRS2("MaxFields"))
sp.WriteLine("")
sp.WriteLine("  TestSource = Request(""frmSource"")(I)")
sp.WriteLine("  If TestSource = """" Then")
sp.WriteLine("          Exit For")
sp.WriteLine("  Else")
sp.WriteLine("          TheSource(I) = TestSource")
sp.WriteLine("          TheCount = I")
sp.WriteLine("  End If")
sp.WriteLine("'Response.Write(I)")
sp.WriteLine("Next")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("TheECount=0")
sp.WriteLine("For I = 1 to " & objRS2("MaxFields"))
sp.WriteLine("")
sp.WriteLine("  TestESource = Request(""frmESource"")(I)")
sp.WriteLine("  If TestESource = """" Then")
sp.WriteLine("          Exit For")
sp.WriteLine("  Else")
sp.WriteLine("          SplitESource=Split(TestESource, ""*"")")
sp.WriteLine("          TheESource(I) = SplitESource(0)")
sp.WriteLine("          TheType(I) = UCASE(SplitESource(1))")
sp.WriteLine("          Filter1(I) = Request(""frmFilter1"")(I)")
sp.WriteLine("          Filter2(I) = Request(""frmFilter2"")(I)")
sp.WriteLine("          Value1(I) = Request(""frmValue1"")(I)")
sp.WriteLine("          Value2(I) = Request(""frmValue2"")(I)")
sp.WriteLine("          AndOr(I) = Request(""frmAndOr"")(I)")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("          TheECount = I")
sp.WriteLine("  End If")
sp.WriteLine("'Response.Write(I)")
sp.WriteLine("Next")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("          SQLSelect = ""Select """)
sp.WriteLine("          SQLGroupBy = "" Group by """)
sp.WriteLine("          For I = 1 to TheCount")
sp.WriteLine("                  ")
' ------------------------------------ Revised Dec. 20, 2005 ------------------------------------------
sp.WriteLine("                  If I <> TheCount Then")
sp.WriteLine("                          If (InStr(TheSource(I), ""*Y"") = FALSE) and (InStr(TheSource(I), ""*M"") = FALSE) Then")
sp.WriteLine("                                  SQLSelect = SQLSelect & TheSource(I) & "", """)
sp.WriteLine("                                  SQLGroupBy = SQLGroupBy & TheSource(I) & "", """)
sp.WriteLine("                          ElseIf InStr(TheSource(I), ""*Y"") Then ")
sp.WriteLine("                                  SourceYear = """"")
sp.WriteLine("                                  SourceSplit = Split(TheSource(I), ""*Y"")")
sp.WriteLine("                                  SourceYear = SourceSplit(0)")
sp.WriteLine("                                  SQLSelect = SQLSelect & "" DatePart(yyyy, "" & SourceYear & "") As "" & SourceYear & "", """)
sp.WriteLine("                                  SQLGroupBy = SQLGroupBy & "" DatePart(yyyy, "" & SourceYear & ""), """)
sp.WriteLine("                                  TheSource(I) = Replace(TheSource(I), ""*Y"", """")")
sp.WriteLine("                          Else ")
sp.WriteLine("                                  SourceYear = """"")
sp.WriteLine("                                  SourceSplit = Split(TheSource(I), ""*M"")")
sp.WriteLine("                                  SourceYear = SourceSplit(0)")
sp.WriteLine("                                  SQLSelect = SQLSelect & "" Cast(DatePart(yyyy,  "" & SourceYear & "") as varchar(4))+'-'+right(Cast(datepart(mm, "" & SourceYear & "") + 1000 as varchar(4)), 2) As "" & SourceYear & "", """)
sp.WriteLine("                                  SQLGroupBy = SQLGroupBy & "" Cast(DatePart(yyyy,  "" & SourceYear & "") as varchar(4))+'-'+right(Cast(datepart(mm, "" & SourceYear & "") + 1000 as varchar(4)), 2), """)
sp.WriteLine("                                  TheSource(I) = Replace(TheSource(I), ""*M"", """")")
sp.WriteLine("                          End If")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("                  Else")
sp.WriteLine("                          If (InStr(TheSource(I), ""*Y"") = FALSE) and (InStr(TheSource(I), ""*M"") = FALSE) Then")
sp.WriteLine("                                  SQLSelect = SQLSelect & TheSource(I)")
sp.WriteLine("                                  SQLGroupBy = SQLGRoupBy & TheSource(I)")
sp.WriteLine("                          ElseIf InStr(TheSource(I), ""*Y"") Then ")
sp.WriteLine("                                  SourceYear = """"")
sp.WriteLine("                                  SourceSplit = Split(TheSource(I), ""*Y"")")
sp.WriteLine("                                  SourceYear = SourceSplit(0)")
sp.WriteLine("                                  SQLSelect = SQLSelect & "" DatePart(yyyy, "" & SourceYear & "") As "" & SourceYear ")
sp.WriteLine("                                  SQLGroupBy = SQLGroupBy & "" DatePart(yyyy, "" & SourceYear & "")""")
sp.WriteLine("                                  TheSource(I) = Replace(TheSource(I), ""*Y"", """")")
sp.WriteLine("                          Else ")
sp.WriteLine("                                  SourceYear = """"")
sp.WriteLine("                                  SourceSplit = Split(TheSource(I), ""*M"")")
sp.WriteLine("                                  SourceYear = SourceSplit(0)")
sp.WriteLine("                                  SQLSelect = SQLSelect & "" Cast(DatePart(yyyy,  "" & SourceYear & "") as varchar(4))+'-'+right(Cast(datepart(mm, "" & SourceYear & "") + 1000 as varchar(4)), 2) As "" & SourceYear ")
sp.WriteLine("                                  SQLGroupBy = SQLGroupBy & "" Cast(DatePart(yyyy,  "" & SourceYear & "") as varchar(4))+'-'+right(Cast(datepart(mm, "" & SourceYear & "") + 1000 as varchar(4)), 2) """)
sp.WriteLine("                                  TheSource(I) = Replace(TheSource(I), ""*M"", """")")

sp.WriteLine("                          End If")
sp.WriteLine("                  End If")
' ------------------------------------ Revised Dec. 20, 2005 ------------------------------------------
sp.WriteLine("                  ")
sp.WriteLine("          Next")
sp.WriteLine("'End Select                       ")
sp.WriteLine("                           ")
sp.WriteLine("")
sp.WriteLine("'------------------------------------ See if there's anything to filter")
sp.WriteLine("")
sp.WriteLine("If TheESource(1) <> """" Then")
sp.WriteLine("")
sp.WriteLine("' make the sql where clause")
sp.WriteLine("")
sp.WriteLine("sqlWhere = ""WHERE """)
sp.WriteLine("For I = 1 to TheECount")
sp.WriteLine("WhereHolder1 = """"")
sp.WriteLine("WhereHolder2 = """"")
sp.WriteLine("")
sp.WriteLine("'Response.Write(""The ESource is: "" & TheESource(I) & ""<br>"")")
sp.WriteLine("")
sp.WriteLine("' see if there's an and/or clause")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("  If AndOr(I) = """" Then")
sp.WriteLine("          Select Case TheType(I)")
sp.WriteLine("                  Case ""N""")
sp.WriteLine("                          WhereHolder1 = MakeWhereNum(TheESource(I), Filter1(I), Value1(I))")
sp.WriteLine("                  Case ""T""")
sp.WriteLine("                          WhereHolder1 = MakeWhereText(TheESource(I), Filter1(I), Value1(I))")
sp.WriteLine("                  Case ""D""")
sp.WriteLine("                          WhereHolder1 = MakeWhereDate(TheESource(I), Filter1(I), Value1(I))")
sp.WriteLine("          End Select")
sp.WriteLine("")
sp.WriteLine("          If I <> TheECount Then")
sp.WriteLine("                  sqlWhere = sqlWhere & "" ("" & WhereHolder1 & "") AND """)
sp.WriteLine("          Else")
sp.WriteLine("                  sqlWhere = sqlWhere & "" ("" & WhereHolder1 & "") """)
sp.WriteLine("          End If")
sp.WriteLine("                  ")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("  Else ' there's an and or or     ")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("          Select Case TheType(I)")
sp.WriteLine("                  Case ""N""")
sp.WriteLine("                          WhereHolder1 = MakeWhereNum(TheESource(I), Filter1(I), Value1(I))")
sp.WriteLine("                          WhereHolder2 = MakeWhereNum(TheESource(I), Filter2(I), Value2(I))")
sp.WriteLine("                  Case ""T""")
sp.WriteLine("                          WhereHolder1 = MakeWhereText(TheESource(I), Filter1(I), Value1(I))")
sp.WriteLine("                          WhereHolder2 = MakeWhereText(TheESource(I), Filter2(I), Value2(I))")
sp.WriteLine("                  Case ""D""")
sp.WriteLine("                          WhereHolder1 = MakeWhereDate(TheESource(I), Filter1(I), Value1(I))")
sp.WriteLine("                          WhereHolder2 = MakeWhereDate(TheESource(I), Filter2(I), Value2(I))")
sp.WriteLine("          End Select")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("          If I <> TheECount Then")
sp.WriteLine("                  sqlWhere = sqlWhere & "" (("" & WhereHolder1 & "") "" & AndOr(I) & "" ("" & WhereHolder2 & "")) AND """)
sp.WriteLine("          Else")
sp.WriteLine("                  sqlWhere = sqlWhere & "" (("" & WhereHolder1 & "") "" & AndOr(I) & "" ("" & WhereHolder2& "")) """)
sp.WriteLine("          End If")
sp.WriteLine("                  ")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("  End If")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("Next")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("End If ' ---------------------------- end of where clause")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("'Test to see if the " + chr(34) + "New Query" + chr(34) + " button was clicked")
sp.WriteLine("  If Request(" + chr(34) + "GoBack" + chr(34) + ") = " + chr(34) + "New Query" + chr(34) + " Then")
sp.WriteLine("                  Response.Redirect(" + chr(34) + ProjectName + "_groupby.asp" + chr(34) + ")")
sp.WriteLine("  End If")
sp.WriteLine("' If there is no sql statement from asp form")
sp.WriteLine("' test for input from default.asp")
sp.WriteLine("' test for three characters")
sp.WriteLine("If Len(""Duck"") < 2 Then")
sp.WriteLine("  Response.Redirect(""default.asp"")")
sp.WriteLine("Else")
sp.WriteLine("' build a Select statement")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("'response.Write(""TorB = "" & TorB & ""<br>"")")
'new
sp.WriteLine("ShowOrder=Request(""frmShowOrder"")")
sp.WriteLine("")
sp.WriteLine("If ShowOrder = ""default"" Then")
sp.WriteLine("     SQLOrderBy = Replace(SQLGroupBy, "" Group "", "" Order "")")
sp.WriteLine("Else")
sp.WriteLine("     SQLOrderBy = ShowOrder")
sp.WriteLine("End If")
'end new
sp.WriteLine("")
'sp.WriteLine("SQL = SQLSelect & "", count(*) as Total from "" & theTable & "" "" & sqlWhere &  SQLGroupBy & SQLOrderBy")
sp.WriteLine("SQL = SQLSelect & "", count(*) as Total from " & objRS2("TheTable")  & " " & chr(34) & " &  sqlWhere &  SQLGroupBy & SQLOrderBy")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
If objRS2("DisplaySQL") = "YES" Then
        sp.WriteLine("response.write(""The SQL is "" & SQL & ""<br>"")")
Else
        sp.WriteLine("'response.write(""The SQL is "" & SQL & ""<br>"")")
End If
sp.WriteLine("")
sp.WriteLine("End If")
sp.WriteLine("")

sp.WriteLine("")
sp.WriteLine("dbName = " & chr(34)  & objRS2("TheDB") & chr(34) & " "  & chr(37) + ">")
sp.WriteLine("")
sp.WriteLine("")

sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("<!" & "--#include virtual = ""/thetimes/longrcon.inc"" -->")
sp.WriteLine("")
sp.WriteLine(Chr(60) + Chr(37))


sp.WriteLine("")
'new
sp.WriteLine("MaxRecs = """" & Request(""frmMaxRet"")")
sp.WriteLine("If MaxRecs <> """"  Then")
sp.WriteLine("  If MaxRecs > 5000 Then MaxRecs = 5000")
sp.WriteLine("  RS.MaxRecords = MaxRecs")
sp.WriteLine("Else")
sp.WriteLine("  RS.MaxRecords = 5000")
sp.WriteLine("End If")

'end new

sp.WriteLine("")
sp.WriteLine("RS.Open sql, Conn, adOpenStatic, adLockReadOnly, adCmdText")
sp.WriteLine("")
'new
sp.WriteLine("RecordCount = RS.RecordCount")
'new
sp.WriteLine("")

sp.WriteLine(chr(37) & ">")
sp.WriteLine("<HTML>")
sp.WriteLine("<head>")
sp.WriteLine("<title>Group By Results</title>")
sp.WriteLine("<LINK REL=""stylesheet"" TYPE=""text/css"" HREF=""/thetimes/shboom.css"">")
sp.WriteLine("</head>")
sp.WriteLine("<body  bgcolor=""white"">")
sp.WriteLine("<!" & "--#include file = """ & ProjectName & "_links.inc"" -->")
'new
sp.WriteLine("<H1 class=pageTitle>" & objRS2("ProjectAlias") & " Group By Results</H1>")
sp.WriteLine("<" & chr(37))
sp.WriteLine("If RecordCount > 5000 Then " & chr(37) & ">")
sp.WriteLine("<center><b>Your query returned 5,000 records, the maximum allowed.</b></center>")
sp.WriteLine("<" & chr(37) & " End If " & chr(37) & ">")
'end new




sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("<table class=resultsTable cellpadding=0 cellspacing=0  >")
sp.WriteLine("  <tr>")
sp.WriteLine("")
sp.WriteLine("<" & chr(37))
sp.WriteLine("")
sp.WriteLine("For I = 1 to TheCount ")
sp.WriteLine("TheFieldName = TextTitle(TheSource(I))")
sp.WriteLine("")
sp.WriteLine(Chr(37) & ">")
sp.WriteLine("")
sp.WriteLine("<td class=resultsHeading><" & Chr(37) & " = TheFieldName " & chr(37) & "></td>")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("<" & chr(37) & " Next " & Chr(37) & ">")
sp.WriteLine("")
sp.WriteLine("    <td class=resultsHeading>")
sp.WriteLine("        Count</td>")
sp.WriteLine("    <td class=resultsHeading>")
sp.WriteLine("        Show Records</td>")
sp.WriteLine("  </tr>")
sp.WriteLine("")
sp.WriteLine("<" & chr(37))
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine(" Do While Not RS.EOF " & chr(37) & ">")
sp.WriteLine("  <tr>")
sp.WriteLine("")
sp.WriteLine("<" & chr(37))
sp.WriteLine("")
sp.WriteLine("SumHref=" & chr(34) & chr(34))
sp.WriteLine("")
sp.WriteLine("For I = 1 to TheCount")
sp.WriteLine("")
sp.WriteLine("ColHRef=" & chr(34) & chr(34))
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("If Trim(rs(TheSource(I))) = """" Then ")
sp.WriteLine("")
sp.WriteLine("  colHref=""frm"" & TheSource(I) & ""=NULL""")
sp.WriteLine("")
sp.WriteLine("  If I = 1 Then")
sp.WriteLine("          SumHref = SumHref & ColHref")
sp.WriteLine("  Else")
sp.WriteLine("          SumHref = SumHref & ""&"" & ColHref")
sp.WriteLine("  End If")
sp.WriteLine("")
sp.WriteLine(chr(37) & ">")
sp.WriteLine("    <td class=resultsData><a href=" & chr(34) & ProjectName & "_drilldown.asp?<" & chr(37) & " = ColHref & ""&frmsqlWhere="" & Server.URLEncode(sqlWhere)"  & chr(37) & ">"" class=resultsDrillDown>Blank</a></td>")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("<" & chr(37) & " ElseIf  IsNull(rs(TheSource(I))) Then ")

'-------------------------------
sp.WriteLine("")
sp.WriteLine("  colHref=""frm"" & TheSource(I) & ""=NULL""")
sp.WriteLine("")
sp.WriteLine("  If I = 1 Then")
sp.WriteLine("          SumHref = SumHref & ColHref")
sp.WriteLine("  Else")
sp.WriteLine("          SumHref = SumHref & ""&"" & ColHref")
sp.WriteLine("  End If")
sp.WriteLine("")
sp.WriteLine(chr(37) & ">")
sp.WriteLine("    <td class=resultsData><a href=" & chr(34) & ProjectName & "_drilldown.asp?<" & chr(37) & " = ColHref & ""&frmsqlWhere="" & Server.URLEncode(sqlWhere)" & chr(37) & ">"" class=resultsDrillDown>Null</a></td>")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("<" & chr(37) & " Else ")
sp.WriteLine("  colHref=""frm"" & TheSource(I) & ""="" & Server.URLEncode(rs(TheSource(I)))")
sp.WriteLine("")
sp.WriteLine("  If I = 1 Then")
sp.WriteLine("          SumHref = SumHref & ColHref")
sp.WriteLine("  Else")
sp.WriteLine("          SumHref = SumHref & ""&"" & ColHref")
sp.WriteLine("  End If")
sp.WriteLine("")
sp.WriteLine(chr(37) & ">")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("    <td class=resultsData><a href=" & chr(34) & ProjectName & "_drilldown.asp?<" & chr(37) & " = ColHref & ""&frmsqlWhere="" & Server.URLEncode(sqlWhere)"  & chr(37) & ">"" class=resultsDrillDown><" & chr(37) & " = RS(TheSource(I)) " & chr(37) & "></td>")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("<" & chr(37) & " End If " & chr(37) & ">")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("<" & chr(37))
sp.WriteLine("Next")
sp.WriteLine("On error Resume Next")
sp.WriteLine(chr(37) & ">")
sp.WriteLine("")
sp.WriteLine("<td class=resultsData align=""right""><" & chr(37) & " = FormatNumber(rs(""Total""),0) " & chr(37) & "></td>")
sp.WriteLine("<td class=resultsData align=""right""><a href=" & chr(34) & ProjectName & "_drilldown.asp?<" & chr(37) & " = SumHref  & ""&frmsqlWhere="" & Server.URLEncode(sqlWhere)" & chr(37) & ">"" class=resultsDrillDown>All in Row</td>")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("</tr>")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("<" & chr(37) & " rs.MoveNext")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("Loop")
sp.WriteLine(chr(37) & ">")
sp.WriteLine("</TABLE>")
sp.WriteLine("</center>")
sp.WriteLine("</p>")
sp.WriteLine("")
sp.WriteLine("<" & chr(37))
sp.WriteLine("rs.Close")
sp.WriteLine("conn.Close")
sp.WriteLine("set rs = nothing")
sp.WriteLine("set Conn = nothing")
sp.WriteLine(chr(37) & ">")
sp.WriteLine("</BODY>")
sp.WriteLine("</HTML>")
sp.WriteLine("")

sp.WriteLine("<" & chr(37))


sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("Function MakeWhereNum(TheFieldName, TheOperator, TheValue)")
sp.WriteLine("")
sp.WriteLine("Select Case TheOperator")
sp.WriteLine("  Case ""eq""")
sp.WriteLine("          MakeWhereNum = "" "" & TheFieldName & "" = "" & TheValue")
sp.WriteLine("  Case ""noteq""")
sp.WriteLine("          MakeWhereNum = "" "" & TheFieldName & "" <> "" & TheValue")
sp.WriteLine("  Case ""gt""")
sp.WriteLine("          MakeWhereNum = "" "" & TheFieldName & "" > "" & TheValue")
sp.WriteLine("  Case ""gtet""")
sp.WriteLine("          MakeWhereNum = "" "" & TheFieldName & "" >= "" & TheValue")
sp.WriteLine("  Case ""lt""")
sp.WriteLine("          MakeWhereNum = "" "" & TheFieldName & "" < "" & TheValue")
sp.WriteLine("  Case ""ltet""")
sp.WriteLine("          MakeWhereNum = "" "" & TheFieldName & "" <= "" & TheValue")
sp.WriteLine("  Case ""bw""")
sp.WriteLine("          MakeWhereNum = "" Cast("" & TheFieldName & "" as varchar(25)) Like '"" & TheValue & ""%'""")
sp.WriteLine("  Case ""dnbw""")
sp.WriteLine("          MakeWhereNum = "" Cast("" & TheFieldName & "" as varchar(25)) NOT Like '"" & TheValue & ""%'""")
sp.WriteLine("  Case ""ew""")
sp.WriteLine("          MakeWhereNum = "" Cast("" & TheFieldName & "" as varchar(25)) Like '%"" & TheValue & ""'""")
sp.WriteLine("  Case ""dnew""")
sp.WriteLine("          MakeWhereNum = "" Cast("" & TheFieldName & "" as varchar(25)) NOT Like '%"" & TheValue & ""'""")
sp.WriteLine("  Case ""con""")
sp.WriteLine("          MakeWhereNum = "" Cast("" & TheFieldName & "" as varchar(25)) Like '%"" & TheValue & ""%'""")
sp.WriteLine("  Case ""dncon""")
sp.WriteLine("          MakeWhereNum = "" Cast("" & TheFieldName & "" as varchar(25)) NOT Like '%"" & TheValue & ""%'""")
sp.WriteLine("  Case ""blank""")
sp.WriteLine("          MakeWhereNum = "" "" & TheFieldName & "" is null """)
sp.WriteLine("  Case ""notblank""")
sp.WriteLine("          MakeWhereNum = "" "" & TheFieldName & "" is NOT null  "" ")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("End Select")
sp.WriteLine("")
sp.WriteLine("End Function")
sp.WriteLine("")
sp.WriteLine("Function MakeWhereText(TheFieldName, TheOperator, TheValue)")
sp.WriteLine("")
sp.WriteLine("If InStr(TheValue, ""'"") Then")
sp.WriteLine("  TheValue = Replace(TheValue, ""'"", ""''"")")
sp.WriteLine("End If")
sp.WriteLine("Select Case TheOperator")
sp.WriteLine("  Case ""eq""")
sp.WriteLine("          MakeWhereText = "" "" & TheFieldName & "" = '"" & TheValue & ""'""")
sp.WriteLine("          If InStr(MakeWhereText, ""]"") and InStr(MakeWhereText, ""["") Then")
sp.WriteLine("                  MakeWhereText = Replace(MakeWhereText, ""'"", """")")
sp.WriteLine("          End If")
sp.WriteLine("  Case ""noteq""")
sp.WriteLine("          MakeWhereText = "" "" & TheFieldName & "" <> '"" & TheValue & ""'""")
sp.WriteLine("          If InStr(MakeWhereText, ""]"") and InStr(MakeWhereText, ""["") Then")
sp.WriteLine("                  MakeWhereText = Replace(MakeWhereText, ""'"", """")")
sp.WriteLine("          End If")
sp.WriteLine("  Case ""gt""")
sp.WriteLine("          MakeWhereText = "" "" & TheFieldName & "" > '"" & TheValue & ""'""")
sp.WriteLine("          If InStr(MakeWhereText, ""]"") and InStr(MakeWhereText, ""["") Then")
sp.WriteLine("                  MakeWhereText = Replace(MakeWhereText, ""'"", """")")
sp.WriteLine("          End If")
sp.WriteLine("  Case ""gtet""")
sp.WriteLine("          MakeWhereText = "" "" & TheFieldName & "" >=  '"" & TheValue & ""'""")
sp.WriteLine("          If InStr(MakeWhereText, ""]"") and InStr(MakeWhereText, ""["") Then")
sp.WriteLine("                  MakeWhereText = Replace(MakeWhereText, ""'"", """")")
sp.WriteLine("          End If")
sp.WriteLine("  Case ""lt""")
sp.WriteLine("          MakeWhereText = "" "" & TheFieldName & "" <  '"" & TheValue & ""'""")
sp.WriteLine("          If InStr(MakeWhereText, ""]"") and InStr(MakeWhereText, ""["") Then")
sp.WriteLine("                  MakeWhereText = Replace(MakeWhereText, ""'"", """")")
sp.WriteLine("          End If")
sp.WriteLine("  Case ""ltet""")
sp.WriteLine("          MakeWhereText = "" "" & TheFieldName & "" <= '"" & TheValue & ""'""")
sp.WriteLine("          If InStr(MakeWhereText, ""]"") and InStr(MakeWhereText, ""["") Then")
sp.WriteLine("                  MakeWhereText = Replace(MakeWhereText, ""'"", """")")
sp.WriteLine("          End If")
sp.WriteLine("  Case ""bw""")
sp.WriteLine("          MakeWhereText = "" "" & TheFieldName & "" Like '"" & TheValue & ""%'""")
sp.WriteLine("  Case ""dnbw""")
sp.WriteLine("          MakeWhereText = "" "" & TheFieldName & "" NOT Like '"" & TheValue & ""%'""")
sp.WriteLine("  Case ""ew""")
sp.WriteLine("          MakeWhereText = "" "" & TheFieldName & "" Like '%"" & TheValue & ""'""")
sp.WriteLine("  Case ""dnew""")
sp.WriteLine("          MakeWhereText = "" "" & TheFieldName & "" NOT Like '%"" & TheValue & ""'""")
sp.WriteLine("  Case ""con""")
sp.WriteLine("          MakeWhereText = "" "" & TheFieldName & "" Like '%"" & TheValue & ""%'""")
sp.WriteLine("  Case ""dncon""")
sp.WriteLine("          MakeWhereText = "" "" & TheFieldName & "" NOT Like '%"" & TheValue & ""%'""")
sp.WriteLine("  Case ""blank""")
sp.WriteLine("          MakeWhereText = "" ("" & TheFieldName & "" is Null or "" & TheFieldName & "" = '') """)
sp.WriteLine("  Case ""notblank""")
sp.WriteLine("          MakeWhereText = "" ("" & TheFieldName & "" is NOT Null and "" & TheFieldName & "" <> '') """)
sp.WriteLine("          ")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("End Select")
sp.WriteLine("")
sp.WriteLine("End Function")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("Function MakeWhereDate(TheFieldName, TheOperator, TheValue)")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("Select Case TheOperator")
sp.WriteLine("  Case ""eq""")
sp.WriteLine("          MakeWhereDate = "" "" & TheFieldName & "" = '"" & TheValue & ""'""")
sp.WriteLine("          If InStr(MakeWhereDate, ""]"") and InStr(MakeWhereDate, ""["") Then")
sp.WriteLine("                  MakeWhereDate = Replace(MakeWhereDate, ""'"", """")")
sp.WriteLine("          End If")
sp.WriteLine("  Case ""noteq""")
sp.WriteLine("          MakeWhereDate = "" "" & TheFieldName & "" <> '"" & TheValue & ""'""")
sp.WriteLine("          If InStr(MakeWhereDate, ""]"") and InStr(MakeWhereDate, ""["") Then")
sp.WriteLine("                  MakeWhereDate = Replace(MakeWhereDate, ""'"", """")")
sp.WriteLine("          End If")
sp.WriteLine("  Case ""gt""")
sp.WriteLine("          MakeWhereDate = "" "" & TheFieldName & "" > '"" & TheValue & ""'""")
sp.WriteLine("          If InStr(MakeWhereDate, ""]"") and InStr(MakeWhereDate, ""["") Then")
sp.WriteLine("                  MakeWhereDate = Replace(MakeWhereDate, ""'"", """")")
sp.WriteLine("          End If")
sp.WriteLine("  Case ""gtet""")
sp.WriteLine("          MakeWhereDate = "" "" & TheFieldName & "" >=  '"" & TheValue & ""'""")
sp.WriteLine("          If InStr(MakeWhereDate, ""]"") and InStr(MakeWhereDate, ""["") Then")
sp.WriteLine("                  MakeWhereDate = Replace(MakeWhereDate, ""'"", """")")
sp.WriteLine("          End If")
sp.WriteLine("  Case ""lt""")
sp.WriteLine("          MakeWhereDate = "" "" & TheFieldName & "" <  '"" & TheValue & ""'""")
sp.WriteLine("          If InStr(MakeWhereDate, ""]"") and InStr(MakeWhereDate, ""["") Then")
sp.WriteLine("                  MakeWhereDate = Replace(MakeWhereDate, ""'"", """")")
sp.WriteLine("          End If")
sp.WriteLine("  Case ""ltet""")
sp.WriteLine("          MakeWhereDate = "" "" & TheFieldName & "" <= '"" & TheValue & ""'""")
sp.WriteLine("          If InStr(MakeWhereDate, ""]"") and InStr(MakeWhereDate, ""["") Then")
sp.WriteLine("                  MakeWhereDate = Replace(MakeWhereDate, ""'"", """")")
sp.WriteLine("          End If")
sp.WriteLine("  Case ""bw""")
sp.WriteLine("          MakeWhereDate = "" Cast("" & TheFieldName & "" as varchar(25)) Like '"" & TheValue & ""%'""")
sp.WriteLine("  Case ""dnbw""")
sp.WriteLine("          MakeWhereDate = "" Cast("" & TheFieldName & "" as varchar(25)) NOT Like '"" & TheValue & ""%'""")
sp.WriteLine("  Case ""ew""")
sp.WriteLine("          MakeWhereDate = "" Cast("" & TheFieldName & "" as varchar(25)) Like '%"" & TheValue & ""'""")
sp.WriteLine("  Case ""dnew""")
sp.WriteLine("          MakeWhereDate = "" Cast("" & TheFieldName & "" as varchar(25)) NOT Like '%"" & TheValue & ""'""")
sp.WriteLine("  Case ""con""")
sp.WriteLine("          MakeWhereDate = "" Cast("" & TheFieldName & "" as varchar(25)) Like '%"" & TheValue & ""%'""")
sp.WriteLine("  Case ""dncon""")
sp.WriteLine("          MakeWhereDate = "" Cast("" & TheFieldName & "" as varchar(25)) NOT Like '%"" & TheValue & ""%'""")
sp.WriteLine("  Case ""blank""")
sp.WriteLine("          MakeWhereDate = "" "" & TheFieldName & "" is NULL "" ")
sp.WriteLine("  Case ""notblank""")
sp.WriteLine("          MakeWhereDate = "" "" & TheFieldName & "" is NOT NULL "" ")
sp.WriteLine("")
sp.WriteLine("End Select")
sp.WriteLine("")
sp.WriteLine("End Function")
sp.WriteLine("")
sp.WriteLine("Function TextTitle(gaga)")
sp.WriteLine("  Select Case gaga")

rs.Filter = ""
rs.Filter = ("ShowPulldown = TRUE")
rs.MoveFirst

Do While NOT rs.EOF

sp.WriteLine("          Case " & chr(34) & rs("TheField") & chr(34))
sp.WriteLine("                  TextTitle=" & chr(34) & rs("TheAlias") & chr(34))

rs.MoveNext
Loop

sp.WriteLine("End Select")

sp.WriteLine("End Function")


sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine(chr(37) & ">")

' ----------------------------- 
' make appropriate _links.inc page




Set sp = fs.CreateTextFile(TheFullPath & ProjectName & "_links.inc", True)


sp.WriteLine "<!-- #include virtual=""/thetimes/nytTabBar.asp"" -->"
sp.WriteLine "<%"
sp.WriteLine "Dim tb"
sp.WriteLine "Set tb = new nytTabBar"

'sp.WriteLine("<center>")
'sp.WriteLine("<small>")

rs.Filter=""
rs.Filter=("ShowCalc=TRUE")
rs.Sort=("TheOrder")

CountNumerics = rs.RecordCount

'sp.WriteLine("<a href=""" & ProjectName & "_filter.asp"">" & objRS2("FilterTitle") & "</a> | ")
'sp.WriteLine("<a href=""" & ProjectName & "_groupby.asp"">" & objRS2("GroupByTitle") & "</a>")

sp.WriteLine "tb.AddTab ""fil"", """ & ProjectName & "_filter.asp"", """ & objRS2("FilterTitle") & """"
sp.WriteLine "tb.AddTab ""gro"", """ & ProjectName & "_groupby.asp"", """ & objRS2("GroupByTitle") & """"

If CountNumerics > 0 Then 
        '       sp.WriteLine(" | <a href=""" & ProjectName & "_expressions.asp"">" & objRS2("CalcTitle") & "</a>")

        sp.WriteLine "tb.AddTab ""exp"", """ & ProjectName & "_expressions.asp"", """ & objRS2("CalcTitle") & """"

End If

'       sp.WriteLine(" | <a href=""" & ProjectName & "_tips.asp"" target=""_blank"">Help</a>")
        sp.WriteLine "tb.AddTab ""tip"", """ & ProjectName & "_tips.asp"", ""Help"""



'Response.Write("Allow User Mod: " & objRS2("AllowUserMod") & "<br>")

If objRS2("AllowUserMod") = "YES" Then
        'sp.WriteLine(" | <a href=""/shboom/modify.asp?TheFullPath=" & Server.URLEncode(TheFullPath) & "&ProjectName="& ProjectName & Chr(34) & ">Modify</a>")
        sp.WriteLine "tb.AddTab ""modify"", ""/shboom/modify.asp?TheFullPath=" & Server.URLEncode(TheFullPath) & "&ProjectName="& ProjectName & """, ""Modify"""

        
End If
sp.WriteLine "tb.SetImagePrefix ""/theTimes/"" "
sp.WriteLine "tb.WriteCSS "
sp.WriteLine "tb.DrawTopper """ & objRS2("ProjectAlias") & """"


sp.WriteLine "Dim j, c, sn, inkey, key"
sp.WriteLine "sn = request.servervariables(""script_name"")"
sp.WriteLine "for j = 1 to len(sn)"
sp.WriteLine "  c = mid( sn, j, 1)"
sp.WriteLine "  if c = ""_"" then"
sp.WriteLine "          inkey = -1"     
sp.WriteLine "          key = """" "
sp.WriteLine "  elseif inkey then"
sp.WriteLine "          key = key & c"
sp.WriteLine "          if len( key ) = 3 then"
sp.WriteLine "                  inkey = 0"
sp.WriteLine "          end if"
sp.WriteLine "  end if"
sp.WriteLine "next"



sp.WriteLine "tb.Draw key"

sp.WriteLine chr(37) & ">"
sp.WriteLine key
'sp.WriteLine("</small>")
'sp.WriteLine("</center>")

sp.Close




' ---------------   write an _expressions.asp pages if necessary ------------------------------------


If CountNumerics > 0 Then



'---------------------------- make mathoptions ----------------------------------

' write the _mathoptions page
Set sp = fs.CreateTextFile(TheFullPath  & ProjectName & "_mathoptions.inc", True)

sp.WriteLine("<select size=""1"" name=""frmExp"">")
sp.WriteLine("<option value="""""" selected></option>")


Do While Not rs.EOf

                sp.WriteLine("<option value=""" & rs("TheField")  & """>" & rs("TheAlias") & "</option>")


rs.MoveNext

Loop

sp.WriteLine("</select>")

sp.Close

' ------------------------- ended with creation of _mathoptions.inc


Set sp = fs.CreateTextFile(TheFullPath  & ProjectName & "_expressions.asp", True)


sp.WriteLine(Chr(60) + Chr(37) + "@ LANGUAGE=VBScript" + chr(37) + ">")
sp.WriteLine("")
sp.WriteLine("<head>")
sp.WriteLine("<title>Calculations Page</title>")
sp.WriteLine("<LINK REL=""stylesheet"" TYPE=""text/css"" HREF=""/thetimes/shboom.css"">")
sp.WriteLine("<style>")
sp.WriteLine("<!--")
sp.WriteLine("select       { font-size: 7pt }")
sp.WriteLine("-->")
sp.WriteLine("</style>")
sp.WriteLine("</head>")
sp.WriteLine("<html>")
sp.WriteLine("")
sp.WriteLine("<body  bgcolor=""white"">")
sp.WriteLine("<!" & "--#include file = """ & ProjectName & "_links.inc"" -->")
sp.WriteLine("<!" & "--#include file = """ & ProjectName & "_header.inc"" -->")

sp.WriteLine("<form method=""POST"" action=""" & ProjectName & "_expressionresults.asp"">")

sp.WriteLine("<p>")
'new
sp.WriteLine("<table border=""0"">")
sp.WriteLine("<tr><td colspan=3 class=pageTitle>" & objRS2("CalcTitle") & "</td></tr>")
sp.WriteLine("<tr><td colspan=3 class=instructions>Note: You should read the ""Caution"" section on the Help link before reporting the results of a calculation.</td></tr>")

sp.WriteLine("<tr>")
sp.WriteLine("  <td class=fieldInstructions>Calculate the: </td>")
sp.WriteLine("  <td> " )
sp.WriteLine("<select size=""1"" name=""frmCalc"">")
sp.WriteLine("<option value=""AVG"">Average</option>")
sp.WriteLine("<option value=""COUNT"">Count</option>")
sp.WriteLine("<option value=""MIN"">Minimum</option>")
sp.WriteLine("<option value=""Max"">Maximum</option>")
sp.WriteLine("<option value=""SUM"">Sum</option>")
sp.WriteLine("</select>")
sp.WriteLine(" of ")

sp.WriteLine("<!" & "--#include file = """ & ProjectName & "_mathoptions.inc"" -->")
sp.WriteLine("</td>")
sp.WriteLine("<td align=right><input type=submit value=Submit></td>")
sp.WriteLine("</tr>")

sp.WriteLine("<tr>")
sp.WriteLine("  <td class=fieldInstructions>Select one or more 	fields to group by:</td>")

sp.WriteLine("")
sp.WriteLine("<td><table cellspacing=0 cellpadding=0 border=0>")
sp.WriteLine("<tr>")
sp.WriteLine("<" & chr(37) & " For I = 1 to " & objRS2("MaxFields") & " " & chr(37) & ">")
sp.WriteLine("")
sp.WriteLine("<tr>")
sp.WriteLine("<!" & "--#include file = """ & ProjectName & "_inputstuff.inc"" -->")
sp.WriteLine("</tr>")
sp.WriteLine("")
sp.WriteLine("<" & chr(37) & " Next " & chr(37) & ">")
sp.WriteLine("</tr>")
sp.WriteLine("</table>")
sp.WriteLine("</td></tr>")

' new
sp.WriteLine("<tr><td class=fieldInstructions>Show results in:</td>")
sp.WriteLine("<td><select name=""frmShowOrder"">")
sp.WriteLine("<option value=""default"" selected>The Same Order as Selected Above</option>")
sp.WriteLine("<option value=""asc"" >From Lowest/Earliest to Highest/Latest</option>")
sp.WriteLine("<option value=""desc"">From Highest/Latest to the Lowest/Earliest</option>")
sp.WriteLine("</select>")
sp.WriteLine("</td></tr>")

sp.WriteLine("<tr><td class=fieldInstructions>Maximum number of rows to show:</td>")

sp.WriteLine("<td><input type=""text"" name=""frmMaxRet""  size=""10""></td></tr>")
sp.WriteLine("<tr><td class=fieldDetails colspan=2>If you leave this blank or enter a value greater than 5,000, a maximum of 5,000 records will be returned.")
sp.WriteLine("</td>")
sp.WriteLine("</tr>")
sp.WriteLine("</table>")

' end new
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("<h4>Select records that meet the following conditions:</h4>")
sp.WriteLine("<table>")
sp.WriteLine("<tr>")
sp.WriteLine("  <td class=fieldHeading>Field</td>")
sp.WriteLine("  <td class=fieldHeading>Show Rows Where</td>")
sp.WriteLine("  <td class=fieldHeading>Value 1</td>")
sp.WriteLine("  <td class=fieldHeading>And / or</td>")
sp.WriteLine("  <td class=fieldHeading>Show Rows Where</td>")
sp.WriteLine("  <td class=fieldHeading>Value 2</td>")
sp.WriteLine("")
sp.WriteLine("</tr>")
sp.WriteLine("")
sp.WriteLine("<% For I = 1 to " & objRS2("MaxFields") & " " & chr(37) & ">")
sp.WriteLine("")
sp.WriteLine("<tr>")
sp.WriteLine("<!" & "--#include file = """ & ProjectName & "_efilter.inc"" -->")
sp.WriteLine("</tr>")
sp.WriteLine("")
sp.WriteLine("<% Next " & chr(37) & ">")
sp.WriteLine("</table>")
sp.WriteLine("  <center><p><input type=""submit"" value=""Submit""><input")
sp.WriteLine("  type=""reset"" value=""Reset"" name=""B2""></p></center>")
sp.WriteLine("")
sp.WriteLine("</form>")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("<!" & "--#include file = """ & ProjectName & "_footer.inc"" -->")
sp.WriteLine("</body>")
sp.WriteLine("</html>")
sp.Close



'------------------ write the _expressionresults.asp page
Set sp = fs.CreateTextFile(TheFullPath & ProjectName & "_expressionresults.asp", True)

sp.WriteLine(Chr(60) + Chr(37) + "@ LANGUAGE=VBScript" + chr(37) + ">")
sp.WriteLine("")
sp.WriteLine(Chr(60) + Chr(37))
sp.WriteLine("'---- CursorTypeEnum Values ----")
sp.WriteLine("Const adOpenForwardOnly = 0")
sp.WriteLine("Const adOpenKeyset = 1")
sp.WriteLine("Const adOpenDynamic = 2")
sp.WriteLine("Const adOpenStatic = 3")
sp.WriteLine("'---- LockTypeEnum Values ----")
sp.WriteLine("Const adLockReadOnly = 1")
sp.WriteLine("Const adLockPessimistic = 2")
sp.WriteLine("Const adLockOptimistic = 3")
sp.WriteLine("Const adLockBatchOptimistic = 4")
sp.WriteLine("'---- CommandTypeEnum Values ----")
sp.WriteLine("Const adCmdUnknown = &H0008")
sp.WriteLine("Const adCmdText = &H0001")
sp.WriteLine("Const adCmdTable = &H0002")
sp.WriteLine("Const adCmdStoredProc = &H0004")
sp.WriteLine("Const adCmdFile = &H0100")
sp.WriteLine("Const adCmdTableDirect = &H0200")
sp.WriteLine("Server.ScriptTimeout = 600")
sp.WriteLine("")
sp.WriteLine("Dim TheSource(" & NP & "), TheESource(" & NP & "), TheType(" & NP & "), Filter1(" & NP & "), Filter2(" & NP & "), Value1(" & NP & "), Value2(" & NP & "), AndOr(" & NP & ")")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("' Read in Variable for Calc")
sp.WriteLine("TheExp = Request(""frmExp"")")
sp.WriteLine("TheCalc = Request(""frmCalc"")")
sp.WriteLine("")
sp.WriteLine("")

sp.WriteLine("' Read in Variable from Group By")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("TheCount=0")
sp.WriteLine("For I = 1 to " & objRS2("MaxFields"))
sp.WriteLine("")
sp.WriteLine("  TestSource = Request(""frmSource"")(I)")
sp.WriteLine("  If TestSource = """" Then")
sp.WriteLine("          Exit For")
sp.WriteLine("  Else")
sp.WriteLine("          TheSource(I) = TestSource")
sp.WriteLine("          TheCount = I")
sp.WriteLine("  End If")
sp.WriteLine("'Response.Write(I)")
sp.WriteLine("Next")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("' Read in variables from Filter")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("TheECount=0")
sp.WriteLine("For I = 1 to " & objRS2("MaxFields"))
sp.WriteLine("")
sp.WriteLine("  TestESource = Request(""frmESource"")(I)")
sp.WriteLine("  If TestESource = """" Then")
sp.WriteLine("          Exit For")
sp.WriteLine("  Else")
sp.WriteLine("          SplitESource=Split(TestESource, ""*"")")
sp.WriteLine("          TheESource(I) = SplitESource(0)")
sp.WriteLine("          TheType(I) = UCASE(SplitESource(1))")
sp.WriteLine("          Filter1(I) = Request(""frmFilter1"")(I)")
sp.WriteLine("          Filter2(I) = Request(""frmFilter2"")(I)")
sp.WriteLine("          Value1(I) = Request(""frmValue1"")(I)")
sp.WriteLine("          Value2(I) = Request(""frmValue2"")(I)")
sp.WriteLine("          AndOr(I) = Request(""frmAndOr"")(I)")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("          TheECount = I")
sp.WriteLine("  End If")
sp.WriteLine("'Response.Write(I)")
sp.WriteLine("Next")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("'----------------------------------- All variables are read.....")
sp.WriteLine("")
sp.WriteLine("' start building the sql, which will be SQLSelectStart, SQLSelectMid, SQLSelectEnd, SQLWhere, SQLGroupby and SQLOrderby")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("SQLSelectStart = ""Select """)
sp.WriteLine("SQLSelectMid = """"")
sp.WriteLine("SQLSelectEnd = """"")
sp.WriteLine("SQLWhere = """"")
sp.WriteLine("SQLGroupBy = """"")
sp.WriteLine("SQLOrderBy = """"")
sp.WriteLine("")
sp.WriteLine("' Build sql end with the expression variables ")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("  AsField= TheCalc  & "" of " & chr(34) & " & TextTitle(TheExp)")
sp.WriteLine("")
sp.WriteLine("If TheCalc <> ""AVG"" Then")
sp.WriteLine("  sqlSelectEnd  = theCalc & ""("" & TheExp & "") As ["" & AsField & ""] from " & objRS2("TheTable") & " " & chr(34))
sp.WriteLine("  NumDec = 0")
sp.WriteLine("ELSE")
sp.WriteLine("  sqlSelectEnd  = theCalc & ""(Cast("" & theExp & "" as Float)) As ["" & AsField & ""] from " & objRS2("TheTable") & " "  & chr(34))
sp.WriteLine("  NumDec = 4")
sp.WriteLine("End If")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("' --------------------------------  test to see if there is a group by")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("If TheSource(1) <> """" Then ' There is a group by ")
sp.WriteLine("'REsponse.Write(""TheCount is: "" & TheCount & ""<br>"")")
' ------------------------------------ Revised Dec. 20, 2005 ------------------------------------------

sp.WriteLine("  For I = 1 to TheCount")
sp.WriteLine("                  ")
sp.WriteLine("                 If (InStr(TheSource(I), ""*Y"") = FALSE) and (InStr(TheSource(I), ""*M"") = FALSE) Then")
sp.WriteLine("                          SQLSelectMid = SQLSelectMid & TheSource(I) & "", """)
sp.WriteLine("                          SQLGroupBy = SQLGroupBy & TheSource(I) & "", """)
sp.WriteLine("                          SQLOrderBy = SQLOrderBy & TheSource(I) & "", """)
sp.WriteLine("'                         Response.Write(""The mid is: "" & sqlselectmid & ""<br>"")")
sp.WriteLine("")
sp.WriteLine("                 ElseIf InStr(TheSource(I), ""*Y"") Then ")
sp.WriteLine("                          SourceYear = """"")
sp.WriteLine("                          SourceSplit = Split(TheSource(I), ""*Y"")")
sp.WriteLine("                          SourceYear = SourceSplit(0)")
sp.WriteLine("                          SQLSelectMid = SQLSelectMid & "" DatePart(yyyy, "" & SourceYear & "") As "" & SourceYear & "", """)
sp.WriteLine("                          TheSource(I) = Replace(TheSource(I), ""*Y"", """")")
sp.WriteLine("                          SQLGroupBy = SQLGroupBy & "" DatePart(yyyy, "" & SourceYear & ""), """)
sp.WriteLine("                          SQLOrderBy = SQLOrderBy & "" DatePart(yyyy, "" & SourceYear & ""), """)
sp.WriteLine("                  Else")
sp.WriteLine("                          SourceYear = """"")
sp.WriteLine("                          SourceSplit = Split(TheSource(I), ""*M"")")
sp.WriteLine("                          SourceYear = SourceSplit(0)")
sp.WriteLine("                          SQLSelectMid = SQLSelectMid & "" Cast(DatePart(yyyy,  "" & SourceYear & "") as varchar(4))+'-'+right(Cast(datepart(mm, "" & SourceYear & "") + 1000 as varchar(4)), 2) As "" & SourceYear & "", """)
sp.WriteLine("                          TheSource(I) = Replace(TheSource(I), ""*M"", """")")
sp.WriteLine("                          SQLGroupBy = SQLGroupBy & "" Cast(DatePart(yyyy,  "" & SourceYear & "") as varchar(4))+'-'+right(Cast(datepart(mm, "" & SourceYear & "") + 1000 as varchar(4)), 2), """)
sp.WriteLine("                          SQLOrderBy = SQLOrderBy & "" Cast(DatePart(yyyy,  "" & SourceYear & "") as varchar(4))+'-'+right(Cast(datepart(mm, "" & SourceYear & "") + 1000 as varchar(4)), 2), """)
sp.WriteLine("                  End If")
sp.WriteLine("  Next")
' ------------------------------------ Revised Dec. 20, 2005 ------------------------------------------

sp.WriteLine("")
sp.WriteLine("SQLGroupBy = Left(SQLGroupBy, Len(SQLGroupBy) - 2)")
sp.WriteLine("SQLOrderBy = Left(SQLOrderBy, Len(SQLOrderBy) - 2)")
sp.WriteLine("  SQLGroupBy = "" Group by "" & SQLGroupBy")
'new
'sp.WriteLine(" SQLOrderBy = "" Order by "" & SQLOrderBy")

sp.WriteLine(" ShowOrder = Request(""frmShowOrder"")")
sp.WriteLine(" If ShowOrder = ""default"" Then")
sp.WriteLine("  SQLOrderBy = "" Order by "" & SQLOrderBy")
sp.WriteLine(" Elseif ShowOrder = ""desc"" Then")
sp.WriteLine("       SQLOrderBy = "" Order by "" & TheCalc & ""("" & TheExp & "") desc """)
sp.WriteLine(" Else")
sp.WriteLine("       SQLOrderBy = "" Order by "" & TheCalc & ""("" & TheExp & "") """)
sp.WriteLine("End If")



'end new




sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("End If")
sp.WriteLine("")
sp.WriteLine("'------------------------------------ See if there's anything to filter")
sp.WriteLine("")
sp.WriteLine("If TheESource(1) <> """" Then")
sp.WriteLine("")
sp.WriteLine("' make the sql where clause")
sp.WriteLine("")
sp.WriteLine("sqlWhere = ""WHERE """)
sp.WriteLine("For I = 1 to TheECount")
sp.WriteLine("WhereHolder1 = """"")
sp.WriteLine("WhereHolder2 = """"")
sp.WriteLine("")
sp.WriteLine("'Response.Write(""The ESource is: "" & TheESource(I) & ""<br>"")")
sp.WriteLine("")
sp.WriteLine("' see if there's an and/or clause")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("  If AndOr(I) = """" Then")
sp.WriteLine("          Select Case TheType(I)")
sp.WriteLine("                  Case ""N""")
sp.WriteLine("                          WhereHolder1 = MakeWhereNum(TheESource(I), Filter1(I), Value1(I))")
sp.WriteLine("                  Case ""T""")
sp.WriteLine("                          WhereHolder1 = MakeWhereText(TheESource(I), Filter1(I), Value1(I))")
sp.WriteLine("                  Case ""D""")
sp.WriteLine("                          WhereHolder1 = MakeWhereDate(TheESource(I), Filter1(I), Value1(I))")
sp.WriteLine("          End Select")
sp.WriteLine("")
sp.WriteLine("          If I <> TheECount Then")
sp.WriteLine("                  sqlWhere = sqlWhere & "" ("" & WhereHolder1 & "") AND """)
sp.WriteLine("          Else")
sp.WriteLine("                  sqlWhere = sqlWhere & "" ("" & WhereHolder1 & "") """)
sp.WriteLine("          End If")
sp.WriteLine("                  ")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("  Else ' there's an and or or     ")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("          Select Case TheType(I)")
sp.WriteLine("                  Case ""N""")
sp.WriteLine("                          WhereHolder1 = MakeWhereNum(TheESource(I), Filter1(I), Value1(I))")
sp.WriteLine("                          WhereHolder2 = MakeWhereNum(TheESource(I), Filter2(I), Value2(I))")
sp.WriteLine("                  Case ""T""")
sp.WriteLine("                          WhereHolder1 = MakeWhereText(TheESource(I), Filter1(I), Value1(I))")
sp.WriteLine("                          WhereHolder2 = MakeWhereText(TheESource(I), Filter2(I), Value2(I))")
sp.WriteLine("                  Case ""D""")
sp.WriteLine("                          WhereHolder1 = MakeWhereDate(TheESource(I), Filter1(I), Value1(I))")
sp.WriteLine("                          WhereHolder2 = MakeWhereDate(TheESource(I), Filter2(I), Value2(I))")
sp.WriteLine("          End Select")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("          If I <> TheECount Then")
sp.WriteLine("                  sqlWhere = sqlWhere & "" (("" & WhereHolder1 & "") "" & AndOr(I) & "" ("" & WhereHolder2 & "")) AND """)
sp.WriteLine("          Else")
sp.WriteLine("                  sqlWhere = sqlWhere & "" (("" & WhereHolder1 & "") "" & AndOr(I) & "" ("" & WhereHolder2& "")) """)
sp.WriteLine("          End If")
sp.WriteLine("                  ")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("  End If")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("Next")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("End If ' ---------------------------- end of where clause")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("sql = SQLSelectStart & "" "" & SQLSelectMid & "" "" & SQLSelectEnd & "" "" & SQLWhere & "" """)
sp.WriteLine("sql = sql & SQLGroupBy & "" "" & SQLOrderBy")
sp.WriteLine("")
sp.WriteLine("")
If objRS2("DisplaySQL") = "YES" Then
        sp.WriteLine("Response.Write(""The SQL is: "" & sql & ""<br>"")")
Else
        sp.WriteLine("'Response.Write(""The SQL is: "" & sql & ""<br>"")")

End If

sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("'------------------------------------------------------")
sp.WriteLine("")

sp.WriteLine("")
sp.WriteLine("dbName = " & chr(34)  & objRS2("TheDB") & chr(34) & " "  & chr(37) + ">")
sp.WriteLine("")
sp.WriteLine("")

sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("<!" & "--#include virtual = ""/thetimes/longrcon.inc"" -->")
sp.WriteLine("")
sp.WriteLine(Chr(60) + Chr(37))


sp.WriteLine("")
'new
sp.WriteLine("MaxRecs = """" & Request(""frmMaxRet"")")
sp.WriteLine("If MaxRecs <> """"  Then")
sp.WriteLine("  If MaxRecs > 5000 Then MaxRecs = 5000")
sp.WriteLine("  RS.MaxRecords = MaxRecs")
sp.WriteLine("Else")
sp.WriteLine("  RS.MaxRecords = 5000")
sp.WriteLine("End If")
'end new


sp.WriteLine("")
sp.WriteLine("RS.Open sql, Conn, adOpenStatic, adLockReadOnly, adCmdText")
sp.WriteLine("")
sp.WriteLine("")



sp.WriteLine(chr(37) & ">")
sp.WriteLine("<HTML>")
sp.WriteLine("<head>")
sp.WriteLine("<title>Calculation Results</title>")
sp.WriteLine("<LINK REL=""stylesheet"" TYPE=""text/css"" HREF=""/thetimes/shboom.css"">")
sp.WriteLine("</head>")
sp.WriteLine("<body  bgcolor=""white"">")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("<!" & "--#include file = """ & ProjectName & "_links.inc"" -->")


sp.WriteLine("<H1 class=pageTitle>" & objRS2("ProjectAlias") & " Calculation Results</H1>")

sp.WriteLine("<table cellpadding=0 cellspacing=0  >")
sp.WriteLine("  <tr>")
sp.WriteLine("")
sp.WriteLine("<" & chr(37))
sp.WriteLine("")
sp.WriteLine("For I = 1 to TheCount ")
sp.WriteLine("TheFieldName = TextTitle(TheSource(I))")
sp.WriteLine("")
sp.WriteLine(Chr(37) & ">")
sp.WriteLine("")
sp.WriteLine("<td class=resultsHeading><" & Chr(37) & " = TheFieldName " & chr(37) & "></td>")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("<" & chr(37) & " Next " & Chr(37) & ">")
sp.WriteLine("")

' ---------------------------------------- new --------------------------------

sp.WriteLine("    <td class=resultsHeading><" & Chr(37) & " = AsField " & chr(37) & "></td>")
sp.WriteLine("<" & chr(37) & " If UCase(TheCalc) = ""MIN"" or UCASE(TheCalc) = ""MAX"" Then " & chr(37) & ">")
sp.WriteLine("    <td class=resultsHeading><" & chr(37) & " = TheCalc " & Chr(37) & "> Value Link</td>")
sp.WriteLine("<" & chr(37) & " End If " & chr(37) & ">")
sp.WriteLine("   <td class=resultsHeading>Show Records</td>")

' ------------------------------------------ end new --------------------------------------------------------------

sp.WriteLine("  </tr>")
sp.WriteLine("")
sp.WriteLine("<" & chr(37))
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine(" Do While Not RS.EOF " & chr(37) & ">")
sp.WriteLine("  <tr>")
sp.WriteLine("")
sp.WriteLine("<" & chr(37))
sp.WriteLine("")
sp.WriteLine("SumHref=" & chr(34) & chr(34))
sp.WriteLine("")
sp.WriteLine("For I = 1 to TheCount")
sp.WriteLine("")
sp.WriteLine("ColHRef=" & chr(34) & chr(34))
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("If Trim(rs(TheSource(I))) = """" Then ")
sp.WriteLine("")
sp.WriteLine("  colHref=""frm"" & TheSource(I) & ""=NULL""")
sp.WriteLine("")
sp.WriteLine("  If I = 1 Then")
sp.WriteLine("          SumHref = SumHref & ColHref")
sp.WriteLine("  Else")
sp.WriteLine("          SumHref = SumHref & ""&"" & ColHref")
sp.WriteLine("  End If")
sp.WriteLine("")
sp.WriteLine(chr(37) & ">")
sp.WriteLine("    <td class=resultsData><a href=" & chr(34) & ProjectName & "_drilldown.asp?<" & chr(37) & " = ColHref & ""&frmsqlWhere="" & Server.URLEncode(sqlWhere)"  & chr(37) & ">"" class=resultsDrillDown>Blank</a></td>")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("<" & chr(37) & " ElseIf  IsNull(rs(TheSource(I))) Then ")

'-------------------------------
sp.WriteLine("")
sp.WriteLine("  colHref=""frm"" & TheSource(I) & ""=NULL""")
sp.WriteLine("")
sp.WriteLine("  If I = 1 Then")
sp.WriteLine("          SumHref = SumHref & ColHref")
sp.WriteLine("  Else")
sp.WriteLine("          SumHref = SumHref & ""&"" & ColHref")
sp.WriteLine("  End If")
sp.WriteLine("")
sp.WriteLine(chr(37) & ">")
sp.WriteLine("    <td class=resultsData><a href=" & chr(34) & ProjectName & "_drilldown.asp?<" & chr(37) & " = ColHref & ""&frmsqlWhere="" & Server.URLEncode(sqlWhere)"  & chr(37) & ">"" class=resultsDrillDown>Null</a></td>")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("<" & chr(37) & " Else ")
sp.WriteLine("  colHref=""frm"" & TheSource(I) & ""="" & Server.URLEncode(rs(TheSource(I)))")
sp.WriteLine("")
sp.WriteLine("  If I = 1 Then")
sp.WriteLine("          SumHref = SumHref & ColHref")
sp.WriteLine("  Else")
sp.WriteLine("          SumHref = SumHref & ""&"" & ColHref")
sp.WriteLine("  End If")
sp.WriteLine("")
sp.WriteLine(chr(37) & ">")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("    <td class=resultsData><a href=" & chr(34) & ProjectName & "_drilldown.asp?<" & chr(37) & " = ColHref & ""&frmsqlWhere="" & Server.URLEncode(sqlWhere)"  & chr(37) & ">"" class=resultsDrillDown><" & chr(37) & " = RS(TheSource(I)) " & chr(37) & "></td>")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("<" & chr(37) & " End If " & chr(37) & ">")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("<" & chr(37))
sp.WriteLine("Next")

'----------------------------------- new ------------------------------------------------

sp.WriteLine(" If UCase(TheCalc) = ""MIN"" or UCASE(TheCalc) = ""MAX"" Then ")

sp.WriteLine("")
sp.WriteLine("  If Not IsNull(rs(AsField)) Then")
sp.WriteLine("          colHref=""frm"" & TheExp & ""="" & Server.URLEncode(rs(AsField))")
sp.WriteLine("  Else")
sp.WriteLine("          colHref=""frm"" & TheExp & ""=NULL""")
sp.WriteLine("  End If          ")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("  If I = 1 Then")
sp.WriteLine("          SumHref = SumHref & ColHref")
sp.WriteLine("  Else")
sp.WriteLine("          SumHref = SumHref & ""&"" & ColHref")
sp.WriteLine("  End If")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")


sp.WriteLine("If rs(AsField).Type = 135 Then ")

'------------------ stopped here ------------------------------------

sp.WriteLine(chr(37) & ">")
sp.WriteLine("          <td class=resultsDataRight><" & chr(37) & " = rs(AsField) " & chr(37) & "></td>")
sp.WriteLine("          <" & chr(37) & " On Error Resume Next " & chr(37) & ">")
sp.WriteLine("          <td class=resultsDataRight><a href=" & chr(34) & ProjectName & "_drilldown.asp?<" & chr(37) & " = ColHref & ""&frmsqlWhere="" & Server.URLEncode(sqlWhere)"  & chr(37) & ">"" class=resultsDrillDown><" & chr(37) & " = TheCalc " & chr(37) & "> Date</a></td>")
sp.WriteLine("          <td class=resultsDataRight><a href=" & chr(34) & ProjectName & "_drilldown.asp?<" & chr(37) & " = SumHref & ""&frmsqlWhere="" & Server.URLEncode(sqlWhere)"  & chr(37) & ">"" class=resultsDrillDown>All in Row</a></td>")
sp.WriteLine("")
sp.WriteLine("          <" & chr(37) & " Else ")
sp.WriteLine("          On error resume next " & chr(37) & ">")
sp.WriteLine("          <td class=resultsDataRight><" & chr(37) & " = FormatNumber(rs(AsField),2) " & chr(37) & "></td>")
sp.WriteLine("          <% On Error Resume Next " & chr(37) & ">")
sp.WriteLine("          <td class=resultsDataRight><a href=" & chr(34) & ProjectName & "_drilldown.asp?<" & chr(37) & " = ColHref & ""&frmsqlWhere="" & Server.URLEncode(sqlWhere)"  & chr(37) & ">"" class=resultsDrillDown><" & chr(37) & " = TheCalc " & chr(37) & "> Number</a></td>")
sp.WriteLine("          <td class=resultsDataRight><a href=" & chr(34) & ProjectName & "_drilldown.asp?<" & chr(37) & " = SumHref & ""&frmsqlWhere="" & Server.URLEncode(sqlWhere)"  & chr(37) & ">"" class=resultsDrillDown>All in Row</a></td>")
sp.WriteLine("")
sp.WriteLine("  <" & chr(37) & " End If ")
sp.WriteLine("")
sp.WriteLine("Else ")
sp.WriteLine("")
sp.WriteLine("  If rs(AsField).Type = 135 Then ")
sp.WriteLine("")
sp.WriteLine(chr(37) & ">")
sp.WriteLine("          <td class=resultsDataRight><" & chr(37) & " = rs(AsField) " & chr(37) & "></td>")
sp.WriteLine("")
sp.WriteLine("  <" & chr(37) & " Else ")
sp.WriteLine("  On error Resume Next " & chr(37) & ">")
sp.WriteLine("          <td class=resultsDataRight><" & chr(37) & " = FormatNumber(rs(AsField), 2) " & chr(37) & "></td>")
sp.WriteLine("")
sp.WriteLine("  ")
sp.WriteLine("  <" & chr(37) & " End If " & chr(37) & ">")
sp.WriteLine("  <td class=resultsDataRight><a href=" & chr(34) & ProjectName & "_drilldown.asp?<" & chr(37) & " = SumHref & ""&frmsqlWhere="" & Server.URLEncode(sqlWhere)"  & chr(37) &  ">"" class=resultsDrillDown>All in Row</a></td>")
sp.WriteLine("")
sp.WriteLine("  ")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("<" & chr(37) & " End If " & chr(37) & ">")
sp.WriteLine("")
sp.WriteLine("</tr>")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("<" & chr(37) & " rs.MoveNext")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("Loop")
sp.WriteLine(chr(37) & ">")
sp.WriteLine("</TABLE>")
sp.WriteLine("</center>")
sp.WriteLine("</p>")
sp.WriteLine("")
sp.WriteLine("<" & chr(37))
sp.WriteLine("rs.Close")
sp.WriteLine("conn.Close")
sp.WriteLine("set rs = nothing")
sp.WriteLine("set Conn = nothing")
sp.WriteLine(chr(37) & ">")
sp.WriteLine("</BODY>")
sp.WriteLine("</HTML>")
sp.WriteLine("")

sp.WriteLine("<" & chr(37))


sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("Function MakeWhereNum(TheFieldName, TheOperator, TheValue)")
sp.WriteLine("")
sp.WriteLine("Select Case TheOperator")
sp.WriteLine("  Case ""eq""")
sp.WriteLine("          MakeWhereNum = "" "" & TheFieldName & "" = "" & TheValue")
sp.WriteLine("  Case ""noteq""")
sp.WriteLine("          MakeWhereNum = "" "" & TheFieldName & "" <> "" & TheValue")
sp.WriteLine("  Case ""gt""")
sp.WriteLine("          MakeWhereNum = "" "" & TheFieldName & "" > "" & TheValue")
sp.WriteLine("  Case ""gtet""")
sp.WriteLine("          MakeWhereNum = "" "" & TheFieldName & "" >= "" & TheValue")
sp.WriteLine("  Case ""lt""")
sp.WriteLine("          MakeWhereNum = "" "" & TheFieldName & "" < "" & TheValue")
sp.WriteLine("  Case ""ltet""")
sp.WriteLine("          MakeWhereNum = "" "" & TheFieldName & "" <= "" & TheValue")
sp.WriteLine("  Case ""bw""")
sp.WriteLine("          MakeWhereNum = "" Cast("" & TheFieldName & "" as varchar(25)) Like '"" & TheValue & ""%'""")
sp.WriteLine("  Case ""dnbw""")
sp.WriteLine("          MakeWhereNum = "" Cast("" & TheFieldName & "" as varchar(25)) NOT Like '"" & TheValue & ""%'""")
sp.WriteLine("  Case ""ew""")
sp.WriteLine("          MakeWhereNum = "" Cast("" & TheFieldName & "" as varchar(25)) Like '%"" & TheValue & ""'""")
sp.WriteLine("  Case ""dnew""")
sp.WriteLine("          MakeWhereNum = "" Cast("" & TheFieldName & "" as varchar(25)) NOT Like '%"" & TheValue & ""'""")
sp.WriteLine("  Case ""con""")
sp.WriteLine("          MakeWhereNum = "" Cast("" & TheFieldName & "" as varchar(25)) Like '%"" & TheValue & ""%'""")
sp.WriteLine("  Case ""dncon""")
sp.WriteLine("          MakeWhereNum = "" Cast("" & TheFieldName & "" as varchar(25)) NOT Like '%"" & TheValue & ""%'""")
sp.WriteLine("  Case ""blank""")
sp.WriteLine("          MakeWhereNum = "" "" & TheFieldName & "" is null """)
sp.WriteLine("  Case ""notblank""")
sp.WriteLine("          MakeWhereNum = "" "" & TheFieldName & "" is NOT null  "" ")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("End Select")
sp.WriteLine("")
sp.WriteLine("End Function")
sp.WriteLine("")
sp.WriteLine("Function MakeWhereText(TheFieldName, TheOperator, TheValue)")
sp.WriteLine("")
sp.WriteLine("If InStr(TheValue, ""'"") Then")
sp.WriteLine("  TheValue = Replace(TheValue, ""'"", ""''"")")
sp.WriteLine("End If")
sp.WriteLine("Select Case TheOperator")
sp.WriteLine("  Case ""eq""")
sp.WriteLine("          MakeWhereText = "" "" & TheFieldName & "" = '"" & TheValue & ""'""")
sp.WriteLine("          If InStr(MakeWhereText, ""]"") and InStr(MakeWhereText, ""["") Then")
sp.WriteLine("                  MakeWhereText = Replace(MakeWhereText, ""'"", """")")
sp.WriteLine("          End If")
sp.WriteLine("  Case ""noteq""")
sp.WriteLine("          MakeWhereText = "" "" & TheFieldName & "" <> '"" & TheValue & ""'""")
sp.WriteLine("          If InStr(MakeWhereText, ""]"") and InStr(MakeWhereText, ""["") Then")
sp.WriteLine("                  MakeWhereText = Replace(MakeWhereText, ""'"", """")")
sp.WriteLine("          End If")
sp.WriteLine("  Case ""gt""")
sp.WriteLine("          MakeWhereText = "" "" & TheFieldName & "" > '"" & TheValue & ""'""")
sp.WriteLine("          If InStr(MakeWhereText, ""]"") and InStr(MakeWhereText, ""["") Then")
sp.WriteLine("                  MakeWhereText = Replace(MakeWhereText, ""'"", """")")
sp.WriteLine("          End If")
sp.WriteLine("  Case ""gtet""")
sp.WriteLine("          MakeWhereText = "" "" & TheFieldName & "" >=  '"" & TheValue & ""'""")
sp.WriteLine("          If InStr(MakeWhereText, ""]"") and InStr(MakeWhereText, ""["") Then")
sp.WriteLine("                  MakeWhereText = Replace(MakeWhereText, ""'"", """")")
sp.WriteLine("          End If")
sp.WriteLine("  Case ""lt""")
sp.WriteLine("          MakeWhereText = "" "" & TheFieldName & "" <  '"" & TheValue & ""'""")
sp.WriteLine("          If InStr(MakeWhereText, ""]"") and InStr(MakeWhereText, ""["") Then")
sp.WriteLine("                  MakeWhereText = Replace(MakeWhereText, ""'"", """")")
sp.WriteLine("          End If")
sp.WriteLine("  Case ""ltet""")
sp.WriteLine("          MakeWhereText = "" "" & TheFieldName & "" <= '"" & TheValue & ""'""")
sp.WriteLine("          If InStr(MakeWhereText, ""]"") and InStr(MakeWhereText, ""["") Then")
sp.WriteLine("                  MakeWhereText = Replace(MakeWhereText, ""'"", """")")
sp.WriteLine("          End If")
sp.WriteLine("  Case ""bw""")
sp.WriteLine("          MakeWhereText = "" "" & TheFieldName & "" Like '"" & TheValue & ""%'""")
sp.WriteLine("  Case ""dnbw""")
sp.WriteLine("          MakeWhereText = "" "" & TheFieldName & "" NOT Like '"" & TheValue & ""%'""")
sp.WriteLine("  Case ""ew""")
sp.WriteLine("          MakeWhereText = "" "" & TheFieldName & "" Like '%"" & TheValue & ""'""")
sp.WriteLine("  Case ""dnew""")
sp.WriteLine("          MakeWhereText = "" "" & TheFieldName & "" NOT Like '%"" & TheValue & ""'""")
sp.WriteLine("  Case ""con""")
sp.WriteLine("          MakeWhereText = "" "" & TheFieldName & "" Like '%"" & TheValue & ""%'""")
sp.WriteLine("  Case ""dncon""")
sp.WriteLine("          MakeWhereText = "" "" & TheFieldName & "" NOT Like '%"" & TheValue & ""%'""")
sp.WriteLine("  Case ""blank""")
sp.WriteLine("          MakeWhereText = "" ("" & TheFieldName & "" is Null or "" & TheFieldName & "" = '') """)
sp.WriteLine("  Case ""notblank""")
sp.WriteLine("          MakeWhereText = "" ("" & TheFieldName & "" is NOT Null and "" & TheFieldName & "" <> '') """)
sp.WriteLine("          ")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("End Select")
sp.WriteLine("")
sp.WriteLine("End Function")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("Function MakeWhereDate(TheFieldName, TheOperator, TheValue)")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("Select Case TheOperator")
sp.WriteLine("  Case ""eq""")
sp.WriteLine("          MakeWhereDate = "" "" & TheFieldName & "" = '"" & TheValue & ""'""")
sp.WriteLine("          If InStr(MakeWhereDate, ""]"") and InStr(MakeWhereDate, ""["") Then")
sp.WriteLine("                  MakeWhereDate = Replace(MakeWhereDate, ""'"", """")")
sp.WriteLine("          End If")
sp.WriteLine("  Case ""noteq""")
sp.WriteLine("          MakeWhereDate = "" "" & TheFieldName & "" <> '"" & TheValue & ""'""")
sp.WriteLine("          If InStr(MakeWhereDate, ""]"") and InStr(MakeWhereDate, ""["") Then")
sp.WriteLine("                  MakeWhereDate = Replace(MakeWhereDate, ""'"", """")")
sp.WriteLine("          End If")
sp.WriteLine("  Case ""gt""")
sp.WriteLine("          MakeWhereDate = "" "" & TheFieldName & "" > '"" & TheValue & ""'""")
sp.WriteLine("          If InStr(MakeWhereDate, ""]"") and InStr(MakeWhereDate, ""["") Then")
sp.WriteLine("                  MakeWhereDate = Replace(MakeWhereDate, ""'"", """")")
sp.WriteLine("          End If")
sp.WriteLine("  Case ""gtet""")
sp.WriteLine("          MakeWhereDate = "" "" & TheFieldName & "" >=  '"" & TheValue & ""'""")
sp.WriteLine("          If InStr(MakeWhereDate, ""]"") and InStr(MakeWhereDate, ""["") Then")
sp.WriteLine("                  MakeWhereDate = Replace(MakeWhereDate, ""'"", """")")
sp.WriteLine("          End If")
sp.WriteLine("  Case ""lt""")
sp.WriteLine("          MakeWhereDate = "" "" & TheFieldName & "" <  '"" & TheValue & ""'""")
sp.WriteLine("          If InStr(MakeWhereDate, ""]"") and InStr(MakeWhereDate, ""["") Then")
sp.WriteLine("                  MakeWhereDate = Replace(MakeWhereDate, ""'"", """")")
sp.WriteLine("          End If")
sp.WriteLine("  Case ""ltet""")
sp.WriteLine("          MakeWhereDate = "" "" & TheFieldName & "" <= '"" & TheValue & ""'""")
sp.WriteLine("          If InStr(MakeWhereDate, ""]"") and InStr(MakeWhereDate, ""["") Then")
sp.WriteLine("                  MakeWhereDate = Replace(MakeWhereDate, ""'"", """")")
sp.WriteLine("          End If")
sp.WriteLine("  Case ""bw""")
sp.WriteLine("          MakeWhereDate = "" Cast("" & TheFieldName & "" as varchar(25)) Like '"" & TheValue & ""%'""")
sp.WriteLine("  Case ""dnbw""")
sp.WriteLine("          MakeWhereDate = "" Cast("" & TheFieldName & "" as varchar(25)) NOT Like '"" & TheValue & ""%'""")
sp.WriteLine("  Case ""ew""")
sp.WriteLine("          MakeWhereDate = "" Cast("" & TheFieldName & "" as varchar(25)) Like '%"" & TheValue & ""'""")
sp.WriteLine("  Case ""dnew""")
sp.WriteLine("          MakeWhereDate = "" Cast("" & TheFieldName & "" as varchar(25)) NOT Like '%"" & TheValue & ""'""")
sp.WriteLine("  Case ""con""")
sp.WriteLine("          MakeWhereDate = "" Cast("" & TheFieldName & "" as varchar(25)) Like '%"" & TheValue & ""%'""")
sp.WriteLine("  Case ""dncon""")
sp.WriteLine("          MakeWhereDate = "" Cast("" & TheFieldName & "" as varchar(25)) NOT Like '%"" & TheValue & ""%'""")
sp.WriteLine("  Case ""blank""")
sp.WriteLine("          MakeWhereDate = "" "" & TheFieldName & "" is NULL "" ")
sp.WriteLine("  Case ""notblank""")
sp.WriteLine("          MakeWhereDate = "" "" & TheFieldName & "" is NOT NULL "" ")
sp.WriteLine("")
sp.WriteLine("End Select")
sp.WriteLine("")
sp.WriteLine("End Function")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")

sp.WriteLine("")
sp.WriteLine("Function TextTitle(gaga)")
sp.WriteLine("  Select Case gaga")

rs.Filter = ""
rs.MoveFirst

Do While NOT rs.EOF

sp.WriteLine("          Case " & chr(34) & rs("TheField") & chr(34))
sp.WriteLine("                  TextTitle=" & chr(34) & rs("TheAlias") & chr(34))

rs.MoveNext
Loop

sp.WriteLine("End Select")

sp.WriteLine("End Function")


sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine(chr(37) & ">")








sp.Close





End If ' end of branch for creating an expressions page

' -------------- end expressions.asp page 







' -----------------------------



' write the _profile.asp page
Set sp = fs.CreateTextFile(TheFullPath & ProjectName & "_profile.asp", True)

sp.WriteLine(Chr(60) + Chr(37) + "@ LANGUAGE=VBScript" + chr(37) + ">")
sp.WriteLine("")
sp.WriteLine(Chr(60) + Chr(37))
sp.WriteLine("'---- CursorTypeEnum Values ----")
sp.WriteLine("Const adOpenForwardOnly = 0")
sp.WriteLine("Const adOpenKeyset = 1")
sp.WriteLine("Const adOpenDynamic = 2")
sp.WriteLine("Const adOpenStatic = 3")
sp.WriteLine("'---- LockTypeEnum Values ----")
sp.WriteLine("Const adLockReadOnly = 1")
sp.WriteLine("Const adLockPessimistic = 2")
sp.WriteLine("Const adLockOptimistic = 3")
sp.WriteLine("Const adLockBatchOptimistic = 4")
sp.WriteLine("'---- CommandTypeEnum Values ----")
sp.WriteLine("Const adCmdUnknown = &H0008")
sp.WriteLine("Const adCmdText = &H0001")
sp.WriteLine("Const adCmdTable = &H0002")
sp.WriteLine("Const adCmdStoredProc = &H0004")
sp.WriteLine("Const adCmdFile = &H0100")
sp.WriteLine("Const adCmdTableDirect = &H0200")



' get the datatype of the linking field

LField = objRS2("LinkingField")
sp.WriteLine(Lfield & " = Request(" & chr(34) & LField & chr(34) & ")")


rs.Filter=""
rs.Filter=("TheField = '" & LField & "'")
If NOT rs.EOF Then
        rs.MoveFirst
        LinkType = rs("TheType")

Else
        LinkType = ""
End If

Select Case LinkType
        Case "bigint", "bit", "decimal", "float", "int", "money", "numeric", "real", "smallint", "smallmoney", "tinyint"
                sp.WriteLine("sql = ""Select * from " & objRS2("TheTable") & " where " & LField & " = " & chr(34) & " & " & LField)
        Case Else
                sp.WriteLine("sql = ""Select * from " & objRS2("TheTable") & " where " & LField & " = '" & chr(34)  & " & " & LField & " & " & chr(34) &  "'" & chr(34))
End Select


rs.Filter=""
rs.Filter=("ShowProfile=TRUE")
rs.Sort=("TheOrder")
rs.MoveFirst


sp.WriteLine("")
sp.WriteLine("dbName = " & chr(34)  & objRS2("TheDB") & chr(34) & " "  & chr(37) + ">")
sp.WriteLine("")
sp.WriteLine("")

sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("<!" & "--#include virtual = ""/thetimes/rcon.inc"" -->")
sp.WriteLine("")
sp.WriteLine(Chr(60) + Chr(37))

sp.WriteLine("RS.Open sql, Conn, adOpenStatic, adLockReadOnly, adCmdText")
                
sp.WriteLine(chr(37) & ">")
sp.WriteLine("<HTML>")
sp.WriteLine("<HEAD>")
sp.WriteLine("<TITLE>Record Details</TITLE>")
sp.WriteLine("<LINK REL=""stylesheet"" TYPE=""text/css"" HREF=""/thetimes/shboom.css"">")
sp.WriteLine("</HEAD>")
sp.WriteLine("<body bgcolor=""white"">")
sp.WriteLine("<!" & "--#include file = """ & ProjectName & "_links.inc"" -->")
sp.WriteLine("<!" & "--#include file = """ & ProjectName & "_header.inc"" -->")

sp.WriteLine("<H1 class=pageTitle>Record Details</H1>")
sp.WriteLine "<table class=detailsTable cellpadding=0 cellspacing=0>"
sp.WriteLine("<" & chr(37))


Do While Not RS.EOF

TheLine = "str" & rs("TheField") & "= rs(" & chr(34) & rs("TheField") & chr(34) & ")"
sp.WriteLine(TheLine)
TheLine = "If str" & rs("TheField") & " <> " & chr(34) & Chr(34) & " Then " & chr(37) & ">"
sp.WriteLine(TheLine)
TheLine = "<tr><td class=detailsHeading>" & rs("TheAlias") & ": </td><td class=detailsData>" & Chr(60) & Chr(37) & " = str" & rs("TheField")  & " " & chr(37) + "></td></tr>"
sp.WriteLine(TheLine)
TheLine = Chr(60) & Chr(37) & " End If"
sp.WriteLine(TheLine)

rs.MoveNext
Loop


sp.WriteLine(chr(37) & ">")
sp.WriteLine("</table>")
sp.WriteLine("</body>")

sp.WriteLine("</html>")



sp.close

'---------------- write the tips page
' write the _profile.asp page
Set sp = fs.CreateTextFile(TheFullPath & ProjectName & "_tips.asp", True)

sp.WriteLine("<html>")
sp.WriteLine("<head>")
sp.WriteLine("<title>Caution and Tips</title>")
sp.WriteLine("<LINK REL=""stylesheet"" TYPE=""text/css"" HREF=""/thetimes/shboom.css"">")
sp.WriteLine("</head>")
sp.WriteLine("<body>")
sp.WriteLine("<!" & "--#include file = """ & ProjectName & "_links.inc"" -->")


sp.WriteLine("<h1 class=pageTitle>Cautions and Tips</h1>")




' -------------------- Added Dec 21
sp.WriteLine("<p>If you've never seen pages like these before, click <a href=""http://technews.nytimes.com/help"" target=""_blank"">here</a> for general help on how to use their many unique features.</p>")
' --------------------- Added Dec 21



If CountNumerics > 0 Then




sp.WriteLine("<h2 align=""center"">Cautions</h2>")
sp.WriteLine("<p align=""left"">Before reporting calculations derived from the page titled")
sp.WriteLine(chr(34) & objRS2("calctitle") & chr(34) & " you should first, look at the distribution of values from the numeric or")
sp.WriteLine("date field on the " & chr(34) & objRS2("groupbytitle") & chr(34) & " page. For example, if you wanted to calculate the")
sp.WriteLine("average age, you should first go to the &quot;" & objRS2("groupbytitle") & "&quot; page and select")
sp.WriteLine("the age field in the first pulldown menu. The results will show you the")
sp.WriteLine("distribution of ages throughout the database and this information can help you")
sp.WriteLine("avoid making meaningless calculations from the &quot;" & objRS2("CalcTitle")  & "&quot; page.</p>")
sp.WriteLine("<p align=""left"">While examining the distribution of ages, you might notice many")
sp.WriteLine("zeros or many entries for 99. Both are common entries when an age is unknown and")
sp.WriteLine("both can skew an average in a meaningless direction. Also, some agencies use")
sp.WriteLine("dates as codes. For example, a bureau of prisons might change a prisoner's")
sp.WriteLine("birthdate from 5/27/1948 to to 1/1/11 to indicate that a prisoner has been")
sp.WriteLine("released. You might erroneously report that the person in custody in a")
TheAge = abs(DateDiff("yyyy", Now, "1/1/1911"))
sp.WriteLine("state&nbsp; is " & TheAge &  ". If you first look at the distribution of dates on")
sp.WriteLine("the &quot;" & objRS2("GroupbyTitle") & "&quot; you might notice an unusually high number of")
sp.WriteLine("prisoners with the birthdate of 1/1/11 and avoid the mistake.</p>")
sp.WriteLine("<p align=""left"">Once you notice anomalies in the distribution, you can use the")
sp.WriteLine("area under the heading &quot;Select records that meet the following")
sp.WriteLine("conditions&quot; to filter out the bad entries. For example you might calculate")
sp.WriteLine("the average age where the age field is not equal to zero and is not equal to 99.")
sp.WriteLine("Or to find someone who is truly the oldest prisoner, you can calculate the")
sp.WriteLine("minimum date after filtering out dates that equal 1/1/11. As an extra")
sp.WriteLine("precaution, if you don't find any prisoners older than " & TheAge & ", you should")
sp.WriteLine("check with the bureau of prisons to ensure that no one was indeed born on")
sp.WriteLine("1/1/11.</p>")



RS.Filter="TheType like *date* AND ShowCalc = TRUE"

If Not RS.EOF Then
        rs.MoveFirst
        CountDates = rs.RecordCount
Else
        CountDates = 0
End If



If CountDates > 0 Then

        DateExample = RS("TheAlias")

sp.WriteLine("<blockquote>")

sp.WriteLine("")
sp.WriteLine("<h2 align=""center"">Date Searches</h2>")
sp.WriteLine("<p align=""left"">Finding items by dates or parts of dates can be difficult")
sp.WriteLine("because dates&nbsp; frequently are stored as dates and times. </p>")
sp.WriteLine("<p align=""left"">For example, the contents of the " & chr(34) & DateExample & chr(34) & " field may be stored as <i>4/22/2003")
sp.WriteLine("5:24:38 PM</i>. If you were to search for<i> 4/22/2003</i>, you would not find <i>4/22/2003")
sp.WriteLine("5:24:38 PM </i>or any other records that also recorded a time for that date.</p>")
sp.WriteLine("<p align=""left"">To see if dates also contain times, click on the " & chr(34) & objRS2("GroupByTitle") & chr(34))
sp.WriteLine("link and, in the pulldown under &quot;Select one or more fields to group")
sp.WriteLine("by&quot; select a date field, " & chr(34) & DateExample & chr(34) & " for example. If you're dealing with a")
sp.WriteLine("large table, you might want to limit the results set. In the area under")
sp.WriteLine("&quot;Select records that meet the following conditions,&quot; select " & chr(34) & DateExample & chr(34))
sp.WriteLine(" under &quot;Field;&quot; select &quot;contains&quot; under &quot;Show")
sp.WriteLine("Rows Where;&quot; and type 2003 under &quot;Value 1.&quot;</p>")
sp.WriteLine("<p align=""left"">If results show dates and times, you'll have to deal with date")
sp.WriteLine("searches accordingly. Generally, you'll use one of two methods: a <b>range</b>")
sp.WriteLine("search, which is faster, or a <b>cast </b>search which is easier to set up.</p>")
sp.WriteLine("<h3 align=""left"">Range Searches</h3>")
sp.WriteLine("<p align=""left"">To search for any records containing the date <i>4/22/2003 </i>in")
sp.WriteLine("the " & chr(34) & DateExample & chr(34) & " field, first click on the " & chr(34) & objRS2("FilterTitle") & chr(34) & " link. Under")
sp.WriteLine("&quot;Field,&quot; select " & chr(34) & DateExample & chr(34) & ". Under &quot;Show Rows Where,&quot;")
sp.WriteLine("select &quot;is greater than or equal to.&quot; Type in &quot;4/22/2003&quot;")
sp.WriteLine("(without the quotes) under &quot;Value 1.&quot;</p>")
sp.WriteLine("<p align=""left"">Next, select &quot;And&quot; from the first &quot;And / or&quot;")
sp.WriteLine("pulldown. Staying on the same row, select &quot;is less than&quot; from the")
sp.WriteLine("&quot;Show Rows Where&quot; pulldown to the right of the &quot;And / or&quot;")
sp.WriteLine("pulldown. Then, under &quot;Value 2&quot; type in &quot;4/3/2003&quot; (without")
sp.WriteLine("the quotes). Click &quot;Submit&quot; and the program will search for any")
sp.WriteLine("records that have <i>4/2/2003 </i>in the " & chr(34) & DateExample & chr(34) & " field.</p>")
sp.WriteLine("<h3 align=""left"">Cast Searches</h3>")
sp.WriteLine("<p align=""left"">When you select a date field, such as " & chr(34) & DateExample & chr(34) & ", from the pulldown menu under Fields on the filter page,")
sp.WriteLine("then select any of the following functions in the &quot;Show Rows Where&quot;")
sp.WriteLine("pulldown:</p>")
sp.WriteLine("<p align=""center"">begins with<br>")
sp.WriteLine("does not begin with<br>")
sp.WriteLine("ends with<br>")
sp.WriteLine("does not end with<br>")
sp.WriteLine("contains<br>")
sp.WriteLine("does not contain</p>")
sp.WriteLine("<p align=""left"">The program will <i>cast</i> the stored date as a string of")
sp.WriteLine("text. When this happens the program converts the date to a format that is not")
sp.WriteLine("intuitive. Unless you know the quirks of the conversion process you will have a")
sp.WriteLine("difficult time find dates using any of the above mentioned operators.</p>")
sp.WriteLine("<p align=""left"">Our example of <i>4/22/2003 5:24:38 PM </i> would be cast as <i>Apr")
sp.WriteLine("22 2003&nbsp; 5:24:38 PM. </i>So, if you wanted to find all records containing")
sp.WriteLine("the date 4/22/2003, you would select the &quot;begins with&quot; operator and")
sp.WriteLine("enter &quot;Apr 22 2003&quot; (without quotes) under Value 1. (You could also")
sp.WriteLine("select the &quot;contains&quot; operator, but &quot;begins with&quot; is <i>always</i>")
sp.WriteLine("faster than &quot;contains.&quot;) <b>Note: </b>There is no comma separating the")
sp.WriteLine("date and the year.</p>")
sp.WriteLine("<p align=""left"">Things get a little trickier when you're dealing with")
sp.WriteLine("single-digit dates, such as April 2. The conversion program inserts a space")
sp.WriteLine("before any single-digit date (or single-digit times). So if you were to search")
sp.WriteLine("for all dates that happened on April 2, regardless of the year, you would select")
sp.WriteLine("the &quot;begins with&quot; operator and enter &quot;Apr&nbsp; 2&quot; (that's <i>a-p-r-space-space-2,")
sp.WriteLine("</i>without the quotes) in the Value 1 box. If you enter <i>a-p-r-space-2</i>,")
sp.WriteLine("that is Apr 2 with a single space, you would find every date in the twenties")
sp.WriteLine("during April, but <i>not</i> April 2nd.</p>")
sp.WriteLine("<p align=""left"">If you're looking for records in a particular year, it")
sp.WriteLine("safer&nbsp; to use the &quot;contains&quot; operator, rather than the &quot;ends")
sp.WriteLine("with&quot; operator. If the date field also is storing times, the stored date")
sp.WriteLine("will not end with the year, but with the time. So, to look for dates in 2003,")
sp.WriteLine("use &quot;contains&quot; and enter 2003 in the Value box.</p>")
sp.WriteLine("</blockquote>")

End If ' end if for test of date fields
End If ' end of help entries for numeric fields

sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("<h2 align=""center"">Comparing one field to another on the page titled " & chr(34) &  objRS2("FilterTitle") &  chr(34) & ":</h2>")
sp.WriteLine("<blockquote>")
sp.WriteLine("  <ol>")
sp.WriteLine("    <li>On the "  &  chr(34) & objRS2("FilterTitle") & chr(34) &  " Page, select one of the fields from the pull-down menu")
sp.WriteLine("      under &quot;Fields&quot;.</li>")
sp.WriteLine("    <li>Select the comparison you want to make from the pull-down menu under")
sp.WriteLine("      &quot;Show Rows Where.&quot; (Note: only the first six selections (equals,")
sp.WriteLine("      does not equal, is greater than or equal to, is greater than, is less")
sp.WriteLine("      than, is less than or equal to) will work with field comparisons.</li>")
sp.WriteLine("    <li>On this page, highlight&nbsp; the &quot;Field Name&quot; (and <b>not</b>")
sp.WriteLine("      the &quot;Field Alias&quot;) you want to compare, copy it (Ctrl-V on the")
sp.WriteLine("      keyboard or Edit | Copy from the Internet Explorer menu bar) and paste it")
sp.WriteLine("      into the &quot;Value 1&quot; box on the " & chr(34) & objRS2("FilterTitle") & chr(34) & " Page. Be sure to include the brackets around the field name. Capitalization does not matter; the")
sp.WriteLine("      application is not case sensitive.</li>")
sp.WriteLine("  </ol>")
sp.WriteLine("  <p>&nbsp;</p>")
sp.WriteLine("</blockquote>")


sp.WriteLine("<table align=""center"" border=""1"">")
sp.WriteLine("<tr>")
sp.WriteLine("<td align=""center""><b>Field Name</b></td>")
sp.WriteLine("<td align=""center""><b>Field Alias</b></td>")
sp.WriteLine("<td align=""center""><b>Field Type</b></td>")
sp.WriteLine("</tr>")

rs.filter=""
rs.Sort=("TheOrder")
rs.MoveFirst

Do While NOT RS.EOF

sp.WriteLine("<tr>")
sp.WriteLine("<td>[" & RS("TheField") & "]</td>")
sp.WriteLine("<td>" & RS("TheAlias") & "</td>")
sp.WriteLine("<td>" & RS("TheType") & "</td>")
sp.WriteLine("</tr>")

rs.MoveNext
Loop
sp.WriteLine("<table>")







sp.WriteLine("</body>")
sp.WriteLine("</html>")





sp.Close


%>

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>Pages Completed</title>
</HEAD>
<body bgColor="#FFCC99">

<!--Written by Tom Torok of The New York Times
tomtorok@nytimes.com-->


<h1 align="center">Pages Completed</h1>
<P align=center>Pages for <% = ProjectName %> were written to the following directory:</P>
<P align=center><STRONG><% = TheFullPath %></STRONG></P>
<% If InStr(TheFullPath,"wwwroot")>0 Then
        WhereStart = InStr(TheFullPath,"wwwroot")
        TheLink = Right(TheFullPath,Len(TheFullPath)-(WhereStart+7))
        HowLong = Len(TheLink)
        If Right(TheLink,1) = "/" Then
                TheLink = Left(TheLink,HowLong-1)
        End If
        TheLink = TheLink & "/" & ProjectName & "_filter.asp"
        TheLink = "../" & TheLink
        TheLink=Replace(TheLink,"\","/")
        TheLink=Replace(TheLink,"//","/")
%>

<P align=center>To check them out <A href="<% = TheLink %>">click here.</A></P>
<% End If %>
<blockquote>
<p>Remember, the application created two include files, <% = ProjectName %>_header.inc, which is blank, and
<% = ProjectName %>_footer.inc, which cater's to Tom Torok's egomania. These files can be used to store 
static information about the application, such as "Current as of <% = Now %>" in the header, and 
"For more information call Joe Schmo at 212-555-1212." in the footer. <b>They will not change if you modify the pages.</b></p>
</blockquote>

<center><i>Shboom was developed by Tom Torok of The New York Times</i></center>

</BODY>
</HTML>


<%

set fs = nothing
rs.close



%>





