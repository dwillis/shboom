




'sql = "select column_name from information_schema.columns where table_name = '" & TheTable & "' and numeric_precision > 5 order by ordinal_position"
'RS.Open sql, Conn, adOpenStatic, adLockReadOnly, adCmdText
'CountNumerics = RS.RecordCount
'Response.Write("There are " & CountNumerics & " numeric fields.<br>")






'------------------ write the _expressionresults.asp page
Set sp = fs.CreateTextFile(TheFullPath & "\" & ProjectName & "_expressionresults.asp", True)

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
sp.WriteLine("Dim TheSource(11), TheESource(11), TheType(11), Filter1(11), Filter2(11), Value1(11), Value2(11), AndOr(11)")
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
sp.WriteLine("For I = 1 to 10")
sp.WriteLine("")
sp.WriteLine("	TestSource = Request(""frmSource"")(I)")
sp.WriteLine("	If TestSource = """" Then")
sp.WriteLine("		Exit For")
sp.WriteLine("	Else")
sp.WriteLine("		TheSource(I) = TestSource")
sp.WriteLine("		TheCount = I")
sp.WriteLine("	End If")
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
sp.WriteLine("For I = 1 to 10")
sp.WriteLine("")
sp.WriteLine("	TestESource = Request(""frmESource"")(I)")
sp.WriteLine("	If TestESource = """" Then")
sp.WriteLine("		Exit For")
sp.WriteLine("	Else")
sp.WriteLine("		SplitESource=Split(TestESource, ""*"")")
sp.WriteLine("		TheESource(I) = SplitESource(0)")
sp.WriteLine("		TheType(I) = UCASE(SplitESource(1))")
sp.WriteLine("		Filter1(I) = Request(""frmFilter1"")(I)")
sp.WriteLine("		Filter2(I) = Request(""frmFilter2"")(I)")
sp.WriteLine("		Value1(I) = Request(""frmValue1"")(I)")
sp.WriteLine("		Value2(I) = Request(""frmValue2"")(I)")
sp.WriteLine("		AndOr(I) = Request(""frmAndOr"")(I)")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("		TheECount = I")
sp.WriteLine("	End If")
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
sp.WriteLine("	AsField= TheCalc  & "" of " & chr(34) & " & TheExp")



sp.WriteLine("")
sp.WriteLine("If TheCalc <> ""AVG"" Then")
sp.WriteLine("	sqlSelectEnd  = theCalc & ""("" & TheExp & "") As ["" & AsField & ""] from " & TheTable & " " & chr(34))
sp.WriteLine("	NumDec = 0")
sp.WriteLine("ELSE")
sp.WriteLine("	sqlSelectEnd  = theCalc & ""(Cast("" & theExp & "" as Float)) As ["" & AsField & ""] from " & TheTable & " "  & chr(34))
sp.WriteLine("	NumDec = 4")
sp.WriteLine("End If")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("' --------------------------------  test to see if there is a group by")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("If TheSource(1) <> """" Then ' There is a group by ")
sp.WriteLine("'REsponse.Write(""TheCount is: "" & TheCount & ""<br>"")")
sp.WriteLine("	For I = 1 to TheCount")
sp.WriteLine("			")
sp.WriteLine("			If InStr(TheSource(I), ""*Y"") = FALSE Then")
sp.WriteLine("				SQLSelectMid = SQLSelectMid & TheSource(I) & "", """)
sp.WriteLine("				SQLGroupBy = SQLGroupBy & TheSource(I) & "", """)
sp.WriteLine("				SQLOrderBy = SQLOrderBy & TheSource(I) & "", """)
sp.WriteLine("'				Response.Write(""The mid is: "" & sqlselectmid & ""<br>"")")
sp.WriteLine("")
sp.WriteLine("			Else")
sp.WriteLine("				SourceYear = """"")
sp.WriteLine("				SourceSplit = Split(TheSource(I), ""*Y"")")
sp.WriteLine("				SourceYear = SourceSplit(0)")
sp.WriteLine("				SQLSelectMid = SQLSelectMid & "" DatePart(yyyy, "" & SourceYear & "") As "" & SourceYear & "", """)
sp.WriteLine("				TheSource(I) = Replace(TheSource(I), ""*Y"", """")")
sp.WriteLine("				SQLGroupBy = SQLGroupBy & "" DatePart(yyyy, "" & SourceYear & ""), """)
sp.WriteLine("				SQLOrderBy = SQLOrderBy & "" DatePart(yyyy, "" & SourceYear & ""), """)
sp.WriteLine("			End If")
sp.WriteLine("	Next")
sp.WriteLine("")
sp.WriteLine("SQLGroupBy = Left(SQLGroupBy, Len(SQLGroupBy) - 2)")
sp.WriteLine("SQLOrderBy = Left(SQLOrderBy, Len(SQLOrderBy) - 2)")
sp.WriteLine("	SQLGroupBy = "" Group by "" & SQLGroupBy")
sp.WriteLine("	SQLOrderBy = "" Order by "" & SQLOrderBy")
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
sp.WriteLine("	If AndOr(I) = """" Then")
sp.WriteLine("		Select Case TheType(I)")
sp.WriteLine("			Case ""N""")
sp.WriteLine("				WhereHolder1 = MakeWhereNum(TheESource(I), Filter1(I), Value1(I))")
sp.WriteLine("			Case ""T""")
sp.WriteLine("				WhereHolder1 = MakeWhereText(TheESource(I), Filter1(I), Value1(I))")
sp.WriteLine("			Case ""D""")
sp.WriteLine("				WhereHolder1 = MakeWhereDate(TheESource(I), Filter1(I), Value1(I))")
sp.WriteLine("		End Select")
sp.WriteLine("")
sp.WriteLine("		If I <> TheECount Then")
sp.WriteLine("			sqlWhere = sqlWhere & "" ("" & WhereHolder1 & "") AND """)
sp.WriteLine("		Else")
sp.WriteLine("			sqlWhere = sqlWhere & "" ("" & WhereHolder1 & "") """)
sp.WriteLine("		End If")
sp.WriteLine("			")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("	Else ' there's an and or or	")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("		Select Case TheType(I)")
sp.WriteLine("			Case ""N""")
sp.WriteLine("				WhereHolder1 = MakeWhereNum(TheESource(I), Filter1(I), Value1(I))")
sp.WriteLine("				WhereHolder2 = MakeWhereNum(TheESource(I), Filter2(I), Value2(I))")
sp.WriteLine("			Case ""T""")
sp.WriteLine("				WhereHolder1 = MakeWhereText(TheESource(I), Filter1(I), Value1(I))")
sp.WriteLine("				WhereHolder2 = MakeWhereText(TheESource(I), Filter2(I), Value2(I))")
sp.WriteLine("			Case ""D""")
sp.WriteLine("				WhereHolder1 = MakeWhereDate(TheESource(I), Filter1(I), Value1(I))")
sp.WriteLine("				WhereHolder2 = MakeWhereDate(TheESource(I), Filter2(I), Value2(I))")
sp.WriteLine("		End Select")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("		If I <> TheECount Then")
sp.WriteLine("			sqlWhere = sqlWhere & "" (("" & WhereHolder1 & "") "" & AndOr(I) & "" ("" & WhereHolder2 & "")) AND """)
sp.WriteLine(" 		Else")
sp.WriteLine("			sqlWhere = sqlWhere & "" (("" & WhereHolder1 & "") "" & AndOr(I) & "" ("" & WhereHolder2& "")) """)
sp.WriteLine("		End If")
sp.WriteLine("			")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("	End If")
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
sp.WriteLine("Response.Write(""The SQL is: "" & sql & ""<br>"")")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("'------------------------------------------------------")
sp.WriteLine("")

sp.WriteLine("")
sp.WriteLine("dbName = " & chr(34)  & TheDB & chr(34) & " "  & chr(37) + ">")
sp.WriteLine("")
sp.WriteLine("")

sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("<!" & "--#include virtual = ""/thetimes/longrcon.inc"" -->")
sp.WriteLine("")
sp.WriteLine(Chr(60) + Chr(37))


sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("RS.Open sql, Conn, adOpenStatic, adLockReadOnly, adCmdText")
sp.WriteLine("")
sp.WriteLine("")



sp.WriteLine(chr(37) & ">")
sp.WriteLine("<HTML>")
sp.WriteLine("<head>")
sp.WriteLine("<title>Calculation Results</title>")
sp.WriteLine("</head>")
sp.WriteLine("<body  bgcolor=""white"">")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("<table border=""1""  >")
sp.WriteLine("  <tr>")
sp.WriteLine("")
sp.WriteLine("<" & chr(37))
sp.WriteLine("")
sp.WriteLine("For I = 1 to TheCount ")
sp.WriteLine("TheFieldName = TheSource(I)")
sp.WriteLine("")
sp.WriteLine(Chr(37) & ">")
sp.WriteLine("")
sp.WriteLine("<td bgcolor=""#009933""><font size=""-1""><b><" & Chr(37) & " = TheFieldName " & chr(37) & "></b></font></td>")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("<" & chr(37) & " Next " & Chr(37) & ">")
sp.WriteLine("")

' ---------------------------------------- new --------------------------------

sp.WriteLine("    <td bgcolor=""009933""><font size=""-1""><b><" & Chr(37) & " = AsField " & chr(37) & "></b></font></td>")
sp.WriteLine("<" & chr(37) & " If UCase(TheCalc) = ""MIN"" or UCASE(TheCalc) = ""MAX"" Then " & chr(37) & ">")
sp.WriteLine("    <td bgcolor=""009933""><font size=""-1""><b><" & chr(37) & " = TheCalc " & Chr(37) & "> Value Link</b></font></td>")
sp.WriteLine("<" & chr(37) & " End If " & chr(37) & ">")
sp.WriteLine("   <td bgcolor=""009933""><font size=""-1""><b>Show Records</b></font></td>")

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
sp.WriteLine("	colHref=""frm"" & TheSource(I) & ""=NULL""")
sp.WriteLine("")
sp.WriteLine("	If I = 1 Then")
sp.WriteLine("		SumHref = SumHref & ColHref")
sp.WriteLine("	Else")
sp.WriteLine("		SumHref = SumHref & ""&"" & ColHref")
sp.WriteLine("	End If")
sp.WriteLine("")
sp.WriteLine(chr(37) & ">")
sp.WriteLine("    <td bgcolor=""#CCFFCC""><font size=""-1""><a href=" & chr(34) & ProjectName & "_drilldown.asp?<" & chr(37) & " = ColHref & ""&frmsqlWhere="" & Server.URLEncode(sqlWhere)"  & chr(37) & ">"">Blank</a></font></td>")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("<" & chr(37) & " ElseIf  IsNull(rs(TheSource(I))) Then ")

'-------------------------------
sp.WriteLine("")
sp.WriteLine("	colHref=""frm"" & TheSource(I) & ""=NULL""")
sp.WriteLine("")
sp.WriteLine("	If I = 1 Then")
sp.WriteLine("		SumHref = SumHref & ColHref")
sp.WriteLine("	Else")
sp.WriteLine("		SumHref = SumHref & ""&"" & ColHref")
sp.WriteLine("	End If")
sp.WriteLine("")
sp.WriteLine(chr(37) & ">")
sp.WriteLine("    <td bgcolor=""#CCFFCC""><font size=""-1""><a href=" & chr(34) & ProjectName & "_drilldown.asp?<" & chr(37) & " = ColHref & ""&frmsqlWhere="" & Server.URLEncode(sqlWhere)"  & chr(37) & ">"">Null</a></font></td>")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("<" & chr(37) & " Else ")
sp.WriteLine("	colHref=""frm"" & TheSource(I) & ""="" & Server.URLEncode(rs(TheSource(I)))")
sp.WriteLine("")
sp.WriteLine("	If I = 1 Then")
sp.WriteLine("		SumHref = SumHref & ColHref")
sp.WriteLine("	Else")
sp.WriteLine("		SumHref = SumHref & ""&"" & ColHref")
sp.WriteLine("	End If")
sp.WriteLine("")
sp.WriteLine(chr(37) & ">")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("    <td bgcolor=""#CCFFCC""><font size=""-1""><a href=" & chr(34) & ProjectName & "_drilldown.asp?<" & chr(37) & " = ColHref & ""&frmsqlWhere="" & Server.URLEncode(sqlWhere)"  & chr(37) & ">""><" & chr(37) & " = RS(TheSource(I)) " & chr(37) & "></font></td>")
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
sp.WriteLine("	If Not IsNull(rs(AsField)) Then")
sp.WriteLine("		colHref=""frm"" & TheExp & ""="" & Server.URLEncode(rs(AsField))")
sp.WriteLine("	Else")
sp.WriteLine("		colHref=""frm"" & TheExp & ""=NULL""")
sp.WriteLine("	End If		")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("	If I = 1 Then")
sp.WriteLine("		SumHref = SumHref & ColHref")
sp.WriteLine("	Else")
sp.WriteLine("		SumHref = SumHref & ""&"" & ColHref")
sp.WriteLine("	End If")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")


sp.WriteLine("If rs(AsField).Type = 135 Then ")

'------------------ stopped here ------------------------------------

sp.WriteLine(chr(37) & ">")
sp.WriteLine("		<td bgcolor=""#CCFFCC"" align=""right""><font size=""-1""><" & chr(37) & " = rs(AsField) " & chr(37) & "></font></td>")
sp.WriteLine("		<" & chr(37) & " On Error Resume Next " & chr(37) & ">")
sp.WriteLine("		<td bgcolor=""#CCFFCC"" align=""right""><font size=""-1""><a href=" & chr(34) & ProjectName & "_drilldown.asp?<" & chr(37) & " = ColHref & ""&frmsqlWhere="" & Server.URLEncode(sqlWhere)"  & chr(37) & ">""><" & chr(37) & " = TheCalc " & chr(37) & "> Date</a></font></td>")
sp.WriteLine("		<td bgcolor=""#CCFFCC"" align=""right""><font size=""-1""><a href=" & chr(34) & ProjectName & "_drilldown.asp?<" & chr(37) & " = SumHref & ""&frmsqlWhere="" & Server.URLEncode(sqlWhere)"  & chr(37) & ">"">All in Row</a></font></td>")
sp.WriteLine("")
sp.WriteLine("		<" & chr(37) & " Else ")
sp.WriteLine("		On error resume next " & chr(37) & ">")
sp.WriteLine("		<td bgcolor=""#CCFFCC"" align=""right""><font size=""-1""><" & chr(37) & " = FormatNumber(rs(AsField),2) " & chr(37) & "></td>")
sp.WriteLine("		<% On Error Resume Next " & chr(37) & ">")
sp.WriteLine("		<td bgcolor=""#CCFFCC"" align=""right""><font size=""-1""><a href=" & chr(34) & ProjectName & "_drilldown.asp?<" & chr(37) & " = ColHref & ""&frmsqlWhere="" & Server.URLEncode(sqlWhere)"  & chr(37) & ">""><" & chr(37) & " = TheCalc " & chr(37) & "> Number</a></font></td>")
sp.WriteLine("		<td bgcolor=""#CCFFCC"" align=""right""><font size=""-1""><a href=" & chr(34) & ProjectName & "_drilldown.asp?<" & chr(37) & " = SumHref & ""&frmsqlWhere="" & Server.URLEncode(sqlWhere)"  & chr(37) & ">"">All in Row</a></font></td>")
sp.WriteLine("")
sp.WriteLine("	<" & chr(37) & " End If ")
sp.WriteLine("")
sp.WriteLine("Else ")
sp.WriteLine("")
sp.WriteLine("	If rs(AsField).Type = 135 Then ")
sp.WriteLine("")
sp.WriteLine(chr(37) & ">")
sp.WriteLine("		<td bgcolor=""#CCFFCC"" align=""right""><font size=""-1""><" & chr(37) & " = rs(AsField) " & chr(37) & "></font></td>")
sp.WriteLine("")
sp.WriteLine("	<" & chr(37) & " Else ")
sp.WriteLine("	On error Resume Next " & chr(37) & ">")
sp.WriteLine("		<td bgcolor=""#CCFFCC"" align=""right""><font size=""-1""><" & chr(37) & " = FormatNumber(rs(AsField), 2) " & chr(37) & "></font></td>")
sp.WriteLine("")
sp.WriteLine("	")
sp.WriteLine("	<" & chr(37) & " End If " & chr(37) & ">")
sp.WriteLine("	<td bgcolor=""#CCFFCC"" align=""right""><font size=""-1""><a href=" & chr(34) & ProjectName & "_drilldown.asp?<" & chr(37) & " = SumHref & ""&frmsqlWhere="" & Server.URLEncode(sqlWhere)"  & chr(37) &  ">"">All in Row</a></font></td>")
sp.WriteLine("")
sp.WriteLine("	")
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
sp.WriteLine("	Case ""eq""")
sp.WriteLine("		MakeWhereNum = "" "" & TheFieldName & "" = "" & TheValue")
sp.WriteLine("	Case ""noteq""")
sp.WriteLine("		MakeWhereNum = "" "" & TheFieldName & "" <> "" & TheValue")
sp.WriteLine("	Case ""gt""")
sp.WriteLine("		MakeWhereNum = "" "" & TheFieldName & "" > "" & TheValue")
sp.WriteLine("	Case ""gtet""")
sp.WriteLine("		MakeWhereNum = "" "" & TheFieldName & "" >= "" & TheValue")
sp.WriteLine("	Case ""lt""")
sp.WriteLine("		MakeWhereNum = "" "" & TheFieldName & "" < "" & TheValue")
sp.WriteLine("	Case ""ltet""")
sp.WriteLine("		MakeWhereNum = "" "" & TheFieldName & "" <= "" & TheValue")
sp.WriteLine("	Case ""bw""")
sp.WriteLine("		MakeWhereNum = "" Cast("" & TheFieldName & "" as varchar(25)) Like '"" & TheValue & ""%'""")
sp.WriteLine("	Case ""dnbw""")
sp.WriteLine("		MakeWhereNum = "" Cast("" & TheFieldName & "" as varchar(25)) NOT Like '"" & TheValue & ""%'""")
sp.WriteLine("	Case ""ew""")
sp.WriteLine("		MakeWhereNum = "" Cast("" & TheFieldName & "" as varchar(25)) Like '%"" & TheValue & ""'""")
sp.WriteLine("	Case ""dnew""")
sp.WriteLine("		MakeWhereNum = "" Cast("" & TheFieldName & "" as varchar(25)) NOT Like '%"" & TheValue & ""'""")
sp.WriteLine("	Case ""con""")
sp.WriteLine("		MakeWhereNum = "" Cast("" & TheFieldName & "" as varchar(25)) Like '%"" & TheValue & ""%'""")
sp.WriteLine("	Case ""dncon""")
sp.WriteLine("		MakeWhereNum = "" Cast("" & TheFieldName & "" as varchar(25)) NOT Like '%"" & TheValue & ""%'""")
sp.WriteLine("	Case ""blank""")
sp.WriteLine("		MakeWhereNum = "" "" & TheFieldName & "" is null """)
sp.WriteLine("	Case ""notblank""")
sp.WriteLine("		MakeWhereNum = "" "" & TheFieldName & "" is NOT null  "" ")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("End Select")
sp.WriteLine("")
sp.WriteLine("End Function")
sp.WriteLine("")
sp.WriteLine("Function MakeWhereText(TheFieldName, TheOperator, TheValue)")
sp.WriteLine("")
sp.WriteLine("If InStr(TheValue, ""'"") Then")
sp.WriteLine("	TheValue = Replace(TheValue, ""'"", ""''"")")
sp.WriteLine("End If")
sp.WriteLine("Select Case TheOperator")
sp.WriteLine("	Case ""eq""")
sp.WriteLine("		MakeWhereText = "" "" & TheFieldName & "" = '"" & TheValue & ""'""")
sp.WriteLine("		If InStr(MakeWhereText, ""]"") and InStr(MakeWhereText, ""["") Then")
sp.WriteLine("			MakeWhereText = Replace(MakeWhereText, ""'"", """")")
sp.WriteLine("		End If")
sp.WriteLine("	Case ""noteq""")
sp.WriteLine("		MakeWhereText = "" "" & TheFieldName & "" <> '"" & TheValue & ""'""")
sp.WriteLine("		If InStr(MakeWhereText, ""]"") and InStr(MakeWhereText, ""["") Then")
sp.WriteLine("			MakeWhereText = Replace(MakeWhereText, ""'"", """")")
sp.WriteLine("		End If")
sp.WriteLine("	Case ""gt""")
sp.WriteLine("		MakeWhereText = "" "" & TheFieldName & "" > '"" & TheValue & ""'""")
sp.WriteLine("		If InStr(MakeWhereText, ""]"") and InStr(MakeWhereText, ""["") Then")
sp.WriteLine("			MakeWhereText = Replace(MakeWhereText, ""'"", """")")
sp.WriteLine("		End If")
sp.WriteLine("	Case ""gtet""")
sp.WriteLine("		MakeWhereText = "" "" & TheFieldName & "" >=  '"" & TheValue & ""'""")
sp.WriteLine("		If InStr(MakeWhereText, ""]"") and InStr(MakeWhereText, ""["") Then")
sp.WriteLine("			MakeWhereText = Replace(MakeWhereText, ""'"", """")")
sp.WriteLine("		End If")
sp.WriteLine("	Case ""lt""")
sp.WriteLine("		MakeWhereText = "" "" & TheFieldName & "" <  '"" & TheValue & ""'""")
sp.WriteLine("		If InStr(MakeWhereText, ""]"") and InStr(MakeWhereText, ""["") Then")
sp.WriteLine("			MakeWhereText = Replace(MakeWhereText, ""'"", """")")
sp.WriteLine("		End If")
sp.WriteLine("	Case ""ltet""")
sp.WriteLine("		MakeWhereText = "" "" & TheFieldName & "" <= '"" & TheValue & ""'""")
sp.WriteLine("		If InStr(MakeWhereText, ""]"") and InStr(MakeWhereText, ""["") Then")
sp.WriteLine("			MakeWhereText = Replace(MakeWhereText, ""'"", """")")
sp.WriteLine("		End If")
sp.WriteLine("	Case ""bw""")
sp.WriteLine("		MakeWhereText = "" "" & TheFieldName & "" Like '"" & TheValue & ""%'""")
sp.WriteLine("	Case ""dnbw""")
sp.WriteLine("		MakeWhereText = "" "" & TheFieldName & "" NOT Like '"" & TheValue & ""%'""")
sp.WriteLine("	Case ""ew""")
sp.WriteLine("		MakeWhereText = "" "" & TheFieldName & "" Like '%"" & TheValue & ""'""")
sp.WriteLine("	Case ""dnew""")
sp.WriteLine("		MakeWhereText = "" "" & TheFieldName & "" NOT Like '%"" & TheValue & ""'""")
sp.WriteLine("	Case ""con""")
sp.WriteLine("		MakeWhereText = "" "" & TheFieldName & "" Like '%"" & TheValue & ""%'""")
sp.WriteLine("	Case ""dncon""")
sp.WriteLine("		MakeWhereText = "" "" & TheFieldName & "" NOT Like '%"" & TheValue & ""%'""")
sp.WriteLine("	Case ""blank""")
sp.WriteLine("		MakeWhereText = "" ("" & TheFieldName & "" is Null or "" & TheFieldName & "" = '') """)
sp.WriteLine("	Case ""notblank""")
sp.WriteLine("		MakeWhereText = "" ("" & TheFieldName & "" is NOT Null and "" & TheFieldName & "" <> '') """)
sp.WriteLine("		")
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
sp.WriteLine("	Case ""eq""")
sp.WriteLine("		MakeWhereDate = "" "" & TheFieldName & "" = '"" & TheValue & ""'""")
sp.WriteLine("		If InStr(MakeWhereDate, ""]"") and InStr(MakeWhereDate, ""["") Then")
sp.WriteLine("			MakeWhereDate = Replace(MakeWhereDate, ""'"", """")")
sp.WriteLine("		End If")
sp.WriteLine("	Case ""noteq""")
sp.WriteLine("		MakeWhereDate = "" "" & TheFieldName & "" <> '"" & TheValue & ""'""")
sp.WriteLine("		If InStr(MakeWhereDate, ""]"") and InStr(MakeWhereDate, ""["") Then")
sp.WriteLine("			MakeWhereDate = Replace(MakeWhereDate, ""'"", """")")
sp.WriteLine("		End If")
sp.WriteLine("	Case ""gt""")
sp.WriteLine("		MakeWhereDate = "" "" & TheFieldName & "" > '"" & TheValue & ""'""")
sp.WriteLine("		If InStr(MakeWhereDate, ""]"") and InStr(MakeWhereDate, ""["") Then")
sp.WriteLine("			MakeWhereDate = Replace(MakeWhereDate, ""'"", """")")
sp.WriteLine("		End If")
sp.WriteLine("	Case ""gtet""")
sp.WriteLine("		MakeWhereDate = "" "" & TheFieldName & "" >=  '"" & TheValue & ""'""")
sp.WriteLine("		If InStr(MakeWhereDate, ""]"") and InStr(MakeWhereDate, ""["") Then")
sp.WriteLine("			MakeWhereDate = Replace(MakeWhereDate, ""'"", """")")
sp.WriteLine("		End If")
sp.WriteLine("	Case ""lt""")
sp.WriteLine("		MakeWhereDate = "" "" & TheFieldName & "" <  '"" & TheValue & ""'""")
sp.WriteLine("		If InStr(MakeWhereDate, ""]"") and InStr(MakeWhereDate, ""["") Then")
sp.WriteLine("			MakeWhereDate = Replace(MakeWhereDate, ""'"", """")")
sp.WriteLine("		End If")
sp.WriteLine("	Case ""ltet""")
sp.WriteLine("		MakeWhereDate = "" "" & TheFieldName & "" <= '"" & TheValue & ""'""")
sp.WriteLine("		If InStr(MakeWhereDate, ""]"") and InStr(MakeWhereDate, ""["") Then")
sp.WriteLine("			MakeWhereDate = Replace(MakeWhereDate, ""'"", """")")
sp.WriteLine("		End If")
sp.WriteLine("	Case ""bw""")
sp.WriteLine("		MakeWhereDate = "" Cast("" & TheFieldName & "" as varchar(25)) Like '"" & TheValue & ""%'""")
sp.WriteLine("	Case ""dnbw""")
sp.WriteLine("		MakeWhereDate = "" Cast("" & TheFieldName & "" as varchar(25)) NOT Like '"" & TheValue & ""%'""")
sp.WriteLine("	Case ""ew""")
sp.WriteLine("		MakeWhereDate = "" Cast("" & TheFieldName & "" as varchar(25)) Like '%"" & TheValue & ""'""")
sp.WriteLine("	Case ""dnew""")
sp.WriteLine("		MakeWhereDate = "" Cast("" & TheFieldName & "" as varchar(25)) NOT Like '%"" & TheValue & ""'""")
sp.WriteLine("	Case ""con""")
sp.WriteLine("		MakeWhereDate = "" Cast("" & TheFieldName & "" as varchar(25)) Like '%"" & TheValue & ""%'""")
sp.WriteLine("	Case ""dncon""")
sp.WriteLine("		MakeWhereDate = "" Cast("" & TheFieldName & "" as varchar(25)) NOT Like '%"" & TheValue & ""%'""")
sp.WriteLine("	Case ""blank""")
sp.WriteLine("		MakeWhereDate = "" "" & TheFieldName & "" is NULL "" ")
sp.WriteLine("	Case ""notblank""")
sp.WriteLine("		MakeWhereDate = "" "" & TheFieldName & "" is NOT NULL "" ")
sp.WriteLine("")
sp.WriteLine("End Select")
sp.WriteLine("")
sp.WriteLine("End Function")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine("")
sp.WriteLine(chr(37) & ">")









sp.Close
rs.Close




End If ' end of branch for creating an expressions page

' -------------- end expressions.asp page 




' -----------------------------


sql = "select data_type from information_schema.columns where table_name = '" & TheTable & "' and column_name = '" & LField & "'"





' write the _profile.asp page
Set sp = fs.CreateTextFile(TheFullPath & "\" & ProjectName & "_profile.asp", True)

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
sql = "select data_type from information_schema.columns where table_name = '" & TheTable & "' and column_name = '" & LField & "'"
sp.WriteLine(Lfield & " = Request(" & chr(34) & LField & chr(34) & ")")

RS.Open sql, Conn, adOpenStatic, adLockReadOnly, adCmdText



Select Case rs("data_type")
	Case "bigint", "bit", "decimal", "float", "int", "money", "numeric", "real", "smallint", "smallmoney", "tinyint"
		sp.WriteLine("sql = ""Select * from " & TheTable & " where " & LField & " = " & chr(34) & " & " & LField)
	Case Else
		sp.WriteLine("sql = ""Select * from " & TheTable & " where " & LField & " = '" & chr(34)  & " & " & LField & " & " & chr(34) &  "'" & chr(34))
End Select




sp.WriteLine("")
sp.WriteLine("dbName = " & chr(34)  & TheDB & chr(34) & " "  & chr(37) + ">")
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
sp.WriteLine("<TITLE>Query Results Page</TITLE>")
sp.WriteLine("</HEAD>")
sp.WriteLine("<body bgcolor=""#FFFFFF"">")
sp.WriteLine("<!" & "--#include file = """ & ProjectName & "_links.inc"" -->")
sp.WriteLine("<blockquote>")
sp.WriteLine("<" & chr(37))


rs.Close


sql="select column_name as TheField from information_schema.columns where table_name = '" & TheTable & "' "
sql = sql & " order by ordinal_position"

'Response.Write("The sql is: " & sql & "<br>")

RS.Open sql, Conn, adOpenStatic, adLockReadOnly, adCmdText


Do While Not RS.EOF

TheLine = "str" & rs("TheField") & "= rs(" & chr(34) & rs("TheField") & chr(34) & ")"
sp.WriteLine(TheLine)
TheLine = "If str" & rs("TheField") & " <> " & chr(34) & Chr(34) & " Then " & chr(37) & ">"
sp.WriteLine(TheLine)
TheLine = "<p align=""left""><font size=+1><b>" & rs("TheField") & ": </b></font>" & Chr(60) & Chr(37) & " = str" & rs("TheField")  & " " & chr(37) + "></p>"
sp.WriteLine(TheLine)
TheLine = Chr(60) & Chr(37) & " End If"
sp.WriteLine(TheLine)

rs.MoveNext
Loop


sp.WriteLine(chr(37) & ">")
sp.WriteLine("</body>")
sp.WriteLine("</html>")



sp.close




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

<center><i>Shboom Express was developed by Tom Torok of The New York Times</i></center>

</BODY>
</HTML>


<%

set fs = nothing
rs.close



%>




