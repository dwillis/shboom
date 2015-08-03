<%@ Language=VBScript %>
<% On Error Resume Next %>
<% 
  ' Declare all ADO Constants
  Const adUseClient = 3	    ' Client Side Cursors
  Const adOpenStatic = 3    ' Static Cursor
  Const adVarChar = 200     ' Variable Character

  ' Declare all ADO Objects to be Used
  Dim objRS                 ' ADO Recordset Obejct



Dim TheLast(5)
Dim TheFirst(5)
Dim TheE(5)

TheLast(1)="Torok"
TheLast(2)="Lentini"
TheLast(3)="Gibson"
TheLast(4)="Smith"
TheLast(5)="Geise"
TheFirst(1)="Tom"
TheFirst(2)="Cece"
TheFirst(3)="Kevin"
TheFirst(4)="Robert"
TheFirst(5)="Arlene"
TheE(1)="Torok@yahoo.com"
TheE(2)="Lentini@yahoo.com"
TheE(3)="Gibson@yahoo.com"
TheE(4)="Smith@yahoo.com"
TheE(5)="Geise@yahoo.com"










  Set objRS = Server.CreateObject("ADODB.Recordset")    ' Create the Recordset
  objRS.CursorType = adOpenStatic       ' Specify the cursor location and type
  objRS.CursorLocation = adUseClient


    objRS.Fields.Append "FName", adVarChar, 25     ' Create the Recordset
    objRS.Fields.Append "LName", adVarChar, 35
    objRS.Fields.Append "EMail", adVarChar, 65
    objRS.Open



For I = 1 to 5

    objRS.AddNew
    objRS("FName") = TheFirst(I)
    objRS("LName") = TheLast(I)
    objRS("EMail") = TheE(I)
    objRS.Update       ' Commit new record


Next


  objRS.Sort = "LName"     ' Sort the Recorset

  objRS.MoveFirst

%>
<html>
<body>
<table align="center" border="1">
<tr>
<td>First</td>
<td>last</td>
<td>email</td>
</tr>


<%
Do Until objRS.EOF %>
<TR>
	<TD><%=objRS("FName")%></TD>
	<TD><%=objRS("LName")%></TD>
	<TD><%=objRS("EMail")%></TD>
</TR>
<% objRS.MoveNext %>
<% Loop %>
</TABLE>

</BODY>
</HTML>

<%
  ' Close the Recodset
  objRS.Close
  '-- Set the Recordset object to nothing
  Set objRS = Nothing
%>
