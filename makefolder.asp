<%@ Language=VBScript %>

<% 
Const ForReading = 1, ForWriting = 2, ForAppending = 3

' Set a flag for folders that exist
BumFolder = False
ANewFolder = "N"

TheDrive = Request("frmDrive")
ThePath = Trim(Request("frmPath"))
TheFullPath = TheDrive & "\" & ThePath
TheFullPath = Replace(TheFullPath,"\\","\")
' Set flags for changed attribute and added adovbs.inc
AttribChanged = "N"
AddedAdo = "N"
Set fs = CreateObject("Scripting.FileSystemObject")

If ThePath <> "" Then ' if the path isn't empty
	
	If fs.FolderExists(TheFullPath) Then ' if the folder exits
		BumFolder = True
	Else ' if the folder needs to be created
		fs.CreateFolder(TheFullPath)
		ANewFolder = "Y"
		TheAsa = TheFullPath & "\global.asa"
		TheAsa = Replace(TheAsa,"\\","\")
		If Not fs.FileExists(TheAsa) Then ' if there is no global asa
			Set ma = fs.CreateTextFile(TheAsa, True)
			ma.WriteLine("<SCRIPT LANGUAGE=" & Chr(34) & "VBScript" & Chr(34) & " RUNAT=" & Chr(34) & "Server" & Chr(34) & ">")
			ma.WriteLine(" ")
			ma.WriteLine("' Put some stuff here if you want ")
			ma.WriteLine(" ")
			ma.WriteLine("</SCRIPT>")
			ma.Close
		End If	' end creating global asa
	End If ' end if folder exists

End If


%>

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>Make or Select Folder</title>
</HEAD>
<body bgColor="white">
<!--Written by Tom Torok of The New York Times
tomtorok@nytimes.com-->


<% If ANewFolder = "Y" Then %>
<h1 align="center">Directory Created</h1>

<p align="center">The following directory has been created:</p>

<p align="center"><strong><% = TheFullPath %></strong></p>

<p align="center">A blank global.asa file was created in the directory</p>

<form method="POST" action="makefiles.asp">
  <center><p><input type="submit" value="Next" name="B1"></p>
  </center>
<input type="hidden" value="<% = TheFullPath %>" name="TheFullPath">
<input type="hidden" value="Y" name="ANewFolder">
</form>
<% Else
	If BumFolder = True Then %>
<h1 align="center">Directory Exits</h1>

<p align="center">There is a directory with that name. Click the back key and try again.</p>

	<% Else %>

<h1 align="center">Directory Selected</h1>

<p align="center">The following directory has been selected:</p>

<p align="center"><strong><% = TheFullPath %></strong></p>


<form method="POST" action="makefiles.asp">
  <center><p><input type="submit" value="Next" name="B1"></p>
  </center>
<input type="hidden" value="<% = TheFullPath %>" name="TheFullPath">
<input type="hidden" value="N" name="ANewFolder">
	<% End If 
End If %>

</BODY>
</HTML>
