<% 

Const ForReading = 1, ForWriting = 2, ForAppending = 3
Set fs = CreateObject("Scripting.FileSystemObject")


' write the search page 
Set sp = fs.CreateTextFile("C:\inetpub\wwwroot\shboomexpress\devtext_filter.inc", True)

sp.WriteLine("<td align=""center""><select size=""1"" name=""frmSource"">")
sp.WriteLine("<option value="""""" selected></option>")

sp.Close

set fs = nothing

%>


