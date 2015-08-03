<%@ LANGUAGE="VBSCRIPT" %>
<% option explicit %>
<!--#include file ="adovbs.inc"-->


<html>

<head>
<title>Shboom</title>
</head>

<body bgColor="white">
<!--Written by Tom Torok of The New York Times
tomtorok@nytimes.com-->

<% dim fs, d, dc, s, n

Set fs = CreateObject("Scripting.FileSystemObject")
Set dc = fs.Drives %>

<center><small>(If you know the drill, to start, select a drive letter for the files you will build.)</small><br>

<% For Each d in dc %>

<%
If d.DriveType = 2 Then %>

<a href="folders.asp?dr=<% = d.RootFolder %>"><%  = d.DriveLetter %></a>|

							
<% End IF %>

<% Next %>
</center>
<h1 align="center">Shboom</h1>
<P align=center><b>Developed by </b></P>
  <p align="center"><b style="FONT-SIZE: larger">Tom 
    Torok</b></p>
<P align=center><b>of The 
New York Times</b></P>

<blockquote>
<p align="Left">If you're in a hurry, use Shboom Express. This program, plain ole Shboom, has more functionality but
takes longer to set up. It will not work on Windows NT (it might if it has a new version of MDAC installed, by I haven't tried it).</p>
Point Shboom to a table or a view and it will generate three sets of
web pages that will give you much of the functionality of a database program. It also will generate a help page specific to your data.</p>
<p>The &quot;filter&quot; pages will let you query as many fields as you choose
in a fashion similar to the &quot;custom&quot; option in Microsoft
Excel's data filter. You can select, order and alias the fields to be shown in the pulldown menus. You also can select and order the columns and rows to be shown in the
results page. If your table or view contains a unique identifier, that
identifier will be used in a link on the results page to a profile page. You can preselect and alias the fields to be shown on the profile page.</p>
<p>The &quot;group by&quot; pages will let you group by and count as many fields as you choose. This page also provides the same filtering capabilities as the filter
page and will display your results in the order of the fields selected.</p>
<p>On the calculations page you can select any numeric or date field&nbsp; to appear in a pulldown for 
calculations of minimum, maximum, average, sum or count. You also can use the
expression in conjunction with the group-by and filter features.</p>
<p>For this program to work you must have two include files, rcon.inc and longrcon.inc, in a directory on your web called "thetimes".</p>
<p>Also you must have your web configured to allow this program to write to web folders</p>
<p>This program is intended for development or project servers. If you intend to use the resulting pages on a more secure server, simply 
transfer the pages over to the secure server. Be sure the secure server has an rcon.inc and longrcon.inc in a directory called "thetimes".
If you intend to use this on a secure server do not use the funtion "Allow User to Modify Application."
<p>The function "Allow User to Modify Application" will create a link called "Modify." The link will let users change 
page titles; the max number of fields for filtering and group bys; the fields, order and aliases to include in pulldowns, the max number of records to display; 
and the fields and order used in the results and profile pages.</p>
<p>In addition to web and include files, this program also stores project and table information in two disconnected recordsets 
that are stored in the web directory as [projectname]_progspecs.gaga and [projectname]_tablespecs.gaga. These files are used during 
the creation process and are used when users modify the application.</p>

</blockquote>
<center>
<% For Each d in dc %>

<%
If d.DriveType = 2 Then %>

<a href="folders.asp?dr=<% = d.RootFolder %>"><%  = d.DriveLetter %></a>|

							
<% End IF %>

<% Next %>
</center>

</body>
</html>
