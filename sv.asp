<%


for each i in request.servervariables
	'response.write "<li>"& i & " : " & request.servervariables(i)
next

sn = request.servervariables("script_name")

Dim j, c, sn, key

for j = 1 to len(sn)
	c = mid( sn, j, 1)
	if c = "_" then
		if len(key) < 3then
			key = key & c
		end if
	end if
	response.write "<li>" & j & " " & c
next
response.write key

%>