<%


for each i in request.servervariables
	'response.write "<li>"& i & " : " & request.servervariables(i)
next

sn = request.servervariables("script_name")

Dim j, c, sn, inkey, key

for j = 1 to len(sn)
	c = mid( sn, j, 1)
	if c = "_" then
		inkey = -1	
		key = ""
	elseif inkey then
		key = key & c
		if len( key ) = 3 then
			inkey = 0
		end if
	end if

	response.write "<li>" & j & " " & c
next
response.write "Key " & key

%>