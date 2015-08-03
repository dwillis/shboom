<%

Dim TheField()



HowMany = Request("C1").Count


Redim TheField(HowMany + 1)


For I = 1 to HowMany

	FieldName = Request("C1")(I)
	TheField(I) = Request(FieldName & "Text")


	Response.Write(TheField(I) & "<br>")

Next

%>
