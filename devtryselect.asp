<%

Dim TheType(24)

TheType(1)="bigint"
TheType(2)="binary"
TheType(3)="bit"
TheType(4)="char"
TheType(5)="datetime"
TheType(6)="decimal"
TheType(7)="float"
TheType(8)="image"
TheType(9)="int"
TheType(10)="money"
TheType(11)="nchar"
TheType(12)="ntext"
TheType(13)="ntext"
TheType(14)="numeric"
TheType(15)="nvarchar"
TheType(16)="smalldatetime"
TheType(17)="real"
TheType(18)="smalldatetime"
TheType(19)="smallint"
TheType(20)="smallmoney"
TheType(21)="text"
TheType(22)="tinyint"
TheType(23)="varchar"


'TheType = "tinyint"

For I = 1 to 23

Select Case TheType(I)

	Case "bigint", "bit", "decimal", "float", "int", "money", "numeric", "real", "smallint", "smallmoney", "tinyint"
		Response.Write(TheType(I) & " is a number.<br>")
	Case "datetime", "smalldatetime"
		Response.Write(TheType(I) & " is a date.<br>")
	Case "char", "nchar", "ntext", "nvarchar", "text", "varchar"
		Response.Write(TheType(I) & " is text.<br>")
End Select
Next


%>
	

