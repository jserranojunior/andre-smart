<html>
<head>
<title>Teste ASP</title>
</head>
<body bgcolor="#E6E7C0">

<p><font face="Tahoma" size="2">
<%
'session.LCID = 1046

dim dia_extenso(7)
dia_extenso(1)="Domingo"
dia_extenso(2)="Segunda-feira"
dia_extenso(3)="Terça-feira"
dia_extenso(4)="Quarta-feira"
dia_extenso(5)="Quinta-feira"
dia_extenso(6)="Sexta-feira"
dia_extenso(7)="Sábado"



Response.write dia_extenso(weekday(now()))
response.write (weekday("10/07/2008"))
%>
</font></p>

</body>
</html>