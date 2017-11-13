<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">

<html>
<head>
	<title>Untitled</title>
</head>
<%nome = request("nome")%>
<body>
<form action="teste.asp" name="teste" id="teste">
	<input type="text" name="teste" value="<%=nome%>">
	<input type="submit" name="ok" value="ok">
</form>

</body>
</html>
