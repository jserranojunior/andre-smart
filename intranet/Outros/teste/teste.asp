<html>
	<head>
		<title>
		</title>
		
		<%if a="" Then
		virg=""
	else
		virg=","
	end if%>
	
	
	
<script language="JavaScript">
<!--
function soma(a,subtotal){
	a=a;
		if (subtotal != ''){
			virg = ',';
			}
		else
		{
		virg= ''
		}
	subtotal=subtotal+virg+a;
	document.demo.res.value=a;
	document.demo.subtotal.value=subtotal;
	document.demo.num1.value="";
}
//-->
</script>
	</head>
<body>
<form name="demo">
	Número1:<input type="text" name="num1" size="5"><br><br>
	&nbsp;&nbsp;&nbsp;&nbsp;<input type="text" name="subtotal">
	
	
	Resultado:<input type="text" name="res" size="5"><br>
	<br>
	<input type="button" name="somar" value="Inserir" 
	onclick="soma(document.demo.num1.value,document.demo.subtotal.value)">
</form>

</body>
</html>