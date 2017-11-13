<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<%procedimentos = 1'request("procedimentos")%>
<script language="Javascript">
function soma(a,b,procedimentos){
	a=a;
	b=b;
	procedimentos=procedimentos
	
		if (procedimentos != ''){
			virg =  ',';			
			}
		else
		{
		virg= ''
		}		
	procedimentos = procedimentos + virg
		if (a != ""){
			procedimentos = procedimentos + a;
			}
		else
			{
			procedimentos = procedimentos + b;
			}
			

	document.Form.res.value=a;
	//document.form.procedimentos.value=procedimentos;
	document.Form.procedimentos.value=(procedimentos.replace(",,", ","));
	//document.Form.procedimento.value="";
	//document.Form.cd_procedimento_1.value="";
	}
</script>
<html>
<head>
	<title>Untitled</title>
</head>

<body>
<form method="post" name="Form" id="Form">
	<input type="text" name="a" size="5" maxlength="4"> - 
	<input type="text" name="b" size="5"> = 
	<input type="text" name="res" size="5">
	(<input type="text" name="procedimentos" value="<%=procedimentos%>">)
	<input type="button" name="subtrair" value="-" onclick="soma(document.Form.a.value,document.Form.procedimentos.value)";"><br>
</form>



<script language="Javascript">
	
</script>
</body>
</html>
