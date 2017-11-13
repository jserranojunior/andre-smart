<head>
	<!--#include file="../../includes/inc_open_connection.asp"-->
	<!--#include file="../../includes/util.asp"-->
    
	<%tipo = request("tipo")
	
	
	cd_unidade = request("cd_unidade")
	cd_protocolo = request("cd_protocolo")
	cd_digito = request("cd_digito")
	%>

<title>Confirmação de exclusão - VDLAP</title>
	  <LINK href="css/estilo.css " type=text/css rel=stylesheet>
	  <style type="text/css" media="screen">
<!--
#ok_print{ visibility:hidden; display: none;}
#no_print{ visibility:visible; display:block;}
table{
	background-color: #ffffff; 
    border: 1px solid #cccccc;
	font-family: verdana;
	font-size: 9px;
	text-decoration:none;}

.txt_titulo {color: #000000;font-family: verdana;font-size:16px;}
.cabecalho {color: #000000;font-family: verdana;font-size:14px;}
.txt_menu {color: #FFFFFF;font-family: verdana;font-size:10px;}
.txt_menu:hover {color: #FFFFFF;font-family: verdana;font-size:10px;}
.txt_menu:link {color: #FFFFFF;font-family: verdana;font-size:10px;}
.txt_menu:visited {color: #FFFFFF;font-family: verdana;font-size:10px;}
.txt {color: #000000;font-family: verdana;font-size:10px;text-decoration:none;}
.txt_cinza {color: #6c6c6c;font-family: verdana;font-size:12px;text-decoration:none; }
.inputs { background-color: #cdcdcd; font: 12px verdana, arial, helvetica, sans-serif; color:#000000; border:1px solid #cdcdcd; }  
-->
</style>
<body>
<!--table style="border:1px solid silver; border-collapse:collapse; width:400px;"-->
<table>
	<%xsql ="SELECT * FROM TBL_protocolo WHERE A002_numpro = '"&cd_protocolo&"' AND A053_codung = '"&cd_unidade&"'"
			Set rs = dbconn.execute(xsql)
			do while Not rs.EOF
			
				cd_codigo = rs("a002_numseq")
				protocolo = rs("A002_numpro")
				cd_prot = rs("A002_numpro")
				cd_unidade = rs("A053_codung")
				nm_nome = rs("A002_pacnom")
				cd_idade = rs("A002_pacida")
				cd_registro = rs("A002_pacreg")
				cd_convenio = rs("A054_codcon")
				nm_sexo = rs("A002_pacsex")
				nm_cirug_realizada = rs("A002_limarm")
				dt_procedimento = rs("A002_datpro")
				nm_agenda = rs("A002_tipage")
				dt_hr_agenda = rs("A002_horage")
				dt_inicio = rs("A002_horini")
				dt_fim = rs("A002_horfin")
					dt_duracao = total_mes(datediff("n",dt_inicio,dt_fim))'&" ["&dt_inicio&" - "&dt_fim&"]"
				cd_med = rs("A055_codmed")
				cd_rack = rs("a056_codrac")
				cd_especialidade = rs("A057_codesp")
				cd_co = rs("cd_co")
				
			rs.movenext
			loop
			
			rs.close
			set rs = nothing%>
	<tr bgcolor="#808080">
		<td align="center" colspan="100%"><font size="2" color="white"><b>&nbsp;Confirmação de exclusão</b></font></td>
	</tr>				
	<tr><td colspan="100%">&nbsp;</td></tr>
	<tr bgcolor="#e0e0e0"">
		<td>&nbsp;Cód.: 
		<td colspan="2"><b><%=digitov(zero(cd_unidade)&"."&proto(cd_protocolo))%> &nbsp;</b></td>
	</tr>
	<tr>
		<td>&nbsp;Hospital: 
		<%strsql ="Select * from TBL_unidades Where cd_codigo='"&cd_unidade&"'"
			Set rs_und = dbconn.execute(strsql)
			if not rs_und.EOF Then
				nm_unidade = rs_und("nm_unidade")
			end if
			
			rs_und.close
			set rs_und = nothing%>			
		<td><b><%=nm_unidade%></b></td>
	</tr>	
	<tr bgcolor="#e0e0e0">
		<td align="left">&nbsp;Nome:</td>
		<td align="left" colspan="2">&nbsp;<b><%=nm_nome%></b></td>
	</tr>
	<tr>
		<td align="left">&nbsp;Registro</td>
		<td align="left" colspan="2">&nbsp;<b><%=cd_registro%></b></td>
	</tr>	
	<tr bgcolor="#e0e0e0">
		<td>&nbsp;Data</td>
		<td colspan="2">&nbsp;<b><%=dt_procedimento%></b></td>
	</tr>
	<tr>
		<td>&nbsp;CRM/Médico</td>
		<td colspan="3">&nbsp;<b><%'=cd_med%> 
						<%strsql ="Select * from TBL_medicos where a055_codmed = '"&cd_med&"'"
					  		Set rs_med = dbconn.execute(strsql)
							Do while Not rs_med.EOF%>
							<%=rs_med("a055_crmmed")%> - 
							<%=rs_med("a055_desmed")%>
							<%rs_med.movenext
					 		Loop
							
							rs_med.close
							set rs_med = nothing%></b></td>
	</tr>
	<tr bgcolor="#e0e0e0"><td colspan="3">&nbsp;</td></tr>	
	<tr>
		<td colspan="3" align="center"><br><img src="../../imagens/ic_del.gif" onclick="javascript:JsProcDelete('<%=cd_codigo%>','<%=cd_protocolo%>','<%=cd_unidade%>','<%=cd_digito%>')" type="button" value="A" title="Apagar"></td>
	</tr>
	<tr>
		<td><img src="../../imagens/px.gif" alt="" width="100" height="1" border="0"></td>
		<td><img src="../../imagens/px.gif" alt="" width="165" height="1" border="0"></td>
		<td><img src="../../imagens/px.gif" alt="" width="140" height="1" border="0"></td>
	</tr>
	</table> 
<SCRIPT LANGUAGE="javascript">
	function JsProcDelete(cod1,cod2,cod3,cod4)
	{
	  if (confirm ("Tem certeza que deseja excluir este registro?"))
		  {
			document.location.href('../acoes/protocolo_acao.asp?cd_codigo='+cod1+'&cd_protocolo='+cod2+'&cd_unidade='+cod3+'&cd_digito='+cod4+'&acao=excluir&obj_excl=1');
		  }
	}
</script>               
</body>
</html>