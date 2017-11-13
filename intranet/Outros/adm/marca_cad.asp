<!--#include file="../includes/util.asp"-->
<!--#include file="../includes/inc_open_connection.asp"-->
<!--include file="includes/inc_area_restrita.asp"-->


<script language="JavaScript" type="text/javascript" src="js/richtext.js"></script>
<SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript" SRC="js/mascarainput.js"></SCRIPT>

<%
cd_codigo = request("cd_codigo")
campo = request("campo")
list = request("list")
action = request("action")
acao = "inserir"
%>
<table width="580" border="0" cellspacing="1" cellpadding="0">
<tr><td class="txt_cinza">&nbsp;</td></tr>

<table>	   
	<tr>		
		<td class="txt_cinza">
		<b>Admin &raquo; - <font color="red">Cadastro de marcas</font></b></td>
	</tr>
	<tr>
		<td valign="top" align=center>
			<table width="400" border="0" cellspacing="2" cellpadding="0" class="textopadrao">
				<tr>			
					
					<form method="post" action="adm/adm2_acao.asp" name="form"  id="form" onsubmit="return validar_cad(document.form);">
					<input type="hidden" name="acao" value="<%=acao%>">
					<input type="hidden" name="tipo" value="marca">
					<input type="hidden" name="cd_codigo" value="<%=codigo%>">
							<tr bgcolor="#b4b4b4"><td colspan="5" align="center" class="textopadrao"></td></tr>
							<tr bgcolor="#f5f5f5">	
								<td  align="left">Marca</td>
								<td><input type="text" name="nm_marca" size="40" maxlength="35" class="inputs" value=""></td>
								
							</tr>
							<tr><td colspan="5" align="center">
									<input type="submit" value="inserir marca."><br><br>
									
									<%=action%><br>
									<br>
									<br>
							</td></tr>
							<tr>
								<td>&nbsp;</td>
								<%
								selecao = " top 100 percent * "
								tabela = " TBL_marca"
								condicoes = ""
								criterios = " Order by nm_marca"
								
								strsql ="up_listagem @selecao='"&selecao&"', @tabela='"&tabela&"', @condicoes='"&condicoes&"', @criterios='"&criterios&"'"
							  	Set rs_marca = dbconn.execute(strsql)%>
								<td colspan="2" align="center">** Lista **<br><select name="cd_marca" size="10">
								<%Do while Not rs_marca.EOF%>
								<option value="<%=rs_marca("cd_codigo")%>"><%=rs_marca("nm_marca")%></option>	<%rs_marca.movenext
							  	Loop%>
								</td>
							</tr>
							
							<!--tr>
								<td colspan="7" bgcolor="#f5f5f5">
								/ <a href="manutencao.asp?cd_codigo=0&list=x">Nova O.S. </a>
							    / <a href="manutencao.asp">Listagem </a> 
								/ <a href="manutencao_pend.asp?sig=vazio&campo=dt_encerramento">Pendentes</a>
								<a href="manutencao.asp">Listagem </a></td>
							</tr-->
							
							</table>						

						</form>
						 <SCRIPT LANGUAGE="javascript">
 {
    MaskInput(document.form.dt_data, "99/99/9999"); 
    MaskInput(document.form.dt_encaminhado, "99/99/9999");
    MaskInput(document.form.ret_data, "99/99/9999");
    MaskInput(document.form.dt_prev, "99/99/9999");
 }
</SCRIPT>