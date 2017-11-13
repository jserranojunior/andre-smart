<!--#include file="includes/util.asp"-->
<!--#include file="includes/inc_open_connection.asp"-->
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
		<b>Admin &raquo; - <font color="red">Cadastro de assistências técnicas</font></b></td>
	</tr>
	<tr>
		<td valign="top" align=center>
			<table width="400" border="1" cellspacing="2" cellpadding="0" class="textopadrao">
				<tr>			
					
					<form method="post" action="admin_acao.asp" name="form"  id="form" onsubmit="return validar_cad(document.form);">
					<input type="hidden" name="acao" value="<%=acao%>">
					<input type="hidden" name="tipo" value="forn">
					<input type="hidden" name="cd_codigo" value="<%=codigo%>">
							<tr bgcolor="#b4b4b4"><td colspan="5" align="center" class="textopadrao"></td></tr>
							<tr bgcolor="#f5f5f5">	
								<td  align="left">Assistência</td>
								<td><input type="text" name="nm_fornecedor" size="50" maxlength="35" class="inputs" value=""></td>
							</tr>
							<tr bgcolor="#f5f5f5">
								<td  align="left">1º Contato</td>
								<td><input type="text" name="nm_contato" size="40" maxlength="35" class="inputs" value=""></td>
							</tr>
							<tr bgcolor="#f5f5f5">
								<td  align="left">2º Contato</td>
								<td><input type="text" name="nm_contato2" size="40" maxlength="35" class="inputs" value=""></td>
							</tr>	
							<tr bgcolor="#f5f5f5">
								<td  align="left">Endereço</td>
								<td><input type="text" name="nm_endereco" size="40" maxlength="35" class="inputs" value=""></td>
							</tr>	
							<tr bgcolor="#f5f5f5">
								<td  align="left">Telefone</td>
								<td><input type="text" name="nm_telefone_cod1" size="4" maxlength="5" class="inputs" value="">
									<input type="text" name="nm_telefone1" size="10" maxlength="9" class="inputs" value=""></td>
							</tr>
							<tr bgcolor="#f5f5f5">
								<td  align="left">Telefone</td>
								<td><input type="text" name="nm_telefone_cod2" size="4" maxlength="5" class="inputs" value="">
									<input type="text" name="nm_telefone2" size="10" maxlength="9" class="inputs" value=""></td>
							
							</tr>
							<tr><td colspan="5" align="center">
									<input type="submit" value="inserir assistência."><br><br>
									
									<%=action%><br>
									<br>
									<br>
							</td></tr>
							<tr>
								<td>&nbsp;</td>
								<%
								selecao = " top 100 percent * "
								tabela = " TBL_fornecedor"
								condicoes = ""
								criterios = " Order by nm_fornecedor"
								
								strsql ="up_listagem @selecao='"&selecao&"', @tabela='"&tabela&"', @condicoes='"&condicoes&"', @criterios='"&criterios&"'"
							  	Set rs_forn = dbconn.execute(strsql)%>
								<td colspan="2" align="center">** Lista **<br><select name="cd_marca" size="20">
								<%Do while Not rs_forn.EOF%>
								<option value="<%=rs_forn("cd_codigo")%>"><%=rs_forn("nm_fornecedor")%></option>	<%rs_forn.movenext
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