<!--#include file="../includes/util.asp"-->
<!--#include file="../includes/inc_open_connection.asp"-->

<script language="JavaScript" type="text/javascript" src="js/richtext.js"></script>
<SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript" SRC="js/mascarainput.js"></SCRIPT>

<%
cd_codigo = request("cd_codigo")
campo = request("campo")
list = request("list")
action = request("action")
acao = "inserir"
janela = request("janela")
%>
<table width="580" border="0" cellspacing="1" cellpadding="0">
<tr><td class="txt_cinza">&nbsp;</td></tr>

<table>	   
	<tr>		
		<td class="txt_cinza">
		<b>Admin &raquo; - <font color="red">Cadastro de unidades</font></b></td>
	</tr>
	<tr>
		<td valign="top" align=center>
			<table width="400" border="0" cellspacing="2" cellpadding="0" class="textopadrao">
				<tr>			
					
					<form method="post" action="../adm/adm2_acao.asp" name="form"  id="form" onsubmit="return validar_cad(document.form);">
					<input type="hidden" name="acao" value="<%=acao%>">
					<input type="hidden" name="tipo" value="unidade">
					<input type="hidden" name="janela" value="<%=janela%>">
							<tr bgcolor="#b4b4b4"><td colspan="5" align="center" class="textopadrao"></td></tr>
							<tr bgcolor="#f5f5f5">
								<td  align="left">Sigla</td>
								<td><input type="text" name="nm_sigla" size="10" maxlength="50" class="inputs" value=""></td>
							</tr>
							</tr>
							<tr bgcolor="#f5f5f5">	
								<td  align="left">Unidade</td>
								<td><input type="text" name="nm_unidade" size="45" maxlength="50" class="inputs" value=""></td>
								
							</tr>
							<tr><td colspan="5" align="center">
									<input type="submit" value="inserir unidade."><br><br>
									
									<%=action%><br>
									<br>
							</td></tr>
							</form>
							<tr>
								<td>&nbsp;</td>
								<%
								selecao = " top 100 percent * "
								tabela = " TBL_unidades"
								condicoes = ""
								criterios = " Order by nm_sigla"'nm_unidade"
								
								strsql ="up_listagem @selecao='"&selecao&"', @tabela='"&tabela&"', @condicoes='"&condicoes&"', @criterios='"&criterios&"'"     '"TBL_equipamento"
							  	Set rs_unid = dbconn.execute(strsql)%>
								<td colspan="2">** Listagem de Unidades **<br><select name="nm_unidade" size="10">
								<%Do while Not rs_unid.EOF%>
								<option value="<%=rs_unid("cd_codigo")%>"><%=rs_unid("nm_sigla")%> &nbsp; - &nbsp; <%=rs_unid("nm_unidade")%></option>	<%rs_unid.movenext
							  	Loop%>
								</td>
							</tr>
							
						</table>						

						
						 <SCRIPT LANGUAGE="javascript">
 {
    MaskInput(document.form.dt_data, "99/99/9999"); 
    MaskInput(document.form.dt_encaminhado, "99/99/9999");
    MaskInput(document.form.ret_data, "99/99/9999");
    MaskInput(document.form.dt_prev, "99/99/9999");
 }
</SCRIPT>