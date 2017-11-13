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
		<b>Admin &raquo; - <font color="red">Cadastro de equipamentos</font></b></td>
	</tr>
	<tr>
		<td valign="top" align=center>
			<table width="400" border="0" cellspacing="2" cellpadding="0" class="textopadrao">
				<tr>			
					
					<form method="post" action="../adm/adm2_acao.asp" name="form"  id="form" onsubmit="return validar_cad(document.form);">
					<input type="hidden" name="acao" value="<%=acao%>">
					<input type="hidden" name="tipo" value="equip">
					<input type="hidden" name="cd_codigo" value="<%=codigo%>">
					<input type="hidden" name="janela" value="<%=janela%>">
							<tr bgcolor="#b4b4b4"><td colspan="5" align="center" class="textopadrao"></td></tr>
							<tr bgcolor="#f5f5f5">
								<td  align="left">Cód. fabricante</td>
								<td><input type="text" name="cd_pn" size="10" maxlength="50" class="inputs" value=""> Deixe em branco se não souber.</td>
							</tr>
							</tr>
							<tr bgcolor="#f5f5f5">	
								<td  align="left">Equipamento</td>
								<td><input type="text" name="nm_equipamento" size="45" maxlength="50" class="inputs" value=""></td>
								
							</tr>
							<!--tr bgcolor="#f5f5f5">
								<td  align="left">Marca</td>
								<%strsql ="TBL_Marca"
							  	Set rs_marca = dbconn.execute(strsql)%>
								<td colspan="4"><select name="nm_marca" class="inputs"><option value="">Selecione</option><%Do while Not rs_marca.EOF %><%If marca = rs_marca("nm_marca") Then%><%check2 = "selected"%><%else check2 = ""%><%end If%>
								<option value="<%=rs_marca("cd_codigo")%>" <%=check2%>><%=rs_marca("nm_marca")%></option>	<%rs_marca.movenext
							  	Loop%></td>	
							</tr-->
							<tr><td colspan="5" align="center">
									<input type="submit" value="inserir equip./instr."><br><br>
									
									<%=action%><br>
									<br>
							</td></tr>
							<tr>
								<td>&nbsp;</td>
								<%
								selecao = " top 100 percent * "
								tabela = " TBL_equipamento"
								condicoes = ""
								criterios = " Order by nm_equipamento"
								
								strsql ="up_listagem @selecao='"&selecao&"', @tabela='"&tabela&"', @condicoes='"&condicoes&"', @criterios='"&criterios&"'"     '"TBL_equipamento"
							  	Set rs_equip = dbconn.execute(strsql)%>
								<td colspan="2">** Listagem de equipamentos/Instrumentos **<br><select name="cd_equipamento" size="10">
								<%Do while Not rs_equip.EOF%>
								<option value="<%=rs_equip("cd_codigo")%>"><%=rs_equip("nm_equipamento")%></option>	<%rs_equip.movenext
							  	Loop%>
								</td>
							</tr>
							
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