<!--#include file="../includes/util.asp"-->
<!--#include file="../includes/inc_open_connection.asp"-->

<script language="JavaScript" type="text/javascript" src="js/richtext.js"></script>
<SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript" SRC="js/mascarainput.js"></SCRIPT>
<%
Function prod_horas(min) 'funcionarios
	Hora = Int(min/60)   '60 minutos
	Minutos = min Mod 60 '60 segundos
	NovaHora = Right(Hora,4) & ":" & Right("0" & Minutos,2)
	prod_horas = novahora
	End Function
%>
		<table width="320" border="0" cellspacing="2" cellpadding="0" class="textopadrao">   
			<tr>		
				<td class="txt_cinza" colspan=5>
				<b>RELATÓRIOS &raquo; - <font color="red">Ajuste de produtividade</font></b></td>
			</tr>
			<tr bgcolor="#b4b4b4"><td colspan="5" align="center" class="textopadrao"></td></tr>
			<tr>
				<td colspan=3>
				<%ano = request("ano")
				mes = request("mes")
				cd_codigo = request("cd_codigo")
				str_condicoes = " AND cd_codigo="&cd_codigo&""
				
				'*** Lista os registros ***
				xsql = "up_relatorio_cooperativa @mes="&mes&", @ano="&ano&", @str_condicoes='"&str_condicoes&"'"
				Set rs = dbconn.execute(xsql)%>
				<%Do while Not rs.EOF

				cd_codigo = rs("cd_codigo")
				nm_nome = rs("nm_nome")
				nm_sobrenome = rs("nm_sobrenome")
				total_trabalhado = rs("total_trabalhado")
				cd_funcao = rs("cd_funcao")
				cd_status = rs("nm_status")
				cd_unidades = rs("cd_unidades")
				nm_unidade = rs("nm_unidade")
				
				rs.movenext
				Loop%>
				
				<br>
				<b>Codigo:</b> <%=cd_codigo%><br>
				<b>Cooperado:</b> <%=nm_nome%> <%=nm_sobrenome%><br>
				<b>Total Horas:</b> <%=prod_horas(total_trabalhado)%><br>&nbsp;</td>
			</tr>
			<tr bgcolor="#b4b4b4"><td colspan="5" align="center" class="textopadrao"></td></tr>
			<tr>
				<td colspan=5>
				<%xsql = "Select * From TBL_funcoes Where cd_codigo = "&cd_funcao&""
				Set rs_funcoes = dbconn.execute(xsql)%><%nm_funcao = rs_funcoes("nm_funcao")%>
				<b>Função:</b> <%=nm_funcao%><br>
				<%xsql = "Select * From TBL_Status Where cd_codigo = "&cd_status&""
				Set rs_status = dbconn.execute(xsql)%>
				<%nm_status = rs_status("nm_status")%>
				<b>Status:</b> <%=nm_status%><br>
				<b>Unidade:</b> <%=nm_unidade%><br>	
				<br>
				</td>
			</tr>
			<tr bgcolor="#b4b4b4"><td colspan="5" align="center" class="textopadrao"></td></tr>
			<tr bgcolor="#f5f5f5">
				<form method="post" action="../adm/adm2_acao.asp" name="form"  id="form" onsubmit="return validar_cad(document.form);">
			<input type="hidden" name="acao" value="editar">
			<input type="hidden" name="tipo" value="ajuste">
			<input type="hidden" name="cd_codigo" value="<%=cd_codigo%>">
			<input type="hidden" name="mes" value="<%=mes%>">
			<input type="hidden" name="ano" value="<%=ano%>">
				<td  align="left">Nova Função:</td>
				<td><select name="cd_funcao">
					<%xsql = "Select * From TBL_funcoes Order By nm_funcao"
					Set rs_funcoes = dbconn.execute(xsql)%>
					<%do while not rs_funcoes.EOF
					cod_funcao = rs_funcoes("cd_codigo")
					nm_funcao = rs_funcoes("nm_funcao")
					
					if cd_funcao = cod_funcao then%><%sel_funcao = "selected"%><%end if%>
					<option value="<%=cod_funcao%>" <%=sel_funcao%>><%=nm_funcao%></option>
					<%rs_funcoes.movenext
					sel_funcao=""
					Loop%>
					</select></td>
			</tr>
			</tr>
			<tr bgcolor="#f5f5f5">	
				<td  align="left">Novo Status:</td>
				<td><select name="nm_status">
					<%xsql = "Select * From TBL_Status Order By nm_status"
					Set rs_status = dbconn.execute(xsql)%>
					<%do while not rs_status.EOF
					cod_status = rs_status("cd_codigo")
					nm_status = rs_status("nm_status")
					
					if cd_status = cod_status then%><%sel_status = "selected"%><%end if%>
					<option value="<%=cod_status%>" <%=sel_status%>><%=nm_status%></option>
					<%rs_status.movenext
					sel_status=""
					Loop%>
					</select></td>
				
			</tr>
			<tr bgcolor="#f5f5f5">	
				<td  align="left">Nova Unidade:</td>
				<td><select name="cd_unidade">
					<%xsql = "Select * From TBL_Unidades Order By nm_unidade"
					Set rs_unidade = dbconn.execute(xsql)%><%nm_unidade = rs_unidade("nm_unidade")%>
					<%do while not rs_unidade.EOF
					
					cod_unidade = rs_unidade("cd_codigo")
					nm_unidade = rs_unidade("nm_unidade")
					
					if cd_unidades = cod_unidade then%><%sel_unidade = "selected"%><%end if%>
					<option value="<%=cod_unidade%>" <%=sel_unidade%>><%=nm_unidade%></option>
					<%rs_unidade.movenext
					sel_unidade=""
					Loop%>
					</select></td>
			</tr>
			<tr bgcolor="#b4b4b4"><td colspan="5" align="center" class="textopadrao"></td></tr>
			<tr><td colspan="5" align="Left"><br>
			
					<input type="submit" value="Alterar.">
					
					<%=action%>
					<br>
			</td></tr>
			
			
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