<!--#include file="../includes/inc_area_restrita.asp"-->
<!--#include file="../includes/inc_open_connection.asp"-->
<!--#include file="../includes/util.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<%cd_codigo_usr = request("cd_codigo")
msg = request("msg")
%>

<html>

<%if cd_codigo_usr <> "" Then
	xsql ="up_adm_lista @cd_codigo='"&cd_codigo_usr&"'"
	Set rs = dbconn.execute(xsql)
%>
	
<head>
	<title>Untitled</title>
</head>
<SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript" SRC="js/formValidator.js"></SCRIPT>
<body>
 <SCRIPT LANGUAGE="javascript">
    {       
       objVld=getValidator('form');       
	   objVld.addRule('nm_nome',['required','texto'],'"Dia Entrada"');

		
   }
</SCRIPT>


<%do while not rs.EOF
cd_codigo = rs("cd_codigo")
nm_nome_usr = rs("nm_nome")
'nm_email = decrypt(rs("nm_email"))
'nm_senha = decrypt(rs("nm_senha"))
nm_email_usr = rs("nm_email")
nm_senha_usr = rs("nm_senha")
cd_grupo_usr = rs("cd_grupo")

rs.movenext
loop
End if
%>

<%if cd_codigo = "" Then
acao = "1"
Else
acao = "2"
End if%>

<table width="400" cellspacing="2" cellpadding="2" border="1" class="textopadrao">
<tr bgcolor="#dfdfdf">
    <td colspan="100%"><b>Administração de usuários &raquo; - <font color="red">Cadastro <%'=acao%></font></b></td>
</tr>
		<form name="form" id="form" action="adm/adm_acao.asp" method="post">
		<input type="hidden" name="acao" value="<%=acao%>">
		<input type="hidden" name="cd_codigo" value="<%=cd_codigo%>">
<tr>
    <td>Nome</td>
    <td><input type="text" name="nm_nome" size="35" maxlength="30" value="<%=nm_nome_usr%>"></td>
</tr>
<tr>
    <td>E-mail</td>
    <td><input type="text" name="nm_email" size="35" maxlength="100" value="<%=nm_email_usr%>"></td>
</tr>
<tr>
    <td>Senha</td>
    <td><input type="password" name="nm_senha" size="35" maxlength="100" value="<%=nm_senha_usr%>"></td>
</tr>
<tr>
    <td>Grupo</td>
    <td><select name="cd_grupo"><!--selected-->
	<%If cd_grupo_usr = "1" Then%><%grp1_sel = "selected"%>
	<%elseif cd_grupo_usr = "2" Then%><%grp2_sel = "selected"%>
	<%elseif cd_grupo_usr = "3" Then%><%grp3_sel = "selected"%>
	<%elseif cd_grupo_usr = "4" Then%><%grp4_sel = "selected"%>
	<%elseif cd_grupo_usr = "5" Then%><%grp5_sel = "selected"%>
	<%end if%>
	
	<option value="">Selecione</option>
	<option value="1" <%=grp1_sel%>>Administração</option>
	<option value="2" <%=grp2_sel%>>Desenvolvedores</option>
	<option value="3" <%=grp3_sel%>>Supervisão</option>
	<option value="4" <%=grp4_sel%>>Agenda</option>
	<option value="5" <%=grp5_sel%>>Básico</option>
	<option value="5" <%=grp6_sel%>>Enfermagem</option>
</select></td>
</tr>
<tr>
    <td colspan="100%"><b>Módulos permitidos</b></td>
</tr>
<tr bgcolor="#dfdfdf">
    <td colspan="100%"><b>** Cooperativa **</b></td>
</tr>
<%
xsql2 = "SELECT * FROM TBL_ADM_modulos WHERE cd_usuario = '"&cd_codigo&"'"
Set rs2 = dbconn.execute(xsql2)

If rs2.EOF Then

Else
'Campos disponíveis
cd_cod_mod = rs2("cd_codigo")
cd_coop_prod = rs2("cd_coop_prod")
cd_coop_ben = rs2("cd_coop_ben")
cd_coop_rel = rs2("cd_coop_rel")
cd_coop_rh = rs2("cd_coop_rh")
cd_coop_fecha = rs2("cd_coop_fecha")

cd_empr_horas_trab = rs2("cd_empr_horas_trab")
cd_empr_ben = rs2("cd_empr_ben")
cd_empr_rel = rs2("cd_empr_rel")
cd_empr_rh = rs2("cd_empr_rh")
cd_empr_fecha = rs2("cd_empr_fecha")
cd_empr_bh = rs2("cd_empr_bh")

cd_manut_lst = rs2("cd_manut_lst")
cd_manut_cad = rs2("cd_manut_cad")
cd_manut_rel = rs2("cd_manut_rel")

cd_prot_dig = rs2("cd_prot_dig")
cd_prot_rel = rs2("cd_prot_rel")
cd_prot_cad = rs2("cd_prot_cad")

cd_folha_pag = rs2("cd_folha_pag")

cd_invent_adm = rs2("cd_invent_adm")

cd_adm_usr = rs2("cd_adm_usr")

End if
'Campos selecionáveis
if cd_coop_prod = "1" Then
	coop1 = "checked"
End if
if cd_coop_ben = "1" Then
	coop2 = "checked"
End if
if cd_coop_rel = "1" Then
	coop3 = "checked"
End if
if cd_coop_rh = "1" Then
	coop4 = "checked"
End if
if cd_coop_fecha = "1" Then
	coop5 = "checked"
End if

if cd_empr_horas_trab = "1" Then
	empr1 = "checked"
End if
if cd_empr_ben = "1" Then
	empr2 = "checked"
End if
if cd_empr_rel = "1" Then
	empr3 = "checked"
End if
if cd_empr_rh = "1" Then
	empr4 = "checked"
End if
if cd_empr_fecha = "1" Then
	empr5 = "checked"
End if
if cd_empr_bh = "1" Then
	empr6 = "checked"
End if

if cd_manut_lst = "1" Then
	manut1 = "checked"
End if
if cd_manut_cad = "1" Then
	manut2 = "checked"
End if
if cd_manut_rel = "1" Then
	manut3 = "checked"
End if

if cd_prot_dig = "1" Then
	prot1 = "checked"
End if
if cd_prot_rel = "1" Then
	prot2 = "checked"
End if
if cd_prot_cad = "1" Then
	prot3 = "checked"
End if

if cd_folha_pag = "1" Then
	folha1 = "checked"
End if

if cd_invent_adm = "1" Then
	invent1 = "checked"
End if

if cd_adm_usr = "1" Then
	adm1 = "checked"
End if
%>
<tr>
    <td colspan="100%"><input type="hidden" name="cd_cod_mod" value="<%=cd_cod_mod%>">
					   <input type="checkbox" name="cd_coop_prod" value="1" <%=coop1%>> Produtividade<br>
					   <input type="checkbox" name="cd_coop_ben" value="1" <%=coop2%>> Benefícios<br>
					   <input type="checkbox" name="cd_coop_rel" value="1" <%=coop3%>> Relatórios<br>
					   <input type="checkbox" name="cd_coop_rh" value="1" <%=coop4%>> Colaboradores<br>
					   <input type="checkbox" name="cd_coop_fecha" value="1" <%=coop5%>> Fecha o mês Atual
	</td>
</tr>
<tr bgcolor="#dfdfdf">
    <td colspan="100%"><b>** Empresa **</b></td>
</tr>
<tr>
    <td colspan="100%"><input type="checkbox" name="cd_empr_horas_trab" value="1" <%=empr1%>> Horas trabalhadas<br>
					   <input type="checkbox" name="cd_empr_ben" value="1" <%=empr2%>> Benefícios<br>
					   <input type="checkbox" name="cd_empr_rel" value="1" <%=empr3%>> Relatórios<br>
  					   <input type="checkbox" name="cd_empr_rh" value="1" <%=empr4%>> Funcionários<br>
					   <input type="checkbox" name="cd_empr_fecha" value="1" <%=empr5%>> Fecha o mês atual<br>
					   <input type="checkbox" name="cd_empr_bh" value="1" <%=empr6%>> Banco de horas
	</td>
</tr>
<tr bgcolor="#dfdfdf">
    <td colspan="100%"><b>** Manutenção **</b></td>
</tr>
<tr>
    <td colspan="100%"><input type="checkbox" name="cd_manut_lst" value="1" <%=manut1%>> Manutenção<br>
					   <input type="checkbox" name="cd_manut_cad" value="1" <%=manut2%>> Cadastramento<br>
					   <input type="checkbox" name="cd_manut_rel" value="1" <%=manut3%>> Relatórios
	</td>
</tr>
<tr bgcolor="#dfdfdf">
    <td colspan="100%"><b>** Protocolos **</b></td>
</tr>
<tr>
    <td colspan="100%"><input type="checkbox" name="cd_prot_dig" value="1" <%=prot1%>> Digitação<br>
					   <input type="checkbox" name="cd_prot_rel" value="1" <%=prot2%>> Relatórios<br>
					   <input type="checkbox" name="cd_prot_cad" value="1" <%=prot3%>> Cadastros
	</td>
</tr>
<tr bgcolor="#dfdfdf">
    <td colspan="100%"><b>** Folha de pagamento **</b></td>
</tr>
<tr>
    <td colspan="100%"><input type="checkbox" name="cd_folha_pag" value="1" <%=folha1%>> Folha de pagamento
	</td>
</tr>
<tr bgcolor="#dfdfdf">
    <td colspan="100%"><b>** Invetário **</b></td>
</tr>
<tr>
    <td colspan="100%"><input type="checkbox" name="cd_invent_adm" value="1" <%=invent1%>> Inventário
	</td>
</tr>
<tr bgcolor="#dfdfdf">
    <td colspan="100%"><b>** ADM **</b></td>
</tr>
<tr>
    <td colspan="100%"><input type="checkbox" name="cd_adm_usr" value="1" <%=adm1%>> Usuários
	</td>
</tr>
<tr>
    <td colspan="100%"><%if msg = "ok" Then%>Cadastramento efetuado com Sucesso!<%end if%></td>
</tr>
<tr>
    <td colspan="100%"><input type="submit" name="Cadastra" value="Cadastra"></td>
</tr>

</form>
</table>

<%'response.write(modulos)%>	
</body>
</html>




