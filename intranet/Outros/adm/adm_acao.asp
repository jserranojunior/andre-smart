
  <!--#include file="../includes/util.asp"-->
<!--#include file="../includes/inc_open_connection.asp"-->
   <!--#include file="../includes/inc_area_restrita.asp"-->

<%

cd_codigo = request("cd_codigo")
nm_nome = request("nm_nome")
nm_email = request("nm_email")
nm_senha = request("nm_senha")
cd_grupo = request("cd_grupo")	
acao = request("acao")

cd_cod_mod = request("cd_cod_mod")
cd_coop_prod = request("cd_coop_prod")
cd_coop_ben = request("cd_coop_ben")
cd_coop_rel = request("cd_coop_rel")
cd_coop_rh = request("cd_coop_rh")
cd_coop_fecha = request("cd_coop_fecha")

cd_empr_horas_trab = request("cd_empr_horas_trab")
cd_empr_ben = request("cd_empr_ben")
cd_empr_rel = request("cd_empr_rel")
cd_empr_rh = request("cd_empr_rh")
cd_empr_fecha = request("cd_empr_fecha")
cd_empr_bh = request("cd_empr_bh")

cd_manut_lst = request("cd_manut_lst")
cd_manut_cad = request("cd_manut_cad")
cd_manut_rel = request("cd_manut_rel")

cd_prot_dig = request("cd_prot_dig")
cd_prot_rel = request("cd_prot_rel")
cd_prot_cad = request("cd_prot_cad")

cd_folha_pag = request("cd_folha_pag")

cd_invent_adm = request("cd_invent_adm")

cd_adm_usr = request("cd_adm_usr")

'nm_email = encrypt(nm_email)
'nm_senha = encrypt(nm_senha)


'****************************** Cadastra *********************************************************

dt_ultimo_acesso = "01/01/1950"
nr_qtd_erro = "0"

If acao = "1" AND cd_codigo = "" Then
'Insere dados - Usuário
xsql = "up_adm_insert @nm_nome='"&nm_nome&"', @nm_email='"&nm_email&"', @nm_senha='"&nm_senha&"', @cd_grupo='"&cd_grupo&"', @dt_ultimo_acesso='"&dt_ultimo_acesso&"', @nr_qtd_erro='"&nr_qtd_erro&"', @nr_qtd_acasso='"&nr_qtd_acasso&"'"
dbconn.execute(xsql)

'Verifica e pega o código do usuário para inserção dos módulos de permissão
xsql2 = "SELECT * FROM TBL_ADM_usuario WHERE nm_nome='"&nm_nome&"' AND nm_email = '"&nm_email&"'"
Set rs2 = dbconn.execute(xsql2)

cd_cod = rs2("cd_codigo")

'Insere módulos
xsql = "up_adm_modulos_insert @cd_cod='"&cd_cod&"', @cd_coop_prod='"&cd_coop_prod&"',@cd_coop_ben='"&cd_coop_ben&"',@cd_coop_rel='"&cd_coop_rel&"',@cd_coop_rh='"&cd_coop_rh&"',@cd_coop_fecha='"&cd_coop_fecha&"', @cd_empr_horas_trab='"&cd_empr_horas_trab&"',@cd_empr_ben='"&cd_empr_ben&"',@cd_empr_rel='"&cd_empr_rel&"',@cd_empr_rh='"&cd_empr_rh&"',@cd_empr_fecha='"&cd_empr_fecha&"',@cd_empr_bh='"&cd_empr_bh&"',@cd_manut_lst='"&cd_manut_lst&"',@cd_manut_cad='"&cd_manut_cad&"',@cd_manut_rel='"&cd_manut_rel&"',@cd_prot_dig='"&cd_prot_dig&"',@cd_prot_rel='"&cd_prot_rel&"',@cd_prot_cad='"&cd_prot_cad&"', @cd_folha_pag='"&cd_folha_pag&"', @cd_invent_adm='"&cd_invent_adm&"', @cd_adm_usr='"&cd_adm_usr&"'"
dbconn.execute(xsql)

response.redirect("../ADM.asp?cd_codigo=&msg=ok")
'response.redirect("ADM_cadastro.asp?cd_codigo="&cd_codigo&"&msg=ok")
'response.Write("Usuário cadastrado com Sucesso!")

'******************************* Modifica **********************************************
Elseif acao = "2" AND cd_codigo <> "" Then
'Modifica os dados
xsql = "up_adm_update @cd_codigo='"&cd_codigo&"', @nm_senha='"&nm_senha&"', @cd_grupo='"&cd_grupo&"'"
dbconn.execute(xsql)

'Modifica os módulos permitidos
xsql = "up_adm_modulos_update @cd_cod_mod='"&cd_cod_mod&"', @cd_coop_prod='"&cd_coop_prod&"', @cd_coop_ben='"&cd_coop_ben&"', @cd_coop_rel='"&cd_coop_rel&"',@cd_coop_rh='"&cd_coop_rh&"',@cd_coop_fecha='"&cd_coop_fecha&"', @cd_empr_horas_trab='"&cd_empr_horas_trab&"',@cd_empr_ben='"&cd_empr_ben&"',@cd_empr_rel='"&cd_empr_rel&"',@cd_empr_rh='"&cd_empr_rh&"',@cd_empr_fecha='"&cd_empr_fecha&"', @cd_empr_bh='"&cd_empr_bh&"', @cd_manut_lst='"&cd_manut_lst&"',@cd_manut_cad='"&cd_manut_cad&"',@cd_manut_rel='"&cd_manut_rel&"',@cd_prot_dig='"&cd_prot_dig&"',@cd_prot_rel='"&cd_prot_rel&"',@cd_prot_cad='"&cd_prot_cad&"', @cd_folha_pag='"&cd_folha_pag&"', @cd_invent_adm='"&cd_invent_adm&"', @cd_adm_usr='"&cd_adm_usr&"'"
dbconn.execute(xsql)


response.redirect("../ADM.asp?cd_codigo=&msg=ok")
'response.redirect("ADM_cadastro.asp?cd_codigo="&cd_codigo&"&msg=ok")
'response.write("Usuário modificado com sucesso")

'******************************* Exclui *************************************************

Elseif acao = "3" AND cd_codigo <> "" Then
'Apaga Usuário
xsql = "up_adm_excluir @cd_codigo='"&cd_codigo&"'"
action = "apagado com sucesso"
dbconn.execute(xsql)

response.redirect("../adm.asp")


End If

'='"&Month()&"-"&day()&"-"&year()&"'


'response.redirect("manutencao.asp?cd_codigo="&cd_codigo&"&campo="&campo&"")

'response.redirect("manutencao.asp?cd_codigo=&campo=dt_os")
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
	<title>Untitled</title>
</head>

<body>

Action = <%=action%><br>
Codigo - <%=cd_codigo%><br>
Usuário - <%=cd_cod%><br><br>

nome - <%=nm_nome%><br>
email - <%=nm_email%><br>
senha - <%=nm_senha%><br>
grupo - <%=cd_grupo%><br><br>

cod = <%=cd_cod_mod%><br>
Coop1 = <%=cd_coop_prod%><br>
acao = <%=acao%><br>

</body>
</html>
