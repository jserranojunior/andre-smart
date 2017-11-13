<%@ Language=VBScript %>
<%

'Set Upload = Server.CreateObject("Persits.Upload.1")

'Count = Upload.Save("E:\users\vdlap.com.br\website\foto\temp\")
%>

<!--#include file="../../includes/inc_open_connection.asp"-->

<%
strcod    = request("cod")
stracao   = request("acao")

strcd_matricula = request.form("cd_matricula")
strnome =  request.form("nm_nome")
strnm_sobrenome =  request.form("nm_sobrenome")
strfoto = request.form("foto_h")
Strdt_contratado =  request.form("dt_contratado")
Strdt_demissao =  request.form("dt_demissao")
strdt_nascimento = request.form("dt_nascimento")
strnm_email =  request.form("nm_email")
strnm_rg =  request.form("nm_rg")
strnm_cpf =  request.form("nm_cpf")
strmae = request.form("nm_mae")
strpai = request.form("nm_pai")
strnm_ddd = request.form("nm_ddd")
strnm_fone =  request.form("nm_fone")
strnm_ddd_cel = request.form("nm_ddd_cel")
strnm_cel =  request.form("nm_cel")
strocontato = request.form("nm_o_contato")
strnm_endereco = request.form("nm_endereco")
strnr_numero =  request.form("nr_numero")
strnm_complemento = request.form("nm_complemento")
strnm_cidade =  request.form("nm_cidade")
strnm_estado =  request.form("nm_estado")
strnm_cep = request.form("nm_cep")
strcd_funcao = request.form("cd_funcao")
strstatus =  request.form("cd_status")
strcd_unidades = request.form("cd_unidades")

dt_ano = YEAR(Now)
dt_mes = Month(Now)
dt_dia = Day(now)

if Strdt_demissao = "" AND strstatus <> 4 Then
Strdt_demissao = "12/30/3000"
End if

If strstatus = "4" AND Strdt_demissao = "" Then
'Strdt_demissao = FormatDateTime(now,2)

	if dt_mes < 10 Then
	dt_mes = 0 & dt_mes
	End if

	Strdt_demissao = dt_ano & dt_mes

Elseif strstatus <> "" AND Strdt_demissao <> "" Then

	demissao_ano = right(strdt_demissao,4)
	demissao_mes = left(strdt_demissao,2)

	strdt_demissao = demissao_ano&demissao_mes'Strdt_demissao
End If

'If Count > 0 Then
'	For Each File in Upload.Files	
'		Upload.OverwriteFiles = false 
'	if Upload.FileExists("E:\users\vdlap.com.br\website\foto\funcionarios\"&Trim(Replace(strnome," ",""))&"_"&Replace(Date(),"/","_")&".jpg") then
'		Upload.DeleteFile "E:\users\vdlap.com.br\website\foto\funcionarios\"&Trim(Replace(strnome," ",""))&"_"&Replace(Date(),"/","_")&".jpg" 
'	End if
'		Upload.MoveFile ""&File.Path&"", "E:\users\vdlap.com.br\website\foto\funcionarios\"&Trim(Replace(strnome," ",""))&"_"&Replace(Date(),"/","_")&".jpg"
	
'		If File.Name = "foto" then
'			strfoto = ""&Trim(Replace(strnome," ",""))&"_"&Replace(Date(),"/","_")&".jpg"
'		End if
'	Next
'End If





If stracao = "insert" Then
	xsql = "up_Funcionario_insert @cd_matricula='"&strcd_matricula&"',@nm_nome='"&strnome&"',@nm_sobrenome='"&strnm_sobrenome&"', @nm_foto='"&strfoto&"', @dt_contratado='"&Month(Strdt_contratado)&"-"&day(Strdt_contratado)&"-"&year(Strdt_contratado)&"', @dt_nascimento='"&Month(strdt_nascimento)&"-"&day(strdt_nascimento)&"-"&year(strdt_nascimento)&"', @nm_email='"&strnm_email&"', @nm_rg='"&strnm_rg&"', @nm_cpf='"&strnm_cpf&"', @nm_mae='"&strmae&"', @nm_pai='"&strpai&"', @nm_ddd='"&strnm_ddd&"', @nm_fone='"&strnm_fone&"', @nm_ddd_cel='"&strnm_ddd_cel&"', @nm_o_contato='"&strocontato&"',@nm_cel='"&strnm_cel&"', @nm_endereco='"&strnm_endereco&"', @nr_numero='"&strnr_numero&"', @nm_complemento='"&strnm_complemento&"', @nm_cidade='"&strnm_cidade&"', @nm_estado='"&strnm_estado&"',@nm_cep='"&strnm_cep&"', @cd_funcao='"&strcd_funcao&"', @cd_status='"&strstatus&"', @cd_unidades='"&strcd_unidades&"'"
	dbconn.execute(xsql)

	strmsg = "Colaborador Cadastrado com sucesso!"

ElseIf stracao = "altera" Then
	xsql = "up_Funcionario_update @cd_matricula='"&strcd_matricula&"',@nm_nome='"&strnome&"',@nm_sobrenome='"&strnm_sobrenome&"', @nm_foto='"&strfoto&"', @dt_contratado='"&Month(Strdt_contratado)&"-"&day(Strdt_contratado)&"-"&year(Strdt_contratado)&"', @dt_demissao='"&Strdt_demissao&"', @dt_nascimento='"&Month(strdt_nascimento)&"-"&day(strdt_nascimento)&"-"&year(strdt_nascimento)&"', @nm_email='"&strnm_email&"', @nm_rg='"&strnm_rg&"', @nm_cpf='"&strnm_cpf&"', @nm_mae='"&strmae&"', @nm_pai='"&strpai&"',@nm_ddd='"&strnm_ddd&"', @nm_fone='"&strnm_fone&"', @nm_ddd_cel='"&strnm_ddd_cel&"', @nm_cel='"&strnm_cel&"', @nm_o_contato='"&strocontato&"', @nm_endereco='"&strnm_endereco&"', @nr_numero='"&strnr_numero&"', @nm_complemento='"&strnm_complemento&"', @nm_cidade='"&strnm_cidade&"', @nm_estado='"&strnm_estado&"',@nm_cep='"&strnm_cep&"',@cd_funcao='"&strcd_funcao&"',@cd_status='"&strstatus&"',@cd_unidades='"&strcd_unidades&"', @cd_codigo='"&strcod&"'"
	dbconn.execute(xsql)

	strmsg = "Colaborador Alterado com sucesso!"

Elseif stracao = "excluir" Then
    xsql= "up_Funcionario_delete @cd_codigo='"&strcod&"'"
	dbconn.execute(xsql)

	strmsg = "cooperado excluido com sucesso!!!"
	response.write("exluiiu")

Else
response.write("erro")
End if





'Set Upload = Nothing 


'response.redirect "../../cooperados_cad.asp?msg="&strmsg&""
response.redirect "../../coop_cooperados_lista.asp?msg="&strmsg&""
%>

   

   
<body>
<%=strcod%><br>
<%=stracao%><br>
<%=strcd_matricula%><br>
<%=strnome%><br>
<%=strnm_sobrenome%><br>
<%=Strdt_contratado%><br>
<%=strmsg%><br>
<%=strfoto%><br>


</body>