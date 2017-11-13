<!--#include file="includes/util.asp"-->
<!--#include file="includes/inc_open_connection.asp"-->
<%
sair = request("sair")
strip = request.servervariables("REMOTE_ADDR")
strerro = request("erro")
msg = request("msg")

acesso = session("acessopermitido!")

if sair = 1 Then
Session.Abandon
End if

if acesso = "ok" Then
response.redirect("home.asp")
end if%>

<script language=javascript>
function inicio(){
x.email.focus();}

</script>


<HTML>
<HEAD>
<TITLE></TITLE>
<style type="text/css">
	<!--a:link       { color: #000000; text-decoration:none }
		a:visited    { color: #000000; text-decoration:none }
		a:hover      { color: #365efc; text-decoration:underline }
		#divDrag0{position:absolute; left:0; top:0; height:120; width:40; clip:rect(0,120,120,0); cursor:hand;}
	-->
</style>
</HEAD>
<body topmargin=0 leftmargin=0 bottommargin=0 rightmargin=0 onLoad=inicio()>
<table cellspacing="0" cellpadding="0" width=779 height="100%" align=center valign=center bgcolor=#ffffff style="border-right: 1px solid #cccccc;border-left: 1px solid #cccccc;">
<tr>
	<td valign=top><FONT face=tahoma color=#003366 style=font-size:17px><img src="imagens/acs.gif" hspace="1" vspace="1" border="0" align="left"><B>&nbsp;Login na administração VDLAP.com.br</B></FONT><hr size=1 color=#cccccc width="99%" align=left><br>
		<table cellspacing="0" cellpadding="2">
			<tr>
				<td bgcolor=#eeeeee><center><FONT face=tahoma style=font-size:11px><img src=imagens/ssl.gif width=13 height=15 hspace=0 vspace=0 border=0 align=absbottom>&nbsp; Seu IP é <b><%=request.servervariables("REMOTE_ADDR") %></b> e seu HOST é <b><%=Request.ServerVariables("SERVER_NAME")%></b>, são <b>17:25:26</b> em <b>26/10/2006</b> , por medidas de segurança estas informações serão gravadas em nossa base de dados.</center></td></tr></table>
					<script language=javascript src='layer.js'></script></td></tr>
			<tr>
				<td valign=top><div id='masterdiv'>
					<table cellspacing=2 cellpadding=2 align=center style="border-bottom: 1px solid #cccccc;border-left: 1px solid #cccccc;border-top: 1px solid #cccccc;border-right: 1px solid #cccccc"'>
						<tr><form action=acesso.asp method=post name=x>
						<input name=acao type=hidden value=valida>
						<input name=url type=hidden value="">
							<td colspan=2><FONT face=tahoma style=font-size:11px;font-family:tahoma><b>Login no administrador</td></tr>
						<tr>
							<td colspan=2><FONT face=tahoma style=font-size:10px;font-family:tahoma;color:000000>Esta é uma área de acesso restrito<br>
			<%If strerro = 1 then%>
				<font face=verdana  size=2 color=red>Acesso negado!!! <%'=msg%></font>
<%End if%>
	</td></tr>
	<tr>
		<td><FONT face=tahoma style=font-size:11px;font-family:tahoma>Email:</td>
		<td><input name="email" type=text style=font-size:11px;font-family:tahoma></td>
	</tr>
	<tr>
		<td><FONT face=tahoma style=font-size:11px;font-family:tahoma>Senha:</td>
		<td><input name=senha type=password style=font-size:11px;font-family:tahoma></td>
	</tr>
	<tr>
		<td colspan=2></td>
	<tr>
		<td colspan=2><span class="submenu" id="sim" style="display:none">
		<FONT face=tahoma style=font-size:11px;font-family:tahoma>Digite seu email:<br>
		<input name=esqs type=text style=font-size:11px;font-family:tahoma size=30 value='digite seu email' onclick='this.value=""'></span>
		</td>
	</tr>
	<tr>
		<td align=center  colspan=2>
		<input type=submit value=" Entrar " style=font-size:11px;font-family:tahoma> 
		<input type=button value=Cancelar onclick='javascript:parent.location="default.asp"' style=font-size:11px;font-family:tahoma></td>
	</tr></form>
</table>

<%'response.write(session("acessopermitido!"))%><%'=acesso%>
</td></tr>
<tr>
	<td valign=bottom></div></td></tr>
</table></td></tr></table></body></html>

