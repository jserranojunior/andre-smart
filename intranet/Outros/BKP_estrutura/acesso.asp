<!--#include file="includes/util.asp"-->
<!--#include file="includes/inc_open_connection.asp"-->
<%

 stremail = request("email")
 strsenha = request("senha")


 if stremail <> "" and  strsenha <> "" Then
 

 strsql =" up_ADM_Usuario_por_email @nm_email='"&stremail&"'"
 Set rs_email = dbconn.execute(strsql)
 
 
 If Not rs_email.eof Then


   If rs_email("nr_qtd_erro") <=3 then

	 xsql = "up_ADM_usuario_valida_email_x_senha @nm_email='"&stremail&"',@nm_senha='"&strsenha&"'"
	 Set rs = dbconn.execute(xsql)


		If Not  rs.eof Then
		cd_codigo = rs("cd_codigo")
		
		session("acessopermitido!") = "ok"
		session("cd_codigo") = cd_codigo
		session("nm_nome") = rs("nm_nome")
		session("dt_ultimo_acesso") = rs("dt_ultimo_acesso") 
		session("nr_qtd_acasso") = rs("nr_qtd_acasso") + 1
		session("cd_grupo") = rs("cd_grupo")
		
		'***** Carrega módulos ****%>
		<!--#include file="modulos.asp"-->
		<%		xsql ="up_ADM_usuario_altera_acasso_por_cd_codigo @cd_codigo="& rs("cd_codigo")&""
		dbconn.execute(xsql)

		rs.close
		Set rs = Nothing

		xsql ="up_ADM_acesso_insert @nm_nome='"&Request.ServerVariables("SERVER_NAME")&"', @cd_usuario="&session("cd_codigo")&",@nm_ip='"&request.servervariables("REMOTE_ADDR")&"', @cd_status=20"
		dbconn.execute(xsql)


		response.redirect "home.asp?cd_codigo="&cd_codigo&""

		Else

			xsql ="up_ADM_usuario_altera_erro_por_email @nm_email='"&stremail&"'"
			dbconn.execute(xsql)

			xsql ="up_ADM_acesso_insert @nm_nome='"&Request.ServerVariables("SERVER_NAME")&"', @cd_usuario='',@nm_ip='"&request.servervariables("REMOTE_ADDR")&"',@nm_email='"&stremail&"', @cd_status=19"
			dbconn.execute(xsql)

			response.redirect "default.asp?erro=1"

		End If	 
	Else

		xsql = "up_ADM_IP_lista_por_nm_nome2 @nm_nome='"&request.servervariables("REMOTE_ADDR")&"'"
		Set rs = dbconn.execute(xsql)

		  If rs.eof Then
			 xsql = "up_ADM_IP_insert_Negra @nm_nome='"&request.servervariables("REMOTE_ADDR")&"'"
			 dbconn.execute(xsql)	
		  else		
			 xsql = "up_ADM_IP_altera_Negra @nm_nome='"&request.servervariables("REMOTE_ADDR")&"'"
			 dbconn.execute(xsql)
		End if

		response.redirect "default.asp?erro=1"
	End If 

	Else
	
		
	response.redirect "default.asp?erro=1"

	End if

Else



response.redirect "default.asp?erro=1"

End if


 %>