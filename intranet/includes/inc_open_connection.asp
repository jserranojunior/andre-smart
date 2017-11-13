<%
verip = request.servervariables("REMOTE_ADDR")

'strConnString = "File Name=C:\inetpub\wwwroot\producao\includes\dbConn_vdlapL.udl"
'strConnString = "File Name=C:\inetpub\wwwroot\producao\includes\dbConn_vdlapL.udl"

	strConnString = "File Name=C:\inetpub\wwwroot\producao\intranet\includes\bd\dbConn_vdlap.udl"
	'response.write(verip)

Set dbConn = Server.CreateObject("ADODB.Connection")

dbConn.Open strConnString
%>