<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<%cd_codigo = request("cd_codigo")%>
<html>
<head>
	<title>Insere/Altera Foto</title>
</head>

<!--body onUnload="window.opener.location.reload()"-->
<body onUnload="window.opener.location.reload()"> 
    <form name="frmSend" method="POST" enctype="multipart/form-data" action="../../acoes/upload.asp" onSubmit="return onSubmitForm();">
    Foto: <input name="attach1" type="file" size=35><br>
	<!-- C�d.: --><input type="hidden" name="cd_codigo" value="<%=cd_codigo%>"><!--<br>-->
	<input style="margin-top:4;" type=submit value="Upload">
    </form>
	
</body>
</html>
