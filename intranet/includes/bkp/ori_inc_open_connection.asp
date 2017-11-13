<%

strConnString = "File Name=E:\users\vdlap.com.br\website\includes\dbConn_vdlap.udl"


Set dbConn = Server.CreateObject("ADODB.Connection")

dbConn.Open strConnString
%>