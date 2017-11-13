<!--#include file="../includes/inc_open_connection.asp"-->

<br>
Lista Duplicados
<br>
<%						
strsql = "SELECT top 3000 cd_cod_protocolo, COUNT(*) AS duplos FROM TBL_Protocolo_wait GROUP BY cd_cod_protocolo HAVING (COUNT(*) > 1) ORDER BY cd_cod_protocolo desc"
Set rs = dbconn.execute(strsql)
while not rs.EOF

cd_protocolo = rs("cd_cod_protocolo")

response.write(cd_protocolo)

strsql_2 = "DELETE TOP (1) FROM TBL_Protocolo_wait WHERE  (cd_cod_protocolo = "&cd_protocolo&") "
Set rs_2 = dbconn.execute(strsql_2)



response.write(" --------  DELETE TOP (1) FROM TBL_Protocolo_wait WHERE  (cd_cod_protocolo = "&cd_protocolo&") <br>")


rs.movenext
wend
							
cd_protocolo = ""						
							
							
							
							%>
							
							
							
fim								