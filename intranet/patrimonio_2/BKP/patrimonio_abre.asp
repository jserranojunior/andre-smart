<!--#include file="../includes/util.asp"-->
<%
ok = request("ok")

origem = zero(1)
dt_gravacao = right(year(now),2)&zero(month(now))&zero(day(now))&zero(hour(now))&zero(minute(now))
usuario =  session("cd_codigo")
nome_arquivo = request("nome_arquivo")
nome_gravacao = origem&"_"&dt_gravacao&"_"&zero(usuario)&"-"&nome_arquivo
grava_txt = request("conteudo")%>










<%if ok = "" Then%>
<body>
<table border="1">
	<form action="patrimonio_grava.asp">
	<input type="hidden" name="conteudo" value="<%=grava_txt%>">
	<tr><td colspan="2">Abrir planilha salva</td></tr>
	<tr>
		<td>Escolha a planilha</td>
		<td><input type="text" name="nome_arquivo" size="30" maxlength="30"></td>
	</tr>
	<tr><td colspan="2"><input type="submit" name="ok" value="salva"></td></tr>
	</form>
	
	
	
	<tr>
	<td>Link&nbsp;</td>
	<td><%nome_pasta = "E:\Servidor\Desenvolvimento\producao\intranet\gravados"
			set FSO = server.createObject("Scripting.FileSystemObject") 
			Set pasta = FSO.GetFolder(nome_pasta) 
			Set arquivos = pasta.Files %>
			<select name="link" size="1" class="inputs">
				<option value="" Selected></option>
				<%if link = nome_arquivo then%><%link_ck="selected"%><%else%><%link_ck=""%><%end if%>
				<option value="#" <%=link_ck%>>#</option>
				<%for each nome_arquivo in arquivos
				nome_arquivo = Replace(nome_arquivo,"E:\Servidor\Desenvolvimento\producao\intranet\gravados","")
				tam_nome_arq = len(nome_arquivo)
				if link = nome_arquivo then%><%link_ck="selected"%><%else%><%link_ck=""%><%end if%>
				<option value="<%=nome_arquivo%>" <%=link_ck%>><%'=cd_menu_2&"."%><%=mid(nome_arquivo,"19",int(tam_nome_arq)-22)%></option>
				<%next%>
			</select>
	</td>
</tr>
</table>

<%else
Set FSO = Server.CreateObject("Scripting.FileSystemObject")


caminho = Server.MapPath("../gravados/"&nome_gravacao&".txt") 'especifique aqui o caminho onde ficará/está o TXT
Set GRAVAR = FSO.CreateTextFile(caminho,true)
'Foi criado o objeto e logo após busca o txt em caminho para gravar, se não achar, vai cria-lo (por causa da marcação TRUE)

gravar.write (request("conteudo"))
gravar.close
response.write "GRAVADO!"&nome_gravacao&""
'apos abrir o TXT, gravará a linha com o texto "TESTE DE GRAVAÇÃO" a confirmação no cliente aparecerá como "GRAVADO"
%>

<!--body onload="window.close();"-->
<%end if%>
</body>
