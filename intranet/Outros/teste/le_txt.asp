<%
Set FSO = Server.CreateObject("Scripting.FileSystemObject")
caminho = Server.Mappath("txt/teste.txt")
Set TXT = FSO.OpenTextFile(caminho)
'cria o objeto, e busca pelo TXT indicado pela variável caminho como acima


'response.write txt.readALL
patrimonio_array = split(txt.readALL,";")
	For Each patrimonio_item_junto In patrimonio_array
		if patrimonio_item_junto <> "" Then
			patrimonio_item_len = len(patrimonio_item_junto)
				separador1 = int(instr(patrimonio_item_junto,","))
				separador2 = int(instr(mid(patrimonio_item_junto,(separador1)+1,patrimonio_item_len),","))
				separador3 = int(instr(mid(patrimonio_item_junto,(separador1+separador2)+1,patrimonio_item_len),","))
				separador4 = int(instr(mid(patrimonio_item_junto,(separador1+separador2+separador3)+1,patrimonio_item_len),","))
				separador5 = int(instr(mid(patrimonio_item_junto,(separador1+separador2+separador3+separador4)+1,patrimonio_item_len),","))
		'		separador6 = int(instr(mid(patrimonio_item_junto,(separador1+separador2+separador3+separador4+separador5)+1,patrimonio_item_len),","))
		'		
				patrimonio_equip =			mid(patrimonio_item_junto,1,separador1-1)
				patrimonio_controle = 		mid(patrimonio_item_junto,separador1+1,separador2-1)
				patrimonio_marca = 			mid(patrimonio_item_junto,separador1+separador2+1,separador3-1)
				patrimonio_modelo = 		mid(patrimonio_item_junto,separador1+separador2+separador3+1,separador4-1)
				patrimonio_periodicidade = 	mid(patrimonio_item_junto,separador1+separador2+separador3+separador4+1,separador5-1)
				patrimonio_data = 			mid(patrimonio_item_junto,separador1+separador2+separador3+separador4+separador5+1,patrimonio_item_len)
				%>

				
				<%=patrimonio_equip&" - "&patrimonio_controle&" - "&patrimonio_marca&" - "&patrimonio_modelo&" - "&patrimonio_periodicidade&" - "&patrimonio_data&"<br>"%>
				<%
		end if
	next


'após abrir o TXT, enviará direto ao cliente todo conteúdo do TXT, neste exemplo, retornará "TESTE DE GRAVAÇÃO" como foi gravado acima
txt.close%>
<%'=txt.readALL%>