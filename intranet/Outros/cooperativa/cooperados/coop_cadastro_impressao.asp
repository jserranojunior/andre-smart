<!--#include file="../../includes/inc_open_connection.asp"-->
<!--include file="../includes/util.asp"-->
<script language="JavaScript" type="text/javascript" src="js/richtext.js"></script>
<script type="text/javascript">
	function printpage() {
	self.print();
	self.close()
	}
</script>
<%   
strcod = request("cod")
strmsg = request("msg")
strpag = request("pag")
strtipo = request("tipo")
strbusca = request("busca")
stracao = request("acao")

If strcod <> "" And strmsg = "" Then
	
xsql ="up_Funcionario_Lista_por_cd_codigo @cd_codigo="&strcod&""
Set rs = dbconn.execute(xsql)

strmatricula = rs("cd_matricula") 
strnome = rs("nm_nome")
strnm_sobrenome = rs("nm_sobrenome")
strfoto = rs("nm_foto")
Strdt_contratado = rs("dt_contratado")
strdt_nascimento = rs("dt_nascimento")
strnm_email = rs("nm_email")
strnm_rg = rs("nm_rg")
strnm_cpf = rs("nm_cpf")
strmae = rs("nm_mae")
strpai = rs("nm_pai")
strnm_ddd = rs("nm_ddd")
strnm_fone = rs("nm_fone")
strnm_ddd_cel = rs("nm_ddd_cel")
strnm_cel = rs("nm_cel")
strocontato = rs("nm_o_contato")
strnm_endereco = rs("nm_endereco")
strnr_numero = rs("nr_numero")
strnm_complemento = rs("nm_complemento")
strnm_cidade = rs("nm_cidade")
strnm_estado = rs("nm_estado")
strnm_cep = rs("nm_cep")
strcd_funcao = rs("cd_funcao")
strstatus = rs("cd_status")
strunidades = rs("cd_unidades")
str_demissao = rs("dt_demissao")

	If Str_demissao = "300012" Then
		Str_demissao = ""
	Elseif str_demissao <> "" Then
	str_demissao_ano = left(str_demissao,4)
	str_demissao_mes = right(str_demissao,2)
	str_demissao = str_demissao_mes &"/"& str_demissao_ano 
	end if
End If
%>

<body align="center">
<!--body align="center"  onLoad="printpage()"-->
<SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript" SRC="js/mascarainput.js"></SCRIPT>
<table width="700" border="0" cellspacing="0" cellpadding="0">
	<tr>			
		<td valign="top" align=center>
			<table width="680" border="1" cellspacing="0" cellpadding="1">
				<tr>
					<td colspan="4">
						<table width="100%" border="0" cellspacing="0" cellpadding="0">
							<tr><%if strfoto = "" Then%><%strfoto = "Vdlap.gif"%><%end if%>
								<td valign="middle">&nbsp;<img src="../../foto/funcionarios/<%=strfoto%>" width="111" border="1" alt="<%=strfoto%>"></td>
								<td class="txt_cinza" valign="middle">&nbsp;<br><br><br><br><b>FICHA CADASTRAL DE COOPERADO<br><br><br><br></td>											
							</tr>
						</table>	
					</td>
				</tr>		
				<tr>
					<td colspan="4">&nbsp;</td></tr>
				<tr>
					
					<td class="txt_cinza">&nbsp;<b>Nome:</b></td>
					<td colspan="3">&nbsp;<%=strnome%> <%=strnm_sobrenome%></td>
					<!--td class="txt_cinza">&nbsp;<b>Data desligamento:</b></td>
					<td>&nbsp;<%=str_demissao%></td-->
				</tr>
				<tr>
					<td class="txt_cinza">&nbsp;<b>Matricula:</b></td>
					<td valign="top">&nbsp;<%=strmatricula%></td>
					<td class="txt_cinza">&nbsp;<b>Data contratação: </b></td>
					<td>&nbsp;<%=Strdt_contratado%></td>
				</tr>	
					
				
				<tr>
					<td class="txt_cinza">&nbsp;<b>Funcao:</b></td>
					<td>&nbsp;<%strsql = "Select * FROM TBL_funcoes Where cd_codigo='"&strcd_funcao&"'"
				  			Set rs_func = dbconn.execute(strsql)
							Do While Not rs_func.eof
							response.write(rs_func("nm_funcao"))
					  		rs_func.MoveNext
					  		loop
							rs_func.close
							Set rs_func = nothing%></td>
					<td class="txt_cinza">&nbsp;<b>Unidades:</b></td>
					<td>&nbsp;<%strsql ="Select * From TBL_unidades where cd_codigo = '"&strunidades&"'"
				  			Set rs_uni = dbconn.execute(strsql)
							Do While Not rs_uni.eof
					  		response.write(rs_uni("nm_Unidade"))
					   		rs_uni.movenext
							loop
							rs_uni.close
							Set rs_uni = nothing%></td>
				</tr>
				<tr>
					<td colspan="4" align="center">&nbsp;<br><b>DADOS PESSOAIS</b><br>&nbsp;</tr>
				<tr>
					<td class="txt_cinza">&nbsp;<b>Data Nasc.: </b></td>
					<td>&nbsp;<%=strdt_nascimento%></td>
					<td class="txt_cinza">&nbsp;<b>Email: </b></td>
					<td>&nbsp;<%=strnm_email%></td>
				</tr>
				<tr>
					<td class="txt_cinza">&nbsp;<b>RG: </b></td>
					<td>&nbsp;<%=strnm_rg%></td>
					<td class="txt_cinza">&nbsp;<b>CPF: </b></td>
					<td>&nbsp;<%=strnm_cpf%></td>
				</tr>
				<tr>
					<td class="txt_cinza" valign="top" rowspan="2">&nbsp;<b>Filiação: </b></td>
					<td colspan="3">&nbsp;<%=strmae%></td>
				</tr>
				<tr>
					<td colspan="3">&nbsp;<%=strpai%></td>
				</tr>
				<tr><td colspan="4">&nbsp;</td></tr>
				<tr>
					<td class="txt_cinza">&nbsp;<b>Endereço:</b></td>
					<td colspan="3">&nbsp;<%=strnm_endereco%>, <%=strnr_numero%></td>
				</tr>
				<tr>
					<td class="txt_cinza">&nbsp;<b>Complemento:</b></td>
					<td colspan="3">&nbsp;<%=strnm_complemento%></td>
				</tr>
				<tr>
					<td class="txt_cinza">&nbsp;<b>Cep:</b></td>
					<td>&nbsp;<%=strnm_cep%></td>
					<td class="txt_cinza">&nbsp;<b>Cidade:</b></td>
					<td>&nbsp;<%=strnm_cidade%> - <%=strnm_estado%></td>
				</tr>
				<tr><td colspan="4">&nbsp;</td></tr>
				<tr>
					<td class="txt_cinza">&nbsp;<b>Fone:</b></td>
					<td>&nbsp;<%=strnm_ddd%> - <%=strnm_fone%></td>
					<td class="txt_cinza">&nbsp;<b>Celular:</b></td>
					<td>&nbsp;<%=strnm_ddd_cel%>-<%=strnm_cel%></td>
				</tr>
				<tr>
					<td class="txt_cinza">&nbsp;<b>Outros contatos:</b></td>
					<td colspan="3">&nbsp;<%=strocontato%></td>
				</tr>
											
				<!--tr>
					<td class="txt_cinza">&nbsp;<b>Status:</b></td>
				<%
				  strsql ="Select * From TBL_status Where cd_codigo="&strstatus&""
				  Set rs_cat = dbconn.execute(strsql)
				%>
					<td>
					  <%Do While Not rs_cat.eof%>
					  <%=rs_cat("nm_status")%>
					  <%rs_cat.movenext
						loop%>
					</td>
				</tr-->						

				
				
										
		<!--/td-->
		</tr>	       
	</table> 
			 

	
		</td>
	</tr>
	                                                  
</table>  

			
			
</body>			
			
			