
<%
sessao = strsession

if sessao = "" Then
sessao = "2" 
End if

sair = request("sair")

%>
<!--Informações apenas para os desenvolvedores
Desenvolvedores - 2
- Marcelo
- André

Administração - 1
- Paul
- Luiz

Supervisão - 3
- Sandra

Agenda - 4
- Marli

Básico - 5
- Paula
//-->




<LINK href="css/estilo.css " type=text/css rel=stylesheet >
			<script type="text/javascript" src="js\menu.js"></script>
<table border="1" cellspacing="0" cellpadding="0" class="txt_menu">
     
	 <tr><td height="11" colspan="2">
	 <ul id="nav">
		    <li><a href="#">Home</a></li> 
		    <li><a href="#">Controle</a> 
					<ul> 
		        		<li><a href="#">Produtividade</a></li>
								<!--ul> Sub-sub menu
									<li><a href="#">teste</a></li>
								</ul--> 
		      			<li><a href="#">Benefícios</a></li> 
		    		</ul> 
		    </li> 
		    <li><a href="#">Relatórios</a> 
					<ul> 
				        <li><a href="#">Cooperativa Final</a></li> 
				    </ul> 
		    </li>
		    <li><a href="#">Manutenção</a> 
					<ul> 
				        <li><a href="#">Manutenção</a></li> 
				        <li><a href="#">Cadastramento</a></li> 
								<ul>
									<li><a href="#">Assistências</a></li>
									<li><a href="#">Equipamentos</a></li>
									<li><a href="#">Especialidades</a></li>
									<li><a href="#">Marcass</a></li>
								</ul>
				    </ul> 
		    </li> 
</ul>
	 
	 
	 </td></tr>
	 
	<tr bgcolor="#868686">
		<td>&nbsp;</td>
		<td>&nbsp;</td></td></tr>
	<tr bgcolor="#868686">

		<td class="txt_menu" colspan="2"><a href="default.asp?sair=1">Sair</a></td></tr>
<!--Mostra qual é o arquivo//-->
<%if sessao = "2" Then%>
<font size="-1"><div align="center"><font color=#b4b4b4>
menu.asp</font></div></font>
<%End If%>		
