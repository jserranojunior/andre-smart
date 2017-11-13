<!--html>
<head>
	<title>Untitled</title>
</head-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="pt-br" xml:lang="pt-br">

<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
 <!--#include file="includes/util.asp"-->
<!--#include file="includes/inc_open_connection.asp"-->
<script type="text/javascript"> 
function IEHoverPseudo() {
 
	var navItems = document.getElementById("primary-nav").getElementsByTagName("li");
	
	for (var i=0; i<navItems.length; i++) {
		if(navItems[i].className == "menuparent") {
			navItems[i].onmouseover=function() { this.className += " over"; }
			navItems[i].onmouseout=function() { this.className = "menuparent"; }
		}
	}
 
}
window.onload = IEHoverPseudo;
</script>


<script type="text/javascript" src="js\menu.js"></script>
<script type="text/javascript" src="js\drop_down.js"></script>
<script type="text/javascript"> 
function IEHoverPseudo() {
 
	var navItems = document.getElementById("primary-nav").getElementsByTagName("li");
	
	for (var i=0; i<navItems.length; i++) {
		if(navItems[i].className == "menuparent") {
			navItems[i].onmouseover=function() { this.className += " over"; }
			navItems[i].onmouseout=function() { this.className = "menuparent"; }
		}
	}
 
}
window.onload = IEHoverPseudo;
</script>

<style type="text/css"> 

body {
	font: normal 11px verdana;
	}
 
ul {
	margin: 0;
	padding: 0;
	list-style: none;
	width: 150px; /* Width of Menu Items */
	border-bottom: 1px solid #ccc;
	}
 
ul li {
	position:relative;
	}
 
ul li a {
	display: block;
	text-decoration: none;
	color: #777;
	background: #fff; /* IE6 Bug */
	padding: 5px;
	border: 1px solid #ccc;
	border-bottom: 0;
	}
 
/* Fix IE. Hide from IE Mac \*/
* html ul li { float: left; height: 1%; }
* html ul li a { height: 1%; }
/* End */
 
ul li a:hover { color: #E2144A; background: #f9f9f9; } /* Hover Styles */
 
ul ul {
	position:absolute;
	display:none;
	left: 149px; /* Set 1px less than menu width */
	top: 0;
}
 
li ul li a { padding: 2px 5px; } /* Sub Menu Styles */
 
li:hover ul ul, li.over ul ul { display:none; }
 
li:hover ul, li li:hover ul, li.over ul, li li.over ul { display: block; } /* The magic */
 
</style>


<%
sessao = strsession

if sessao = "" Then
sessao = "2" 
End if

sair = request("sair")
cd_codigo  = 46

%>

<div>
	<ul id="nav">
		    <li><a href="home.asp">Home</a></li>
			<%
			''*** Usuário do sistema ***
			'*** Permissões ***
			xsql_usuario = "SELECT * FROM TBL_ADM_usuario_permissoes Where cd_usuario="&cd_codigo&""
			Set rs_permissao = dbconn.execute(xsql_usuario)
			
			menu1 = rs_permissao("menu_1")
			menu2 = rs_permissao("menu_2")
			menu3 = rs_permissao("menu_3")
			
			'*** Item principal do menu ***
			xsql_m1 ="SELECT * FROM tbl_menu_1 "
			Set rs_menu_1 = dbconn.execute(xsql_m1)
			
			do while not rs_menu_1.EOF
			cd_menu_1 = rs_menu_1("cd_codigo")
			item_menu_1 = rs_menu_1("item_menu")
			link_menu_1 = rs_menu_1("link_menu")
			link_parametro_1 = rs_menu_1("link_parametro")
			
			mprincipal = instr(1,menu1,cd_menu_1,0)
			if mprincipal <> "0" then
			%>
			<li><a href="<%=link_menu_1%><%=link_parametro_1%>"><%=cd_menu_1%> - <%=item_menu_1%></a></li>
				
				<%'*** Item secundário do menu ***
				xsql_m2 ="SELECT * FROM tbl_menu_2 Where menu_principal = "&cd_menu_1&""
				Set rs_menu_2 = dbconn.execute(xsql_m2)%>
				<%if Not rs_menu_2.EOF Then%>
				<ul>
					<%do while not rs_menu_2.EOF
					cd_menu_2 = rs_menu_2("cd_codigo")
					item_menu_2 = rs_menu_2("item_menu")
					link_menu_2 = rs_menu_2("link_menu")
					link_parametro_2 = rs_menu_2("link_parametro")
					
					msecundario = instr(1,menu2,cd_menu_2,0)
					if msecundario <> 0 then%>					
					<li><a href="<%=link_menu_2%><%=link_parametro_2%>"><%=cd_menu_1%>.<%=cd_menu_2%> - <%=item_menu_2%></a> 						 
				       	<%'*** Item terciário do menu ***
						xsql_m3 ="SELECT * FROM tbl_menu_3 Where menu_principal = "&cd_menu_1&" AND menu_medio = "&cd_menu_2&""
						Set rs_menu_3 = dbconn.execute(xsql_m3)%>
						<%if Not rs_menu_3.EOF Then%>
						<ul>
							<%do while not rs_menu_3.EOF
							cd_menu_3 = rs_menu_3("cd_codigo")
							item_menu_3 = rs_menu_3("item_menu")
							link_menu_3 = rs_menu_3("link_menu")
							link_parametro_3 = rs_menu_3("link_parametro")
							
							mterciario = instr(1,menu3,cd_menu_3,0)
							if mterciario <> 0 then%>
							<li><a href="<%=link_menu_3%><%=link_parametro_3%>"><%=cd_menu_1%>.<%=cd_menu_2%>.<%=cd_menu_3%> - <%=item_menu_3%></a></li>							
				    		<%end if
							rs_menu_3.movenext
							loop%>
						</ul>
						<%end if%>
					</li>		 
					<%end if
					rs_menu_2.movenext
					loop%>
				</ul>
				<%end if%>
			</li>
			<%end if
			rs_menu_1.movenext
			loop%>
</div>

<br>
&nbsp;<a href="default.asp?sair=1">Sair</a>
<br>&nbsp;
<%'if int(session_cd_usuario) = int(46) then%>
<!--include file="calendario/calendario.asp"-->
<%'end if%>
<%'=session_cd_usuario%>
<%'**********************************************************************************
'*					Exemplo de item de menu											*
'************************************************************************************
'*					<%if ma1 <> 0 Then% >
'					<li><a href="#">Item principal</a> 
'								<ul> 
'				        			<li><a href="#">Sub-item 1</a></li>
'									<li><a href="#">Sub-item 2</a></li> 
'										<ul>
	'										<li><a href="#">Item final</a></li>
'										</ul>
'				    			</ul>
'					</li>
'					<%end if% >
'***********************************************************************************
%>
<!--/body>
</html-->