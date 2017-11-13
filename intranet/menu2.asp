<%
sessao = strsession

if sessao = "" Then
sessao = "2" 
End if

sair = request("sair")

%>
<!--body-->
<ul id="nav">
  <li><a href="#">Home</a></li>
  <li><a href="#">About</a>
 
    <ul>
      <li><a href="#">History</a></li>
      <li><a href="#">Team</a></li>
      <li><a href="#">Offices</a></li>
    </ul>
  </li>
  <li><a href="#">Services</a>
 
    <ul>
      <li><a href="#">Web Design</a></li>
      <li><a href="#">Internet Marketing</a></li>
      <li><a href="#">Hosting</a>
      	<ul>
      		<li><a href="#">Dedicated</a></li>
      		<li><a href="#">Virtual</a></li>
 
      		<li><a href="#">Shared</a></li>
      		<li><a href="#">Managed</a></li>
      	</ul>
      </li>
      <li><a href="#">Domain Names</a></li>
      <li><a href="#">Broadband</a></li>
    </ul>
 
  </li>
  <li><a href="#">Contact Us</a>
    <ul>
      <li><a href="#">United Kingdom</a></li>
      <li><a href="#">France</a></li>
      <li><a href="#">USA</a></li>
      <li><a href="#">Australia</a></li>
 
    </ul>
  </li>
</ul>

<br>
<br>
<br>







&nbsp;<a href="default.asp?sair=1">Sair</a>
<br>&nbsp;
<!--img src="imagens/blackdot.gif" alt="" width="180" height="1" border="0"-->
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