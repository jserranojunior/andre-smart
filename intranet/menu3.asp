<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="pt-br" xml:lang="pt-br">
<head>
<meta name="keywords" content="tutorial,acessibilidade,css,css menu,estilo,folhas estilo,layout css,layout sites,menu css,paginas web,tutorial css,web design,web standards,webdesign,tableless" />
<meta name="description" content=" Tutoriais, macetes, dicas sobre uso das CSS para projetar sites." />
<meta name="author" content="Nick Rigby" />
<meta name="MSSmartTagsPreventParsing" content="true" />
<meta http-equiv="imagetoolbar" content="no" />
<meta http-equiv="Pragma" content="no-cache" />
<meta name="robots" content="all" />
<meta name="language" content="pt-br" />
<meta name="ICBM" content="-22.975781,-43.193082" />
<meta name="DC.title" content="CSS para Web Design" />
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Horizontal Drop Down Menus</title>

</head>
<style type="text/css" media="screen">
<!--
ul {margin: 0;
	padding: 0;
	list-style: none;
	width: 150px;}	
ul li {position: relative;}
li ul {position: absolute;
	left: 149px;
	top: 0;
	display: none;}
ul li a {display: block;
	text-decoration: none;
	color: #777;
	background: #fff;
	padding: 5px;
	border: 1px solid #ccc;
	border-bottom: 0;}
		/* Fix IE. Hide from IE Mac \*/
		* html ul li { float: left; }
		* html ul li a { height: 1%; }
		/* End */
ul {	margin: 0;
	padding: 0;
	list-style: none;
	width: 150px;
	border-bottom: 1px solid #ccc;}
li:hover ul {display: block;}


-->	
</style>
<body> 
<ul> 
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
      <li><a href="#">Hosting</a></li> 
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
</body>
</html>
