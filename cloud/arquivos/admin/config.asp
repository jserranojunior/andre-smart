<%

' - Constants ------------------------------------------------
	' For disabling password security set wexPassword = "" 
	'Const wexPassword = "admin"
	Dim wexPassword
	Dim password
	Dim user
	Dim wexRoot
	user = request("usuario")
	password = request("pwd")
	
	
	if user <> "" and  password <> "" Then
		'session("wexPassword") = trim(user&Password)
		session("wexPassword") = trim(user&password)
	else
		
	end if
	
	wexPassword = session("wexPassword")
	' Root folder, it can be a physical or virtual folder: "c:\test", "/test"
	' Beware that you will not have web access(no web browsing of files/folders) if you use a physical folder like "c:\test"
	' If you want to have web access with a physical folder, you should create a virtual folder(IIS alias like "/folder") for that physical folder and use it instead.
	

	'Const wexRoot = "/cloud/arquivos/"
	
	if session("wexPassword") = "sandra242829" then
		wexRoot = "/cloud/arquivos/sandra/"
		'wexRoot = "/arquivos/sandra/"
	elseif session("wexPassword") = "luizlasoares" then
		wexRoot = "/cloud/arquivos/Luiz/"
		'wexRoot = "/arquivos/Luiz/"
	elseif session("wexPassword") = "admin1445224" then
		wexRoot = "/cloud/arquivos/admin/"
		'wexRoot = "/arquivos/admin/"
	else
		wexPassword = "*68495s$bx2klgp9*"
	end if
	
	'response.write(session("wexPassword")&"("&user&":"&password&")")
	
	' Show files and folders that have hidden attribute set?
	Const showHiddenItems = true
	' Calculate total size of the current folder? Disable if it takes long time with huge folders.
	Const calculateTotalSize = true
	' Calculate total sizes of the folders in the listing? Disable if it takes long time with huge folders.
	Const calculateFolderSize = true
	' List of file extensions which are showed with the "T" icon and can be edited by clicking the icon.
	Const editableExtensions = "htm,html,asp,asa,txt,inc,css,aspx,js,vbs,shtm,shtml,xml,xsl,log"
	' List of file extensions which are showed with the "P" icon and can be viewed by clicking the icon.
	Const viewableExtensions = "gif,jpg,jpeg,png,bmp,jpe"
	' Display full physical path of the current folder? Set it to false for hiding your server path structure.
	Const displayPath = true
' ------------------------------------------------------------

' - Variables ------------------------------------------------
	' Set script timeout value to higher values (in seconds) if the script fails when uploading large files
	Server.ScriptTimeout = 300
	' Preferred character set, default value is "ISO-8859-1" (Western European character set)
	' Don't bother to change it unless you are having problems with handling text in your language
	' Character sets and supporting code pages can be found at
	' http://msdn.microsoft.com/library/default.asp?url=/workshop/author/dhtml/reference/charsets/charset4.asp
	Response.CharSet = "ISO-8859-1"
	' Preferred code page, default value is 1252 (Western European codepage)
	' Don't bother to change it unless you are having problems with handling text in your language
	Session.CodePage = 1252
	' Preferred locale identifier, default value is 1033 (English - United States)
	' Locale ID (LCID) Chart can be found at
	' http://msdn.microsoft.com/library/default.asp?url=/library/en-us/script56/html/vsmscLCID.asp
	Session.LCID = 1046
' ------------------------------------------------------------

' - Plugins --------------------------------------------------
	' File transfer plugin
	%><!-- #include file="./plugins/file transfer/WexGeneric.asp" --><%
' ------------------------------------------------------------

%>