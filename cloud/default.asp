<%@ Language=VBScript %>
<%Option Explicit%>
<%
' WebExplorer Free - A powerful asp tool for managing remote files with an user-friendly web interface
' November 2002 - Version 2.0
' Copyright © 2000-2003 by Cem Alacayir

' Copyright notice:
' WebExplorer Free is a freeware with open source code. 
' As long as you keep the copyright message, you can 
' play with the code for your own purposes but you can
' not distribute the altered version. You can not sell 
' WebExplorer Free or use it for commercial purposes.
' Please contact me for commercial uses.

' Errors, comments or suggestions : contact@wexfms.com
' Visit the web page for updates and support : http://www.wexfms.com



' - WebExplorer Free Main ------------------------------------

	Const appName = "Gerenciador de Arquivos"
	Const appVersion = "2.0"
	
	%><!-- #include file="config.asp" --><%

	Dim FSO, re
	Dim scriptName, wexId
	Dim wexMessage, wexRootPath, targetPath
	Dim encoding, codepage, charset

	InitApp()

	' Actions in the popup windows
	Select Case Request.Form("command")
		Case "Edit"
			Editor()
		Case "View"
			Viewer()
		Case "FileDetails", "FolderDetails"
			Details()
		Case "Upload"
			Upload(false)
	End Select
	
	' Actions in the main window
	Select Case Request.Form("command")
		Case "NewFile", "NewFolder"
			CreateItem()
		Case "DeleteFile", "DeleteFolder"
			DeleteItem()
		Case "RenameFile", "RenameFolder"
			RenameItem()
		Case "OpenFolder"
			targetPath = WexMapPath(Request.Form("folder") & Request.Form("parameter"))
		Case "LevelUp"
			targetPath = WexMapPath(FSO.GetParentFolderName(Request.Form("folder")))
		Case "Logout"
			Logout()
	End Select

	List()

	DestroyApp()

' ------------------------------------------------------------

' - WebExplorer Free Functions -------------------------------

	' Initializes some variables, creates instances of some objects and ensures security
	Sub InitApp()
		scriptName = Request.ServerVariables("SCRIPT_NAME")
		wexId = appName & appVersion & "-"

		Response.Buffer = true
		
		If not Secure() Then 
			If Request.Form("popup")="true" or Request.QueryString("popup")="true" Then PopupRelogin() Else Login()
		End If
		
		Set FSO = server.CreateObject ("Scripting.FileSystemObject")
		Set re = new regexp

		wexRootPath = RealizePath(wexRoot)

		encoding = -2 'System default encoding

		' Commands with high priority
		' These commands require to be performed before any Request.Form statement
		Select Case Request.QueryString("precommand")
			Case "ProcessUpload"
				Upload(true)
			Case "Download"
				Download()
			Case "Encoding"
				If Request.QueryString("value")<>"" Then encoding = Int(Request.QueryString("value"))
				If encoding=-1 Then 'Unicode encoding
					codepage = Session.CodePage
					Session.CodePage = 65001
					Response.CharSet = "UTF-8"
				End If
		End Select
		
		targetPath = WexMapPath(Request.Form("folder"))
	End Sub
	
	' Frees the objects and ends the application
	Sub DestroyApp()
		If encoding=-1 Then Session.CodePage = codepage
		Set FSO = Nothing
		Set re = Nothing
		Response.End
	End Sub
	
	' Writes the html header
	Sub HtmlHeader(title)
%>
<!--
	<%=appName%> v<%=appVersion%>
	Copyright © 2000-2003 by Cem Alacayir
	http://www.wexfms.com
-->

<html>
<head>
<title><%=title%></title>
<%HtmlStyle%>
<%HtmlJavaScript%>
</head>
<body>
<%	
	End Sub
	
	' Writes the html footer
	Sub HtmlFooter()
%>	
</body>
</html>
<%
	End Sub

	' Writes the copyright message
	Sub HtmlCopyright()
	%>
		<!--table cellspacing=0 cellpadding=0 border=0 align=center><tr><td>
			<a href="mailto:contact@wexfms.com?subject=WebExplorer Free - Contact" title="Email me for errors, comments or suggestions">Copyright &copy; 2000-2003 by Cem Alacayir</a> - <a target="_blank" href="http://www.wexfms.com" title="Visit the web page for updates and support">Web page</a>
		</td></tr></table-->
	<%
	End Sub
	
	' Writes the stylesheet
	Sub HtmlStyle()
%>
<style>
	BODY
		{BACKGROUND-COLOR: #FFFFFF}
	TD
		{FONT-WEIGHT: normal;
		 FONT-SIZE: 10pt;
		 COLOR: white;
		 FONT-FAMILY: Verdana,Tahoma,Arial,Helvetica}
		 
	.formClass
		{BACKGROUND-COLOR: #000000;
	     FONT-WEIGHT: normal;
	     FONT-SIZE: 10pt;
	     COLOR: #FFFFFF;
	     FONT-FAMILY: Verdana,Tahoma,Arial,Helvetica}
	.lightRow 
		{BACKGROUND-COLOR: #c0c0c0}	
	.darkRow 
		{BACKGROUND-COLOR: #dfdfdf}
	.titleRow 
		{BACKGROUND-COLOR: #808080}
	.loginRow 
		{border: black solid 1px;
		BACKGROUND-COLOR: DodgerBlue}
	.boldText
		{FONT-WEIGHT: bold}
		
	A	{FONT-WEIGHT: normal;
	    COLOR: #000000;
	    TEXT-DECORATION: none}
	A:hover {COLOR: #cc0000;
		    TEXT-DECORATION: none}
	A:visited {TEXT-DECORATION: none}
	A:active {COLOR: #330000;
			  TEXT-DECORATION: none}		  
</style>
<%
	End Sub
	
	' Writes the javascript code
	Sub HtmlJavaScript()
%>
<script language=javascript>
	function Command(cmd, param) {
		var str;
		var someWin;
		switch (cmd) {
			case "NewFile":
				str = prompt("Please enter a name for the new file", "New File");
				if(!str) return;
				else if (!CheckName(str)) {alert("File name can not contain any of the\nfollowing characters: \\ / : * ? \" < > |"); return;}
				document.forms.formBuffer.parameter.value = str;
				break;
			case "NewFolder":
				str = prompt("Please enter a name for the new folder", "New Folder");
				if(!str) return;
				else if (!CheckName(str)) {alert("Folder name can not contain any of the\nfollowing characters: \\ / : * ? \" < > |"); return;}
				document.forms.formBuffer.parameter.value = str;
				break;
			case "Edit":
				str = document.forms.formBuffer.folder.value + param;
				someWin = openWin(cmd + str, "", 600, 440, false, false);
				someWin.focus();
				createPage(someWin,cmd,param);
				someWin = null;
				return;
				break;
			case "View":
				str = document.forms.formBuffer.folder.value + param;
				someWin = openWin(cmd + str, "", 600, 440, false, true);
				someWin.focus();
				createPage(someWin,cmd,param);
				someWin = null;
				return;
				break;
			case "FileDetails":
			case "FolderDetails":
				str = document.forms.formBuffer.folder.value + param;
				someWin = openWin(cmd + str, "", 350, 220, false, false);
				someWin.focus();
				createPage(someWin,cmd,param);
				someWin = null;
				return;
				break;
			case "Upload":
				someWin = openWin(cmd, "", 400, 150, true, false);
				someWin.focus(); 
				createPage(someWin,cmd,param);
				someWin = null;
				return;
				break;
			case "DeleteFolder":
				if (!confirm('Are you sure to delete the folder "' + param + '" and all its contents ?')) return;
				document.forms.formBuffer.parameter.value = param;
				break;
			case "DeleteFile":
				if (!confirm('Are you sure to delete "' + param + '" ?')) return;
				document.forms.formBuffer.parameter.value = param;				
				break;
			case "RenameFile":
				str = prompt("Please enter the new name for the file", param);
				if(!str) return;
				else if (!CheckName(str)) {alert("File name can not contain any of the\nfollowing characters: \\ / : * ? \" < > |"); return;}
				document.forms.formBuffer.parameter.value = param + "|" + str;
				break;
			case "RenameFolder":
				str = prompt("Please enter the new name for the folder", param);
				if(!str) return;
				else if (!CheckName(str)) {alert("Folder name can not contain any of the\nfollowing characters: \\ / : * ? \" < > |"); return;}
				document.forms.formBuffer.parameter.value = param + "|" + str;
				break;
			case "NoWebAccess":
				alert("You don't have web access for this folder so\nweb browsing of files/folders will not be available!");
				return;
				break;
			default:
				document.forms.formBuffer.parameter.value = param;
		}
		document.forms.formBuffer.target = "";
		document.forms.formBuffer.command.value = cmd
		document.forms.formBuffer.submit();	
	}
	
	function Check() {
		if (document.forms.formBuffer.pwd.value == "") {
			alert("You haven't entered the password!"); 
			return false;
		} else return true;
	}

	function openWin(winName, urlLoc, w, h, showStatus, isViewer) {
		l = (screen.availWidth - w)/2;
		t = (screen.availHeight - h)/2;
		features  = "toolbar=no";      // yes|no 
		features += ",location=no";    // yes|no 
		features += ",directories=no"; // yes|no 
		features += ",status=" + (showStatus?"yes":"no");  // yes|no 
		features += ",menubar=no";     // yes|no 
		features += ",scrollbars=" + (isViewer?"yes":"no");   // auto|yes|no 
		features += ",resizable=" + (isViewer?"yes":"no");   // yes|no 
		features += ",dependent";      // close the parent, close the popup, omit if you want otherwise 
		features += ",height=" + h;
		features += ",width=" + w;
		features += ",left=" + l;
		features += ",top=" + t;
		winName = winName.replace(/[^a-z]/gi,"_");
		return window.open(urlLoc,winName,features);
	} 
	
	function createPage (theWin, cmd, param){
		document.forms.formBuffer.target = theWin.name;
		document.forms.formBuffer.command.value = cmd;
		document.forms.formBuffer.parameter.value = param;
		document.forms.formBuffer.popup.value = "true";
		document.forms.formBuffer.submit();
		document.forms.formBuffer.popup.value = "false";
	}

	function EditorCommand (cmd) {
		switch (cmd) {
			case "Info":
				alert(document.forms.formBuffer.info.value.replace(/\|/gi,"\n"));
				break;
			case "Reload":
				document.forms.formBuffer.reset();
				break;
			case "Save":
				document.forms.formBuffer.action += "?precommand=Encoding&value=";
				document.forms.formBuffer.action += document.forms.formBuffer.encoding.options[document.forms.formBuffer.encoding.selectedIndex].value;
				document.forms.formBuffer.subcommand.value = "Save";
				document.forms.formBuffer.submit();
				break;
			case "SaveAs":
				var str, oldname;
				oldname = document.forms.formBuffer.parameter.value;
				str = prompt("Save the file as :", oldname);
				if (!str || str==oldname) return;
				document.forms.formBuffer.action += "?precommand=Encoding&value=";
				document.forms.formBuffer.action += document.forms.formBuffer.encoding.options[document.forms.formBuffer.encoding.selectedIndex].value;
				document.forms.formBuffer.parameter.value = str;
				document.forms.formBuffer.subcommand.value = "SaveAs";
				document.forms.formBuffer.submit();
				break;
			case "Encoding":
				document.forms.formBuffer.action += "?precommand=Encoding&value=";
				document.forms.formBuffer.action += document.forms.formBuffer.encoding.options[document.forms.formBuffer.encoding.selectedIndex].value;
				document.forms.formBuffer.subcommand.value = cmd;
				document.forms.formBuffer.submit();
				break;
		}
	}

	function ViewerCommand (cmd) {
		switch (cmd) {
			case "Info":
				alert(document.forms.formBuffer.info.value.replace(/\|/gi,"\n"));
				break;
			case "Reload":
				document.forms.formBuffer.submit();
				break;
		}
	}

	function Upload() {
		document.forms.formBuffer.submit();
	}
	
	function PopupRelogin() {	
		opener.Command('Refresh');
		window.close();
	}
	
	function CheckName(str) {
		var re;
		re = /[\\/:*?"<>|]/gi;
		if (re.test(str)) return false;	
		else return true;
	}	
</script>
<%
	End Sub

	' Writes file listing of the given folder
	Sub List()
		Dim objFolder, virtual, folder
		Dim item, arr
		Dim rowType
		Dim listed

		HtmlHeader appName
		
		on error resume next
		Set objFolder = FSO.GetFolder(targetPath)
		If err.Number<>0 Then wexMessage = "Error opening folder !"
		
		virtual = VirtualPath(targetPath)
		folder = right(targetPath, len(targetPath)-len(wexRootPath))
		%>
<%
'Retorna o caminho físico da pasta atual.
caminho_fisico1 = Request.ServerVariables("APPL_PHYSICAL_PATH")
caminho_fisico2 = Request.ServerVariables("PATH_TRANSLATED")


Const ForReading = 1, ForWriting = 2
Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.OpenTextFile(Request.ServerVariables("APPL_PHYSICAL_PATH")&"data.txt", ForReading)
ReadAllTextFile =  f.ReadAll


%>

<form name="formGlobal" action="noaction" onSubmit="return(false);">
<table cellspacing="0" cellpadding="0" border="0" width=100%>
	<tr class="titleRow">
		<td align="Center" colspan="2">&nbsp;<span class="boldText"><%=appName%> - <%=ReadAllTextFile%><br><%=Request.ServerVariables("APPL_PHYSICAL_PATH")&"data.txt"%></span><!--a href="http://<%=Request.ServerVariables("SERVER_NAME")%>" target="_blank" title="Browse the web root"><%=Request.ServerVariables("SERVER_NAME")%></a--></td>
	</tr>
</table>
<table cellspacing=0 cellpadding=2 border=0 width=100%>
	<tr class=lightRow height=60>
		<td>
			<div style="font-size:12pt;">&nbsp;<img align=absmiddle border=0 width=32 height=27 src="./images/folder_open_big.png">&nbsp;<span class=boldText><%=objFolder.Name%></span></div>
			
			<%If virtual<>"" Then%>
				<!--div style="font-size:8pt;">&nbsp;&nbsp;(<a href="<%=virtual%>" target="_blank" title="Browse the virtual folder"><%=virtual%></a>)</div-->
				<div style="font-size:8pt;">&nbsp;&nbsp;(<%=virtual%>)</div-->
			<%Else%>
				<!--div style="font-size:8pt;">&nbsp;&nbsp;(<a href="javascript:Command('NoWebAccess');">no web access</a>)</div-->
				<div style="font-size:8pt;">&nbsp;&nbsp;(no web access)</div>
			<%End If%>
		</td>
		<td nowrap>
			<span class=boldText><%=objFolder.subfolders.count%></span> Pasta(s)<br>
			<span class=boldText><%=objFolder.files.count%></span> Arquivo(s)
		</td>
		<td nowrap>
			Total: <span class=boldText><%If err.Number<>0 or (not calculateTotalSize) Then Response.Write "N/A" Else Response.Write FormatSize(objFolder.size)%></span>
		</td>
		<td colspan=2 align=center>
			<a href="javascript:Command('Refresh');"><img align=absmiddle border=0 width=21 height=20 src="./images/refresh.png" alt="Refresh file listing"></a>&nbsp;
			<a href="javascript:Command('NewFile');"><img align=absmiddle border=0 width=21 height=20 src="./images/create_file.png" alt="Create new file"></a>&nbsp;
			<a href="javascript:Command('NewFolder');"><img align=absmiddle border=0 width=21 height=20 src="./images/create_folder.png" alt="Create new folder"></a>&nbsp;
			<a href="javascript:Command('Upload');"><img align=absmiddle border=0 width=21 height=20 src="./images/upload.png" alt="Upload to this folder"></a>&nbsp;
<%If wexPassword <> "" Then%>
			<a href="javascript:Command('Logout');"><img align=absmiddle border=0 width=21 height=20 src="./images/logout.png" alt="Log out from WebExplorer Free"></a>
<%End If%>
			<br><input name=wexMessage type=text class=formClass size=20 value="<%=server.HTMLEncode(wexMessage)%>" readonly>
		</td>
	</tr>
	<tr class=titleRow>
		<td>&nbsp;<span class=boldText>Name</span></td>
		<td>&nbsp;<span class=boldText>Size</span></td>
		<td>&nbsp;<span class=boldText>Type</span></td>
		<td>&nbsp;<span class=boldText>Modified</span></td>
		<td>&nbsp;</td>
	</tr>
<%
	rowType = "darkRow"

	If len(targetPath) > len(wexRootPath) Then
%>
	<tr class=<%=rowType%>><td>&nbsp;<a href="javascript:Command('LevelUp');" title="Up One level"><img align=absmiddle border=0 width=15 height=13 src="./images/folder_up.png"></a>&nbsp;<a href="javascript:Command('LevelUp');" title="Up One level">..</a></td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td></tr>
<%
		rowType = "lightRow"
	End If
	
	listed = 0
	If (objFolder.subfolders.Count + objFolder.files.Count) = 0 Then
		' Do nothing when error occurs
	Else
		For each item in objFolder.subfolders
			If showHiddenItems or not item.Attributes and 2 Then
				listed = listed + 1
%>
	<tr class=<%=rowType%>>
		<td>&nbsp;<%=GetIcon(item.Name, true)%>&nbsp;<a href="javascript:Command('OpenFolder',&#34;<%=item.Name%>&#34;);" title="Open Folder"><%=item.Name%></a></td>
		<td>&nbsp;<%If calculateFolderSize Then Response.write FormatSize(item.Size)%></td><td>&nbsp;<%=item.Type%></td>
		<td nowrap>&nbsp;<%=item.DateLastModified%></td>
		<td nowrap>
			<%If virtual<>"" Then%>
			<a href="<%=virtual & item.Name%>" target="_blank"><img align=absmiddle border=0 width=11 height=14 src="./images/browse.png" alt="Browse Folder"></a>
			<%End If%>
			<a href="javascript:Command('RenameFolder', &#34;<%=item.Name%>&#34;);"><img align=absmiddle border=0 width=11 height=14 src="./images/rename.png" alt="Rename Folder"></a>
			<a href="javascript:Command('DeleteFolder', &#34;<%=item.Name%>&#34;);"><img align=absmiddle border=0 width=14 height=14 src="./images/delete.png" alt="Delete Folder"></a>
		</td>
	</tr>
<%
				If rowType = "darkRow" Then rowType = "lightRow" Else rowType = "darkRow"
			End If
		Next

		For each item in objFolder.files
			If showHiddenItems or not item.Attributes and 2 Then
				listed = listed + 1
%>
	<tr class=<%=rowType%>>
		<td>&nbsp;<%=GetIcon(item.Name, false)%>&nbsp;<a href="<%=scriptName & "?precommand=Download&folder=" & Server.URLEncode(folder) & "&file=" & Server.URLEncode(item.Name)%>" title="Download File"><%=item.Name%></a></td>
		<td>&nbsp;<%=FormatSize(item.Size)%></td><td>&nbsp;<%=item.Type%></td>
		<td nowrap>&nbsp;<%=item.DateLastModified%></td>
		<td nowrap>
			<%If virtual<>"" Then%>
			<a href="<%=virtual & item.Name%>" target="_blank"><img align=absmiddle border=0 width=11 height=14 src="./images/browse.png" alt="Browse File"></a>
			<%End If%>
			<a href="javascript:Command('RenameFile', &#34;<%=item.Name%>&#34;);"><img align=absmiddle border=0 width=11 height=14 src="./images/rename.png" alt="Rename File"></a>
			<a href="javascript:Command('DeleteFile', &#34;<%=item.Name%>&#34;);"><img align=absmiddle border=0 width=14 height=14 src="./images/delete.png" alt="Delete File"></a>
		</td>
	</tr>
<%
				If rowType = "darkRow" Then rowType = "lightRow" Else rowType = "darkRow"
			End If	
		Next
	End If
%>
	<tr></tr>
</table>
<table cellspacing=0 cellpadding=0 border=0 width=100%>
	<tr class=titleRow>
		<td>&nbsp;</td>
	</tr>
</table>
</form>
<form method=post action="<%=scriptName%>" name=formBuffer>
	<input type="hidden" name=command value="">
	<input type="hidden" name=parameter value="">
	<input type="hidden" name=virtual value="<%=virtual%>">
	<input type="hidden" name=folder value="<%=folder%>">
	<input type="hidden" name=popup value="false">
</form>
<%
		If wexMessage="" Then 
			If (objFolder.subfolders.Count + objFolder.files.Count) <> listed Then
				wexMessage = "Listed " & listed & " of " & (objFolder.subfolders.Count + objFolder.files.Count) & " item(s) , " & (objFolder.subfolders.Count + objFolder.files.Count - listed) & " item(s) are hidden..."
			Else
				wexMessage = "Listed " & (objFolder.subfolders.Count + objFolder.files.Count) & " item(s)..."
			End If
			Response.Write "<script language=""javascript"">document.forms.formGlobal.wexMessage.value='" & wexMessage & "'</script>"
		End If
		
		Set objFolder = Nothing
		
		HtmlCopyright
		HtmlFooter
	End Sub

	' Writes the given error message
	Sub Error(title, message, popup)
		HtmlHeader appName
%>
<table cellpadding=0 cellspacing=0 border=0 align=center width=300>
	<tr class=titleRow>
		<td>&nbsp;<b>An error occured</b></td>
	</tr>
	<tr class=lightRow>
		<td>
			<table cellpadding=0 cellspacing=5 border=0>
				<tr>
					<td valign=top><img width=32 height=32 border=0 align=absmiddle src="./images/error.png"></td>
					<td><b><%=title%>:</b><br><%=message%></td>
				</tr>
			</table>
		</td>
	</tr>
	<tr class=titleRow align=center>
		<td>
			<%If popup Then%>
			<a href="javascript:this.close();">Close</a>
			<%Else%>
			<a href="javascript:history.back();">Back</a>
			<%End If%>
		</td>
	</tr>
</table>
<%
		HtmlFooter
		DestroyApp()
	End Sub
	
	dim logn, pwd
	' WebExplorer Free login screen
	Sub	Login()
		If Request.Form("command") = "Login" Then
			logn = Request.Form("usuario")
			pwd = Request.Form("pwd")
			If logn&pwd = wexPassword Then
				Session(wexId & "Login") = true
				Exit Sub
			Else
				wexMessage = "Wrong password!"
			End If
		End If
		
		HtmlHeader appName
		If(wexMessage<>"") Then Response.Write "<script language=""javascript"">alert('" & wexMessage & "');</script>"
%>
<form name=formBuffer method=post action="<%=scriptName%>" onSubmit="javascript:return(Check());">
<table border=0 cellspacing=0 cellpadding=0 width=400 align=center>
	<tr><td><br><br><br></td></tr>
	<tr><td>
		<table border=0 cellspacing=0 cellpadding=0 width=100%>
			<tr class=titleRow>
				<td align=left>
					&nbsp;<span class=boldText>Login</span>
				</td>
			</tr>
		</table>
		<table border=0 cellspacing=0 cellpadding=0 width=100%>
			<tr align=center class=lightRow>
				<td>
					<br>
					<span class=boldText>Bem-vindo à Nuvem<%'=appName%> <%'="v"&appVersion%></span>
					<br><br>
					<table cellspacing=0 cellpadding=5 border="0" class=loginRow>
						<tr>
							<td rowspan="3"><img src="images/logo.png" alt="" width="120" height="91" border="0"></td>
							<td align=left>&nbsp;<span class=boldText>Login:</span></td>
							<td align=center><input type="text" class="formClass" name="usuario" value="" size="21"></td>
						</tr>
						<tr>
							<td align=left>&nbsp;<span class=boldText>Senha:</span></td>
							<td align=center><input type="password" class="formClass" name="pwd" value="" size="21"></td>
						</tr>
						<tr>
							<td>&nbsp;</td>
							<td align="center"><input type=submit name=submitter value="Login" class=formClass></td>
						</tr>
					</table>
					<br><br><br>
				</td>
			</tr>
		</table>
		<table border=0 cellspacing=0 cellpadding=0 width=100%>
			<tr class=titleRow>
				<td align=center>&nbsp;</td>
			</tr>
		</table>
	</td></tr>
</table>
<input type="hidden" name=command value="Login">
</form>
<script language="javascript">document.forms.formBuffer.usuario.focus();</script>
<%	
		HtmlFooter
		DestroyApp() 
	End Sub
	
	' Relogin message for the popup windows
	Sub PopupRelogin()
		HtmlHeader appName
%>
		<div style="COLOR: white; FONT-FAMILY: Verdana,Tahoma,Arial,Helvetica; FONT-SIZE: 10pt; FONT-WEIGHT: bold;">
		<%=appName%> session is destroyed, please <a href="javascript:PopupRelogin();">relogin</a>.
		</div>
<%
		HtmlFooter
		DestroyApp() 
	End Sub

	' Checks if there is a valid login
	Function Secure()
		If wexPassword = "" Then
			Secure = true
		Else
			If Session(wexId & "Login") Then Secure = true Else Secure = false
			session(wexPassword) = "teste"
		End If
	End Function
		
	' Logs out from WebExplorer Free
	Sub Logout()
		Session.Abandon()
		Login
	End Sub
		
	' Returns the icon of the file
	Function GetIcon(fileName, isFolder)
		Dim ext

		If isFolder Then
			GetIcon = "<a href=""javascript:Command('FolderDetails', &#34;" & fileName & "&#34;);""><img align=absmiddle border=0 width=15 height=13 src=""./images/folder.png"" alt=""Folder - Click to learn details""></a>"
		Else
			ext = FSO.GetExtensionName(fileName)
			
			re.IgnoreCase = true
			re.Pattern = "^" & ext & ",|," & ext & ",|," & ext & "$"

			If re.test(editableExtensions) Then
				GetIcon = "<a href=""javascript:Command('Edit', &#34;" & fileName & "&#34;);""><img align=absmiddle border=0 width=11 height=14 src=""./images/file_editable.png"" alt=""Text file - Click to edit and learn details""></a>"
			ElseIf re.test(viewableExtensions) Then
				GetIcon = "<a href=""javascript:Command('View', &#34;" & fileName & "&#34;);""><img align=absmiddle border=0 width=11 height=14 src=""./images/file_viewable.png"" alt=""Picture file - Click to view and learn details""></a>"
			Else
				GetIcon = "<a href=""javascript:Command('FileDetails', &#34;" & fileName & "&#34;);""><img align=absmiddle border=0 width=11 height=14 src=""./images/file.png"" alt=""File - Click to learn details""></a>"
			End If
		End If
	End Function
	
	' Formats given size in bytes,KB,MB and GB
	Function FormatSize(givenSize)
		If (givenSize < 1024) Then
			FormatSize = givenSize & " B"
		ElseIf (givenSize < 1024*1024) Then
			FormatSize = FormatNumber(givenSize/1024,2) & " KB"
		ElseIf (givenSize < 1024*1024*1024) Then
			FormatSize = FormatNumber(givenSize/(1024*1024),2) & " MB"
		Else
			FormatSize = FormatNumber(givenSize/(1024*1024*1024),2) & " GB"
		End If
	End Function

	' Adds given type of the slash to the end of the path if required
	Function FixPath(path, slash)
		If Right(path, 1) <> slash Then
            FixPath = path & slash
        Else
			FixPath = path
        End If
	End Function

	' Converts the given path to physical path
	Function RealizePath(path)
		Dim fpath
		fpath = replace(path,"/","\")
		If left(fpath,1) = "\" Then 'Virtual path
			on error resume next
			RealizePath = server.MapPath(fpath)
			If err.Number<>0 Then RealizePath = fpath 'Possibly network path
		Else 'Physical Path
			RealizePath = fpath
		End If
		RealizePath = FixPath(RealizePath, "\")
	End Function	

	' Converts the given path to virtual path
	Function VirtualPath(path)
		Dim webRoot, fpath
		webRoot = FixPath(server.MapPath("/"),"\")
		fpath = FixPath(path,"\")
		VirtualPath = ""
		If left(wexRoot,1) = "/" Then
			VirtualPath = FixPath(wexRoot, "/")
			VirtualPath = VirtualPath & right(fpath, len(fpath) - len(wexRootPath))
			VirtualPath = replace(VirtualPath, "\", "/")
			VirtualPath = FixPath(VirtualPath,"/")
		ElseIf left(lcase(fpath), len(webRoot)) = lcase(webRoot) Then
			VirtualPath = "/" & right(fpath, len(fpath) - len(webRoot))
			VirtualPath = replace(VirtualPath, "\", "/")
			VirtualPath = FixPath(VirtualPath,"/")
		End If
	End Function

	'Maps the given path according to the root path
	Function WexMapPath(path)
		If SecurePath(path) Then WexMapPath = FixPath(wexRootPath & path, "\") Else Error "Security Error", "Relative path syntax is forbidden for security reasons.", false
	End Function
		
	' Checks against relative path syntax (. or .. injection)
	Function SecurePath(path)
		Dim fpath
		fpath = replace(path,"/","\")
		
		If fpath="." Then fpath=".\"

		re.IgnoreCase = false
		re.Pattern = "^\.\.$|^\.\.\\|\\\.\.\\|\\\.\.$"
		re.Pattern = re.Pattern & "|^\.\\|\\\.\\|\\\.$"
		
		If re.Test(fpath) Then SecurePath=false Else SecurePath=true
	End Function
	
	' Makes sure that given file name does not contain path info
	Function SecureFileName(name)
		SecureFileName = replace(name,"/","?")
		SecureFileName = replace(SecureFileName,"\","?")
	End Function

	' Creates a folder or a file
	Function CreateItem()
		Dim itemType, itemName, itemPath
		itemType = Request.Form("command")
		itemName = SecureFileName(Request.Form("parameter"))
		itemPath = targetPath & itemName

		on error resume next
		
		Select Case itemType
			Case "NewFolder"
				If FSO.FolderExists(itemPath) = false and FSO.FileExists(itemPath) = false Then 
					FSO.CreateFolder(itemPath)
					If err.Number <> 0 Then 
						wexMessage = "Unable to create the folder """ & itemName & """, an error occured..." 
					Else
						wexMessage = "Created the folder """ & itemName & """..."
					End If
				Else
					wexMessage = "Unable to create the folder """ & itemName & """, there exists a file or a folder with the same name..."
				End If
			Case "NewFile"
				If FSO.FolderExists(itemPath) = false and FSO.FileExists(itemPath) = false Then 
					FSO.CreateTextFile(itemPath)
					If err.Number <> 0 Then 
						wexMessage = "Unable to create the file """ & itemName & """, an error occured..."
					Else
						wexMessage = "Created the file """ & itemName & """..."
					End If
				Else 
					wexMessage = "Unable to create the file """ & itemName & """, there exists a file or a folder with the same name..."
				End IF
		End Select
	End Function
	
	' Deletes a folder or a file
	Function DeleteItem()
		Dim itemType, itemName, itemPath
		itemType = Request.Form("command")
		itemName = SecureFileName(Request.Form("parameter"))
		itemPath = targetPath & itemName

		on error resume next
		
		Select Case itemType
			Case "DeleteFolder"
				FSO.DeleteFolder itemPath, true
				If err.Number <> 0 Then 
					wexMessage = "Unable to delete the folder """ & itemName & """, an error occured..." 
				Else
					wexMessage = "Deleted the folder """ & itemName & """..."
				End If
			Case "DeleteFile"
				FSO.DeleteFile itemPath, true
				If err.Number <> 0 Then 
					wexMessage = "Unable to delete the file """ & itemName & """, an error occured..." 
				Else
					wexMessage = "Deleted the file """ & itemName & """..."
				End If
		End Select
	End Function

	' Renames a folder or a file
	Function RenameItem()
		Dim item, itemType, itemName, itemPath
		Dim param, newName
		itemType = Request.Form("command")
		param = split(Request.Form("parameter"), "|")
		itemName = SecureFileName(param(0))
		newName = SecureFileName(param(1))
		itemPath = targetPath & newName

		on error resume next
		
		Select Case itemType
			Case "RenameFolder"
				If FSO.FolderExists(itemPath) = false and FSO.FileExists(itemPath) = false Then 
					itemPath = targetPath & itemName
					Set item = FSO.GetFolder(itemPath)
					item.Name = newName
					If err.Number <> 0 Then 
						wexMessage = "Unable to rename the folder """ & itemName & """, an error occured..." 
					Else
						wexMessage = "Renamed the folder """ & itemName & """ to """ & newName & """..."
					End If
				Else
					wexMessage = "Unable to rename the folder """ & itemName & """, there exists a file or a folder with the new name """ & newName & """..."
				End If
			Case "RenameFile"
				If FSO.FolderExists(itemPath) = false and FSO.FileExists(itemPath) = false Then 
					itemPath = targetPath & itemName
					Set item = FSO.GetFile(itemPath)
					item.Name = newName
					If err.Number <> 0 Then 
						wexMessage = "Unable to rename the file """ & itemName & """, an error occured..." 
					Else
						wexMessage = "Renamed the file """ & itemName & """ to """ & newName & """..."
					End If
				Else
					wexMessage = "Unable to rename the file """ & itemName & """, there exists a file or a folder with the new name """ & newName & """..."
				End If
		End Select
		
		Set item = Nothing
	End Function
		
	' WebExplorer Free Editor
	Sub Editor()
		Dim fileName, filePath, file
		
		on error resume next

		Select Case Request.Form("subcommand")
			Case "Save", "SaveAs"
				fileName = SecureFileName(Request.Form("parameter"))
				filePath = targetPath & fileName
				Set file = FSO.OpenTextFile (filePath, 2, true, encoding)
				If (err.Number<>0) Then 
					wexMessage = "Can not write to the file """ & fileName & """, permission denied!"
					err.Clear
				Else
					file.write Request.Form("content")
				End If
				Set file = Nothing
				Set file = FSO.OpenTextFile (filePath, 1, false, encoding)
			Case Else
				fileName = SecureFileName(Request.Form("parameter"))
				filePath = targetPath & fileName
				
				If not FSO.FileExists(filePath) Then
					wexMessage = "The file """ & fileName & """ does not exist"
					Set file = FSO.CreateTextFile (filePath, false)
					If err.Number<>0 Then 
						wexMessage = wexMessage & ", also unable to create new file."
						err.Clear 
					Else
						wexMessage = wexMessage & ", created new file."
					End If
				Else
					Set file = FSO.OpenTextFile (filePath, 1, false, encoding)
					If err.Number<>0 Then 
						wexMessage = "Can not read from the file """ & fileName & """, permission denied!"
						err.Clear 
					End If
				End If
		End Select

		HtmlHeader appName
		If(wexMessage<>"") Then Response.Write "<script language=""javascript"">alert('" & wexMessage & "');</script>"
%>
<form name=formBuffer method=post action="<%=scriptName%>">
		<table border=0 cellspacing=0 cellpadding=0 width=100%>
			<tr class=titleRow>
				<td align=left>
					&nbsp;<span class=boldText>Editing</span> - <%=fileName%>
				</td>
				<td align=right>
					<b>Encoding:</b>
					<select name=encoding class=formClass onChange="EditorCommand('Encoding')">
						<option value="-2"<%If encoding=-2 Then Response.Write " selected"%>>Default</option>
						<option value="-1"<%If encoding=-1 Then Response.Write " selected"%>>Unicode</option>
					</select>
				</td>
			</tr>
		</table>
		<table border=0 cellspacing=0 cellpadding=0 width=100% height=90%>
			<tr align=center class=lightRow>
				<td valign=middle>
<textarea name=content class=formClass rows=22 cols=46 style="width:580; height:370;" wrap="off">
<%=Server.HTMLEncode(file.ReadAll)%></textarea>
				</td>
			</tr>
		</table>
		<table border=0 cellspacing=0 cellpadding=0 width=100%>
			<tr class=titleRow>
				<td align=center>					
					<a href="javascript:EditorCommand('Save');">Save</a> | <a href="javascript:EditorCommand('SaveAs');">Save As</a> | <a href="javascript:EditorCommand('Reload');">Reload</a> | <a href="javascript:EditorCommand('Info');">Info</a> | <a href="javascript:this.close();">Close</a>
				</td>
			</tr>
		</table>
<%
		Set file = Nothing
		Set file = FSO.GetFile (filePath)
%>
<input type="hidden" name=command value="Edit">
<input type="hidden" name=subcommand value="">
<input type="hidden" name=parameter value="<%=fileName%>">
<input type="hidden" name=folder value="<%=Request.Form("folder")%>">
<input type="hidden" name=info value="Size: <%=FormatSize(file.Size)%>|Type: <%=file.Type%>|Created: <%=file.DateCreated%>|Last Accessed: <%=file.DateLastAccessed%>|Last Modified: <%=file.DateLastModified%>">
<input type="hidden" name=popup value="true">
</form>
<%
		Set file = Nothing

		HtmlFooter
		DestroyApp() 
	End Sub

	' WebExplorer Free Viewer
	Sub Viewer()
		Dim filePath, file

		filePath = targetPath & Request.Form("parameter")
		If not FSO.FileExists(filePath) Then Error "Viewer Error", "File not found. Please refresh the listing to see if the file actually exists.", true
		
		on error resume next
		Set file = FSO.GetFile(filePath)

		HtmlHeader appName
%>
<form name=formBuffer method=post action="<%=scriptName%>">
		<table border=0 cellspacing=0 cellpadding=0 width=100%>
			<tr class=titleRow>
				<td align=left>
					&nbsp;<span class=boldText>Viewing</span> - <%=file.Name%>
				</td>
			</tr>
		</table>
		<table border=0 cellspacing=0 cellpadding=0 width=100% height=90%>
			<tr align=center class=lightRow>
				<td valign=middle>
					<img src="<%=scriptName & "?precommand=Download&folder=" & Server.URLEncode(Request.Form("folder")) & "&file=" & Server.URLEncode(file.Name)%>">
				</td>
			</tr>
		</table>
		<table border=0 cellspacing=0 cellpadding=0 width=100%>
			<tr class=titleRow>
				<td align=center>					
					<a href="javascript:ViewerCommand('Reload');">Reload</a> | <a href="javascript:ViewerCommand('Info');">Info</a> | <a href="javascript:this.close();">Close</a>
				</td>
			</tr>
		</table>
<input type="hidden" name=command value="View">
<input type="hidden" name=subcommand value="Refresh">
<input type="hidden" name=parameter value="<%=file.Name%>">
<input type="hidden" name=folder value="<%=Request.Form("folder")%>">
<input type="hidden" name=info value="Size: <%=FormatSize(file.Size)%>|Type: <%=file.Type%>|Created: <%=file.DateCreated%>|Last Accessed: <%=file.DateLastAccessed%>|Last Modified: <%=file.DateLastModified%>">
<input type="hidden" name=popup value="true">
</form>
<%
		Set file = Nothing
		HtmlFooter
		DestroyApp() 
	End Sub

	' File/Folder Details
	Sub Details()
		Dim fileName, filePath, file
		
		on error resume next
		fileName = Request.Form("parameter")
		filePath = targetPath & fileName

		HtmlHeader appName
%>
<form name=formBuffer method=post action="<%=scriptName%>">
		<table border=0 cellspacing=0 cellpadding=0 width=100%>
			<tr class=titleRow>
				<td align=left>
					&nbsp;<span class=boldText>Details</span> - <%=fileName%>
				</td>
			</tr>
		</table>
		<table border=0 cellspacing=0 cellpadding=0 width=100% height=80%>
			<tr align=center class=lightRow>
				<td valign=middle>
<%
		If Request.Form("command") = "FileDetails" Then
				Set file = FSO.GetFile (filePath)
		Else
				Set file = FSO.GetFolder (filePath)
		End If
%>
				<table border=0 cellspacing=5 cellpadding=0>
					<tr><td><span class=boldText>Size:</span></td><td><%=FormatSize(file.Size)%></td></tr>
					<tr><td><span class=boldText>Type:</span></td><td><%=file.Type%></td></tr>
					<tr><td><span class=boldText>Created:</span></td><td><%=file.DateCreated%></td></tr>
					<tr><td><span class=boldText>Last Accessed:</span></td><td><%=file.DateLastAccessed%></td></tr>
					<tr><td><span class=boldText>Last Modified:</span></td><td><%=file.DateLastModified%></td></tr>
				</table>
<%
		Set file = Nothing
%>
				</td>
			</tr>
		</table>
		<table border=0 cellspacing=0 cellpadding=0 width=100%>
			<tr class=titleRow>
				<td align=center>					
					<a href="javascript:this.close();">Close</a>
				</td>
			</tr>
		</table>
<input type="hidden" name=command value="<%=Request.Form("command")%>">
<input type="hidden" name=parameter value="<%=fileName%>">
<input type="hidden" name=folder value="<%=Request.Form("folder")%>">
<input type="hidden" name=popup value="true">
</form>
<%
		HtmlFooter
		DestroyApp() 
	End Sub
	
	' Uploads a file
	Sub Upload(process)
		Dim fileTransfer, result

		on error resume next
		Set fileTransfer = New pluginFileTransfer
		If err.number<>0 Then Error "File Transfer Plugin Error", "Plugin cannot be initialized. Please make sure that the components required by the plugin is installed on the server.", true

		If process Then targetPath = WexMapPath(Request.QueryString("folder"))
		
		HtmlHeader appName
%>
		<table border=0 cellspacing=0 cellpadding=0 width=100%>
			<tr class=titleRow>
				<td align=left>
					&nbsp;<span class=boldText>Upload</span> - <%=FSO.GetBaseName(targetPath)%>
				</td>
			</tr>
		</table>
		<table border=0 cellspacing=0 cellpadding=0 width=100% height=80%>
			<tr align=center class=lightRow>
				<td valign=middle>
					<span class=boldText>
<%
		' Actual upload process
		If process Then
			fileTransfer.path = targetPath
			
			result = fileTransfer.Upload()
			
			Select Case result
				Case 0
					Response.Write fileTransfer.uploadedFileName & " is uploaded<br>"
					Response.Write FormatSize(fileTransfer.uploadedFileSize) & " (" & fileTransfer.uploadedFileSize & " bytes) written<br>"
					Response.Write "Content type: " & fileTransfer.contentType
					Response.Write "<script language=""javascript"">opener.Command('Refresh');</script>"
					Response.Write "<script language=""javascript"">alert('Upload concluido');</script>"
					
				Case 1
					Response.Write "No file sent"
				Case 2
					Response.Write "Path not found"
				Case 3
					Response.Write fileTransfer.uploadedFileName & " can not be written"
			End Select
%>
					</span>
					<form name=formBuffer method=post action="<%=scriptName%>">
						<input type=hidden name=command value="Upload">
						<input type=hidden name=folder value="<%=Request.QueryString("folder")%>">
					</form>
<%	
		Else
%>
					<form enctype="multipart/form-data" name=formBuffer method=post action="<%=scriptName%>?precommand=ProcessUpload&folder=<%=server.URLEncode(Request.Form("folder"))%>&popup=true">
						<input type=file name=file class=formClass>
					</form>
<%
		End If
%>					
				</td>
			</tr>
			<tr class=lightRow>
				<td align=center>.........................................................................</td>
			</tr>
			<tr class=lightRow>
				<td align=center>					
					<a href="javascript:Upload();">Upload</a> | <a href="javascript:this.close();">Close</a>
				</td>
			</tr>
		</table>
		<table border=0 cellspacing=0 cellpadding=0 width=100%>
			<tr class=titleRow>
				<td align=center>&nbsp;			
					
				</td>
			</tr>
		</table>
<%
		Set fileTransfer = Nothing

		HtmlFooter
		DestroyApp() 
	End Sub
	
	' Downloads a file
	Sub Download()
		Dim fileTransfer, result
		
		on error resume next
		Set fileTransfer = New pluginFileTransfer
		If err.number<>0 Then Error "File Transfer Plugin Error", "Plugin cannot be initialized. Please make sure that the components required by the plugin is installed on the server.", false
		
		fileTransfer.path = WexMapPath(Request.QueryString("folder"))
		
		result = fileTransfer.Download(Request.QueryString("file"))
		
		Select Case result
			Case 0
				'Success
			Case 1
				Error "Download Error", "File not found. Please refresh the listing to see if the file actually exists.", false
			Case 2
				Error "Download Error", "File cannot be read. Please make sure that you have read permission on the file.", false
		End Select
		
		Set fileTransfer = Nothing

		DestroyApp() 
	End Sub
	
' ------------------------------------------------------------
%>