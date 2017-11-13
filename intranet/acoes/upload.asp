<%@ Language=VBScript %>
<%

option explicit 
Response.Expires = -1
Server.ScriptTimeout = 600

%>
<!--#include file="../includes/inc_open_connection.asp"-->
<!-- #include file="../includes/freeaspupload.asp" -->

<%' ****************************************************
' Change the value of the variable below to the pathname
' of a directory with write permissions, for example "C:\Inetpub\wwwroot"
  Dim uploadsDirVar
  
  'uploadsDirVar = "E:\Servidor\Desenvolvimento\desenvolvimento\intranet\foto\funcionarios" 'Desenvolvimento
  'uploadsDirVar = "E:\Servidor\Desenvolvimento\producao\intranet\foto\funcionarios" 'Producao
  uploadsDirVar = "C:\inetpub\wwwroot\producao\intranet\foto\funcionarios" 'Producao
  
 
' ****************************************************

' Note: this file uploadTester.asp is just an example to demonstrate
' the capabilities of the freeASPUpload.asp class. There are no plans
' to add any new features to uploadTester.asp itself. Feel free to add
' your own code. If you are building a content management system, you
' may also want to consider this script: http://www.webfilebrowser.com/

    Dim Upload, fileName, fileSize, ks, i, fileKey, nome_arquivo, SaveFiles, cd_codigo,xsql,dbconn,strConnString

    Set Upload = New FreeASPUpload
    Upload.Save(uploadsDirVar)

    SaveFiles = ""
    ks = Upload.UploadedFiles.keys
    if (UBound(ks) <> -1) then
        'SaveFiles = "<B>Files uploaded:</B> "
		
        for each fileKey in Upload.UploadedFiles.keys
            SaveFiles = SaveFiles & Upload.UploadedFiles(fileKey).FileName' & " (" & Upload.UploadedFiles(fileKey).Length & "B) "
			cd_codigo = Upload.form("cd_codigo")
		next
    else
        SaveFiles = "The file name specified in the upload form does not correspond to a valid file in the system."
    end if
	
	'cd_codigo = request("cd_codigo")
%>

<HTML>
<HEAD>
<TITLE></TITLE>

<%
    'OutputForm()
    'response.write SaveFiles()

	xsql = "up_foto_update @cd_codigo='"&cd_codigo&"', @nm_foto='"&SaveFiles&"'"
	dbconn.execute(xsql)


%>
</HEAD>

<BODY Onload="window.close();">
<body>
<br><br>
<%=SaveFiles%>
<%=cd_codigo%>

</BODY>
</HTML-->
