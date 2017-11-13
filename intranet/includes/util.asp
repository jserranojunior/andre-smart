<!-- TYPEx="TypeLib" UUID="{B72DF063-28A4-11D3-BF19-009027438003}" -->
<%
   StrError = 0
 '  On Error Resume Next
   
' StrAuth = Decrypt(request("Auth"))
 StrAuth = request("Auth")
 strbusca = LimpaString(request("buscar"))

 If StrAuth = "" Then
	StrAuth = 0
End if


Function alteraparaponto(strings)

  alteraparaponto = Replace(strings,",",".")

End function


'Quebra a string encryptada e retorna o request
'------------------------------------------------
Function quebra(request,id)
 if StrAuth <> "" Then
    StrString = trim(replace(StrAuth,"&","#")) & "#"	
		  Do While StrString  <> ""
			StrString1 =  Trim(Divide(StrString,"#"))	
				If StrString1 <> "" Then			
					quebra1 =  Replace(StrString1,request&"=","")		   
						If quebra1 <>  StrString1 then	    
							quebra = LimpaString(quebra1)
							Exit do
						End if			
				End if
		 Loop
   Else
	 quebra = 0
  End if
 End function   

'path  do servidor
'---------------------------------------------------
scriptpath=Request.ServerVariables("PATH_TRANSLATED") 

'Divide String
'----------------------------------------------
Function Divide(xString,xSepara)
	xPosi = Instr(xString,xSepara)
	If xPosi = 0 Then
		xTemp = xString
		xString = ""
	Else
		xTemp = Mid(xString,1,xPosi-1)
		xString = Mid(xString,xPosi+Len(xSepara))
	End If
	Divide = xTemp
End Function

' Função Encrypt()
'----------------------------------------------------------------------------------------
Function Encrypt(theNumber)
  On Error Resume Next
Set CM = Server.CreateObject("Persits.CryptoManager")
    Set Context = CM.OpenContext("mycontainer", True )

    ' generate key from password
    Set Key = Context.GenerateKeyFromPassword( "MM##UIJ5RpSbkb07kf/FVEs2Vw==", _
                    calgSHA, calg3DES )
      
   Set Blob = Key.EncryptText(theNumber)
   
  Encrypt = Blob.Hex

  Set CM = Nothing
  Set Context = Nothing
  Set Key = Nothing
  Set Blob = Nothing
  
End Function


' Função Decrypt()
'----------------------------------------------------------------------------------------
Function Decrypt(theNumber)
  On Error Resume Next
 
 Set CM = Server.CreateObject("Persits.CryptoManager")
        Set Context = CM.OpenContext("mycontainer", True )

       ' generate key from password
        Set Key = Context.GenerateKeyFromPassword( "MM##UIJ5RpSbkb07kf/FVEs2Vw==", _
                        calgSHA, calg3DES )
                        
                        
    ' create an empty blob object and fill it with data
     Set Blob = CM.CreateBlob   
     Blob.Hex = theNumber
      
  Decrypt = Key.DecryptText(Blob)   

   Set CM = Nothing
   Set Key = Nothing
   Set Blob = Nothing
  
End Function

        

'SQL Injection
'---------------------------------------------------------------------------------------------------
'#Valida
function validate( input, datatype )
    good_name_chars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ'-"
     good_password_chars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
       good_number_chars = "0123456789"
         good_sessionid_chars = "ABCDEF0123456789"
  
  validate = true
  
    select case datatype
        case "name"
            for i = 1 to len( input )
                c = mid( input, i, 1 )
                    if ( InStr( good_name_chars, c ) = 0 ) then
                        validate = false
                    end if
            next
        case "password"
            for i = 1 to len( input )
                c = mid( input, i, 1 )
                    if ( InStr( good_password_chars, c ) = 0 ) then
                        validate = false
                    end if
            next
        case "number"
             for i = 1 to len( input )
                 c = mid( input, i, 1 ) 
                    if ( InStr( good_number_chars, c ) = 0 ) then
                       validate = false
                    end if
             next
        case "sessionid"
            for i = 1 to len( input )
                c = mid( input, i, 1 )
                if ( InStr( good_sessionid_chars, c ) = 0 ) then
                      validate = false
                end if
            next
    case else
            validate = false
  end select
 end function


'Valida acesso ao login
'------------------------------------------------------------------------------------------------------
'retrieve our form textbox values and assign to variables 
'StrEmail=escape(LimpaString(Request.Form("Email")))
'StrSenha=escape(LimpaString(Request.Form("Senha")))
StrEmail=escape(Request.Form("Email"))
StrSenha=escape(Request.Form("Senha"))

'Call the function IllegalChars to check for illegal characters
If IllegalChars(StrEmail)=True OR IllegalChars(StrSenha)=True Then
    Response.redirect("index.asp")
End If 

'Function IllegalChars to guard against SQL injection
Function IllegalChars(sInput) 
    'Declare variables 
    Dim sBadChars, iCounter 
        'Set IllegalChars to False 
        IllegalChars=False
            'Create an array of illegal characters and words 
            sBadChars=array("select", "drop", ";", "--", "insert", "delete", "xp_","up_","tbl_","view_", _
                     "#", "%", "&", "'", "(", ")", "/", "\", ":", ";", "<", ">", "=", "[", "]", "?", "`", "|") 
    'Loop through array sBadChars using our counter & UBound function
    For iCounter = 0 to uBound(sBadChars) 
        'Use Function Instr to check presence of illegal character in our variable
        If Instr(sInput,sBadChars(iCounter))>0 Then
            IllegalChars=True
        End If
    Next 
	 iCounter = Null
	 sBadChars = Null
End function

'----------------------------------------------------------------------------------------------------------

'Limpa String
'-----------------------------------------------------------------------------------------------------------
Function LimpaString(input) 
    'Declare variables 
    Dim sBadChars, iCounter 
        'Set IllegalChars to False 
        '   IllegalChars=False
            'Create an array of illegal characters and words 
            sBadChars=array("up_", "update","select", "drop", ";", "--", "insert", "delete", "xp_", " or ","tbl_","view_",_
                     "#", "%", "&", "'", "(", ")", "\", ":", ";", "<", ">", "=", "[", "]", "?", "`", "|") 
    'Loop through array sBadChars using our counter & UBound function
    For iCounter = 0 to uBound(sBadChars) 
        'Use Function Instr to check presence of illegal character in our variable
        input =  replace(input,sBadChars(iCounter),"")  
        'Response.Write     StrLimpa
    Next 
    LimpaString = Trim(input)
	 iCounter = Null
	 sBadChars = Null	
End function
'-----------------------------------------------------------------------------------------------------------        


'#Função que tira todos os acentos das palavras
'-----------------------------------------------------------------------------------------------------------
function TiraAcento(StrAcento)
 for i = 1 to len(StrAcento) 
  Letra = mid(StrAcento, i, 1)
  Select Case Letra
   Case "á","Á","à","À","ã","Ã","â","Â","â","ä","Ä"
   Letra = "a"
   Case "é","É","ê","Ê","Ë","ë","È","è"
   Letra = "e"
   Case "í","Í","ï","Ï","Ì","ì"
   Letra = "I"
   Case "ó","Ó","ô","Ô","õ","Õ","ö","Ö","ò","Ò"
   Letra = "o"
   Case "ú","Ú","Ù","ù","ú","û","ü","Ü","Û"
   Letra = "u"
   Case "ç","Ç"
   Letra = "c"
   Case "ñ"
   Letra = "n"
   Case ","
   Letra = ":"
  End Select
  texto = texto & Letra
 next
 TiraAcento = texto
end function  

'#Função que altera os caracteres no ajax
'-----------------------------------------------------------------------------------------------------------
function busca_ajax(Strcatacestranho)
		Strcatacestranho = replace(Strcatacestranho,"Ã§","ç")
		Strcatacestranho = replace(Strcatacestranho,"Âº","º")
		Strcatacestranho = replace(Strcatacestranho,"Ã£","ã")
		Strcatacestranho = replace(Strcatacestranho,"Ã¡","á")
		Strcatacestranho = replace(Strcatacestranho,"Ã ","à")
		Strcatacestranho = replace(Strcatacestranho,"Ã¢","â")
		Strcatacestranho = replace(Strcatacestranho,"Ã©","é")
		Strcatacestranho = replace(Strcatacestranho,"Ã¨","è")
		Strcatacestranho = replace(Strcatacestranho,"Ãª","ê")
		Strcatacestranho = replace(Strcatacestranho,"Ã­","í")
		Strcatacestranho = replace(Strcatacestranho,"Ã¬","ì")
		Strcatacestranho = replace(Strcatacestranho,"Ã®","î")
		Strcatacestranho = replace(Strcatacestranho,"Ã³","ó")
		Strcatacestranho = replace(Strcatacestranho,"Ã²","ò")
		Strcatacestranho = replace(Strcatacestranho,"Ã´","ô")
		Strcatacestranho = replace(Strcatacestranho,"Ãµ","õ")
		Strcatacestranho = replace(Strcatacestranho,"Ãº","ú")
		Strcatacestranho = replace(Strcatacestranho,"Ã¹","ù")
		Strcatacestranho = replace(Strcatacestranho,"Ã»","û")
		
		
		
		
	busca_ajax = Strcatacestranho
end function  

'-----------------------------------------------------------------------------------------------------------        
Function sinais_ajax(frase) 'Não ativado em nenhum aplicativo ainda.
	for i = 1 to len(frase)
		ajax = mid(frase,i,1)
			if ajax = "%20" then
				ajax = " "
			end if
			texto_ajax = texto_ajax & ajax
	next
	sinais_ajax = texto_ajax
end function

MailFrom = "no-replay@autopro.com.br"
MailFromName = "AUTOPRO"
    
' Função de envio de email
'=====================================================================================================
Function EnviaEmail(MailFrom, MailFromName, MailTo, MailToName, Subject, MailBody, MailFormat)

    
     'Const SMTP_SERVER = "progress.netpoint.com.br"
     Const SMTP_SERVER = "localhost"
	

    Set objMail = Server.CreateObject("Persits.MailSender")

    If MailFormat = 1 Then
        objMail.isHTML = true
    ElseIf MailFormat = 0 Then
        objMail.isHTML = false
    End If
    
    'Coloca a ultima barra no caminho http, caso tenha sido omitida
    If Not Right(ImgPath,1) = "/" Then
        ImgPath = ImgPath & "/"
    End If
    If Not Right(stylePath,1) = "/" Then
        stylePath = stylePath & "/"
    End If
    
    
    strID = "Message-Id: <"&year(strdata_servidor)&""&month(strdata_servidor)&""&day(strdata_servidor)&""&hour(strdata_servidor)&""&minute(strdata_servidor)&""&second(strdata_servidor)&">"
    

              
          srcorpo =             "<html>"  & vbCrLf
          strcorpo = strcorpo & "<head>"  & vbCrLf
          strcorpo = strcorpo & "<title>AutoPro</title>"   & vbCrLf
          strcorpo = strcorpo & "</head>"   & vbCrLf
          strcorpo = strcorpo & "<body bottomMargin=0 leftMargin=0  topMargin=0 rightMargin=0 bgcolor=""#ffffff""> "   & vbCrLf
          strcorpo = strcorpo & "<table  cols=""0"" border=""0"" cellpadding=""0"" cellspacing=""0"" width=100% height=""100%"">"   & vbCrLf
          strcorpo = strcorpo & "<tr>"   & vbCrLf
          strcorpo = strcorpo & "<td bgcolor=#cococo align=center  valign=top>"   & vbCrLf
          strcorpo = strcorpo & "<table width=""769"" height=100% cols=""0"" border=""0"" cellpadding=""0"" cellspacing=""0"" align=center bgcolor=""#ffffff"">"   & vbCrLf
          strcorpo = strcorpo & "<Tr> "   & vbCrLf
          strcorpo = strcorpo & "<Td  valign=top>"   & vbCrLf
          strcorpo = strcorpo & "<table width=""769""  cols=""0"" border=""0"" cellpadding=""0"" cellspacing=""0"" align=center bgcolor=""#ffffff"">"   & vbCrLf
          strcorpo = strcorpo & "<Tr>"   & vbCrLf
          strcorpo = strcorpo & "<Td><img src=""http://www.autopro.com.br/images/top_email.jpg"" alt="""" /></td>"   & vbCrLf
          strcorpo = strcorpo & "</tr>"   & vbCrLf
          strcorpo = strcorpo & "<tr>"   & vbCrLf
          strcorpo = strcorpo & "<td background=http://www.autopro.com.br/images/bg_barra_menu_superior.gif>&nbsp;</TD>"   & vbCrLf
          strcorpo = strcorpo & "</tr>"   & vbCrLf
          strcorpo = strcorpo & "<tr>"   & vbCrLf
          strcorpo = strcorpo & "<td background=http://www.autopro.com.br/images/bg_barra_menu2_superior.gif>&nbsp;</td> "   & vbCrLf
          strcorpo = strcorpo & "</tr>"   & vbCrLf
          strcorpo = strcorpo & "<tr>"   & vbCrLf
          strcorpo = strcorpo & "<Td align=center>"   & vbCrLf
          strcorpo = strcorpo & "<table width=""97%"">"   & vbCrLf
          strcorpo = strcorpo & "<Tr>"   & vbCrLf
          strcorpo = strcorpo & "<Td align=center>"   & vbCrLf
          strcorpo = strcorpo & "<table width=""755"">"   & vbCrLf
          strcorpo = strcorpo & "<Tr> "   & vbCrLf
          strcorpo = strcorpo & "<td><font face=verdana size=1>"&StrMailBody&"<br><br><br>Boas compras!<br><br>Equipe AUTOPRO<br>http://www.autopro.com.br<br><br><br><br><br><br><br><br></font>    </td>"   & vbCrLf
          strcorpo = strcorpo & "</tr>"   & vbCrLf
          strcorpo = strcorpo & "</table>"   & vbCrLf
          strcorpo = strcorpo & "<font color=FFFFFF> "&strID&" </font></td>"   & vbCrLf
          strcorpo = strcorpo & "</tr>"   & vbCrLf
          strcorpo = strcorpo & "</table>"   & vbCrLf
          strcorpo = strcorpo & "</td>"   & vbCrLf
          strcorpo = strcorpo & "</tr>"   & vbCrLf
          strcorpo = strcorpo & "</table>"   & vbCrLf
          strcorpo = strcorpo & "</td> "   & vbCrLf
          strcorpo = strcorpo & "</tr>"   & vbCrLf
          strcorpo = strcorpo & "</table>"   & vbCrLf
          strcorpo = strcorpo & "</body>"   & vbCrLf
          strcorpo = strcorpo & "</html>"   & vbCrLf



    objMail.Host = SMTP_SERVER
    objMail.From = MailFrom
    objMail.FromName = MailFromName
    objMail.AddAddress MailTo, MailToName
    objMail.Subject = Subject
    objMail.Body = strcorpo
   ' objMail.Send
    
    
    Set objMail = nothing

End Function

Function EnviaEmailDebug(MailFrom, MailFromName, MailTo, MailToName, Subject, MailBody, imgPath, MailFormat)


    Const SMTP_SERVER = "progress.netpoint.com.br"
    Set objMail = Server.CreateObject("Persits.MailSender")

    If MailFormat = 1 Then
        objMail.isHTML = true
    ElseIf MailFormat = 0 Then
        objMail.isHTML = false
    End If
    
    'Coloca a ultima barra no caminho http, caso tenha sido omitida
    If Not Right(ImgPath,1) = "/" Then
        ImgPath = ImgPath & "/"
    End If
    
    strBody = MailBody 
    
    objMail.Host = SMTP_SERVER
    objMail.From = MailFrom
    objMail.FromName = MailFromName
    objMail.AddAddress MailTo, MailToName
    objMail.Subject = Subject
    objMail.Body = strBody
    objMail.Send

    Set objMail = nothing
End Function
'-----------------------------------------------------------------------------------------------------------
    
   
'-----------------------------------------------------------------
'Funcao:    IsCartaoCredito(ByVal strNumeroCartao, strTipoCartao)
'Sinopse:    Verifica se o cartão de crédito passado por parametro 
'            está no formato correto e se o dígito é correto. 
'            Formatos aceitos: 
'            Cartão            Prefixo                    Tamanho
'             MASTERCARD         51-55                     16
'             VISA             4                         13 ou 16
'             AMEX             34 ou 37                 15
'             DINERSCLUB         300-305 ou 36 ou 38     14
'Parametro: strNumeroCartao: Numero do cartão (somente número)
'            strTipoCartao: Pode assumir os seguintes valores: 
'                           MASTERCARD, VISA, AMEX, DINERSCLUB
'Retorno: Booleano
'Autor: Gabriel Fróes - www.roccofroes.com
'-------------------------------------------------------------
Function IsCartaoCredito(ByVal strNumeroCartao, strTipoCartao)
    'Verificando se o valor passado é todo numerico
    If Not IsNumeric(strNumeroCartao) Then
        Retorno = False
    Else
        Retorno = True
    End If


    
    'Valor é numérico
    If Retorno Then
        'Selecionando o prefixo do cartão
        strTipoCartao    =    Ucase(strTipoCartao)
        Select Case strTipoCartao
            Case 2
                strExpressaoRegular = "^5[1-5]\d{14}$"
            Case 1
                strExpressaoRegular = "^4(\d{12}|\d{15})$"
            Case 6
                strExpressaoRegular = "^3(3|7)\d{14}$"
            Case 3
                strExpressaoRegular = "^3((6|8)\d{12})|(00|01|02|03|04|05)\d{11})$"
        End Select

            
        'Validando o formato do cartão de crédito
        Set regEx = New RegExp                            ' Cria o Objeto Expressão
        regEx.Pattern = strExpressaoRegular                  ' Expressão Regular
        regEx.IgnoreCase = True                           ' Sensitivo ou não
        regEx.Global = True                               
        Retorno = RegEx.Test(strNumeroCartao)
        Set regEx = Nothing
        
        'Formato correto
        If Retorno Then
            '-----------------------------------------
            'Processo de validação do numero do cartão    
            '-----------------------------------------
            intVerificaSoma        = 0
            blnFlagDigito        = False 
            For Cont = Len(strNumeroCartao) To 1 Step -1
                Digito = Asc(Mid(strNumeroCartao, Cont, 1))        'Isola o caracter da vez
                If (Digito > 47) And (Digito < 58) Then            'Somente se for inteiro
                    Digito = Digito - 48                        'Converte novamente para numero (-48)
                    If blnFlagDigito Then                  
                        Digito = Digito + Digito                'Primeiro duplica-o
                        If Digito > 9 Then                        'Verifica se o Digito é maior que 9
                            Digito = Digito - 9                    'Força ser somente um número
                        End If
                    End If
                    blnFlagDigito = Not blnFlagDigito      
                    intVerificaSoma = intVerificaSoma + Digito   
                    If intVerificaSoma > 9 Then                
                        intVerificaSoma = intVerificaSoma - 10   'Mesmo que MOD 10 só que mais rapido
                    End If
                End If
            Next
            If intVerificaSoma <> 0 Then ' Deve totalizar zero
                Retorno = False
            Else
                Retorno = True
            End If
            '-----------------------------------------
        End If
    End If
    'Retornando a função
    IsCartaoCredito = Retorno
    'response.write IsCartaoCredito
End Function


'---------------------------------
' Traduzir Mês do ano
'---------------------------------
'sub mesdoano(mes)
'select case mes
'	 case "0"
'	   mes = ""
'         case "1"
'           mes = "Jan"         
'         case "2"
'           mes = "Fev"
'         case "3"
'           mes = "Mar"
 '        case "4"
'           mes = "Abri"
 '        case "5"
 '          mes = "Mai" 
 '        case "6"
 '          mes = "Jun"
 '  	    case "7"
 '          mes = "Jul"  
 ''        case "8"
 '          mes = "Ago"  
 '        case "9"
  '         mes = "Set"  
  '       case "10"
  '         mes = "Out"  
  '       case "11"
  '         mes = "Nov"  
  '       case "12"
  '         mes = "Dez"                  
 ' end select  
'  response.write  mes

function mesdoano(mes)
		if mes = "0" then
	  		mes_abrv = ""
        elseif mes = "1" then
			mes_abrv = "Jan"
        elseif mes = "2" then
           	mes_abrv = "Fev"
        elseif mes = "3" then
           	mes_abrv = "Mar"
        elseif mes = "4" then
           	mes_abrv = "Abri"
        elseif mes = "5" then
           	mes_abrv = "Mai"
        elseif mes = "6" then
           	mes_abrv = "Jun"
		elseif mes = "7" then
           	mes_abrv = "Jul"
        elseif mes = "8" then
			mes_abrv = "Ago"
        elseif mes = "9" then
			mes_abrv = "Set"
        elseif mes = "10" then
			mes_abrv = "Out"
        elseif mes = "11" then
           	mes_abrv = "Nov"
        elseif mes = "12" then
           	mes_abrv = "Dez"
		end if
	mesdoano = mes_abrv
 end function
'  response.write  mes

'End sub

Function mes_selecionado(mes) 'funcionarios
'prod_valor_tt = FormatCurrency(min_t,,,,0)
If mes = "1" Then
	mes_selecionado = "Janeiro"
ElseIf mes = "2" Then
	mes_selecionado = "Fevereiro"
ElseIf mes = "3" Then
	mes_selecionado = "Março"
ElseIf mes = "4" Then
	mes_selecionado = "Abril"
ElseIf mes = "5" Then
	mes_selecionado = "Maio"
ElseIf mes = "6" Then
	mes_selecionado = "Junho"
ElseIf mes = "7" Then
	mes_selecionado = "Julho"
ElseIf mes = "8" Then
	mes_selecionado = "Agosto"
ElseIf mes = "9" Then
	mes_selecionado = "Setembro"
ElseIf mes = "10" Then
	mes_selecionado = "Outubro"
ElseIf mes = "11" Then
	mes_selecionado = "Novembro"
ElseIf mes = "12" Then
	mes_selecionado = "Dezembro"
End If
End Function

'Function formata_ausencias(min)
'Hora = Int(min/59)   '60 minutos
'Minutos = min Mod 60 '60 segundos
'NovaHora = Right("0" & Hora,2) & ":" & Right("0" & Minutos,2)
'formatahora = NovaHora
'End Function

function formata_horas(horas)'Troca , por :
 for i = 1 to len(horas) 
  Letra = mid(horas, i, 1)
  Select Case Sinal
   Case ","
   Letra = ":"
  End Select
  texto = texto & Letra
 next
 formata_horas = texto
end function

Function formataidade(idade)
if isNumeric(idade) then
	if idade > 12 then
		anos = idade/12
		idade_anos = Int(anos)
		idade_meses = anos Mod 12 
			NovaIdade = idade_anos&"a "&idade_meses&"m"
	elseif idade <= 12 then
		NovaIdade = idade&"m"
	end if
end if
formataidade = NovaIdade
End Function

Function formatahora(min)
Hora = Int(min/60)   '60 minutos
Minutos = min Mod 60 '60 segundos
'NovaHora = Right("0" & Hora,3) & ":" & Right("0" & Minutos,2)
	if Right(Hora,3)< 10 then NHh=Right("0"&Hora,3) Else NHh=Right(Hora,3) end if
	if Right(Minutos,2)< 10 then NHm=Right("0"&Minutos,2) Else NHm=Right(Minutos,2) end if
		NovaHora = NHh& ":" & NHm
formatahora = NovaHora
End Function

Function formataseg(seg)
Hora = Int(seg/3600)   '60 minutos
minutos = seg - hora
'	Minutos = seg Mod 60 '60 segundos
Segundos = seg mod 60  'segundos

Novoseg = Right("0" & Hora,2) & ":" & Right("0" & Minutos,2)
formataseg = Novoseg
End Function

Function total_mes(min)
negativo = Instr(1,min,"-",1)
	if negativo <> 0 Then 'Verifica e retira sinal de negativo
		minuto = replace(min,"-","")
		sinal = "-"
	else
		minuto = min
		sinal = ""
	end if
Hora = Int(minuto/60)   '60 minutos
Minutos = minuto Mod 60 '60 segundos
	'negativo = Instr(1,Minutos,"-",1)
		'if negativo <> 0 Then 'Verifica e retira sinal de negativo dos minutos
		'Minutos = replace(Minutos,"-","")
		'end if
NovaHora = sinal&Right("" & Hora,4) & ":" & Right("0" & Minutos,2)
total_mes = NovaHora
End Function

Function formata_hora_rcm(min)
Hora = Int(min/60)   '60 minutos
Minutos = min Mod 60 '60 segundos
NovaHora = Right("0" & Hora,3) & ":" & Right("0" & Minutos,2)
formatahora = NovaHora
End Function

Function rcm_horas(min)
negativo = Instr(1,min,"-",1)
	if negativo <> 0 Then 'Verifica e retira sinal de negativo
		minuto = replace(min,"-","")
		sinal = "-"
	else
		minuto = min
		sinal = ""
	end if
Hora = Int(minuto/60)   '60 minutos
Minutos = minuto Mod 60 '60 segundos
	'negativo = Instr(1,Minutos,"-",1)
		'if negativo <> 0 Then 'Verifica e retira sinal de negativo dos minutos
		'Minutos = replace(Minutos,"-","")
		'end if
NovaHora = sinal&Right("" & Hora,4) & ":" & Right("0" & Minutos,2)
total_mes = NovaHora
End Function


'Função ultimo dia do mes

Function UltimoDiaMes(iMonth, iYear)
 Dim dTemp
 dTemp = DateAdd("d", -1, DateSerial(iYear, iMonth + 1, 1))
 UltimoDiaMes = Day(dTemp)
End Function
'Response.Write UltimoDiaMes(02,2005)

'Dia da semana por extenso
Function DiaSemanaExtenso(dia)
	if dia = "1" then
	DiaSemanaExtenso = "Domingo"
	elseif dia = "2" then
	DiaSemanaExtenso = "Segunda-Feira"
	elseif dia = "3" then
	DiaSemanaExtenso = "Terça-feira"
	elseif dia = "4" then
	DiaSemanaExtenso = "Quarta-feira"
	elseif dia = "5" then
	DiaSemanaExtenso = "Quinta-Feira"
	elseif dia = "6" then
	DiaSemanaExtenso = "Sexta-Feira"
	elseif dia = "7" then
	DiaSemanaExtenso = "Sábado"
	end if
end function



'-----------------------------------------------------
'Funcao: getTempoExtenso(ByVal Tempo)
'Sinopse: Descreve o tempo por extenso em relação ao
'          tempo atual
'Parametro(s):
'    Tempo: Tempo expressos em Bytes
'Retorno: String
'Autor: Gabriel Fróes - www.codigofonte.com.br
'-----------------------------------------------------
Function getTempoExtenso(ByVal Tempo)
    On Error Resume Next
    Dim Retorno
    Dim Agora
    Agora = Now()
    Tempo = DateDiff("n", Tempo, Agora)
    'Response.Write "<br>Agora:" & Agora
    'Response.Write "<br>Tempo:" & Tempo
    If IsNumeric(Tempo) Then
        If Tempo <= 60 Then
            Retorno = Tempo & " minuto"
            If Tempo > 1 Then
                Retorno = Retorno & "s"
            End If
        ElseIf Tempo > 60  And Tempo <= 1440 Then
            Tempo    = Round(Tempo/60, 0)
            Retorno = Tempo & " hora"
            If Tempo > 1 Then
                Retorno = Retorno & "s"
            End If
        ElseIf Tempo > 1440  And Tempo <= 43200 Then
            Tempo    = Round((Tempo/60)/24, 0)
            Retorno = Tempo & " dia"
            If Tempo > 1 Then
                Retorno = Retorno & "s"
            End If
        ElseIf Tempo > 43200  And Tempo <= 518400 Then
            Tempo    = Round(((Tempo/60)/24)/30, 0)
            Retorno = Tempo & " mês"
            If Tempo > 1 Then
                Retorno = Tempo & " meses"
            End If
        ElseIf Tempo > 518400  Then
            Tempo    = Round((((Tempo/60)/24)/30)/12, 0)
            Retorno = Tempo & " ano"
            If Tempo > 1 Then
                Retorno = Retorno & "s"
            End If
        End If
    Else
        Retorno = "n/a"
    End If
        
    'Retornando a função
    getTempoExtenso = Retorno
End Function
'EXEMPLO DE CHAMADA 
'Response.Write "Exemplo de Chamada:" & getTempoExtenso(1440)

'____________________________________________________________
Function zero(dezena)
	If dezena = "" Then
	Elseif dezena < 10 AND dezena <> "" Then
		'response.write("0"&dezena)
		zero = "0"&dezena
	Else
		'response.write(dezena)
		zero = dezena
	End if
End Function

Function manutencao(os)
	If os = "" Then
		'Não faz nada
	Elseif os < 10 AND os <> "" Then
		manutencao = "000"&os	
	Elseif os < 100 AND os <> "" Then
		manutencao = "00"&os
	Elseif os < 1000 and os <> "" Then
		manutencao = "0"&os
	ELSE
		manutencao = os
	End if
End Function

Function prefx_pat(pat)
	patrimonio = pat
	while not len(patrimonio) = "4"
		patrimonio = "0"&patrimonio	
	wend
	prefx_pat = patrimonio
End Function

'________________________________________________________________
'-----PROTOCOLOS-------------------------------------------------
Function cdun(unidade)
If unidade = "" Then
	cdun = "00"
Elseif unidade < 10 Then
	cdun = "0"&unidade
Else
	cdun = unidade
End if
End Function


Function proto(protocolo)
p_comp = int(len(protocolo))
len_proto = 6-p_comp
for i = 1 to len_proto
	prot_0 = "0"&prot_0
next
proto = prot_0&protocolo
End Function

'Function proto(protoco)
'if protoco <= 9 then
'	prot_0 = "00000"
'elseif protoco <= 99 then
'	prot_0 = "0000"
'elseif protoco <= 999 then
'	prot_0 = "000"
'elseif protoco <= 9999 then
'	prot_0 = "00"
'elseif protoco <= 99999 then
'	prot_0 = "0"
'end if


'proto = prot_0&protoco'&" ("&p_comps&")"
'End Function


function digitov(protocolo)

	str_unidade = mid(protocolo,1,2)
	str_protocolo = mid(protocolo,4,6)
	'response.write("("&str_protocolo&") &nbsp;")
		
	'Verificar se o protocolo é proveniente do siga
	'***********************************************
		if str_unidade = 2 and str_protocolo < 17222 OR str_unidade = 3 and str_protocolo < 114145 OR str_unidade = 11 and str_protocolo < 1476 OR str_unidade = 14 and str_protocolo < 2955  OR str_unidade = 15 and str_protocolo < 3102 OR str_unidade = 19 and str_protocolo < 2412 OR str_unidade = 20 and str_protocolo < 4138 OR str_unidade = 22 and str_protocolo < 65 then
		pula = "1"
		end if
		'elseif  str_unidade = "2" and str_protocolo <= "017222" then
	'-----------------------------------------------
	if pula <> 1 then
	
			str_unid_proto = zero(str_unidade)&proto(str_protocolo)
				multipl = 2
			for i = 1 to len(str_unid_proto)
				caracteres = mid(str_unid_proto,i,1)
				soma_tudo = int(caracteres*multipl) + int(soma_tudo)
				multipl = multipl + 1 
			next		
		
			div_soma = soma_tudo/11
			resultado = round(div_soma,1)
				ver_resto = instr(resultado,",")
				resto = mid(resultado, ver_resto + 1, 1)*10
			resultado2 = round(div_soma,2)
				ver_resto2 = instr(resultado2,",")
				resto2 = mid(resultado2, ver_resto2 + 1,2)
			
			'Sempre Arredonda para cima
			if int(resto2) > int(resto) then 
			resto = int(resto/10+1)
			Else
			resto = resto/10
			end if
			'*************************
			dv =  11 - resto		
					if dv = 10 then
					dv = 0
					end if
		
		digitov = (str_unidade&"."&str_protocolo&"-"&dv)
		
				
	Else'if pula = 1 then
		digitov = (str_unidade&"."&str_protocolo&"-<u style='color:gray;'>X</u>")
	end if
			
		'digitov = (str_unidade&"."&str_protocolo&"-"&dv)
			
end function


Function busca_inteligente(str)
	Dim v  
	v = lcase(str)  
	v = Replace(v,"%","")  
	v = Replace(v,"'","")  
	v = Replace(v,"""","")  
	v = replace(v, "ó" , "o")
	'v = replace(v, "Ã³" , "o")
	v = replace(v, "ò" , "o")  
	v = replace(v, "ô" , "o")  
	v = replace(v, "õ" , "o")  
	v = replace(v, "ö" , "o")  
	v = replace(v, "á" , "a")  
	v = replace(v, "à" , "a")  
	v = replace(v, "â" , "a")  
	v = replace(v, "ã" , "a")  
	v = replace(v, "ä" , "a")  
	v = replace(v, "é" , "e")  
	v = replace(v, "è" , "e")  
	v = replace(v, "ê" , "e")  
	v = replace(v, "ú" , "u")  
	v = replace(v, "ù" , "u")  
	v = replace(v, "û" , "u")  
	v = replace(v, "ü" , "u")  
	v = replace(v, "í" , "i")  
	v = replace(v, "ì" , "i")  
	v = replace(v, "ç" , "c")  
	v = replace(v,"a","[a,á,à,ã,â,ä]")  
	v = replace(v,"e","[e,é,è,ê]")  
	v = replace(v,"i","[i,í,ì]")  
	v = replace(v,"o","[o,ó,ò,õ,ô,ö]")  
	v = replace(v,"u","[u,ú,ù,û,ü]")  
	v = replace(v,"c","[c,ç]")  
	v = replace(v,"'","['']")  
	'Prepara_busca = v
	busca_inteligente = v
End Function

Function busca_inteligente_ajax(str)
	Dim v
	v = lcase(str)
	v = Replace(v,"%","")
	v = Replace(v,"'","")
	v = Replace(v,"""","")
	v = replace(v,"Âº","º")
	v = replace(v,"Ã£","a")
	v = replace(v,"Ã¡","a")
	v = replace(v,"Ã ","a")
	v = replace(v,"Ã¢","a")
	v = replace(v,"ã£","a")
	v = replace(v,"ã¡","a")
	v = replace(v,"ã ","a")
	v = replace(v,"ã¢","a")
	'v = replace(v, "ä" , "a")
	v = replace(v,"Ã©","e")
	v = replace(v,"Ã¨","e")
	v = replace(v,"Ãª","e")
	v = replace(v,"ã©","e")
	v = replace(v,"ã¨","e")
	v = replace(v,"ãª","e")
	v = replace(v,"Ã­","i")
	v = replace(v,"Ã¬","i")
	v = replace(v,"Ã®","i")
	v = replace(v,"Ã³","o")
	v = replace(v,"Ã²","o")
	v = replace(v,"Ã´","o")
	v = replace(v,"Ãµ","o")
	v = replace(v,"ã³","o")
	v = replace(v,"ã²","o")
	v = replace(v,"ã´","o")
	v = replace(v,"ãµ","o")
	'v = replace(v, "ö" , "o")
	v = replace(v,"Ãº","u")
	v = replace(v,"Ã¹","u")
	v = replace(v,"Ã»","u")
	'v = replace(v, "ü" , "u")
	v = replace(v,"Ã§","c")
	v = replace(v,"a","[a,á,à,ã,â,ä]")
	v = replace(v,"e","[e,é,è,ê]")
	v = replace(v,"i","[i,í,ì]")
	v = replace(v,"o","[o,ó,ò,õ,ô,ö]")
	v = replace(v,"u","[u,ú,ù,û,ü]")
	v = replace(v,"c","[c,ç]")
	v = replace(v,"'","['']")
	busca_inteligente_ajax = v
End Function

function retira_virgula(str)
	if instr(1,str,",",1) = 1 then
		str = mid(str,2,len(str))
	end if
	retira_virgula = str
end function

%>
<script language="JavaScript">
// ***** Util.asp ***

if (navigator.appName.indexOf('Microsoft') != -1){   
    clientNavigator = "IE";   
}else{   
    clientNavigator = "Other";   
}  


document.getElementsByClassName = function(cl) {
	var retnode = [];
	var myclass = new RegExp('\\b'+cl+'\\b');
	var elem = this.getElementsByTagName('*');
		for (var i = 0; i < elem.length; i++) {
			var classes = elem[i].className;
				if (myclass.test(classes)) retnode.push(elem[i]);
		}
		return retnode;		
};


function fecha_atualiza(){
window.opener.location.reload();
window.close();
}

function fecha(){
window.close();
}
</script>

<script language="javascript">
<!--
VerifiqueTAB=true;
function Mostra(quem, tammax) {
if ( (quem.value.length == tammax) && (VerifiqueTAB) ) { 
var i=0,j=0, indice=-1;
for (i=0; i<document.forms.length; i++) { 
for (j=0; j<document.forms[i].elements.length; j++) { 
if (document.forms[i].elements[j].name == quem.name) { 
indice=i;
break;
} 
} 
if (indice != -1) break; 
} 
for (i=0; i<=document.forms[indice].elements.length; i++) { 
if (document.forms[indice].elements[i].name == quem.name) { 
while ( (document.forms[indice].elements[(i+1)].type == "hidden") &&
(i < document.forms[indice].elements.length) ) { 
i++;
} 
document.forms[indice].elements[(i+1)].focus();
VerifiqueTAB=false;
break;
} 
} 
} 
} 
function PararTAB(quem) { VerifiqueTAB=false; } 
function ChecarTAB() { VerifiqueTAB=true; } 


  function mOvr(src,clrOver) {
    if (!src.contains(event.fromElement)) {
	  src.style.cursor = 'hand';
	  src.bgColor = clrOver;
	}
  }
  function mOut(src,clrIn) {
	if (!src.contains(event.toElement)) {
	  src.style.cursor = 'default';
	  src.bgColor = clrIn;
	}
  }

  
function visualizarImpressao(){
 var Navegador = '<object id="Navegador1" width="0" height="0" classid="CLSID:8856F961-340A-11D0-A96B-00C04FD705A2"></object>';
 document.body.insertAdjacentHTML('beforeEnd', Navegador);
 Navegador1.ExecWB(7, 1);
 Navegador1.outerHTML = "";
}


function Change(input){
var str=document.getElementById(input).value;
var word=/ /g; // Letra à trocar: entre as barras
var word2=/ /g;
var replacestr=str.replace(word,"_"); // Letra substituta: entre as aspas
var replacestr2=replacestr.replace(word2,"_"); 
document.getElementById(input).value=replacestr2;}


function somente_numeros(evnt){
//Função permite digitação de números   
    if (clientNavigator == "IE"){   
        if (evnt.keyCode < 48 || evnt.keyCode > 57){   
            return false   
        }   
    }else{   
        if ((evnt.charCode < 48 || evnt.charCode > 57) && evnt.keyCode == 0){   
            return false   
        }   
    }   
}

startList = function() {
if (document.all&&document.getElementById) {
	navRoot = document.getElementById("nav");
		for (i=0; i<navRoot.childNodes.length; i++) {
			node = navRoot.childNodes[i];
	if (node.nodeName=="LI") {
		node.onmouseover=function() {
	this.className+=" over";
  }
  		node.onmouseout=function() {
  	this.className=this.className.replace(" over", "");
   }
   }
  }
 }
}
//window.onload=startList;

 /* 
// Pega elemento por classe 
*/ 
var allHTMLTags = new Array(); 
 
 
function hideandshow_a(theClass) { 
        //Cria Array com todas as TAGS HTML 
        var allHTMLTags=document.getElementsByTagName("*"); 
        //Passa por todas as tags usando um FOR 
        for (i=0; i<allHTMLTags.length; i++) { 
                //Pega todas as tags com a classe passada na função. 
                if (allHTMLTags[i].className==theClass) { 
                //Aqui voce coloca o código que você deseja para cada tag com a classe desejada 
                //No exemplo abaixo, mudei a cor do fundo de todas as tags com o nome que passei pela função 
                //allHTMLTags[i].style.background="#DDFFBB"; 
				//var el = document.getElementById( id ); 
		        if( allHTMLTags[i].style.display != 'none' ) 
		                allHTMLTags[i].style.display = 'none'; 
		        else 
		                allHTMLTags[i].style.display = 'inline';
				} 
        } 
}
function hideandshow_b(theClass,situacao) { 
        //Cria Array com todas as TAGS HTML 
        var allHTMLTags=document.getElementsByTagName("*"); 
        //Passa por todas as tags usando um FOR 
        for (i=0; i<allHTMLTags.length; i++) { 
                //Pega todas as tags com a classe passada na função. 
                if (allHTMLTags[i].className==theClass) { 
                //Aqui voce coloca o código que você deseja para cada tag com a classe desejada 
                //No exemplo abaixo, mudei a cor do fundo de todas as tags com o nome que passei pela função 
                //allHTMLTags[i].style.background="#DDFFBB"; 
				//var el = document.getElementById( id ); 
		        /*if( allHTMLTags[i].style.display != situacao1 ) 
		                allHTMLTags[i].style.display = situacao1; 
		        else */
		                allHTMLTags[i].style.display = situacao;
				} 
        } 
}

function hideandshow(theClass,memoria_classe) { 
        var allHTMLTags=document.getElementsByTagName("*");
		//var Mem_class = document.getElementById(memoria_classe);
		var Mem_class = document.getElementsByClassName(memoria_classe);
		  //alert(theClass);
		  //document.getElementById("documentomano").value="x"
		  
		  for (i=0; i<allHTMLTags.length; i++) { 
                
				if (allHTMLTags[i].className==theClass) { 
	                if( allHTMLTags[i].style.display != 'none' )
			            allHTMLTags[i].style.display = 'none';
							//document.getElementById("documentos").value="x"
							//document.getElementById("documentomano").value="x"
			        else
						allHTMLTags[i].style.display = 'inline';	
							if (memoria_classe != ''){
							//if (navigator.appName.indexOf('Microsoft') != -1)||(memoria_classe != ''){ 	
								Mem_class.value = allHTMLTags[i].style.display;
								//document.form.dt_dependente_nascimento_mes.value="";
								//document.getElementById("documentos").value="y"
								//document.getElementById("documentomano").value="y"
							}						
				} 				
        } 
}

/*
function hideandshow(theClass,memoria_classe) { 
        var allHTMLTags=document.getElementsByTagName("*");
		var Mem_class = document.getElementById(memoria_classe);
		  for (i=0; i<allHTMLTags.length; i++) { 
                
				if (allHTMLTags[i].className==theClass) { 
	                if( allHTMLTags[i].style.display != 'none' )
			            allHTMLTags[i].style.display = 'none'; 						
			        else
						allHTMLTags[i].style.display = 'inline';	
							if (memoria_classe != ''){
								Mem_class.value = allHTMLTags[i].style.display;
							}						
				} 				
        } 
}

*/


function Identifica(campo) {
	var nome = campo.name;
	alert(nome);;
	/*coloque este cod no campo que quer descobrir o nome:
	 onClick="Identifica(this);"*/

}

// Formata valores
addEvent = function(o, e, f, s){
	var r = o[r = "_" + (e = "on" + e)] = o[r] || (o[e] ? [[o[e], o]] : []), a, c, d;
	r[r.length] = [f, s || o], o[e] = function(e){
		try{
			(e = e || event).preventDefault || (e.preventDefault = function(){e.returnValue = false;});
			e.stopPropagation || (e.stopPropagation = function(){e.cancelBubble = true;});
			e.target || (e.target = e.srcElement || null);
			e.key = (e.which + 1 || e.keyCode + 1) - 1 || 0;
		}catch(f){}
		for(d = 1, f = r.length; f; r[--f] && (a = r[f][0], o = r[f][1], a.call ? c = a.call(o, e) : (o._ = a, c = o._(e), o._ = null), d &= c !== false));
		return e = null, !!d;
    }
};

removeEvent = function(o, e, f, s){
	for(var i = (e = o["_on" + e] || []).length; i;)
		if(e[--i] && e[i][0] == f && (s || o) == e[i][1])
			return delete e[i];
	return false;
};

function formataMoeda(o, n, dig, dec){
	o.c = !isNaN(n) ? Math.abs(n) : 2;
	o.dec = typeof dec != "string" ? "," : dec, o.dig = typeof dig != "string" ? "." : dig;
	addEvent(o, "keypress", function(e){
		if(e.key > 47 && e.key < 58){
			var o, s, l = (s = ((o = this).value.replace(/^0+/g, "") + String.fromCharCode(e.key)).replace(/\D/g, "")).length, n;
			if(o.maxLength + 1 && l >= o.maxLength) return false;
			l <= (n = o.c) && (s = new Array(n - l + 2).join("0") + s);
			for(var i = (l = (s = s.split("")).length) - n; (i -= 3) > 0; s[i - 1] += o.dig);
			n && n < l && (s[l - ++n] += o.dec);
			o.value = s.join("");
		}
		e.key > 30 && e.preventDefault();
	});
}

// Pula de uma Campo a Outro Automático
function JumpField(fields,next_field) {
	campo_data = fields;
	prox_campo = next_field;
	tam_campo_data = fields.value.length;
	max_campo_data = fields.maxLength;
		if (tam_campo_data == max_campo_data){
			document.getElementById(prox_campo).focus();		
		}
}

-->
</script>

