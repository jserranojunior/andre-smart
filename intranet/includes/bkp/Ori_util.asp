<!--METADATA TYPE="TypeLib" UUID="{B72DF063-28A4-11D3-BF19-009027438003}"-->
<%

   StrError = 0
 '  On Error Resume Next
   
 StrAuth = Decrypt(request("Auth"))
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
StrEmail=escape(LimpaString(Request.Form("Email")))
StrSenha=escape(LimpaString(Request.Form("Senha")))

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
'-----------------------------------------------------------------------------------------------------------        


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
sub mesdoano(mes)

select case mes
	 case "0"
	   mes = ""
         case "1"
           mes = "Jan"         
         case "2"
           mes = "Fev"
         case "3"
           mes = "Mar"
         case "4"
           mes = "Abri"
         case "5"
           mes = "Mai" 
         case "6"
           mes = "Jun"
   	    case "7"
           mes = "Jul"  
         case "8"
           mes = "Ago"  
         case "9"
           mes = "Set"  
         case "10"
           mes = "Out"  
         case "11"
           mes = "Nov"  
         case "12"
           mes = "Dez"                  
  end select
  
  response.write  mes

End sub

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

function formata_horas(horas)'Troca , por ;
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


%>