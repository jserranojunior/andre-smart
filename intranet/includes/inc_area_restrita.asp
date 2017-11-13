<%

	If  session("acessopermitido!")  <> "ok" Then
		
		
		
	   Response.Redirect("../../default.asp?erro=1") 

	End if
	
	if user = 46 then
	
	response.write(" teste ")
	
	end if


%>