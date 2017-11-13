	<%
		xsql_mod = "up_ADM_modulos @cd_codigo='"&cd_codigo&"'"
		Set rs_mod = dbconn.execute(xsql_mod)
		
		coop1 = rs_mod("cd_coop_prod")
		coop2 = rs_mod("cd_coop_ben")
		coop3 = rs_mod("cd_coop_rel")
		coop4 = rs_mod("cd_coop_rh")
		coop5 = rs_mod("cd_coop_fecha")
		
		empr1 = rs_mod("cd_empr_horas_trab")
		empr2 = rs_mod("cd_empr_ben")
		empr3 = rs_mod("cd_empr_rel")
		empr4 = rs_mod("cd_empr_rh")
		empr5 = rs_mod("cd_empr_fecha")
		empr6 = rs_mod("cd_empr_bh")
		
		manut1 = rs_mod("cd_manut_lst")
		manut2 = rs_mod("cd_manut_cad")
		manut3 = rs_mod("cd_manut_rel")
		
		prot1 = rs_mod("cd_prot_dig")
		prot2 = rs_mod("cd_prot_rel")
		prot3 = rs_mod("cd_prot_cad")
		
		folha1 = rs_mod("cd_folha_pag")
		
		invent1 = rs_mod("cd_invent_adm")
		
		adm1 = rs_mod("cd_adm_usr")
		
		modulos = "ca"&coop1&","&"cb"&coop2&","&"cc"&coop3&","&"cd"&coop4&","&"ce"&coop5&","&"ea"&empr1&","&"eb"&empr2&","&"ec"&empr3&","&"ed"&empr4&","&"ee"&empr5&","&"ef"&empr6&","&"ma"&manut1&","&"mb"&manut2&","&"mc"&manut3&","&"pa"&prot1&","&"pb"&prot2&","&"pc"&prot3&","&"fa"&folha1&","&"ia"&invent1&","&"aa"&adm1
		
		'session("modulos") = modulos
		
		' Exemplo: "ca"&coop1&"
		
	%>