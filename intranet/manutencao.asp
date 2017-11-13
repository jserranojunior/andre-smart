	<%tipo = request("tipo")
	
	'if tipo = "relatorio" Then%>	
	<!--script language="javascript">
		function foco_protocolo()		{
		h_data = new Date();
			h_dia = document.calendario.dt_dia.value;
			h_mes = document.calendario.dt_mes.value;
			h_ano = document.calendario.dt_ano.value;
				
				h_data = h_dia + "/" + h_mes + "/" + h_ano;
				document.calendario.variavel.value=h_mes + "/" + h_ano;
	
			if (document.Form.cd_unidade.value == ''){			
					document.form.cd_unidade.focus();
					//alert("unidade");
				}
			else{
				document.Form.nm_nome.focus();
				}
			}
		}		
	</script-->
	<%'end if%>	
    <!--#include file="topo.asp"-->
	<!--#include file="menu.asp"-->
	
   <div id="frame">
		<%if tipo = "pendencias" Then%>
		<!--#include file="manutencao/manutencao_pendencias.asp"-->
		<%elseif tipo = "especialidades" Then%>
		<!--#include file="manutencao/manutencao_pendencias_especialidades.asp"-->
		<%elseif tipo = "andamento" Then%>
		<!--#include file="manutencao/manutencao_andamento.asp"-->
		<%elseif tipo = "nova" Then%>
		<!--#include file="manutencao/manutencao_nova.asp"-->
		<%elseif tipo = "relatorio" Then%>
		<!--#include file="manutencao/manutencao_relatorio.asp"-->
		<%elseif tipo = "periodo" Then%>
		<!--#include file="manutencao/manutencao_relatorio_periodo.asp"-->
		<%end if%>
   </div>                    
</body>
</html>