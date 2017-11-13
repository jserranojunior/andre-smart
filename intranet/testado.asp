<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<!--#include file="/includes/inc_open_connection.asp"-->


<%
sql = "SELECT * dbo.TBL_financeiro_valores3"

executaconexao(sql)

%>

SELECT TOP 1000 [cd_codigo]     
  FROM [VdLap].[dbo].[TBL_financeiro_valores3]