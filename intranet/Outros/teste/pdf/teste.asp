<!--#include file="../../../includes/pdf/fpdf.asp"--> 
<% 
Set pdf=CreateJsObject("FPDF")
	
	pdf.CreatePDF()
	pdf.CreatePDF "L", "mm", "A4"
	pdf.SetPath("../gravados/")
	
	'pdf.SetFont "Arial","",16 
	pdf.Open() 
	pdf.AddPage() 
	
		pdf.Cell 40,10,"Ol� Mundo!"
	
	pdf.Close() 
	pdf.Output() 
%>