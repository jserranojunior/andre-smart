
<script>

var xmlhttp;

function loadXMLDoc(url,cfunc)
	{
		if (window.XMLHttpRequest)
  			{// code for IE7+, Firefox, Chrome, Opera, Safari
  				xmlhttp = new XMLHttpRequest();			
				//alert("IE7+, Firefox, Chrome, Opera, Safari");	
  			}
		else
  			{// code for IE6, IE5
  				xmlhttp = new ActiveXObject("Microsoft.XMLHTTP");
				//alert("IE6, IE5");
  			}
	
		xmlhttp.onreadystatechange = cfunc;
		xmlhttp.open("POST",url,true);
		xmlhttp.send();
	}

</script>