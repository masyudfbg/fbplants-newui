function validate_(vtr)
{
var svsn1;
	if (vtr=="")
	{
	document.getElementById("nonscript1").innerHTML="&nbsp;&nbsp;&nbsp;";
	document.getElementById("nonscript1_a").style = "visibility:visible;";
	$("#s1").removeClass("error_val");
	return;
	}
	if (window.XMLHttpRequest)
		{// code for IE7+, Firefox, Chrome, Opera, Safari
			svsn1=new XMLHttpRequest();
		}
	else
		{// code for IE6, IE5
			svsn1=new ActiveXObject("Microsoft.XMLHTTP");
		}
		
svsn1.onreadystatechange=function()
	{
		if (svsn1.readyState==4 && svsn1.status==200 )
		{
			var ids2 = document.getElementById ( "s2" );
			var ids1 = document.getElementById ( "s1" );
			var ids3 = document.getElementById ( "s3" );
			var ids4 = document.getElementById ( "s4" );
			var n1 = document.getElementById("nonscript1");
			var n1a = document.getElementById("nonscript1_a");
			var id_empcd = document.getElementById ( "empcd" );
			if (svsn1.responseText == "x" ){	
				document.getElementById("s1").value="";
				$("#s1").addClass("error_val");
				n1.innerHTML="&nbsp;&nbsp;&nbsp;";
				alert('" ERROR ! " Kode Employee ' + vtr + ', tidak ada di Employee Master');
				$("#s1").focus();
				n1a.style = "visibility:visible;";
			}else if (ids1.value == ids2.value || ids1.value ==  ids3.value || ids1.value ==  ids4.value || ids1.value ==  id_empcd.value){
				ids1.value="";
				$("#s1").addClass("error_val");
				n1.innerHTML="&nbsp;&nbsp;&nbsp;";
				alert('" ERROR ! " Kode Employee ' + vtr + ', sudah terpakai kolom lain');
				$("#s1").focus();
				n1a.style = "visibility:visible;";
			}else{
			$("#s1").removeClass("error_val");
			n1.innerHTML=svsn1.responseText;
			n1.style = "color:green;font-size:10px; padding-left:5px;"
			n1a.style = "visibility:hidden;"
			}
		}
	}
	svsn1.open("GET","ajax/validate.asp?validate=svr1&q="+vtr,true);
	svsn1.send();
}
function validate_2(vtr)
{
var svsn2;
	if (vtr=="")
	{
	document.getElementById("nonscript2").innerHTML="&nbsp;&nbsp;&nbsp;";
	document.getElementById("nonscript2_a").style = "visibility:visible;";
	$("#s2").removeClass("error_val");
	return;
	}
	if (window.XMLHttpRequest)
		{// code for IE7+, Firefox, Chrome, Opera, Safari
			svsn2=new XMLHttpRequest();
		}
	else
		{// code for IE6, IE5
			svsn2=new ActiveXObject("Microsoft.XMLHTTP");
		}
		
svsn2.onreadystatechange=function()
	{
		if (svsn2.readyState==4 && svsn2.status==200 )
		{	
			var ids2 = document.getElementById ( "s2" );
			var ids1 = document.getElementById ( "s1" );
			var ids3 = document.getElementById ( "s3" );
			var ids4 = document.getElementById ( "s4" );
			var n2 = document.getElementById("nonscript2");
			var n2a = document.getElementById("nonscript2_a");
			var id_empcd = document.getElementById ( "empcd" );
			if (svsn2.responseText == "x2" ){	
				ids2.value="";
				$("#s2").addClass("error_val");
				n2.innerHTML="&nbsp;&nbsp;&nbsp;";
				alert('" ERROR ! " Kode Employee ' + vtr + ', tidak ada di Employee Master');
				$("#s2").focus();
				n2a.style = "visibility:visible;"
			}else if ( ids2.value == ids1.value ||  ids2.value == ids3.value ||  ids2.value == ids4.value ||  ids2.value == id_empcd.value){
				ids2.value="";
				$("#s2").addClass("error_val");
				n2.innerHTML="&nbsp;&nbsp;&nbsp;";
				alert('" ERROR ! " Kode Employee ' + vtr + ', sudah terpakai kolom lain');
				$("#s2").focus();
				n2a.style = "visibility:visible;"
			}else{
			$("#s2").removeClass("error_val");
			n2.innerHTML=svsn2.responseText;
			n2.style = "color:green;font-size:10px; padding-left:5px;"
			n2a.style = "visibility:hidden;"
			return true;
			
			}
		}
	}
	svsn2.open("GET","ajax/validate.asp?validate=svr2&q="+vtr,true);
	svsn2.send();
}

function validate_3(vtr)
{
var svsn3;
	if (vtr=="")
	{
	document.getElementById("nonscript3").innerHTML="&nbsp;&nbsp;&nbsp;";
	document.getElementById("nonscript3_a").style = "visibility:visible;";
	$("#s3").removeClass("error_val");
	return;
	}
	if (window.XMLHttpRequest)
		{// code for IE7+, Firefox, Chrome, Opera, Safari
			svsn3=new XMLHttpRequest();
		}
	else
		{// code for IE6, IE5
			svsn3=new ActiveXObject("Microsoft.XMLHTTP");
		}
		
svsn3.onreadystatechange=function()
	{
		if (svsn3.readyState==4 && svsn3.status==200 )
		{
			var ids2 = document.getElementById ( "s2" );
			var ids1 = document.getElementById ( "s1" );
			var ids3 = document.getElementById ( "s3" );
			var ids4 = document.getElementById ( "s4" );
			var n3 = document.getElementById("nonscript3");
			var n3a = document.getElementById("nonscript3_a");
			var id_empcd = document.getElementById ( "empcd" );
			if (svsn3.responseText == "x3" ){	
				ids3.value="";
				$("#s3").addClass("error_val");
				n3.innerHTML="&nbsp;&nbsp;&nbsp;";
				alert('" ERROR ! " Kode Employee ' + vtr + ', tidak ada di Employee Master');
				$("#s3").focus();
				n3a.style = "visibility:visible;"
			}else if ( ids3.value == ids1.value || ids3.value == ids2.value || ids3.value == ids4.value || ids3.value == id_empcd.value){
				ids3.value="";
				$("#s3").addClass("error_val");
				n3.innerHTML="&nbsp;&nbsp;&nbsp;";
				alert('" ERROR ! " Kode Employee ' + vtr + ', sudah terpakai kolom lain');
				$("#s3").focus();
				n3a.style = "visibility:visible;"
			}else{
			$("#s3").removeClass("error_val");
			n3.innerHTML=svsn3.responseText;
			n3.style = "color:green;font-size:10px; padding-left:5px;"
			n3a.style = "visibility:hidden;"
			return true;
			
			}
		}
	}
	svsn3.open("GET","ajax/validate.asp?validate=svr3&q="+vtr,true);
	svsn3.send();
}

function validate_4(vtr)
{
var svsn4;
	if (vtr=="")
	{
	document.getElementById("nonscript4").innerHTML="&nbsp;&nbsp;&nbsp;";
	document.getElementById("nonscript4_a").style = "visibility:visible;";
	$("#s4").removeClass("error_val");
	return;
	}
	if (window.XMLHttpRequest)
		{// code for IE7+, Firefox, Chrome, Opera, Safari
			svsn4=new XMLHttpRequest();
		}
	else
		{// code for IE6, IE5
			svsn4=new ActiveXObject("Microsoft.XMLHTTP");
		}
		
svsn4.onreadystatechange=function()
	{
		if (svsn4.readyState==4 && svsn4.status==200 )
		{
			var ids2 = document.getElementById ( "s2" );
			var ids1 = document.getElementById ( "s1" );
			var ids3 = document.getElementById ( "s3" );
			var ids4 = document.getElementById ( "s4" );
			var n4 = document.getElementById("nonscript4");
			var n4a = document.getElementById("nonscript4_a");
			var id_empcd = document.getElementById ( "empcd" );
			if (svsn4.responseText == "x4" ){	
				ids4.value="";
				$("#s4").addClass("error_val");
				n4.innerHTML="&nbsp;&nbsp;&nbsp;";
				alert('" ERROR ! " Kode Employee ' + vtr + ', tidak ada di Employee Master');
				$("#s4").focus();
				n4a.style = "visibility:visible;"
			}else if ( ids4.value == ids1.value || ids4.value == ids2.value || ids4.value == ids3.value || ids4.value == id_empcd.value){
				ids4.value="";
				$("#s4").addClass("error_val");
				n4.innerHTML="&nbsp;&nbsp;&nbsp;";
				alert('" ERROR ! " Kode Employee ' + vtr + ', sudah terpakai kolom lain');
				$("#s4").focus();
				n4a.style = "visibility:visible;"
			}else{
			$("#s4").removeClass("error_val");
			n4.innerHTML=svsn4.responseText;
			n4.style = "color:green;font-size:10px; padding-left:5px;"
			n4a.style = "visibility:hidden;"
			return true;
			
			}
		}
	}
	svsn4.open("GET","ajax/validate.asp?validate=svr4&q="+vtr,true);
	svsn4.send();
}

function validate_empcd(vtr)
{
var empcd;
	if (vtr=="")
	{
	document.getElementById("nonscript_empcd").innerHTML="&nbsp;*&nbsp;";
	document.getElementById("nonscriptempcd_a").style = "visibility:visible;";
	$("#empcd").removeClass("error_val");
	document.getElementById("nonscript_empcd").style = "color:red;"
	return;
	}
	if (window.XMLHttpRequest)
		{// code for IE7+, Firefox, Chrome, Opera, Safari
			
			empcd=new XMLHttpRequest();
		}
	else
		{// code for IE6, IE5
			
			empcd=new ActiveXObject("Microsoft.XMLHTTP");
		}
	
empcd.onreadystatechange=function()
	{
		if (empcd.readyState==4 && empcd.status==200)
		{
			var ids2 = document.getElementById ( "s2" );
			var ids1 = document.getElementById ( "s1" );
			var ids3 = document.getElementById ( "s3" );
			var ids4 = document.getElementById ( "s4" );
			var nempcd = document.getElementById("nonscript_empcd");
			var nempcda = document.getElementById("nonscriptempcd_a");
			var id_empcd = document.getElementById ( "empcd" );
			if (empcd.responseText == "x5" ){	
				id_empcd.value="";
				$("#empcd").addClass("error_val");
				nempcd.innerHTML="&nbsp;*&nbsp;";
				alert('" ERROR ! " Kode Employee ' + vtr + ', tidak ada di Employee Master');
				$("#empcd").focus();
				nempcda.style = "visibility:visible;"
				nempcd.style = "color:red;"
			}else if ( id_empcd.value == ids1.value || id_empcd.value == ids3.value || id_empcd.value == ids4.value || id_empcd.value == ids2.value){
				id_empcd.value="";
				$("#empcd").addClass("error_val");
				nempcd.innerHTML="&nbsp;*&nbsp;";
				alert('" ERROR ! " Kode Employee ' + vtr + ', sudah terpakai kolom lain');
				$("#empcd").focus();
				nempcda.style = "visibility:visible;"
				nempcd.style = "color:red;"
			}else{
			$("#empcd").removeClass("error_val");
			nempcd.innerHTML=empcd.responseText;
			nempcd.style = "color:green;font-size:10px; padding-left:5px;"
			nempcda.style = "visibility:hidden;"
			return true;
			
			}
		}
	}
	empcd.open("GET","ajax/validate.asp?validate=empcd&q="+vtr,true);
	empcd.send();
}		