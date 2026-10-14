<% ValidateSession %>
<!--#include virtual="/ConndB1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%
Dim ScriptName, v_txtFcCode, v_txtFcName, gangcode, divisi

v_txtFcCode		=	trim(Request("txtFcCode"))
v_txtFcName		=	trim(Request("txtFcName"))
gangcode	=	trim(Request("gangcode"))
divisi		=	trim(Request("divisi"))
selecteddivision= request("selecteddivision")
selectedgang= request("selectedgang")

' Added Sendy 10/17/2018
			v_txtfcba = Request("txtfcba")
			txtYear   = Request("txtYear")
			txtMonth  = Request("txtMonth")

			txttdate = Request("txttdate")
			' v_txttodaydate		= Request("txttodaydate")
			
			' txtyear 		= YEAR(Request("txttodaydate"))
			' txtmonth 		= MONTH(Request("txttodaydate"))
			' txtday			= DAY(Request("txttodaydate")) 
			' txttodaydate 	= Request("txttodaydate")
			
			' txtComp         = Request("txtCompany")
			nSpanTotal = 8
			' v_txtComp   =   txtComp
			' v_txtyear			=	txtyear
			' v_txtmonth			=	txtmonth
			If "" & txtYear = "" Then txtYear = Year(date)
			If "" & txtMonth = "" Then txtMonth = Month(date)
			If "" & txttdate = "" Then txttdate = textDate(date)

			txttdate = txtMonth & "/1/" & txtYear
			
			
			'response.write  "GP " & v_txtGroupPlantation & "  P " & v_txtPlantation & "  GR " & v_txtGroupRegion & " R " & v_txtRegion & " C " & v_txtcompany & " S " & v_txtfcba  &  " D " & division & " B "  & v_txtBlok & " TT" & v_txtPlantingDate
		
			varDefaultBa = Session("fcBA")
		

		If (v_txtFcba = "" Or Len(Trim(v_txtFcba)) = 0) Then
			v_txtFcba = varDefaultBa
		End If	
			
	       if  v_txtfcba <> "" then
		       whereCons = whereCons & "  AND em.fcba = '" & v_txtfcba & "'"
			end if
		    ' selectCons = selectCons & ",fcba"
		    ' orderCons  = orderCons & ",BUSINESSUNIT_API.get_OrderData(fcba)"
			q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & v_txtfcba & "' "
			set rsCat = connStr.Execute(q)
			if not rsCat.eof then
			   fcbaname = rsCat("fcName")
			   CompanyName=rsCat("FCCOMPANYNAME")
			   MgrName = rsCat("FCMANAGER")
			   KTUName = rsCat("FCKTU")
			   CropName=rsCat("FCCROP")
			   FcbaType=rsCat("fcType")
		else
		    nSpanTotal = nSpanTotal - 1 
		end if

			if selecteddivision <>  "--ALL--" and selecteddivision <> "" then
				whereCons = whereCons & " AND EM.SECTIONNAME = '" & selecteddivision & "'"
			else
				nSpanTotal = nSpanTotal - 1  
			end if

			if selectedgang <>  "--ALL--" and selectedgang <> "" then
				whereCons = whereCons & "  AND EM.GANGCODE = '" & selectedgang & "'"
			else
				nSpanTotal = nSpanTotal - 1  
			end if

' End Addded


Dim v_txtisactive : v_txtisactive = Request("txtisactive")

'response.write "aktive -" & v_txtisactive
ScriptName		=	Request.ServerVariables("SCRIPT_NAME")
%>
<html>

<head>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>ePlaS Reporting - Data Master Employee</title>
<script language="JavaScript" src='/js/x.js'></script>
<script language="JavaScript" src='/js/xtableheaderfixed.js'></script>
<link rel="stylesheet" href="css/jquery-ui.css" />
<!-- HADIR/ font ICON -->
<link rel="stylesheet" href="/css/font/font-awesome.css" />
<!-- END:HADIR -->
<script type="text/javascript" src="/dtree.js"></script>
<script src="/js/jquery-1.9.1.js"></script>
<script src="/js/jquery-ui.js"></script>

<link href="/plugins/tabcontent/tabcontent.css" rel="stylesheet" type="text/css" />
<link href="/plugins/chosen/chosen.min.css" rel="stylesheet" />

<link rel="stylesheet" href="/plugins/responsivegridsystem/col.css" media="all">
<link rel="stylesheet" href="/plugins/responsivegridsystem/12cols.css" media="all">
<script type="text/javascript">
	function hitungrumus() {
		var squantity = $("#quantity").val();
		var samount1 = $("#amount1").val(); 
		var newsquantity = squantity.replace(/,/g, '');
		var newsamount1 = samount1.replace(/,/g, '');
		
		//Untuk menghitung Total Quantity Tolerance
		v_unitprice = newsamount1 / newsquantity; //a kali b
		if (!isNaN(v_unitprice)) {
    	 	$("#unitprice").val(v_unitprice);
    	}
    	else{
    		$("#unitprice").val();
    	} 
	}
</script>
<script language="javascript" type="text/javascript">
/** XHConn - Simple XMLHTTP Interface - bfults@gmail.com - 2005-04-08        **
 ** Code licensed under Creative Commons Attribution-ShareAlike License      **
 ** http://creativecommons.org/licenses/by-sa/2.0/                           **/
function XHConn()
{
  var xmlhttp, bComplete = false;
  try { xmlhttp = new ActiveXObject("Msxml2.XMLHTTP"); }
  catch (e) { try { xmlhttp = new ActiveXObject("Microsoft.XMLHTTP"); }
  catch (e) { try { xmlhttp = new XMLHttpRequest(); }
  catch (e) { xmlhttp = false; }}}
  if (!xmlhttp) return null;
  this.connect = function(sURL, sMethod, sVars, fnDone)
  {
    if (!xmlhttp) return false;
    bComplete = false;
    sMethod = sMethod.toUpperCase();
    try {
      if (sMethod == "GET")
      {
        xmlhttp.open(sMethod, sURL+"?"+sVars, true);
        sVars = "";
      }
      else
      {
        xmlhttp.open(sMethod, sURL, true);
        xmlhttp.setRequestHeader("Method", "POST "+sURL+" HTTP/1.1");
        xmlhttp.setRequestHeader("Content-Type",
          "application/x-www-form-urlencoded");
      }
      xmlhttp.onreadystatechange = function(){
        if (xmlhttp.readyState == 4 && !bComplete)
        {
          bComplete = true;
          fnDone(xmlhttp);
        }};
      xmlhttp.send(sVars);
    }
    catch(z) { return false; }
    return true;
  };
  return this;
}

  	// doAJAXCall : Generic AJAX Handler, used with XHConn
// Author : Bryce Christensen (www.esonica.com)
// PageURL : the server side page we are calling
// ReqType : either POST or GET, typically POST
// PostStr : parameter passed in a query string format 'param1=foo&param2=bar'
// FunctionName : the JS function that will handle the response

var doAJAXCall = function (PageURL, ReqType, PostStr, FunctionName) {

	// create the new object for doing the XMLHTTP Request
	var myConn = new XHConn();

	// check if the browser supports it
	if (myConn)	{
	    
	    // XMLHTTPRequest is supported by the browser, continue with the request
	    myConn.connect('' + PageURL + '', '' + ReqType + '', '' + PostStr + '', FunctionName);    
	} 
	else {
	    // Not support by this browser, alert the user
	    alert("XMLHTTP not available. Try a newer/better browser, this application will not work!");   
	}
}

// launched from button click 
var getlistdivision = function () {
	// build up the post string when passing variables to the server side page
	
	var fcba = document.getElementById("fcba").value;
	var selecteddivision = document.getElementById("selecteddivision").value;
	var PostStr = 'fcba='+fcba+'&divisi='+selecteddivision;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('/ajax/getlistdivision.asp', 'POST', PostStr, showlistdivision);
}

// The function for handling the response from the server
var showlistdivision = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;

    // update the Div to show the result from the server
	document.getElementById("divisilist").innerHTML = "";
	document.getElementById("divisilist").innerHTML = response;

	var config = {
		  '.chosen-select'           : {},
		  '.chosen-select-deselect'  : {allow_single_deselect:true},
		  '.chosen-select-no-single' : {disable_search_threshold:10},
		  '.chosen-select-no-results': {no_results_text:'Oops, nothing found!'},
		  '.chosen-select-width'     : {width:"45%"}
		}
		for (var selector in config) {
		  $(selector).chosen(config[selector]);
		}
}

var getlistgangcode = function () {
	// build up the post string when passing variables to the server side page
	
	var selecteddivision = document.getElementById("selecteddivision").value;
	var fcba = document.getElementById("fcba").value;
	var selectedgang = document.getElementById("selectedgang").value;
	var PostStr = 'divisi='+selecteddivision+'&fcba='+fcba+'&gangcode='+selectedgang;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('/ajax/getlistgangcode.asp', 'POST', PostStr, showlistgangcode);
}

// The function for handling the response from the server
var showlistgangcode = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;

    // update the Div to show the result from the server
	document.getElementById("ganglist").innerHTML = "";
	document.getElementById("ganglist").innerHTML = response;

	var config = {
		  '.chosen-select'           : {},
		  '.chosen-select-deselect'  : {allow_single_deselect:true},
		  '.chosen-select-no-single' : {disable_search_threshold:10},
		  '.chosen-select-no-results': {no_results_text:'Oops, nothing found!'},
		  '.chosen-select-width'     : {width:"45%"}
		}
		for (var selector in config) {
		  $(selector).chosen(config[selector]);
		}
};

</script>
<script language="JavaScript">
xAddEventListener(window, 'load',
  function() {
    new xTableHeaderFixed('xthf-site', window);
    new xTableHeaderFixed('xthf-green', window);
    new xTableHeaderFixed('xthf-blue', window);
  }, false
);
</script>
<script language="JavaScript" src="/js/date-picker.js"></script> 
<script language="JavaScript">
var newwin;
function launchwin(winurl,winname,winfeatures)
{
	//This launches a new window and then
	//focuses it if window.focus() is supported.
	newwin = window.open(winurl,winname,winfeatures);

}
</script>
<style>
@media print
{
body {
font-family: 'Arial Narrow', Verdana, Helvetica, sans-serif;
font-size:12px;
background-color:#fff;
color:#000;
}
h1{font-size:16px;}
h2{font-size:14px;}
h3{font-size:12px;}
h4{font-size:11px;}
h5{font-size:10px;}
h6{font-size:9px;}
.printhide{display:none;}
}
</style>
</head>

<body>
	<div id="wrapper">
	 <!--#include file="../print_excel.asp"-->
		<div id="header-wrapper">
			<div class="printhide">
				<!--#include file="menutop.asp"-->
				<table id="form-table" class="centered-margin">
					<tr>
						<form name="frmAR" method="post" action="<%=ScriptName%>">
							<!-- Added Sendy 09/05/2018 -->
						<tr>
							<th>TAHUN</th>
							<th>BULAN</th>
							<th>KEBUN</th>
							<th>KODE</th>
							<th>NAME</th>
							<th>DIVISI</th>
							<th>GANGCODE</th>
							<th>STATUS AKTIF</th>
							<th>&nbsp;</th>
						</tr>
						<td align="center">
							<% 
								Dim mon, Yr, MinYr, MaxYr, i, MM

								Yr = txtYear
								MinYr = 2015
								MaxYr = Yr  + 2
							%>
										<select name="txtYear">
							<%
								For i = MinYr to MaxYr
							%>
									<Option Value="<%=i%>" <%If cint(i) = cint(txtYear) then%> Selected <%End If%>><%=i%></option>
							<%
								Next
							%>
										</select>
							</td>
							<td align="center">
								<select name="txtMonth">
											<%
												For i = 1 to 12
													if i <= 9 then 
														iBulan = 0&i
													else
														iBulan = i
													end if
													i	= iBulan
											%>
												<Option Value="<%=iBulan%>" <%If cint(i) = cint(txtMonth) then%> Selected <%End If%>><%=MonthName(i)%></option>
											<%
												Next
											%>
										</select>
							</td>
							<td>
								<select name="txtFcba" id="fcba" onchange="javascript:getlistdivision();">
								<%
								Dim rsFcba
								Set rsFcba = Server.CreateObject("ADODB.Recordset")
								rsFcba.CursorLocation = adUseClient
								strsql =   " select fccode fcba from businessunit " & _
											" where " & _
											" USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',fccode) IS NOT NULL " 
								  
								
								rsFcba.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
								if not rsFcba.eof then
									DO while not rsFcba.eof
										if v_txtFcba = rsFcba("fcba") then
										%>			
											<option value="<%=rsFcba("fcba")%>" Selected><%=rsFcba("fcba")%></option>
										<% Else %>	
											<option value="<%=rsFcba("fcba")%>"><%=rsFcba("fcba")%></option>
										<% End If
										
										rsFcba.movenext
									loop
								End if
								rsFcba.Close
								set rsFcba = Nothing
								%>
								</select>	
								<input type="hidden" id="selectedfcba" name= "selectedfcba" value="<%=selectedfcba%>">
							</td>
							<!-- End Added -->
						<td>
							<input type="text" maxlength="10" name="txtFcCode" size="10" value="<%=v_txtFcCode%>" title="Input Code">
						</td>
						<td>
							<input type="text" maxlength="10" name="txtFcName" size="30" value="<%=v_txtFcName%>" title="Input Name">
						</td>
						<td>
							 <div id="divisilist"></div>
							 <input type="hidden" id="selecteddivision" name= "selecteddivision" value="<%=selecteddivision%>">
						</td>						
						
						
						<td>
							 <div id="ganglist"></div>
							 <input type="hidden" id="selectedgang" name= "selectedgang" value="<%=selectedgang%>">
						</td>
						<td>
							<select name="txtisactive" >
								<option value = "" <%If Len(Trim(v_txtisactive)) = 0 Then%> Selected <%End If%>>--SEMUA--</option>
								<option value = "Y" <%If v_txtisactive = "Y" Then%> Selected <%End If%>>AKTIF</option>
								<option value = "N" <%If v_txtisactive = "N" Then%> Selected <%End If%>>TIDAK AKTIF</option>
							</select>
						</td>

						<td><input type="submit" value="View" /></td>
						</form>
					</tr>
				</table>
			</div>
<%
	Dim rsqGang, qGang, txtqGang
	qGang	=	" SELECT FCNAME FROM GANG WHERE FCBA = '" & Session("FCBA") & "' AND FCCODE = '"& gangcode & "'"
	
	Set rsqGang = Server.CreateObject("ADODB.Recordset")
	rsqGang.Open qGang, connstr, adOpenStatic
	
	if not rsqGang.eof then
		txtqGang	=	rsqGang("FCNAME")
	end if
		rsqGang.Close
		set rsqGang = Nothing
%>
<table border="0" width="100%" id="product-table" class="xthf-site">
<thead>
	<tr>
    <th colspan="50" style="background-color:#FFFFFF; border:none;">  
	<div id="header-wrapper">
		<div id="head_judul" style="float:left; text-align:left;">
			<h4>PT. <%=CompanyName%></h4>
			<h5><%=v_txtFcba%></h5>
		</div>
		<div id="subheader" class="centered" style="width:50%; margin:auto;">
			<h2>DATA KARYAWAN</h2>
			<h4>DIVISION : <%if divisi <> "" then%> <%=divisi%> <%end if%></h4>
			<% IF gangcode <> "" then %>
			<h4>GANG : <%=gangcode%> - <%=txtqGang%></h4>
			<%END IF%>
			 
		</div>
	</div>
	
</th>
</tr>
	<tr class="rpt-keep">
		<th style="color:#fff;"></th>
		<th colspan="12" style="color:#fff;">Biodata</th>
		<th colspan="8" style="color:#fff;">Pekerjaan</th>
		<th colspan="4" style="color:#fff;">Pajak</th>
		<th colspan="4" style="color:#fff;">Bank</th>
		<th colspan="4" style="color:#fff;">BPJS</th>
		<th style="color:#fff;"></th>
	</tr>
     <tr>
		<th style="color:#fff;">No</th>
		<th style="color:#fff;">Kode Karyawan</th>
		<th style="color:#fff;">Nama Karyawan</th>
		<th style="color:#fff;">Jenis Kelamin</th>
		<th style="color:#fff;">Tanggal Lahir</th>
		<th style="color:#fff;">Agama</th>
		<th style="color:#fff;">Suku</th>
		<th style="color:#fff;">No. ID</th>
		<th style="color:#fff;">Pendidikan</th>
		<th style="color:#fff;">Status Perkawinan</th>
		<th style="color:#fff;">Pasangan</th>
		<th style="color:#fff;">Alamat</th>
		<th style="color:#fff;">Negara</th>
		<th style="color:#fff;">NIK Perusahaan</th>
		<th style="color:#fff;">Divisi</th>
		<th style="color:#fff;">Gang</th>
		<th style="color:#fff;">Tanggal Bergabung</th>
		<th style="color:#fff;">Status</th>
		<th style="color:#fff;">Golongan</th>
		<th style="color:#fff;">Jabatan</th>
		<th style="color:#fff;">Tipe Beras</th>
		<th style="color:#fff;">Potong Pajak</th>
		<th style="color:#fff;">Status Perkawinan</th>
		<th style="color:#fff;">No. NPWP</th>
		<th style="color:#fff;">Tanggung Perusahaan</th>
		<th style="color:#fff;">Tipe Pembayaran</th>
		<th style="color:#fff;">Nama Bank</th>
		<th style="color:#fff;">Nama</th>
		<th style="color:#fff;">No. Rekening</th>
		<th style="color:#fff;">Kategori BPJS Ketenagakerjaan</th>
		<th style="color:#fff;">No. BPJS Ketenagakerjaan</th>
		<th style="color:#fff;">Kategori BPJS Kesehatan</th>
		<th style="color:#fff;">No. BPJS Kesehatan</th>
		<th style="color:#fff;">Tanggal Berhenti</th>

	</tr>
 </thead>
 <tbody>   
    
	<%
	Dim rsEmployee, strEmployee, noline , strKJ, rsKJ
	status_payroll = "FALSE"
	strKJ = " SELECT NVL(payroll_executed, 'FALSE') detail" & _
	        " FROM empalldedstatus" & _
			" WHERE fcba = '" & v_txtFcba & "'" & _
	        "       AND year = " & txtYear & "" & _
	        "       AND lpad(month,2,0) = '" & txtMonth & "'"

	Set rsKJ = Server.CreateObject("ADODB.Recordset")
	rsKJ.Open strKJ, connstr, adOpenStatic

	IF NOT rsKJ.EOF THEN
		status_payroll = rsKJ("detail")
	ELSE
		status_payroll = "FALSE"
	END IF
	rsKJ.Close
	Set rsKJ = Nothing
	' response.write strKJ
	if  v_txtfcba <> "" then
	'IF status_payroll = "TRUE" THEN
			strEmployee	=	" SELECT EM.FCCODE," & _
							"        EM.FCNAME," & _
							"        EM.GENDER," & _
							"        to_char(EM.DATEOFBIRTH,'dd/mm/yyyy')DATEOFBIRTH," & _
							"        EM.RELIGION," & _
							"        EM.RACE," & _
							"        EM.IDENTITYCARD," & _
							"        EM.HIGHESTEDUCATION," & _
							"        EM.MARITALSTATUS," & _
							"        EM.SPOUSE," & _
							"        EM.ADDRESS," & _
							"        EM.NATIONALITY," & _
							"        EM.COMPANY_NIK," & _
							"        NVL(HIST.SECTIONNAME,EM.SECTIONNAME) SECTIONNAME," & _
							"        NVL(HIST.GANGCODE,EM.GANGCODE) GANGCODE," & _
							"        to_char(EM.DATEOFJOIN,'dd/mm/yyyy')DATEOFJOIN," & _
							"        NVL(HIST.PRESENTGRADE,EM.PRESENTGRADE) PRESENTGRADE," & _
							"        EM.SUBGRADE," & _							
							"        EP.DESCRIPTION," & _
							"        EM.NATURETYPE," & _
							"        EM.INCLUDEDTAXCALCULATION," & _
							"        EM.TAXSTATUS," & _
							"        EM.NPWP," & _
							"        EM.TAXPAIDBYCOMPANY," & _
							"        EM.PAYMENTTYPE," & _
							"        EM.BANKNAME," & _
							"        EM.OWNERSHIPOFACCOUNT," & _
							"        EM.ACCOUNTNO," & _
							"        PARAMETER_API.get_fcname(EM.JAMSOSTEKCATEGORY,'E0007',EM.FCBA) JAMSOSTEKCATEGORY," & _
							"        EM.JAMSOSTEKNO," & _
							"        PARAMETER_API.get_fcname(EM.BPJSKES_CATEGORY,'PAYROLL_BPJSKES',EM.FCBA) BPJSKES_CATEGORY," & _
							"        EM.BPJSKES_NO," & _
							"        to_char(EM.DATETERMINATE,'dd/mm/yyyy')DATETERMINATE" & _
							" FROM EMPLOYEE EM" & _
							" INNER JOIN EMPPOSITION EP ON (EM.FCBA = EP.FCBA AND EM.POSITION = EP.POSITION) " & _
							" LEFT JOIN (SELECT FCBA,FCCODE,SECTIONNAME,GANGCODE,PRESENTGRADE FROM EMPLOYEE_HIST_TAB " & _ 
							"            WHERE year = '"& txtyear &"' and LPAD(month,2,0) = '"& txtmonth &"' and fcba='" & v_txtFcba & "') HIST ON (EM.FCBA=HIST.FCBA AND EM.FCCODE=HIST.FCCODE) " & _
							" WHERE  TO_NUMBER(TO_CHAR(DATEOFJOIN,'YYYYMM')) <= TO_NUMBER(TO_CHAR(TO_DATE('"& txttdate &"', 'mm/dd/yyyy'), 'yyyymm')) "& whereCons & "" 
		if  len(trim(v_txtFcCode)) > 0 then
			strEmployee	=	strEmployee &  " AND EM.FCCODE = '" & v_txtFcCode & "'"
		end if
		if  len(trim(v_txtFcName)) > 0 then
			strEmployee	=	strEmployee &  " AND EM.FCNAME LIKE '%" & v_txtFcName & "%'"
		end if
		' if  len(trim(divisi)) > 0 then
		' 	strEmployee	=	strEmployee &  " AND GG.GANGGROUP = '" & divisi & "'"
		' end if
		' if  len(trim(gangcode)) > 0 then
		' 	strEmployee	=	strEmployee &  " AND EM.GANGCODE = '" & gangcode & "'"
		' end if
		if  len(trim(v_txtisactive)) > 0 then
			If v_txtisactive = "Y" Then
				strEmployee	=	strEmployee &  " AND (em.dateterminate is null or TO_NUMBER(TO_CHAR(em.dateterminate,'YYYYMMDD')) >= TO_NUMBER(TO_CHAR( LAST_DAY(TO_DATE('"& txttdate &"', 'mm/dd/yyyy')), 'YYYYMMDD')) )"
			Else
				strEmployee	=	strEmployee &  " AND (em.dateterminate is not null AND TO_NUMBER(TO_CHAR(TO_DATE('"& txttdate &"', 'mm/dd/yyyy'), 'yyyymm')) >= TO_NUMBER(TO_CHAR(em.dateterminate,'YYYYMM')) )"
			End If
		end if
		strEmployee	=	strEmployee &  " ORDER BY EM.FCCODE ASC"
		
	'	response.write strEmployee
		Set rsEmployee = Server.CreateObject("ADODB.Recordset")
		rsEmployee.Open strEmployee, connstr, adOpenStatic	
	END IF
	' Else
		' strEmployee	=	" SELECT EM.FCCODE," & _
							' "        EM.FCNAME," & _
							' "        EM.GENDER," & _
							' "        EM.DATEOFBIRTH," & _
							' "        EM.RELIGION," & _
							' "        EM.RACE," & _
							' "        EM.IDENTITYCARD," & _
							' "        EM.HIGHESTEDUCATION," & _
							' "        EM.MARITALSTATUS," & _
							' "        EM.SPOUSE," & _
							' "        EM.ADDRESS," & _
							' "        EM.NATIONALITY," & _
							' "        EM.COMPANY_NIK," & _
							' "        EM.SECTIONNAME," & _
							' "        EM.GANGCODE," & _
							' "        EM.DATEOFJOIN," & _
							' "        EM.PRESENTGRADE," & _
							' "        EM.SUBGRADE," & _							
							' "        EP.DESCRIPTION," & _
							' "        EM.NATURETYPE," & _
							' "        EM.INCLUDEDTAXCALCULATION," & _
							' "        EM.TAXSTATUS," & _
							' "        EM.NPWP," & _
							' "        EM.TAXPAIDBYCOMPANY," & _
							' "        EM.PAYMENTTYPE," & _
							' "        EM.BANKNAME," & _
							' "        EM.OWNERSHIPOFACCOUNT," & _
							' "        EM.ACCOUNTNO," & _
							' "        PARAMETER_API.get_fcname(EM.JAMSOSTEKCATEGORY,'E0007',EM.FCBA) JAMSOSTEKCATEGORY," & _
							' "        EM.JAMSOSTEKNO," & _
							' "        PARAMETER_API.get_fcname(EM.BPJSKES_CATEGORY,'PAYROLL_BPJSKES',EM.FCBA) BPJSKES_CATEGORY," & _
							' "        EM.BPJSKES_NO," & _
							' "        EM.DATETERMINATE" & _
							' " FROM EMPLOYEE EM" & _
							' " INNER JOIN EMPPOSITION EP ON (EM.FCBA = EP.FCBA AND EM.POSITION = EP.POSITION)" & _
							' "    WHERE  TO_NUMBER(TO_CHAR(DATEOFJOIN,'YYYYMM')) <= TO_NUMBER(TO_CHAR(TO_DATE('"& txttdate &"', 'mm/dd/yyyy'), 'yyyymm')) "& whereCons & "" 
		' if  len(trim(v_txtFcCode)) > 0 then
			' strEmployee	=	strEmployee &  " AND EM.FCCODE = '" & v_txtFcCode & "'"
		' end if
		' if  len(trim(v_txtFcName)) > 0 then
			' strEmployee	=	strEmployee &  " AND EM.FCNAME LIKE '%" & v_txtFcName & "%'"
		' end if
		' ' if  len(trim(divisi)) > 0 then
		' ' 	strEmployee	=	strEmployee &  " AND GG.GANGGROUP = '" & divisi & "'"
		' ' end if
		' ' if  len(trim(gangcode)) > 0 then
		' ' 	strEmployee	=	strEmployee &  " AND EM.GANGCODE = '" & gangcode & "'"
		' ' end if
		' if  len(trim(v_txtisactive)) > 0 then
			' If v_txtisactive = "Y" Then
				' strEmployee	=	strEmployee &  " AND em.dateterminate is null"
			' Else
				' strEmployee	=	strEmployee &  " AND em.dateterminate is not null"
			' End If
		' end if
		' strEmployee	=	strEmployee &  " ORDER BY EM.FCCODE ASC"
	' END IF	
	' response.write strEmployee
	'---- Perbaikan: query dulu dijalankan dua kali, dan error bila BA kosong (SQL kosong)
	Dim hasEmployee : hasEmployee = IsObject(rsEmployee)
	if hasEmployee then hasEmployee = (rsEmployee.State = 1)
	if not hasEmployee then
%>
	<tr>
		<td colspan="34">PILIH KEBUN TERLEBIH DAHULU</td>
	</tr>
<%
	elseif rsEmployee.bof and rsEmployee.eof then
%>
	<tr>
		<td colspan="34">NO DATA AVAILABLE</td>
	</tr>
<%
	elseif Not rsEmployee.eof then
		noline = 1
		
		while Not rsEmployee.eof  %>
		           
		  <tr>
			<td align="center"><%=noline%></td>
			<td align="left"><%=rsEmployee("FCCODE")%></td>
			<td align="left"><%=rsEmployee("FCNAME")%></td>
			<td align="left"><%=rsEmployee("GENDER")%></td>
			<td align="left"><%=rsEmployee("DATEOFBIRTH")%></td>
			<td align="left"><%=rsEmployee("RELIGION")%></td>
			<td align="left"><%=rsEmployee("RACE")%></td>
			<td align="left"><%=rsEmployee("IDENTITYCARD")%></td>
			<td align="left"><%=rsEmployee("HIGHESTEDUCATION")%></td>
			<td align="left"><%=rsEmployee("MARITALSTATUS")%></td>
			<td align="left"><%=rsEmployee("SPOUSE")%></td>
			<td align="left"><%=rsEmployee("ADDRESS")%></td>
			<td align="left"><%=rsEmployee("NATIONALITY")%></td>
			<td align="left"><%=rsEmployee("COMPANY_NIK")%></td>
			<td align="left"><%=rsEmployee("SECTIONNAME")%></td>
			<td align="left"><%=rsEmployee("GANGCODE")%></td>
			<td align="left"><%=rsEmployee("DATEOFJOIN")%></td>
			<td align="left"><%=rsEmployee("PRESENTGRADE")%></td>
			<td align="left"><%=rsEmployee("SUBGRADE")%></td>
			<td align="left"><%=rsEmployee("DESCRIPTION")%></td>
			<td align="left"><%=rsEmployee("NATURETYPE")%></td>
			<td align="left"><%=rsEmployee("INCLUDEDTAXCALCULATION")%></td>
			<td align="left"><%=rsEmployee("TAXSTATUS")%></td>
			<td align="left"><%=rsEmployee("NPWP")%></td>
			<td align="left"><%=rsEmployee("TAXPAIDBYCOMPANY")%></td>
			<td align="left"><%=rsEmployee("PAYMENTTYPE")%></td>
			<td align="left"><%=rsEmployee("BANKNAME")%></td>
			<td align="left"><%=rsEmployee("OWNERSHIPOFACCOUNT")%></td>
			<td align="left"><%=rsEmployee("ACCOUNTNO")%></td>
			<td align="left"><%=rsEmployee("JAMSOSTEKCATEGORY")%></td>
			<td align="left"><%=rsEmployee("JAMSOSTEKNO")%></td>
			<td align="left"><%=rsEmployee("BPJSKES_CATEGORY")%></td>
			<td align="left"><%=rsEmployee("BPJSKES_NO")%></td>
			<td align="left"><%=rsEmployee("DATETERMINATE")%></td>
		  </tr>	
		      
    <%
			noline	=	noline + 1			
			rsEmployee.MoveNext
		wend
	end if
	if hasEmployee then rsEmployee.Close
	Set rsEmployee = Nothing
	%> 
</table>

   <table width="90%" class="data">
   	   <tr> 
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	   </tr>
	   <tr>
 		<td width="50%" align="center"><font face="Verdana" size="1">Approved by,</font></td>
		<td width="50%" align="center"><font face="Verdana" size="1">Prepared by,</font></td>
	   </tr>
	   <tr> 
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	   </tr>
	   <tr> 
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	   </tr>
	   <tr> 
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	   </tr>
	   <tr> 
		<td width="50%" align="center"><font face="Verdana" size="2"><strong><u><%=MgrName%></u></strong></td>
		<td width="50%" align="center"><font face="Verdana" size="2"><strong><u><%=KTUName%></u></strong></td>
	   </tr>
	   	  <tr> 
		<td width="50%" align="center"><font face="Verdana" size="1"><strong><i>MANAGER</i></strong></td>
		<td width="50%" align="center"><font face="Verdana" size="1"><strong><i>KTU</i></strong></td>
	   </tr>
       </tbody>
   </table>
   <script src="/js/jquery.inputmask.bundle.min.js"></script>
	<script src="/js/inputmask.binding.js"></script>
  	<script src="/plugins/chosen/chosen.jquery.min.js"></script>
	<script type="text/javascript">	
		// $(".chosen-select").empty();
		
		var config = {
		  '.chosen-select'           : {},
		  '.chosen-select-deselect'  : {allow_single_deselect:true},
		  '.chosen-select-no-single' : {disable_search_threshold:10},
		  '.chosen-select-no-results': {no_results_text:'Oops, nothing found!'},
		  '.chosen-select-width'     : {width:"45%"}
		}
		for (var selector in config) {
		  $(selector).chosen(config[selector]);
		}
		// $(".chosen-select").trigger("chosen:updated");
		
	</script>
 	<script type="text/javascript">
		function handleClick(cb){
			if(cb.checked == 1){
				cb.value = "Y";
			} else {
				cb.value = "N";
			}
        }
	</script>

<script>

function getselectedfcba() {
	var sfcba = document.getElementById("selectedfcba");
	sfcba.value = document.getElementById("fcba").value;
}

function getselecteddivisi() {
	var sdivisi = document.getElementById("selecteddivision");
	if (sdivisi && document.getElementById("divisi")) sdivisi.value = document.getElementById("divisi").value;   // dropdown divisi dimuat lewat AJAX
}

function getselectedgang() {
	var sgang = document.getElementById("selectedgang");
	if (sgang && document.getElementById("gangcode")) sgang.value = document.getElementById("gangcode").value;   // dropdown gang dimuat lewat AJAX
}

function getdeliverydestination() {
	var sdeliverydestination = document.getElementById("selecteddeliverydestination");
	sdeliverydestination.value = document.getElementById("deliverydestination").value;
}

Document.prototype.ready = function(callback) {
  if(callback && typeof callback === 'function') {
    document.addEventListener("DOMContentLoaded", function() {
      if(document.readyState === "interactive" || document.readyState === "complete") {
        return callback();
      }
    });
  }
};

// usage
document.ready(function() {
	getlistgangcode();
	getlistdivision();
	getselecteddivisi();
	getselectedgang();
});
</script>
   </body>
</html>
   </html>