<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
	Dim v_txtYear, v_txtMonth
	Dim ScriptName
	Dim rsGrnInv, strGrnInv
	Dim v_txtPOCode
	Dim NO, cClass
	
	ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

	If "" & txtYear = "" Then txtYear = Year(Session("v_Costbook"))
	If "" & txtMonth = "" Then txtMonth = Month(Session("v_Costbook"))
	If "" & txttdate = "" Then txttdate = textDate(date)
	v_txtFcba = Request("txtFcba")
	v_txtFcba2 = Request("txtFcba2")

	txttdate = txtMonth & "" & txtYear
	v_txtPOCode		=	trim(request("txtPOCode"))
	compare1 = trim(request("compare1"))

	varDefaultBa = Session("fcBA")

	v_txtsectionname = request("txtsectionname")
	v_txtvehiclecode = request("txtvehiclecode")
	v_txtdatefrom = Request("txtdatefrom")
	v_txtdateuntil = Request("txtdateuntil")
	v_txtown = Request("txtown")
	v_txtempcode = Request("txtempcode")
	If (v_txtdatefrom = "" Or Len(Trim(v_txtdatefrom)) = 0) Then v_txtdatefrom = textdate(date)
	If (v_txtdateuntil = "" Or Len(Trim(v_txtdateuntil)) = 0) Then v_txtdateuntil = textdate(date)

	If (v_txtFcba2 = "" Or Len(Trim(v_txtFcba2)) = 0) Then
		v_txtFcba2 = varDefaultBa
	End If	
		
       if  v_txtfcba <> "" then
	       whereCons = whereCons & "  AND fp.fcba = '" & v_txtfcba & "'"
		end if
	    ' selectCons = selectCons & ",fcba"
	    ' orderCons  = orderCons & ",BUSINESSUNIT_API.get_OrderData(fcba)"
		q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & v_txtfcba2 & "' "
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
	'Update Feb 21 AMH-Data Double jd di tambah DISTINCT cek export datanya...
	
	strGrnInv		=	" SELECT DISTINCT fcba_mill," & _
						"        fcba_estate," & _
						"        spbno," & _
						"        chitno," & _
						"        vehiclecode," & _
						"        vehicleno," & _
						"        driver," & _
						"        driver_name as drivername," & _
						"        datein," & _
						"        timein," & _
						"        timeout," & _
						"        tbstype," & _
						"        estate_bunches," & _
						"        mill_bunches," & _
						"        bunch_mill_weight_bruto," & _
						"        kg_potongan," & _
						"        bunch_mill_weight_netto," & _
						"        bjr_per_chitno," & _
						"        total_bucket," & _
						"        kg_brondolan," & _
						"        kg_per_bucket" & _
						"   FROM wb_harvesting" & _
						" WHERE fcba_mill = '"& v_txtfcba &"'"
  
	if  len(trim(v_txtfcba2)) > 0 then
		strGrnInv = strGrnInv & " AND fcba_estate = '"& v_txtfcba2 &"'"
	end if
	if  len(trim(v_txtdatefrom)) > 0 then
		strGrnInv = strGrnInv & " AND datein >= TO_DATE('" & v_txtdatefrom & "', 'mm/dd/yyyy')"
	end if
	if  len(trim(v_txtdateuntil)) > 0 then
		strGrnInv = strGrnInv & " AND datein <= TO_DATE('" & v_txtdateuntil & "', 'mm/dd/yyyy')"
	end if
	if  len(trim(v_txtsectionname)) > 0 then
		strGrnInv = strGrnInv & " AND spbno like '%"& v_txtsectionname &"%'"
	end if
	if  len(trim(v_txtvehiclecode)) > 0 then
		strGrnInv = strGrnInv & " AND vehiclecode ='"& v_txtvehiclecode &"'"
	end if 
	if  len(trim(v_txtown)) > 0 then
		strGrnInv = strGrnInv & " AND own ='"& v_txtown &"'"
	end if 
	if  len(trim(v_txtempcode)) > 0 then
		strGrnInv = strGrnInv & " AND driver ='"& v_txtempcode &"'"
	end if 


	' if  cdbl(to_nol(compare1)) = 1 then
	'     strGrnInv	=	strGrnInv & " and attendance_gad.gad_employeecode != vehicleactivity.vh_employeecode"
	' end if
	' if  cdbl(to_nol(compare1)) = 2 then
	'     strGrnInv	=	strGrnInv & " and attendance_gad.gad_employeecode != harvestingspb.harvesting_employeecode" 
	' end if
	' if  cdbl(to_nol(compare1)) = 3 then
	'     strGrnInv	=	strGrnInv & " and vehicleactivity.vh_employeecode != harvestingspb.harvesting_employeecode" 
	' end if

	' strGrnInv	=	strGrnInv &	" GROUP BY" & _
	' 							"     FCBA" & _
	' 							"     , NOPO" & _
	' 							"     , STOCKCODE"
								
	strGrnInv	=	strGrnInv & " ORDER BY spbno," & _
								"          chitno," & _
								"          datein," & _
								"          tbstype" 

	' response.write strGrnInv	
	Set rsGrnInv = Server.CreateObject("ADODB.Recordset")
	rsGrnInv.Open strGrnInv, connstr, adOpenStatic 
	
%>


<html>

<head>
<title>ePlaS Reporting - Weighbridge Harvesting</title>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<link rel="stylesheet" href="/css/jquery-ui.css" />
<script src="/js/jquery-1.9.1.js"></script>
<script src="/js/jquery-ui.js"></script>
<link href="/plugins/tabcontent/tabcontent.css" rel="stylesheet" type="text/css" />
<link href="/plugins/chosen/chosen.min.css" rel="stylesheet" />

<link rel="stylesheet" href="/plugins/responsivegridsystem/col.css" media="all">
<link rel="stylesheet" href="/plugins/responsivegridsystem/12cols.css" media="all">
<script language="JavaScript" src='/js/x.js'></script>
<script language="JavaScript" src='/js/xtableheaderfixed.js'></script>
<script language="JavaScript">
xAddEventListener(window, 'load',
  function() {
    new xTableHeaderFixed('xthf-site', window);
    new xTableHeaderFixed('xthf-green', window);
    new xTableHeaderFixed('xthf-blue', window);
  }, false
);
</script>
 <script>
$(function() {
$( ".datepicker" ).datepicker();
});
</script>
<!--<script language="JavaScript" src="/js/date-picker.js"></script> 
<script language="JavaScript">

var newwin;
function launchwin(winurl,winname,winfeatures)
{
	//This launches a new window and then
	//focuses it if window.focus() is supported.
	newwin = window.open(winurl,winname,winfeatures);

}
</script>-->

<style type="text/css">
@media print
{
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
			</div>
            	<form action="<%=ScriptName%>" method="post" name="frmEntry">
			<div class="printhide">
				<table id="form-table" class="centered-margin">
					<tr>
						<td align="center">
							<font face="verdana" size="1" style="font-weight: bold;">PABRIK</font>
						</td>
						<td align="center">
							<font face="verdana" size="1" style="font-weight: bold;">KEBUN</font>
						</td>
						<td align="center">
							<font face="verdana" size="1" style="font-weight: bold;">DARI</font>
						</td>
						<td align="center">
							<font face="verdana" size="1" style="font-weight: bold;">SAMPAI</font>
						</td>
						<td align="center">
							<font face="verdana" size="1" style="font-weight: bold;">DIVISION</font>
						</td>
						<td align="center">
							<font face="verdana" size="1" style="font-weight: bold;">VEHICLE CODE</font>
						</td>
						<td align="center">
							<font face="verdana" size="1" style="font-weight: bold;">OWN</font>
						</td>
						<td align="center">
							<font face="verdana" size="1" style="font-weight: bold;">EMPLOYEE CODE</font>
						</td>
					</tr>
					<tr>						
						<td>
							<select name="txtFcba" id="vFcba" onchange="javascript:filtering()">
							<%
							Dim rsFcba
							Set rsFcba = Server.CreateObject("ADODB.Recordset")
							rsFcba.CursorLocation = adUseClient
							strsql =   " select fccode fcba from businessunit " & _
										" where  fctype ='M' and fccode not in ('PBSU','PMKS')"  
							' response.write strsql			
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
						</td> 
						<td>
							<select name="txtFcba2" id="vFcba2" onchange="javascript:filtering()">
							<option></option>
							<%
							Set rsFcba = Server.CreateObject("ADODB.Recordset")
							rsFcba.CursorLocation = adUseClient
							strsql =   " select fccode fcba from businessunit " & _
										" where " & _
										" USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',fccode) IS NOT NULL " 
							  
							
							rsFcba.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
							if not rsFcba.eof then
								DO while not rsFcba.eof
									if v_txtFcba2 = rsFcba("fcba") then
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
						</td> 
						<td width="4">
							<input class="datepicker" type="text" maxlength="10" name="txtdatefrom" size="10" value="<%=v_txtdatefrom%>">
						</td> 
						<td width="2">
							<input class="datepicker" type="text" maxlength="10" name="txtdateuntil" size="10" value="<%=v_txtdateuntil%>">
						</td>  
						<td>
							<select name="txtsectionname" >
								<option value="">--ALL--</option>
									<% 
									strsectionname	=	" select fccode from parameter where fcba = '"& v_txtFcba2 &"' and parcode = 'RPT_WB_HARVESTING' order by fccode"
									
									Set rssectionname = Server.CreateObject("ADODB.Recordset")
									rssectionname.Open strsectionname, connstr, adOpenStatic 
									
									if not rssectionname.eof then
									Do While Not rssectionname.eof
									if rssectionname("fccode") = v_txtsectionname then%>
											<option value="<%=rssectionname("fccode")%>" Selected><%=rssectionname("fccode")%> </option> 
										 <%Else%>
											<option value="<%=rssectionname("fccode")%>"><%=rssectionname("fccode")%> </option>  
										<%End if
										rssectionname.movenext
									loop
									end if
									rssectionname.Close
									set rssectionname = Nothing
									%>
							</select>
						</td>  
						<td>
							<select name="txtvehiclecode" >
								<option value="">--ALL--</option>
									<%
									Dim rssectionname, strsectionname
									strsectionname	=	" select fccode, fcname from vehicle where fcba = '"& v_txtFcba2 &"' and activation = 'Y' order by fccode"
									
									Set rssectionname = Server.CreateObject("ADODB.Recordset")
									rssectionname.Open strsectionname, connstr, adOpenStatic 
									
									if not rssectionname.eof then
									Do While Not rssectionname.eof
									if rssectionname("fccode") = v_txtvehiclecode then%>
											<option value="<%=rssectionname("fccode")%>" Selected><%=rssectionname("fccode")%> : <%=rssectionname("fcname")%></option> 
										 <%Else%>
											<option value="<%=rssectionname("fccode")%>"><%=rssectionname("fccode")%> : <%=rssectionname("fcname")%></option>  
										<%End if
										rssectionname.movenext
									loop
									end if
									rssectionname.Close
									set rssectionname = Nothing
									%>
							</select>
						</td> 
						<td>
							<select name="txtown" >
								<option value="">--ALL--</option>
									<% 
									strsectionname	=	" select distinct own from harvestingspb "
									
									Set rssectionname = Server.CreateObject("ADODB.Recordset")
									rssectionname.Open strsectionname, connstr, adOpenStatic 
									
									if not rssectionname.eof then
									Do While Not rssectionname.eof
									if rssectionname("own") = v_txtown then%>
											<option value="<%=rssectionname("own")%>" Selected><%=rssectionname("own")%></option> 
										 <%Else%>
											<option value="<%=rssectionname("own")%>"><%=rssectionname("own")%> </option>  
										<%End if
										rssectionname.movenext
									loop
									end if
									rssectionname.Close
									set rssectionname = Nothing
									%>
							</select>
						</td> 
						<td>
							<select name="txtempcode" class="chosen-select" >
								<option value="">--ALL--</option>
									<% 
									strsectionname	=	" select fccode, fcname from employee where fcba = '"& v_txtfcba2 &"' order by fcCode"
									
									Set rssectionname = Server.CreateObject("ADODB.Recordset")
									rssectionname.Open strsectionname, connstr, adOpenStatic 
									
									if not rssectionname.eof then
									Do While Not rssectionname.eof
									if rssectionname("fccode") = v_txtempcode then%>
											<option value="<%=rssectionname("fccode")%>" Selected><%=rssectionname("fccode")%> : <%=rssectionname("fcname")%></option> 
										 <%Else%>
											<option value="<%=rssectionname("fccode")%>"><%=rssectionname("fccode")%> : <%=rssectionname("fcname")%> </option>  
										<%End if
										rssectionname.movenext
									loop
									end if
									rssectionname.Close
									set rssectionname = Nothing
									%>
							</select>
						</td>
						<!--
						<td>
							<font face="verdana" size="1">VEHICLE CODE</font>
						</td>
						<td width="2">
							<input type="text" name="txtvehiclecode" value="<%=txtvehiclecode%>">
						</td> -->
						<td>
							<input type="submit" value="View" />
						</td>
					</tr>
				</table>
			</div>
			</form>
			
			</div>
<table border="0" width="100%" id="product-table" class="xthf-site">
<thead>
	<tr>
    <th colspan="21" style="background-color:#FFFFFF; border:none;">  
	<div id="header-wrapper">
		<div id="head_judul" style="float:left; text-align:left;">
				<h4>PT. <%=CompanyName%></h4>
				<h5><%=fcbaname%></h5>
			</div>
			
			<div id="subheader" class="centered" tyle="width:50%; margin:auto;">
				<h2>WEIGHBRIDGE HARVESTING</h2>
                <br/>
			</div>
		</div>
</th>
		</tr>
	<tr>
			<th style="background-color:#00A300;color:#fff;">Pabrik</th>
			<th style="background-color:#00A300;color:#fff;">Kebun</th>
			<th style="background-color:#00A300;color:#fff;">SPB No</th>
			<th style="background-color:#00A300;color:#fff;">Chitno</th>
			<th style="background-color:#00A300;color:#fff;">Driver Code</th>
			<th style="background-color:#00A300;color:#fff;">Driver Name</th>
			<th style="background-color:#00A300;color:#fff;">Vehicle Code</th>
			<th style="background-color:#00A300;color:#fff;">Vehicle No</th>
			<th style="background-color:#00A300;color:#fff;">Date In</th>
			<th style="background-color:#00A300;color:#fff;">Time In</th>
			<th style="background-color:#00A300;color:#fff;">Time Out</th>
			<th style="background-color:#00A300;color:#fff;">TBS Type</th>
			<th style="background-color:#00A300;color:#fff;">Estate Bunches</th>
			<th style="background-color:#00A300;color:#fff;">Mill Bunches</th>
			<th style="background-color:#00A300;color:#fff;">Bunch Mill Weight Bruto</th>
			<th style="background-color:#00A300;color:#fff;">KG Potongan</th>
			<th style="background-color:#00A300;color:#fff;">Bunch Mill Weight Netto</th>
			<th style="background-color:#00A300;color:#fff;">BJR Per Chitno</th>
			<th style="background-color:#00A300;color:#fff;">Total Bucket</th>
			<th style="background-color:#00A300;color:#fff;">Kg Brondolan</th>
			<th style="background-color:#00A300;color:#fff;">KG Per Bucket</th>
		</tr>
    </thead>
	<tbody>
<%if rsGrnInv.BOF and rsGrnInv.EOF then%>
		<tr>
			<td colspan="21" align="left">NO DATA AVAILABLE</td>
		</tr>
<%end if
		NO = 1
		v_estate_bunches = 0
		v_mill_bunches = 0
		v_bunch_mill_weight_bruto = 0
		v_kg_potongan = 0
		v_bunch_mill_weight_netto = 0
		v_bjr_per_chitno = 0
		v_total_bucket = 0
		v_kg_brondolan = 0
		v_kg_per_bucket = 0
		while not rsGrnInv.eof
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
%>
		<tr class="<%=cClass%>">
			<td align="center"><%=rsGrnInv("fcba_mill")%></td>
			<td align="center"><%=rsGrnInv("fcba_estate")%></td>
			<td align="left"><%=rsGrnInv("spbno")%></td>
			<td align="left"><%=rsGrnInv("chitno")%></td>
			<td align="center"><%=rsGrnInv("driver")%></td>
			<td align="center"><%=rsGrnInv("drivername")%></td>
			<td align="center"><%=rsGrnInv("vehiclecode")%></td>
			<td align="center"><%=rsGrnInv("vehicleno")%></td>
			<td align="right"><%=rsGrnInv("datein")%></td>
			<td align="right"><%=rsGrnInv("timein")%></td>
			<td align="right"><%=rsGrnInv("timeout")%></td>
			<td align="right"><%=rsGrnInv("tbstype")%></td>
			<td align="right"><%=formatnumber(cdbl(to_nol(rsGrnInv("estate_bunches"))),2)%></td>
			<td align="right"><%=formatnumber(cdbl(to_nol(rsGrnInv("mill_bunches"))),2)%></td>
			<td align="right"><%=formatnumber(cdbl(to_nol(rsGrnInv("bunch_mill_weight_bruto"))),2)%></td>
			<td align="right"><%=formatnumber(cdbl(to_nol(rsGrnInv("kg_potongan"))),2)%></td>
			<td align="right"><%=formatnumber(cdbl(to_nol(rsGrnInv("bunch_mill_weight_netto"))),2)%></td>
			<td align="right"><%=formatnumber(cdbl(to_nol(rsGrnInv("bjr_per_chitno"))),2)%></td>
			<td align="right"><%=formatnumber(cdbl(to_nol(rsGrnInv("total_bucket"))),2)%></td>
			<td align="right"><%=formatnumber(cdbl(to_nol(rsGrnInv("kg_brondolan"))),2)%></td>
			<td align="right"><%=formatnumber(cdbl(to_nol(rsGrnInv("kg_per_bucket"))),2)%></td>
		</tr>
<%			
			v_estate_bunches = v_estate_bunches + cdbl(to_nol(rsGrnInv("estate_bunches")))
			v_mill_bunches = v_mill_bunches + cdbl(to_nol(rsGrnInv("mill_bunches")))	
			v_bunch_mill_weight_bruto = v_bunch_mill_weight_bruto + cdbl(to_nol(rsGrnInv("bunch_mill_weight_bruto")))
			v_kg_potongan = v_kg_potongan + cdbl(to_nol(rsGrnInv("kg_potongan")))
			v_bunch_mill_weight_netto = v_bunch_mill_weight_netto + cdbl(to_nol(rsGrnInv("bunch_mill_weight_netto")))	
			v_bjr_per_chitno =v_bjr_per_chitno + cdbl(to_nol(rsGrnInv("bjr_per_chitno")))	
			v_total_bucket = v_total_bucket + cdbl(to_nol(rsGrnInv("total_bucket")))	
			v_kg_brondolan = v_kg_brondolan + cdbl(to_nol(rsGrnInv("kg_brondolan")))	
			v_kg_per_bucket = v_kg_per_bucket + cdbl(to_nol(rsGrnInv("kg_per_bucket")))	

			rsGrnInv.MoveNext
		Wend
	rsGrnInv.Close
	set rsGrnInv = Nothing
%>
		<tr bgcolor="#F4F4A0"> 
			<td align="center" colspan="12"><font style="font-weight: bold;">TOTAL</font></td>				
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_estate_bunches)),2)%></font></td> 
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_mill_bunches)),2)%></font></td>  
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_bunch_mill_weight_bruto)),2)%></font></td> 
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_kg_potongan)),2)%></font></td> 
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_bunch_mill_weight_netto)),2)%></font></td> 
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_bjr_per_chitno)),2)%></font></td> 
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_total_bucket)),2)%></font></td> 
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_kg_brondolan)),2)%></font></td> 
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_kg_per_bucket)),2)%></font></td> 
		</tr>
	</tbody>
</table>

<!--#include file="../Bottom_Report.asp"-->
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