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
	v_txtempcode = Request("txtempcode")
	txttdate = txtMonth & "" & txtYear
	v_txtPOCode		=	trim(request("txtPOCode"))
	compare1 = trim(request("compare1"))

	varDefaultBa = Session("fcBA")

	v_txtsectionname = request("txtsectionname")
	v_txtgangcode = request("txtgangcode")
	v_txtvehiclecode = request("txtvehiclecode")
	v_txtdatefrom = Request("txtdatefrom")
	v_txtdateuntil = Request("txtdateuntil")
	If (v_txtdatefrom = "" Or Len(Trim(v_txtdatefrom)) = 0) Then v_txtdatefrom = textdate(date)
	If (v_txtdateuntil = "" Or Len(Trim(v_txtdateuntil)) = 0) Then v_txtdateuntil = textdate(date)

	If (v_txtFcba = "" Or Len(Trim(v_txtFcba)) = 0) Then
		v_txtFcba = varDefaultBa
	End If	
		
       if  v_txtfcba <> "" then
	       whereCons = whereCons & "  AND fp.fcba = '" & v_txtfcba & "'"
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
	
		strGrnInv	=	" SELECT fcba," & _
						"        fddate," & _
						"        division," & _
						"        gangcode," & _ 
						"        presentgrade," & _
						"        employeecode," & _
						"		 employee_api.get_EmployeeName(fcba, employeecode) empname, " & _
						"        supervision_1," & _
						"		 employee_api.get_EmployeeName(fcba, supervision_1) supervision_1_name, " & _
						"        supervision_2," & _
						"		 employee_api.get_EmployeeName(fcba, supervision_2) supervision_2_name, " & _
						"        supervision_3," & _
						"		 employee_api.get_EmployeeName(fcba, supervision_3) supervision_3_name, " & _
						"        supervision_4," & _
						"		 employee_api.get_EmployeeName(fcba, supervision_4) supervision_4_name, " & _
						"        attendance," & _
						"        SUM(nvl(unit,0)) unit," & _
						"        SUM(nvl(quantity_bydaily,0)) quantity_bydaily," & _
						"        SUM(nvl(quantity_bymonthly,0)) quantity_bymonthly," & _
						"        AVG(nvl(basis,0)) basis," & _
						"        MAX(nvl(extraquantity_bydaily,0)) extraquantity_bydaily," & _
						"        MAX(ROUND(nvl(deductionqty_bydaily,0),5)) deductionqty_bydaily," & _
						"        MAX(nvl(extraamount_bydaily,0)) extraamount_bydaily," & _
						"        MAX(round(NVL(deductionamount_bydaily,0), 5)) deductionamount_bydaily," & _
						"        MAX(nvl(extraquantity_bymonthly,0)) extraquantity_bymonthly," & _
						"        MAX(ROUND(nvl(deductionqty_bymonthly,0),5)) deductionqty_bymonthly," & _
						"        MAX(nvl(extraamount_bymonthly,0)) extraamount_bymonthly," & _
						"        MAX(ROUND(nvl(deductionamount_bymonthly,0),5)) deductionamount_bymonthly," & _
						"        SUM(nvl(under_ripe,0)) under_ripe," & _
						"        SUM(nvl(under_ripe_amount,0)) under_ripe_amount" & _
						"   FROM harvesting_daily_bjr" & _
						"  WHERE fcba = '"& v_txtfcba &"'" 

	if  len(trim(v_txtdatefrom)) > 0 then
		strGrnInv = strGrnInv & " AND to_date (fddate,'mm/dd/yyyy') between TO_DATE('" & v_txtdatefrom & "','mm/dd/yyyy') and TO_DATE('" & v_txtdateuntil & "','fmMM/dd/yyyy')"
	end if
	'if  len(trim(v_txtdateuntil)) > 0 then
	'	strGrnInv = strGrnInv & " AND fddate <= to_Char(TO_DATE('" & v_txtdateuntil & "','fmMM/dd/yyyy'))"
	' end if
	if  len(trim(v_txtsectionname)) > 0 then
		strGrnInv = strGrnInv & " AND division ='"& v_txtsectionname &"'"
	end if
	if  len(trim(v_txtempcode)) > 0 then
		strGrnInv = strGrnInv & " AND employeecode ='"& v_txtempcode &"'"
	end if
	if  len(trim(v_txtgangcode)) > 0 then
		strGrnInv = strGrnInv & " AND gangcode ='"& v_txtgangcode &"'"
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

	strGrnInv	=	strGrnInv &	" GROUP BY fcba," & _
								"        fddate," & _
								"        division," & _
								"        gangcode," & _ 
								"        presentgrade," & _
								"        employeecode," & _
								"		 employee_api.get_EmployeeName(fcba, employeecode) , " & _
								"        supervision_1," & _
								"		 employee_api.get_EmployeeName(fcba, supervision_1) , " & _
								"        supervision_2," & _
								"		 employee_api.get_EmployeeName(fcba, supervision_2) , " & _
								"        supervision_3," & _
								"		 employee_api.get_EmployeeName(fcba, supervision_3) , " & _
								"        supervision_4," & _
								"		 employee_api.get_EmployeeName(fcba, supervision_4) , " & _
								"        attendance" 
								
	strGrnInv	=	strGrnInv & " ORDER BY fcba," & _
								"        fddate," & _ 
								"        division " 

	' response.write strGrnInv	
	Set rsGrnInv = Server.CreateObject("ADODB.Recordset")
	rsGrnInv.Open strGrnInv, connstr, adOpenStatic 
	
%>


<html>

<head>
<title>ePlaS Reporting - Harvesting Daily BJR</title>
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
						<td>
							<font face="verdana" size="1">Business Unit</font>
						</td>
						<td>
							<select name="txtFcba" id="vFcba" onchange="javascript:filtering()">
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
						</td>
						<td>
							<font face="verdana" size="1">From</font>
						</td>
						<td width="4">
							<input class="datepicker" type="text" maxlength="10" name="txtdatefrom" size="10" value="<%=v_txtdatefrom%>">
						</td>
						<td>
							<font face="verdana" size="1">To</font>
						</td>
						<td width="2">
							<input class="datepicker" type="text" maxlength="10" name="txtdateuntil" size="10" value="<%=v_txtdateuntil%>">
						</td>
						<td>
							<font face="verdana" size="1">Gang</font>
						</td>
						<td>
							<select name="txtgangcode" class="chosen-select" >
								<option value="">-- ALL --</option>
									<%
									strsectionname	=	"select fccode, fcname from gang where fcba = '"& v_txtFcba &"'order by fccode"
									
									Set rssectionname = Server.CreateObject("ADODB.Recordset")
									rssectionname.Open strsectionname, connstr, adOpenStatic 
									
									if not rssectionname.eof then
									Do While Not rssectionname.eof
									if rssectionname("fccode") = v_txtgangcode then%>
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
							<font face="verdana" size="1">Division</font>
						</td>
						<td>
							<select name="txtsectionname" >
								<option value="">-- ALL --</option>
									<%
									Dim rssectionname, strsectionname
									strsectionname	=	"select fccode, fcname from sectionname where fccode in ('AFD-01','AFD-02','AFD-03','AFD-04','AFD-05','AFD-06','AFD-07','AFD-08','AFD-09','AFD-10','AFD-11','AFD-12','AFD-13') order by fccode"
									
									Set rssectionname = Server.CreateObject("ADODB.Recordset")
									rssectionname.Open strsectionname, connstr, adOpenStatic 
									
									if not rssectionname.eof then
									Do While Not rssectionname.eof
									if rssectionname("fccode") = v_txtsectionname then%>
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
							<font face="verdana" size="1">Employee</font>
						</td>
						<td>
							<select name="txtempcode" class="chosen-select">
								<option value="">-- ALL --</option>
									<% 
									strsectionname	=	"select fccode, fcname from employee where fcba = '"& v_txtFcba &"' and dateterminate is null order by fccode"
									
									Set rssectionname = Server.CreateObject("ADODB.Recordset")
									rssectionname.Open strsectionname, connstr, adOpenStatic 
									
									if not rssectionname.eof then
									Do While Not rssectionname.eof
									if rssectionname("fccode") = v_txtempcode then%>
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
    <th colspan="26" style="background-color:#FFFFFF; border:none;">  
	<div id="header-wrapper">
		<div id="head_judul" style="float:left; text-align:left;">
				<h4>PT. <%=CompanyName%></h4>
				<h5><%=fcbaname%></h5>
			</div>
			
			<div id="subheader" class="centered" tyle="width:50%; margin:auto;">
				<h2>HARVESTING DAILY BJR</h2>
                <br/>
			</div>
		</div>
</th>
		</tr>
	<tr>
			<th style="background-color:#00A300;color:#fff;">Business Unit</th>
			<th style="background-color:#00A300;color:#fff;">Date</th>
			<th style="background-color:#00A300;color:#fff;">Gang Code</th>
			<th style="background-color:#00A300;color:#fff;">Division</th>
			<th style="background-color:#00A300;color:#fff;">Employee Code</th>
			<th style="background-color:#00A300;color:#fff;">Employee Name</th>
			<th style="background-color:#00A300;color:#fff;">Supervision 1</th>
			<th style="background-color:#00A300;color:#fff;">Supervision 2</th>
			<th style="background-color:#00A300;color:#fff;">Supervision 3</th>
			<th style="background-color:#00A300;color:#fff;">Supervision 4</th>
			<th style="background-color:#00A300;color:#fff;">Attendance</th>
			<th style="background-color:#00A300;color:#fff;">Present Grade</th>
			<th style="background-color:#00A300;color:#fff;">Unit</th>
			<th style="background-color:#00A300;color:#fff;">Quantity By Daily</th>
			<th style="background-color:#00A300;color:#fff;">Quantity By Monthly</th>
			<th style="background-color:#00A300;color:#fff;">Basis</th>
			<th style="background-color:#00A300;color:#fff;">Extra Quantity By Daily</th>
			<th style="background-color:#00A300;color:#fff;">Deduction Quantity By Daily</th>
			<th style="background-color:#00A300;color:#fff;">Extra Amount By Daily</th>
			<th style="background-color:#00A300;color:#fff;">Deduction Amount By Daily</th>
			<th style="background-color:#00A300;color:#fff;">Extra Quantity By Monthly</th>
			<th style="background-color:#00A300;color:#fff;">Deduction Quantity By Monthly</th>
			<th style="background-color:#00A300;color:#fff;">Extra Amount By Monthly</th>
			<th style="background-color:#00A300;color:#fff;">Deduction Amount By Monthly</th>
			<th style="background-color:#00A300;color:#fff;">Under Ripe</th>
			<th style="background-color:#00A300;color:#fff;">Under Ripe Amount</th>
		</tr>
    </thead>
	<tbody>
<%if rsGrnInv.BOF and rsGrnInv.EOF then%>
		<tr>
			<td colspan="26" align="left">NO DATA AVAILABLE</td>
		</tr>
<%end if
		NO = 1
		v_harvestingbased_abw = 0
		v_harvestingbased_abw_month = 0
		v_unit = 0
		v_quantity_bydaily = 0
		v_quantity_bymonthly = 0
		v_basis = 0
		v_extraquantity_bydaily = 0
		v_deductionqty_bydaily = 0
		v_extraamount_bydaily = 0
		v_deductionamount_bydaily = 0
		v_extraquantity_bymonthly = 0
		v_deductionqty_bymonthly = 0
		v_extraamount_bymonthly = 0
		v_deductionamount_bymonthly = 0
		v_under_ripe = 0 
		v_under_ripe_amount = 0
		while not rsGrnInv.eof
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
%>
		<tr class="<%=cClass%>">
			<td align="center"><%=rsGrnInv("fcBA")%></td>
			<td align="center"><%=rsGrnInv("fddate")%></td> 
			<td align="center"><%=rsGrnInv("gangcode")%></td>
			<td align="center"><%=rsGrnInv("division")%></td> 
			<td align="right"><%=rsGrnInv("employeecode")%></td>
			<td align="center"><%=rsGrnInv("empname")%></td>
			<td align="center"><%=rsGrnInv("supervision_1")%> : <%=rsGrnInv("supervision_1_name")%></td>
			<td align="center"><%=rsGrnInv("supervision_2")%> : <%=rsGrnInv("supervision_2_name")%></td>
			<td align="center"><%=rsGrnInv("supervision_3")%> : <%=rsGrnInv("supervision_3_name")%></td>
			<td align="center"><%=rsGrnInv("supervision_4")%> : <%=rsGrnInv("supervision_4_name")%></td>
			<td align="center"><%=rsGrnInv("attendance")%></td>
			<td align="right"><%=FbpNum2(rsGrnInv("presentgrade"))%></td> 
			<td align="right"><%=formatnumber(cdbl(to_nol(rsGrnInv("unit"))),2)%></td>
			<td align="right"><%=formatnumber(cdbl(to_nol(rsGrnInv("quantity_bydaily"))),2)%></td>
			<td align="right"><%=formatnumber(cdbl(to_nol(rsGrnInv("quantity_bymonthly"))),2)%></td>
			<td align="right"><%=formatnumber(cdbl(to_nol(rsGrnInv("basis"))),2)%></td>
			<td align="right"><%=formatnumber(cdbl(to_nol(rsGrnInv("extraquantity_bydaily"))),2)%></td>
			<td align="right"><%=formatnumber(cdbl(to_nol(rsGrnInv("deductionqty_bydaily"))),2)%></td>
			<td align="right"><%=formatnumber(cdbl(to_nol(rsGrnInv("extraamount_bydaily"))),2)%></td>
			<td align="right"><%=formatnumber(cdbl(to_nol(rsGrnInv("deductionamount_bydaily"))),2)%></td>
			<td align="right"><%=formatnumber(cdbl(to_nol(rsGrnInv("extraquantity_bymonthly"))),2)%></td>
			<td align="right"><%=formatnumber(cdbl(to_nol(rsGrnInv("deductionqty_bymonthly"))),2)%></td>
			<td align="right"><%=formatnumber(cdbl(to_nol(rsGrnInv("extraamount_bymonthly"))),2)%></td>
			<td align="right"><%=formatnumber(cdbl(to_nol(rsGrnInv("deductionamount_bymonthly"))),2)%></td>
			<td align="right"><%=formatnumber(cdbl(to_nol(rsGrnInv("under_ripe"))),2)%></td>
			<td align="right"><%=formatnumber(cdbl(to_nol(rsGrnInv("under_ripe_amount"))),2)%></td>
		</tr>
<%			 
			v_unit = v_unit + cdbl(to_nol(rsGrnInv("unit")))
			v_quantity_bydaily = v_quantity_bydaily + cdbl(to_nol(rsGrnInv("quantity_bydaily")))
			v_quantity_bymonthly = v_quantity_bymonthly + cdbl(to_nol(rsGrnInv("quantity_bymonthly")))
			v_basis = v_basis + cdbl(to_nol(rsGrnInv("basis")))
			v_extraquantity_bydaily = v_extraquantity_bydaily + cdbl(to_nol(rsGrnInv("extraquantity_bydaily")))
			v_deductionqty_bydaily = v_deductionqty_bydaily + cdbl(to_nol(rsGrnInv("deductionqty_bydaily")))
			v_extraamount_bydaily = v_extraamount_bydaily + cdbl(to_nol(rsGrnInv("extraamount_bydaily")))
			v_deductionamount_bydaily = v_deductionamount_bydaily + cdbl(to_nol(rsGrnInv("deductionamount_bydaily")))
			v_extraquantity_bymonthly = v_extraquantity_bymonthly + cdbl(to_nol(rsGrnInv("extraquantity_bymonthly")))
			v_deductionqty_bymonthly = v_deductionqty_bymonthly + cdbl(to_nol(rsGrnInv("deductionqty_bymonthly")))
			v_extraamount_bymonthly = v_extraamount_bymonthly + cdbl(to_nol(rsGrnInv("extraamount_bymonthly")))
			v_deductionamount_bymonthly = v_deductionamount_bymonthly + cdbl(to_nol(rsGrnInv("deductionamount_bymonthly")))
			v_under_ripe = v_under_ripe + cdbl(to_nol(rsGrnInv("under_ripe")))
			v_under_ripe_amount = v_under_ripe_amount + cdbl(to_nol(rsGrnInv("under_ripe_amount")))
			rsGrnInv.MoveNext
		Wend
	rsGrnInv.Close
	set rsGrnInv = Nothing
%>
	</tbody>
		<tr style="background-color: #F4F4A0">
			<td align="center" colspan="12"><font style="font-weight: bold;">TOTAL</font></td> 
			<td align="right"><font style="font-weight: bold;"><%=formatnumber(cdbl(to_nol(v_unit)),2)%></font></td>
			<td align="right"><font style="font-weight: bold;"><%=formatnumber(cdbl(to_nol(v_quantity_bydaily)),2)%></font></td>
			<td align="right"><font style="font-weight: bold;"><%=formatnumber(cdbl(to_nol(v_quantity_bymonthly)),2)%></font></td>
			<td align="right"><font style="font-weight: bold;"><%=formatnumber(cdbl(to_nol(v_basis)),2)%></font></td>
			<td align="right"><font style="font-weight: bold;"><%=formatnumber(cdbl(to_nol(v_extraquantity_bydaily)),2)%></font></td>
			<td align="right"><font style="font-weight: bold;"><%=formatnumber(cdbl(to_nol(v_deductionqty_bydaily)),2)%></font></td>
			<td align="right"><font style="font-weight: bold;"><%=formatnumber(cdbl(to_nol(v_extraamount_bydaily)),2)%></font></td>
			<td align="right"><font style="font-weight: bold;"><%=formatnumber(cdbl(to_nol(v_deductionamount_bydaily)),2)%></font></td>
			<td align="right"><font style="font-weight: bold;"><%=formatnumber(cdbl(to_nol(v_extraquantity_bymonthly)),2)%></font></td>
			<td align="right"><font style="font-weight: bold;"><%=formatnumber(cdbl(to_nol(v_deductionqty_bymonthly)),2)%></font></td>
			<td align="right"><font style="font-weight: bold;"><%=formatnumber(cdbl(to_nol(v_extraamount_bymonthly)),2)%></font></td>
			<td align="right"><font style="font-weight: bold;"><%=formatnumber(cdbl(to_nol(v_deductionamount_bymonthly)),2)%></font></td>
			<td align="right"><font style="font-weight: bold;"><%=formatnumber(cdbl(to_nol(v_under_ripe)),2)%></font></td>
			<td align="right"><font style="font-weight: bold;"><%=formatnumber(cdbl(to_nol(v_under_ripe_amount)),2)%></font></td>
		</tr>
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
