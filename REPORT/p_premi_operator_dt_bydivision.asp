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

	txttdate = txtMonth & "" & txtYear

	varDefaultBa = Session("fcBA")
	
	v_txtDivision		= Request("txtDivision")
	v_txtsectionname = request("txtsectionname")
	v_txtvehiclecode = request("txtvehiclecode")
	v_txtdatefrom = Request("txtdatefrom")
	v_txtdateuntil = Request("txtdateuntil")
	v_txtempcode = Request("txtempcode")
	If (v_txtdatefrom = "" Or Len(Trim(v_txtdatefrom)) = 0) Then v_txtdatefrom = textdate(date)
	If (v_txtdateuntil = "" Or Len(Trim(v_txtdateuntil)) = 0) Then v_txtdateuntil = textdate(date)

		
       ' if  v_txtfcba <> "" then
	       ' whereCons = whereCons & "  AND fp.fcba = '" & v_txtfcba & "'"
		' end if
	    ' selectCons = selectCons & ",fcba"
	    ' orderCons  = orderCons & ",BUSINESSUNIT_API.get_OrderData(fcba)"
		q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & session("fcba") & "' "
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
	
	' If (v_txtFcba = "" Or Len(Trim(v_txtFcba)) = 0) Then
		' v_txtFcba = varDefaultBa
	' End If	
	
	' if v_txtDivision <> "--ALL--"  And v_txtDivision <> "" then
		' whereCons = whereCons & " AND DIVISION = '" & v_txtDivision & "'"
	' else
		' nSpanTotal = nSpanTotal - 1  
	' end if   
	
	strGrnInv		=	" SELECT fcba," & _
						"        transdate," & _ 
						"        operatorcode," & _
						"        operatorname," & _
						"        division," & _
						"        kg_tbs," & _
						"        kg_brondolan," & _
						"        kg_pupuk," & _
						"        kg_lain_lain," & _
						"        total_kg," & _
						"        basis," & _
						"        kg_premi_lebih_basis," & _
						"        total_premi" & _
						"   FROM premi_operator_dt_bydivision" & _
						"  WHERE fcba = '"& session("fcba") &"'"  

	if  len(trim(v_txtdatefrom)) > 0 then
		strGrnInv = strGrnInv & " AND transdate >= TO_DATE('" & v_txtdatefrom & "', 'mm/dd/yyyy')"
	end if
	if  len(trim(v_txtdateuntil)) > 0 then
		strGrnInv = strGrnInv & " AND transdate <= TO_DATE('" & v_txtdateuntil & "', 'mm/dd/yyyy')"
	end if
	if  len(trim(v_txtvehiclecode)) > 0 then
		strGrnInv = strGrnInv & " AND vehiclecode ='"& v_txtvehiclecode &"'"
	end if 
	if  len(trim(v_txtempcode)) > 0 then
		strGrnInv = strGrnInv & " AND operatorcode ='"& v_txtempcode &"'"
	end if 
	if  len(trim(v_txtsectionname)) > 0 then
		strGrnInv = strGrnInv & " AND division ='"& v_txtsectionname &"'"
	end if  
	 

	' strGrnInv	=	strGrnInv &	" GROUP BY" & _
	' 							"     FCBA" & _
	' 							"     , NOPO" & _
	' 							"     , STOCKCODE"
								
	strGrnInv	=	strGrnInv & " ORDER BY fcba," & _
								"          transdate," & _
								"          division," & _
								"          operatorcode" 

	' response.write strGrnInv	
	Set rsGrnInv = Server.CreateObject("ADODB.Recordset")
	rsGrnInv.Open strGrnInv, connstr, adOpenStatic 
	
%>


<html>

<head>
<title>ePlaS Reporting - Premi Operator By Division</title>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<link rel="stylesheet" href="/css/jquery-ui.css" />
<script src="/js/jquery-1.9.1.js"></script>
<script src="/js/jquery-ui.js"></script>
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
							<font face="verdana" size="1">FROM</font>
						</td>
						<td width="4">
							<input class="datepicker" type="text" maxlength="10" name="txtdatefrom" size="10" value="<%=v_txtdatefrom%>">
						</td>
						<td>
							<font face="verdana" size="1">TO</font>
						</td>
						<td width="2">
							<input class="datepicker" type="text" maxlength="10" name="txtdateuntil" size="10" value="<%=v_txtdateuntil%>">
						</td>
						<td>
							<font face="verdana" size="1">DIVISON</font>
						</td>
						<td>
							<select name="txtsectionname" >
								<option value="">-- ALL --</option>
									<%
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
							<font face="verdana" size="1">VEHICLE CODE</font>
						</td>
						<td>
							<select name="txtvehiclecode" >
								<option value="">--ALL--</option>
									<%
									Dim rssectionname, strsectionname
									strsectionname	=	" select fccode, fcname from vehicle where fcba = '"& session("fcba") &"' and activation = 'Y' order by fccode"
									
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
							<font face="verdana" size="1">EMPLOYEE</font>
						</td>
						<td>
							<select name="txtempcode" >
								<option value="">--ALL--</option>
									<% 
									strsectionname	=	"select fccode, fcname from employee where fcba = '"& session("fcba") &"' and dateterminate is null order by fccode"
									
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
    <th colspan="12" style="background-color:#FFFFFF; border:none;">  
		<div id="header-wrapper">
			<div id="head_judul" style="float:left; text-align:left;">
				<h4>PT. <%=CompanyName%></h4>
				<h5><%=fcbaname%></h5>
			</div>	
			<div id="subheader" class="centered" tyle="width:50%; margin:auto;">
				<h2>PREMI OPERATOR DT</h2>
				<h3>BY DIVISION</h3>
				<br/>
			</div>
		</div>
	</th>
		</tr>
	<tr>
			<th style="background-color:#00A300;color:#fff;">Date</th> 
			<th style="background-color:#00A300;color:#fff;">Operator Code</th>
			<th style="background-color:#00A300;color:#fff;">Operator Name</th>
			<th style="background-color:#00A300;color:#fff;">Division</th>
			<th style="background-color:#00A300;color:#fff;">KG TBS</th>
			<th style="background-color:#00A300;color:#fff;">KG Brondolan</th>
			<th style="background-color:#00A300;color:#fff;">KG Pupuk</th>
			<th style="background-color:#00A300;color:#fff;">KG Lain - Lain</th>
			<th style="background-color:#00A300;color:#fff;">Total KG</th>
			<th style="background-color:#00A300;color:#fff;">Basis</th>
			<th style="background-color:#00A300;color:#fff;">KG Premi Lebih Basis</th>
			<th style="background-color:#00A300;color:#fff;">Total Premi</th>
		</tr>
    </thead>
	<tbody>
<%if rsGrnInv.BOF and rsGrnInv.EOF then%>
		<tr>
			<td colspan="13" align="left">NO DATA AVAILABLE</td>
		</tr>
<%end if
		NO = 1
		v_kg_tbs = 0
		v_kg_brondolan = 0
		v_kg_pupuk = 0
		v_kg_lain_lain = 0
		v_total_kg = 0
		v_basis = 0 
		v_kg_premi_lebih_basis = 0
		v_total_premi = 0
		while not rsGrnInv.eof
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
%>
		<tr class="<%=cClass%>">
			<td align="center"><%=rsGrnInv("transdate")%></td> 
			<td align="center"><%=rsGrnInv("operatorcode")%></td>
			<td align="center"><%=rsGrnInv("operatorname")%></td>
			<td align="center"><%=rsGrnInv("division")%></td>
			<td align="right"><%=formatnumber(cdbl(to_nol(rsGrnInv("kg_tbs"))),2)%></td>
			<td align="right"><%=formatnumber(cdbl(to_nol(rsGrnInv("kg_brondolan"))),2)%></td>
			<td align="right"><%=formatnumber(cdbl(to_nol(rsGrnInv("kg_pupuk"))),2)%></td>
			<td align="right"><%=formatnumber(cdbl(to_nol(rsGrnInv("kg_lain_lain"))),2)%></td>
			<td align="right"><%=formatnumber(cdbl(to_nol(rsGrnInv("total_kg"))),2)%></td>
			<td align="right"><%=formatnumber(cdbl(to_nol(rsGrnInv("basis"))),2)%></td>
			<td align="right"><%=formatnumber(cdbl(to_nol(rsGrnInv("kg_premi_lebih_basis"))),2)%></td>
			<td align="right"><%=formatnumber(cdbl(to_nol(rsGrnInv("total_premi"))),2)%></td>
		</tr>
<%			
			v_kg_tbs = v_kg_tbs + cdbl(to_nol(rsGrnInv("kg_tbs")))
			v_kg_brondolan = v_kg_brondolan + cdbl(to_nol(rsGrnInv("kg_brondolan")))
			v_kg_pupuk = v_kg_pupuk + cdbl(to_nol(rsGrnInv("kg_pupuk")))
			v_kg_lain_lain = v_kg_lain_lain + cdbl(to_nol(rsGrnInv("kg_lain_lain")))
			v_total_kg = v_total_kg + cdbl(to_nol(rsGrnInv("total_kg")))
			v_basis = v_basis + cdbl(to_nol(rsGrnInv("basis")))
			v_kg_premi_lebih_basis = v_kg_premi_lebih_basis + cdbl(to_nol(rsGrnInv("kg_premi_lebih_basis")))
			v_total_premi = v_total_premi + cdbl(to_nol(rsGrnInv("total_premi")))
			rsGrnInv.MoveNext
		Wend
	rsGrnInv.Close
	set rsGrnInv = Nothing
%>
	</tbody>
		<tr style="background-color: #F4F4A0">
			<td align="center" colspan="4"><font style="font-weight: bold;">TOTAL</font></td>
			<td align="right"><font style="font-weight: bold;"><%=formatnumber(cdbl(to_nol(v_kg_tbs)),2)%></font></td>
			<td align="right"><font style="font-weight: bold;"><%=formatnumber(cdbl(to_nol(v_kg_brondolan)),2)%></font></td>
			<td align="right"><font style="font-weight: bold;"><%=formatnumber(cdbl(to_nol(v_kg_pupuk)),2)%></font></td>
			<td align="right"><font style="font-weight: bold;"><%=formatnumber(cdbl(to_nol(v_kg_lain_lain)),2)%></font></td>
			<td align="right"><font style="font-weight: bold;"><%=formatnumber(cdbl(to_nol(v_total_kg)),2)%></font></td>
			<td align="right"><font style="font-weight: bold;"><%=formatnumber(cdbl(to_nol(v_basis)),2)%></font></td> 
			<td align="right"><font style="font-weight: bold;"><%=formatnumber(cdbl(to_nol(v_kg_premi_lebih_basis)),2)%></font></td>
			<td align="right"><font style="font-weight: bold;"><%=formatnumber(cdbl(to_nol(v_total_premi)),2)%></font></td> 
		</tr>
</table>

<!--#include file="../Bottom_Report.asp"-->
