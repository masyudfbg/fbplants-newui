<%@Language=VBScript%>
<%Option Explicit%>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
	Dim v_txtBudgetYear, v_txtBudgetVersion, v_txtModuleVersion, v_txtdatefrom, v_txtMonth, v_txtYear, v_txtgrade, v_txtDivisi, v_txtemployeetype, v_txtjobcode
	Dim no
	Dim rsDefaultBa, strDefaultBa, varDefaultBa, rssearchfirstba, strsearchfirstba
	Dim ScriptName
	Dim strMain, sqlLHPP, sqlLHPP2, rsLHPP
	Dim q, CompanyName, fcbaname, MgrName, KTUName, CropName, FcbaType
	Dim OrderSection : OrderSection = 0
	Dim TotalMandays : TotalMandays = 0

    v_txtGroupPlantation = Request("txtGroupPlantation")
	v_txtPlantation = Request("txtPlantation")
	v_txtGroupRegion = Request("txtGroupRegion")
	v_txtRegion = Request("txtRegion")
	v_txtcompany = Request("txtCompany")
	v_txtfcba = Request("txtfcba")
	
	v_txtDivisi = Request("txtDivisi")
	v_txtjobcode = Request("txtjobcode")
	v_txtemployeetype = Request("txtemployeetype")
	
	v_txtBudgetYear = Request("txtBudgetYear")
	v_txtBudgetVersion = Request("txtBudgetVersion")
	v_txtModuleVersion = Request("txtModuleVersion")
	
	v_txtdatefrom		= Request("txtdatefrom")
	
	v_txtYear		= Request("txtYear")
	v_txtMonth		= Request("txtMonth")
	v_txtgrade		= Request("txtgrade")
	
	If Len(Trim(v_txtYear)) = 0 Then v_txtYear = Year(Session("v_Costbook"))
	If Len(Trim(v_txtMonth)) = 0 Then v_txtMonth = Month(Session("v_Costbook"))
	
	If Len(Trim(v_txtBudgetYear)) = 0 Then
		v_txtBudgetYear = Year(textDate(Date))
	Else
		v_txtBudgetYear = v_txtYear
	End If
	If Len(Trim(v_txtBudgetVersion)) = 0 Then v_txtBudgetVersion = 1
	If Len(Trim(v_txtModuleVersion)) = 0 Then v_txtModuleVersion = 1
	
	If Len(Trim(v_txtdatefrom)) = 0 Then v_txtdatefrom = textDate(Date)
	
	
	strDefaultBa = " SELECT fccode as fcba" & _
				   " FROM businessunit bu" & _
				   " INNER JOIN user_fcba_tab uft on (bu.fccode = uft.fcba) " & _
				   " WHERE isdefault = 'TRUE'" & _
				   "       AND fcusercode = '" & Session("CodeName") & "'" & _
				   "       AND fctype = 'M'"

	Set rsDefaultBa = Server.CreateObject("ADODB.Recordset")
	rsDefaultBa.Open strDefaultBa, connstr, adOpenStatic

	If Not rsDefaultBa.EOF Then
		varDefaultBa = rsDefaultBa("fcba")
	Else
	
		strsearchfirstba = " SELECT fcba" & _
		                   " FROM (" & _
		                   " SELECT fccode AS fcba" & _
		                   "     FROM businessunit bu" & _
		                   "     INNER JOIN user_fcba_tab uft ON (bu.fccode = uft.fcba) " & _
		                   "     WHERE fcusercode = '" & Session("CodeName") & "'" & _
		                   "           AND fctype = 'M'" & _
		                   "     ORDER BY orderdata" & _
		                   " ) WHERE ROWNUM = 1"
		Set rssearchfirstba = Server.CreateObject("ADODB.Recordset")
		rssearchfirstba.Open strsearchfirstba, connstr, adOpenStatic
		
		IF NOT rssearchfirstba.EOF THEN
			varDefaultBa = rssearchfirstba("fcba")
		ELSE
			varDefaultBa = Session("fcba")
		END IF
		rssearchfirstba.Close
		set rssearchfirstba = Nothing
		
	End If
	rsDefaultBa.Close
	set rsDefaultBa = Nothing

	if not isnull(v_txtFcba) then
		If (v_txtFcba = "" Or Len(Trim(v_txtFcba)) = 0) Then
			v_txtFcba = varDefaultBa
		End If
	End If
	
	Dim var_colspan : var_colspan = 1
	
Dim selectFields, strFieldCons, selectCons, whereCons, GroupbyCons, orderCons

strFieldCons	=	" USER_COMPANY_API.GetGroupPlantationCode('" & Session("CodeName") & "',fcba) GroupPlantation" & _
					" , USER_COMPANY_API.GetPlantationCode('" & Session("CodeName") & "',fcba) Plantation" & _
					" , USER_COMPANY_API.GetGroupRegionCode('" & Session("CodeName") & "',fcba) groupregion" & _
					" , USER_COMPANY_API.GetRegionCode('" & Session("CodeName") & "',fcba) region" & _
					" , USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',fcba) company" & _
					" , BUSINESSUNIT_API.get_OrderData(fcba) NoUrut, fcba"
					
wherecons	=	" and businessunit_api.get_fctype(fcba) = 'M' AND USER_COMPANY_API.IsAuthorizedFcba('" & Session("CodeName") & "', fcba) =  'TRUE'"

If Len(Trim(v_txtGroupPlantation)) > 0 Then
	if v_txtGroupPlantation = "--SUMMARY--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
	Elseif v_txtGroupPlantation = "--ALL--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
		var_colspan	=	var_colspan	+ 1
	Else
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " And GroupPlantation = '" & v_txtGroupPlantation & "'"
		orderCons	=	orderCons	& " "
		var_colspan	=	var_colspan	+ 1
	End If
End If

If Len(Trim(v_txtPlantation)) > 0 Then
	if v_txtPlantation = "--SUMMARY--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
	Elseif v_txtPlantation = "--ALL--" then
		selectCons	=	selectCons	& " , Plantation"
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " , PLANTATION_API.get_OrderData(Plantation)"
		var_colspan	=	var_colspan	+ 1
	Else
		selectCons	=	selectCons	& " , Plantation"
		whereCons	=	whereCons	& " And Plantationcode = '" & v_txtPlantation & "'"
		orderCons	=	orderCons	& " , PLANTATION_API.get_OrderData(Plantation)"
		var_colspan	=	var_colspan	+ 1
	End If
End If


If Len(Trim(v_txtGroupRegion)) > 0 Then
	if v_txtGroupRegion = "--SUMMARY--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
	Elseif v_txtGroupRegion = "--ALL--" then
		selectCons	=	selectCons	& " , GroupRegion"
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " , GROUPREGION_API.get_OrderData(GroupRegion)"
		var_colspan	=	var_colspan	+ 1
	Else
		selectCons	=	selectCons	& " , GroupRegion"
		whereCons	=	whereCons	& " And GroupRegioncode = '" & v_txtGroupRegion & "'"
		orderCons	=	orderCons	& " , GROUPREGION_API.get_OrderData(GroupRegion)"
		var_colspan	=	var_colspan	+ 1
	End If
End If

If Len(Trim(v_txtRegion)) > 0 Then
	if v_txtRegion = "--SUMMARY--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
	Elseif v_txtRegion = "--ALL--" then
		selectCons	=	selectCons	& " , Region"
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " , REGION_API.get_OrderData(Region)"
		var_colspan	=	var_colspan	+ 1
	Else
		selectCons	=	selectCons	& " , Region"
		whereCons	=	whereCons	& " And Regioncode = '" & v_txtRegion & "'"
		orderCons	=	orderCons	& " , REGION_API.get_OrderData(Region)"
		var_colspan	=	var_colspan	+ 1
	End If
End If


If Len(Trim(v_txtCompany)) > 0 Then
	if v_txtCompany = "--SUMMARY--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
	Elseif v_txtCompany = "--ALL--" then
		selectCons	=	selectCons	& " , company"
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " , COMPANY_API.get_OrderData(company)"
		var_colspan	=	var_colspan	+ 1
	Else
		selectCons	=	selectCons	& " , company"
		whereCons	=	whereCons	& " And companycode = '" & v_txtCompany & "'"
		orderCons	=	orderCons	& " , COMPANY_API.get_OrderData(company)"
		var_colspan	=	var_colspan	+ 1
		
		q	=	" SELECT FCCODE" & _
				"        , FCCOMPANYCODE" & _
				"        , FCNAME" & _
				"        , FCCOMPANYNAME" & _
				"        , FCMANAGER" & _
				"        , FCKTU" & _
				"        , FCCROP" & _
				"        , FCTYPE" & _
				" FROM BUSINESSUNIT WHERE" & _
				" FCCOMPANYCODE = '" & v_txtCompany & "'"	
		set rsCat = connStr.Execute(q)		
		if not rsCat.eof then
			CompanyName = rsCat("FCCOMPANYNAME")	
		end if
		q	=	""
		rsCat.Close
		set rsCat = Nothing
	End If
End If

If Len(Trim(v_txtfcba)) > 0 Then
	if v_txtfcba = "--SUMMARY--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
	Elseif v_txtfcba = "--ALL--" then
		selectCons	=	selectCons	& " , nourut, fcba"
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " , nourut"
		var_colspan	=	var_colspan	+ 1
	Else
		selectCons	=	selectCons	& " , nourut, fcba"
		whereCons	=	whereCons	& " And fcba = '" & v_txtfcba & "'"
		orderCons	=	orderCons	& " , nourut"
		var_colspan	=	var_colspan	+ 1
		
		q	=	" SELECT FCCODE" & _
				"        , FCCOMPANYCODE" & _
				"        , FCNAME" & _
				"        , FCCOMPANYNAME" & _
				"        , FCMANAGER" & _
				"        , FCKTU" & _
				"        , FCCROP" & _
				"        , FCTYPE" & _
				" FROM BUSINESSUNIT WHERE" & _
				" FCCODE = '" & v_txtfcba & "'"	
		set rsCat = connStr.Execute(q)		
		if not rsCat.eof then
			fcbaname = rsCat("FCNAME")
			CompanyName = rsCat("FCCOMPANYNAME")
			MgrName = rsCat("FCMANAGER")
			KTUName = rsCat("FCKTU")
			CropName = rsCat("FCCROP")
			FcbaType = rsCat("FCTYPE")
		end if
		q	=	""
		rsCat.Close
		set rsCat = Nothing
	End If
End If

ScriptName = Request.ServerVariables("SCRIPT_NAME")

Dim array_data, array_eof, array_i

array_eof = true

strMain = " "
		  
sqlLHPP2 = " "

sqlLHPP = " SELECT TO_CHAR(WB.POSTINGDATE,'dd.mm.yyyy') POSTINGDATE" & _
                 "        , WB.TIMEOUT" & _
                 "        , WB.CHITNO" & _
                 "        , WB.CUST_CUSTOMERCODE CUSTOMERCODE" & _
                 "        , CU.DESCRIPTION CUSTOMERNAME" & _
				 "        , Businessunit_api.get_SapLocation(WB.FCBA) SapLocation" & _
                 "        , WB.PRODUCT_PRODUCTCODE" & _
                 "        , WB.WEIGHTIN - WB.WEIGHTOUT BRUTO" & _
                 "        , WB.WEIGHTIN" & _
                 "        , WB.WEIGHTOUT" & _
                 "        , (WB.WEIGHTIN - WB.WEIGHTOUT) - NVL(GR.WEIGHT, 0) NETTO" & _
                 "        , WB.VEHICLENO" & _
                 "        , WB.DRIVERNAME" & _
				 "		  , WB.SPBNO" & _
                 "        , NVL(GR.BUNCHES, 0) JJG" & _
				 "        , NVL(GR.Air, 0) AIR" & _
				 "        , NVL(GR.Sampah, 0) Sampah" & _
				 "        , NVL(GR.Pasir, 0) Pasir" & _
				 "        , NVL(GR.dura, 0) Dura" & _
				 "        , nvl(GR.EATENBYRAT,0) + nvl(GR.Brondolan,0) + nvl(GR.Bkecil,0) Lain" & _
                                 "        , wb.vhtype, " & _
       				 " case wb.vhtype  " & _
         			 "	when 'DT' then wb.tarif " & _
         			 "      when 'TR' then wb.tarif *  ( wb.weightin - wb.weightout) " & _
         			 "	else 0 " & _
       				 " end  total_tarif,	GR.Remarks" & _	
                 " FROM WEIGHBRIDGETICKET WB" & _
                 " LEFT JOIN GRADINGRECORD GR ON (WB.FCBA = GR.FCBA AND WB.CHITNO = GR.WBT_CHITNO)" & _
                 " LEFT JOIN CUSTOMER CU ON (WB.FCBA = CU.FCBA AND WB.CUST_CUSTOMERCODE = CU.CUSTOMERCODE)" & _
                 " WHERE WB.FCBA = '"& v_txtFcba & "'" & _
                 "       AND TO_CHAR(WB.POSTINGDATE, 'yyyymm') = '" & v_txtYear & lpad(v_txtMonth, "0", "2") & "'" & _
                 "       AND WB.PRODUCT_PRODUCTCODE = 'FFB'" 
				 '& _
                 '"       AND CU.CUSTFFBTYPE = 'TBSLUAR'"
				 
sqlLHPP = sqlLHPP & " ORDER BY WB.CHITNO"

		  
'response.write sqlLHPP
%>


<html>

	<head>
		<title>ePlaS Reporting - Penerimaan TBS Luar SAP</title>
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
            $( "#datepicker" ).datepicker();
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
				@media print{.printhide{display:none;}}
			</style>
       </head>


	<body>
		<div id="wrapper">
		<div class="printhide">
		<!--#include file="../print_excel.asp"-->
		</div>
			<form name="frmEntry" method="post" action="<%=ScriptName%>">
				<div class="printhide">
				<!--#include file="menutop.asp"-->
				<!--#include virtual="/REPORT/FormSelectFcbaOnlyM.html"-->
				
					<table id="form-table" class="centered-margin">
						<tr>
							<td><font face="verdana" size="1">YEAR</font></td>
							<td><font face="verdana" size="1">MONTH</font></td>
							<td>&nbsp;</td>
						</tr>
						<tr>
							<% 
							Dim mon, Yr, MinYr, MaxYr, i, MM
							mon = Month(Session("v_Costbook"))
							If len(v_txtMonth)=0 then v_txtMonth=mon

							Yr = Year(Session("v_Costbook"))
							MinYr = 2015
							MaxYr = Yr  + 2
							If len(v_txtYear)= 0 then v_txtYear = Yr
							%>
							<td  align="center" width="4">
								<select name="txtYear">
									<%For i = MinYr to MaxYr%>
									<Option Value="<%=i%>" <%If cint(i) = cint(v_txtYear) then%> Selected <%End If%>><%=i%></option>
									<%Next%>
								</select>
							</td>
							<td width="2">
								<select name="txtMonth">
									<%For i = 1 to 12%>
									<Option Value="<%=i%>" <%If cint(i) = cint(v_txtMonth) then%> Selected <%End If%>><%=Bulan(i)%></option>
									<%Next%>
								</select>
							</td>
							<td align="center">
								<input type="submit" value="View" />
							</td>							
						</tr>
					</table>
				</div>
			</form>

			<table border="1" width="100%" id="product-table" class='xthf-site'>
				<thead>
					<tr>
						<th colspan="<%=37%>" style="background:#ffffff; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
							<div id="header-wrapper">
								<div id="head_judul" style="margin-bottom:-30px; position:static; text-align:left;">
									<h4>PT. <%=CompanyName%></h4>
									<h5><%=fcbaname%></h5>
								</div>
								<div id="subheader" class="centered">
									<h1>AMS PLANTATION</h1>
									<div style="margin-botton:10px;"/>
									<h3>LAPORAN PENERIMAAN TBS LUAR SAP</h3>
									<br/>
								</div>
								<div style="text-align:left;">
									<h4>Periode : <%=Bulan(v_txtMonth)%>&nbsp;<%=v_txtYear%></h4>
								</div>
							</div>
						</th>
					</tr>
					<tr>
						<th>Posting Date</th>
						<th>Posting Date</th>
						<th>Posting Date</th>
						<th>Time Out</th>
						<th>Chit No</th>
						<th>Supplier Name</th>
						<th>Supplier Code</th>
						<th>Location Code</th>
						<th>Batch</th>
						<th>Bruto</th>
						<th>Weight In</th>
						<th>Weight Out</th>
						<th>Bruto</th>
						<th>Netto</th>
						<th>Blank</th>
						<th>Vehicle No</th>
						<th>Driver</th>
						<th>SPB No.</th>
						<th>Grade</th>
						<th>R.CPO</th>
						<th>Janjang</th>
						<th>Fr0&00</th>
						<th>FR5</th>
						<th>TK</th>
						<th>TP</th>
						<th>Air</th>
						<th>Sampah</th>
						<th>Pasir</th>
						<th>Lain</th>
						<th>Dura</th>
						<th>Blank</th>
						<th>Type</th>
						<th>Nego</th>
						<th>Sto</th>
						<th>Sto Item</th>
						<th>Tarif</th>
						<th>Remarks</th>
					</tr>
					
					<tr bgcolor="#ddd9c3">
						<th style="background-color:#ddd9c3;color:#ff0000;">Tanggal</th>
						<th style="background-color:#ddd9c3;color:#ff0000;">Tanggal</th>
						<th style="background-color:#ddd9c3;color:#ff0000;">Tanggal</th>
						<th style="background-color:#ddd9c3;color:#ff0000;">Jam Keluar</th>
						<th style="background-color:#ddd9c3;color:#ff0000;">Tiket Timbang</th>
						<th style="background-color:#ddd9c3;color:#ff0000;">Supplier</th>
						<th style="background-color:#ddd9c3;color:#ff0000;">Supplier Code</th>
						<th style="background-color:#ddd9c3;color:#ff0000;">Kode Lokasi</th>
						<th style="background-color:#ddd9c3;color:#ff0000;">Kode Batch</th>
						<th style="background-color:#ddd9c3;color:#ff0000;">Bruto</th>
						<th style="background-color:#ddd9c3;color:#ff0000;">Berat Masuk</th>
						<th style="background-color:#ddd9c3;color:#ff0000;">Berat Keluar</th>
						<th style="background-color:#ddd9c3;color:#ff0000;">Bruto</th>
						<th style="background-color:#ddd9c3;color:#ff0000;">Netto</th>
						<th style="background-color:#ddd9c3;color:#ff0000;">Satuan</th>
						<th style="background-color:#ddd9c3;color:#ff0000;">No Kendaraan</th>
						<th style="background-color:#ddd9c3;color:#ff0000;">Sopir</th>
						<th style="background-color:#ddd9c3;color:#ff0000;">SPB No</th>
						<th style="background-color:#ddd9c3;color:#ff0000;">Grade</th>
						<th style="background-color:#ddd9c3;color:#ff0000;">R.CPO</th>
						<th style="background-color:#ddd9c3;color:#ff0000;">Janjang</th>
						<th style="background-color:#ddd9c3;color:#ff0000;">Fr0&00</th>
						<th style="background-color:#ddd9c3;color:#ff0000;">FR5</th>
						<th style="background-color:#ddd9c3;color:#ff0000;">TK</th>
						<th style="background-color:#ddd9c3;color:#ff0000;">TP</th>
						<th style="background-color:#ddd9c3;color:#ff0000;">Air</th>
						<th style="background-color:#ddd9c3;color:#ff0000;">Sampah</th>
						<th style="background-color:#ddd9c3;color:#ff0000;">Pasir</th>
						<th style="background-color:#ddd9c3;color:#ff0000;">Lain</th>
						<th style="background-color:#ddd9c3;color:#ff0000;">Dura</th>
						<th style="background-color:#ddd9c3;color:#ff0000;">Blank</th>
						<th style="background-color:#ddd9c3;color:#ff0000;">Type</th>
						<th style="background-color:#ddd9c3;color:#ff0000;">Nego</th>
						<th style="background-color:#ddd9c3;color:#ff0000;">Sto</th>
						<th style="background-color:#ddd9c3;color:#ff0000;">Sto Item</th>
						<th style="background-color:#ddd9c3;color:#ff0000;">Tarif</th>
						<th style="background-color:#ddd9c3;color:#ff0000;">Remarks</th>
					</tr>
				</thead>
				
				<tbody>
<%
Set rsLHPP = Server.CreateObject("ADODB.Recordset")
rsLHPP.Open sqlLHPP, connstr, adOpenStatic

if not rsLHPP.EOF then
	array_data = rsLHPP.getrows()
	array_eof = false
end if
rsLHPP.Close
set rsLHPP = Nothing
%>
<%if array_eof = true then%>
					<tr>
						<td colspan="<%=37%>">NO DATA AVAILABLE</td>
					</tr>
<%end if%>
<%if not array_eof then%>
	<%TotalMandays = 0%>
	<%For array_i = LBound(array_data, 2) To UBound(array_data, 2)%>
					<tr>
						<td align="left"><%=array_data(0, array_i)%></td>
						<td align="left"><%=array_data(0, array_i)%></td>
						<td align="left"><%=array_data(0, array_i)%></td>
						<td align="left"><%=array_data(1, array_i)%></td>
						<td align="left"><%=array_data(2, array_i)%></td>
						<td align="left"><%=array_data(4, array_i)%></td>
						<td align="left"><%=array_data(3, array_i)%></td>
						<td align="left"><%=array_data(5, array_i)%></td>
						<td align="left">TBS</td>
						<td align="right"><%=FbpNum2(array_data(7, array_i))%></td>
						<td align="right"><%=FbpNum2(array_data(8, array_i))%></td>
						<td align="right"><%=FbpNum2(array_data(9, array_i))%></td>
						<td align="right"><%=FbpNum2(array_data(7, array_i))%></td>
						<td align="right"><%=FbpNum2(array_data(10, array_i))%></td>
						<td align="left">KG</td>
						<td align="left"><%=array_data(11, array_i)%></td>
						<td align="left"><%=array_data(12, array_i)%></td>
						<td align="left"><%=array_data(13, array_i)%></td>
						<td align="left">&nbsp;</td>
						<td align="left">&nbsp;</td>
						<td align="right"><%=FbpNum2(array_data(14, array_i))%></td>
						<td align="left">&nbsp;</td>
						<td align="left">&nbsp;</td>
						<td align="left">&nbsp;</td>
						<td align="left">&nbsp;</td>
						<td align="right"><%=FbpNum2(array_data(15, array_i))%></td>
						<td align="right"><%=FbpNum2(array_data(16, array_i))%></td>
						<td align="right"><%=FbpNum2(array_data(17, array_i))%></td>
						<td align="right"><%=FbpNum2(array_data(18, array_i))%></td>
						<td align="right"><%=FbpNum2(array_data(19, array_i))%></td>
						<td align="left">&nbsp;</td>
						<td align="left">A01</td>
						<td align="left">&nbsp;</td>
						<td align="left">&nbsp;</td>
						<td align="left">&nbsp;</td>
					
						<td align="right"><%=FbpNum2(array_data(21, array_i))%></td>
						<td align="right"><%=FbpNum2(array_data(22, array_i))%></td>
					</tr>
	<%next%>
<%end if%>
				</tbody>
			</table>

		</div>
	</body>
</html>
