<%@Language=VBScript%>
<%Option Explicit%>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
	Dim v_txtBudgetYear, v_txtBudgetVersion, v_txtModuleVersion, v_txtdatefrom, v_txtMonth, v_txtYear, v_txtgrade, v_txtProduct, v_txtemployeetype, v_txtjobcode, varproductname, varproductcode
	Dim no
	Dim rsDefaultBa, strDefaultBa, varDefaultBa, rssearchfirstba, strsearchfirstba
	Dim ScriptName
	Dim strMain, sqlLHPP, sqlLHPP2, rsLHPP
	Dim q, CompanyName, fcbaname, MgrName, KTUName, CropName, FcbaType
	Dim OrderSection : OrderSection = 0
	Dim indx : indx = 0
	Dim rsRowspan, strRowspan, posting_rowspan
	Dim row_todate : row_todate = 0
	Dim bal_todate : bal_todate = 0
	
	Dim rsqtysent, strqtysent, posting_qtysent
	
    v_txtGroupPlantation = Request("txtGroupPlantation")
	v_txtPlantation = Request("txtPlantation")
	v_txtGroupRegion = Request("txtGroupRegion")
	v_txtRegion = Request("txtRegion")
	v_txtcompany = Request("txtCompany")
	v_txtfcba = Request("txtfcba")
	
	v_txtProduct = Request("txtProduct")
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
	
	If Len(Trim(v_txtProduct)) = 0 Then v_txtProduct = "CPO"
	
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

sqlLHPP = " SELECT DN.CONTRACTCODE" & _
                 "        , NVL(DC.TOTALWEIGHT, 0) TOTALWEIGHT" & _
				 "        , DN.DELIVERYNOTECODE"& _
                 "        , WB.PRODUCT_PRODUCTCODE" & _
				 "        , TO_CHAR(DN.POSTINGDATE, 'DD.MM.YYYY') POSTINGDATE" & _
                 "        , WB.VEHICLENO" & _
                 "        , DN.WBT_CHITNO" & _
                 "        , WB.TIMEIN" & _
                 "        , WB.TIMEOUT" & _
                 "        , DN.DRIVERNAME" & _
                 "        , DN.TRANSPORTCO" & _
                 "        , WB.CUST_CUSTOMERCODE" & _
                 "        , CU.DESCRIPTION CUSTOMERNAME" & _
                 "        , WB.WEIGHTOUT BRUTO" & _
                 "        , WB.WEIGHTIN TARRA" & _
                 "        , DN.QUANTITY NETTO" & _
                 " FROM DELIVERYNOTE DN" & _
                 " LEFT JOIN WEIGHBRIDGETICKET WB ON (DN.FCBA = WB.FCBA AND DN.WBT_CHITNO = WB.CHITNO)" & _
                 " LEFT JOIN DELIVERYCONTRACT DC ON (DN.FCBA = DC.FCBA AND DN.CONTRACTCODE = DC.CONTRACTCODE)" & _
                 " LEFT JOIN CUSTOMER CU ON (WB.FCBA = CU.FCBA AND WB.CUST_CUSTOMERCODE = CU.CUSTOMERCODE)" & _
                 " WHERE DN.FCBA = '"& v_txtFcba & "'" & _
                 "       AND TO_CHAR(DN.POSTINGDATE, 'yyyymm') = '" & v_txtYear & lpad(v_txtMonth, "0", "2") & "'"
IF LEN(TRIM(v_txtProduct)) > 0 THEN
	sqlLHPP = sqlLHPP & "       AND WB.PRODUCT_PRODUCTCODE = '" & v_txtProduct & "'"
END IF
	
sqlLHPP = sqlLHPP & "       AND DN.QUANTITY > 0"
                 
				 
sqlLHPP = sqlLHPP & " ORDER BY DN.CONTRACTCODE, DN.WBT_CHITNO"

		  
'response.write sqlLHPP
%>


<html>

	<head>
		<title>ePlaS Reporting - Rekap Pengiriman CPO SAP</title>
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
							<td><font face="verdana" size="1">PRODUCT</font></td>
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
							<td>
								<select name="txtProduct" >
							<%
							Dim rsProductCode, strProductCode
							strProductCode	=	" SELECT productcode fccode, productdescription fcname FROM product WHERE fcba = '" & v_txtFcba & "'"

							Set rsProductCode = Server.CreateObject("ADODB.Recordset")
							rsProductCode.Open strProductCode, connstr, adOpenStatic 

							if not rsProductCode.eof then
								Do While Not rsProductCode.eof
									if rsProductCode("fccode") = v_txtProduct then
										varproductname = rsProductCode("fcname")
										varproductcode = rsProductCode("fccode")
									%>
										<option value="<%=rsProductCode("fccode")%>" Selected><%=rsProductCode("fccode")%> - <%=rsProductCode("fcname")%></option> 
									<%Else%>
										<option value="<%=rsProductCode("fccode")%>"><%=rsProductCode("fccode")%> - <%=rsProductCode("fcname")%></option>  
									<%End if
									rsProductCode.movenext
								loop
							end if
							rsProductCode.Close
							set rsProductCode = Nothing
							%>
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
						<th colspan="<%=23%>" style="background:#ffffff; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
							<div id="header-wrapper">
								<div id="head_judul" style="margin-bottom:-30px; position:static; text-align:left;">
									<h4>PT. <%=CompanyName%></h4>
									<h5><%=fcbaname%></h5>
								</div>
								<div id="subheader" class="centered">
									<h1>AMS PLANTATION</h1>
									<div style="margin-botton:10px;"/>
									<h3>LAPORAN REKAP PENGIRIMAN <%=varproductname%> (<%=varproductcode%>) SAP</h3>
									<br/>
								</div>
								<div style="text-align:left;">
									<h4>Periode : <%=Bulan(v_txtMonth)%>&nbsp;<%=v_txtYear%></h4>
								</div>
							</div>
						</th>
					</tr>
					<!--
					<tr>
						<th>Contract Code</th>
						<th>Qty Contract</th>
						<th>Product Code</th>
						<th>Date Out</th>
						<th>Blank</th>
						<th>Vehicle No</th>
						<th>Chit No</th>
						<th>Time In</th>
						<th>Time Out</th>
						<th>Driver</th>
						<th>Transport Co</th>
						<th>Blank</th>
						<th>Weight In</th>
						<th>Weight Out</th>
						<th>Bruto</th>
						<th>Blank</th>
						<th>Blank</th>
						<th>Blank</th>
					</tr>
					
					<tr bgcolor="#ddd9c3">
						<th style="background-color:#ddd9c3;color:#ff0000;">No Kontrak</th>
						<th style="background-color:#ddd9c3;color:#ff0000;">Qty Kontrak</th>
						<th style="background-color:#ddd9c3;color:#ff0000;">Produk</th>
						<th style="background-color:#ddd9c3;color:#ff0000;">Tanggal Kirim PMKS</th>
						<th style="background-color:#ddd9c3;color:#ff0000;">Tanggal Terima Buyer</th>
						<th style="background-color:#ddd9c3;color:#ff0000;">No Kendaraan</th>
						<th style="background-color:#ddd9c3;color:#ff0000;">Tiket Timbang</th>
						<th style="background-color:#ddd9c3;color:#ff0000;">Jam Masuk</th>
						<th style="background-color:#ddd9c3;color:#ff0000;">Jam Keluar</th>
						<th style="background-color:#ddd9c3;color:#ff0000;">Sopir</th>
						<th style="background-color:#ddd9c3;color:#ff0000;">Transporter</th>
						<th style="background-color:#ddd9c3;color:#ff0000;">Buyer</th>
						<th style="background-color:#ddd9c3;color:#ff0000;">Bruto Kg</th>
						<th style="background-color:#ddd9c3;color:#ff0000;">Tarra Kg</th>
						<th style="background-color:#ddd9c3;color:#ff0000;">Netto Kg</th>
						<th style="background-color:#ddd9c3;color:#ff0000;">Todate Kg</th>
						<th style="background-color:#ddd9c3;color:#ff0000;">Balance Kg</th>
						<th style="background-color:#ddd9c3;color:#ff0000;">Keterangan</th>
					</tr>
					-->
					<tr>
						<th rowspan="2">No Kontrak</th>
						<th rowspan="2">Qty Kontrak</th>
						<th rowspan="2">DO Kontrak</th>
						<th rowspan="2">Produk</th>
						<th rowspan="2">Tanggal Kirim PMKS</th>
						<th rowspan="2">Tanggal Terima Buyer</th>
						<th rowspan="2">No Kendaraan</th>
						<th rowspan="2">Tiket Timbang</th>
						<th rowspan="2">Jam Masuk</th>
						<th rowspan="2">Jam Keluar</th>
						<th rowspan="2">Sopir</th>
						<th rowspan="2">Transporter</th>
						<th rowspan="2">Buyer</th>
						<th colspan="3">Timbangan PMKS</th>
						<th colspan="3">Timbangan Buyer</th>						
						<th rowspan="2">Variance</th>
						<th rowspan="2">Keterangan</th>
					</tr>
					<tr>
						<th>Bruto Kg</th>
						<th>Tarra Kg</th>
						<th>Netto Kg</th>
						<!--
						<th>Todate Kg</th>
						<th>Balance Kg</th>
						-->
						<th>Bruto Kg</th>
						<th>Tarra Kg</th>
						<th>Netto Kg</th>
						<!--
						<th>Todate Kg</th>
						-->
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
						<td colspan="<%=23%>">NO DATA AVAILABLE</td>
					</tr>
<%end if%>
<%if not array_eof then%>
	<%indx = 0%>
	<%For array_i = LBound(array_data, 2) To UBound(array_data, 2)%>
		<%indx = indx + 1%>
		<%IF indx = 1 THEN%>
			<%
			' Cari Total Data Row Per Contract
			strRowspan = " SELECT COUNT (1) totaldata FROM (" & sqlLHPP & ") WHERE CONTRACTCODE = '" & array_data(0, array_i) & "'"
			Set rsRowspan = Server.CreateObject("ADODB.Recordset")
			rsRowspan.Open strRowspan, connstr, adOpenStatic
			if not rsRowspan.EOF then
				posting_rowspan = CDbl(rsRowspan("totaldata"))
			else
				posting_rowspan = 0
			end if
			rsRowspan.Close
			set rsRowspan = Nothing
			
			' Cari Total Data Quantity yang sudah dikirim di bulan sebelumnya
			strqtysent = " SELECT NVL(SUM(QUANTITY), 0) qty" & _
			                   " FROM DELIVERYNOTE" & _
							   " WHERE FCBA = '"& v_txtFcba & "'" & _
							   "       AND TO_NUMBER(TO_CHAR(POSTINGDATE, 'yyyymm')) < '" & v_txtYear & lpad(v_txtMonth, "0", "2") & "'" & _
							   "       AND CONTRACTCODE = '" & array_data(0, array_i) & "'"
			Set rsqtysent = Server.CreateObject("ADODB.Recordset")
			rsqtysent.Open strqtysent, connstr, adOpenStatic
			if not rsqtysent.EOF then
				posting_qtysent = CDbl(rsqtysent("qty"))
			else
				posting_qtysent = 0
			end if
			rsqtysent.Close
			set rsqtysent = Nothing
			
			
			row_todate = 0
			row_todate = posting_qtysent + Cdbl(array_data(15, array_i))
			
			bal_todate = 0
			bal_todate = ( posting_qtysent + Cdbl(array_data(1, array_i)) ) - Cdbl(array_data(15, array_i))
			%>
			<%ELSE%>
				<%row_todate = row_todate + Cdbl( array_data(15, array_i)) %>
				<%bal_todate = bal_todate - Cdbl( array_data(15, array_i)) %>
		<%END IF%>
					<tr>
						<td align="left"><%=array_data(0, array_i)%></td>
						<td align="right"><%=FormatNumber(array_data(1, array_i), 0)%></td>
						<td align="left"><%=array_data(2, array_i)%></td>
						<td align="left"><%=array_data(3, array_i)%></td>
						<td align="left"><%=array_data(4, array_i)%></td>
						<td align="right">&nbsp;</td>
						<td align="left"><%=array_data(5, array_i)%></td>
						<td align="left"><%=array_data(6, array_i)%></td>
						<td align="left"><%=array_data(7, array_i)%></td>
						<td align="left"><%=array_data(8, array_i)%></td>
						<td align="left"><%=array_data(9, array_i)%></td>
						<td align="left"><%=array_data(10, array_i)%></td>
						<td align="left"><%=array_data(12, array_i)%></td>
						<td align="right"><%=FormatNumber(array_data(13, array_i), 0)%></td>
						<td align="right"><%=FormatNumber(array_data(14, array_i), 0)%></td>
						<td align="right"><%=FormatNumber(array_data(15, array_i), 0)%></td>
						<!--
						 <td align="right"><%=FormatNumber(Cdbl(row_todate), 0)%></td>
						 <td align="right"><%=FormatNumber(Cdbl(bal_todate), 0)%></td>
						 -->
						<td align="right">&nbsp;</td>
						<td align="right">&nbsp;</td>
						<td align="right">&nbsp;</td>
						<!--
						<td align="right">&nbsp;</td>
						-->
						<td align="right">&nbsp;</td>
						<td align="right">&nbsp;</td>
					</tr>
				<%IF indx = posting_rowspan THEN%>
					<%indx = 0%>
				<%END IF%>
	<%next%>
<%end if%>
				</tbody>
			</table>

		</div>
	</body>
</html>
