<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
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
<title>ePlaS Reporting - Harvesting</title>
</head>
<%
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"

v_txtCHITNO= trim(request("CHITNO"))
v_txtFIELDCODE= trim(request("FIELDCODE"))
v_txtCROPCODE= trim(request("CROPCODE"))
v_txtRECEPTIONDATE_From= trim(request("RECEPTIONDATE_From"))
v_txtRECEPTIONDATE_To= trim(request("RECEPTIONDATE_To"))
v_txtPRODUCTCODE=trim(request("PRODUCTCODE"))
v_txtsectionname=trim(request("txtsectionname"))
v_txtOWN = request("OWN")
v_txtVehicle = request("Vehicle")
v_txtTRANSPORTTYPE = request("TRANSPORTTYPE")
v_txtCONTRACTNUMBER = request("CONTRACTNUMBER")
businessunit= trim(request("businessunit"))

	q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & businessunit & "' "
		' response.write q
		set rsCat = connStr.Execute(q)
		if not rsCat.eof then
		   fcbaname = rsCat("fcName")
		   CompanyName=rsCat("FCCOMPANYNAME")
		   MgrName = rsCat("FCMANAGER")
		   KTUName = rsCat("FCKTU")
		   CropName=rsCat("FCCROP")
		   FcbaType=rsCat("fcType")
		end if

txtSQL = 	" SELECT spb.fcba," & _
			"        spb.driver," & _
			"		 EMPLOYEE_API.GET_EMPLOYEENAME(spb.fcba, spb.driver) drivername," & _
			"        spb.receptiondate," & _
			"        spb.spbno," & _
			"        spb.harvestdate," & _
			"        spb.fieldcode," & _
			"        spb.cropcode," & _
			"        spb.productcode," & _
			"        spb.vehicle," & _
			"        vehicle_api.get_vehiclecode_name(spb.vehicle, spb.fcba) fcname," & _
			"        spb.own," & _
			"        spb.bunch," & _
			"        spb.bucket," & _
			"        spb.pressemester_abw," & _
			"        spb.bunch_estateweight," & _
			"        spb.brondolan," & _
			"        spb.transporttype," & _
			"        spb.agreementcode" & _
			"   FROM harvestingspb spb" & _ 
			"        LEFT JOIN vehicle vh" & _
			"           ON (vh.fccode = spb.vehicle AND vh.fcba = spb.fcba)" & _
			"        INNER JOIN field ff" & _
			"           ON (spb.fieldcode = ff.fccode AND spb.fcba = ff.fcba)" & _
			"  WHERE spb.fcba = '"& businessunit &"'" 

		 
if  len(trim(v_txtRECEPTIONDATE_From)) > 0 then
    txtSQL = txtSQL & " and RECEPTIONDATE >=to_Date('" & v_txtRECEPTIONDATE_From & "','MM/DD/YYYY')"
	txtSelect = " From " & FormatDateTime(v_txtRECEPTIONDATE_From,1)
end if
if  len(trim(v_txtRECEPTIONDATE_To)) > 0 then
    txtSQL =txtSQL &  " and RECEPTIONDATE <=to_Date('" & v_txtRECEPTIONDATE_To & "','MM/DD/YYYY')"
	txtSelect = txtSelect & "  to " & FormatDateTime(v_txtRECEPTIONDATE_To,1)
end if
if  len(trim(v_txtsectionname)) > 0 then
    txtSQL = txtSQL & "  AND ff.division = '"& v_txtsectionname &"'"
	txtSelect = txtSelect & " Division " & v_txtsectionname
end if
if  len(trim(v_txtFIELDCODE)) > 0 then
    txtSQL =txtSQL &  " and spb.FIELDCODE='" & v_txtFIELDCODE & "'"
	txtSelect = txtSelect & " Field Code " & v_txtFIELDCODE
end if 
if  len(trim(v_txtCROPCODE)) > 0 then
    txtSQL =txtSQL &  " and spb.CROPCODE='" & v_txtCROPCODE & "'"
	txtSelect = txtSelect & " Crop Code " & v_txtCROPCODE
end if 
if  len(trim(v_txtPRODUCTCODE)) > 0 then
    txtSQL =txtSQL &  " and spb.PRODUCTCODE='" & v_txtPRODUCTCODE & "'"
	txtSelect = txtSelect & " Product Code " & v_txtPRODUCTCODE
end if

if  len(trim(v_txtOWN)) > 0 then
    txtSQL =txtSQL &  " and spb.own = '" & v_txtOWN & "'"
	txtSelect = txtSelect & " Vehicle Status " & v_txtOWN
end if

if  len(trim(v_txtVehicle)) > 0 then
    txtSQL = txtSQL &  " and spb.Vehicle = '" & v_txtVehicle & "'"
	txtSelect = txtSelect & " Vehicle " & v_txtVehicle
end if

if  len(trim(v_txtTRANSPORTTYPE)) > 0 then
    txtSQL = txtSQL &  " and spb.Transporttype = '" & v_txtTRANSPORTTYPE & "'"
	txtSelect = txtSelect & " Transport Type " & v_txtTRANSPORTTYPE
end if

if  len(trim(v_txtCONTRACTNUMBER)) > 0 then
    txtSQL = txtSQL &  " and spb.AGREEMENTCODE = '" & v_txtCONTRACTNUMBER & "'"
	txtSelect = txtSelect & " Contract Number " & v_txtCONTRACTNUMBER
end if

' txtsql = txtsql & " group by chitno" & _
'                   "          , fieldcode" & _
'                   "          , cropcode" & _
'                   "          , productcode" & _
'                   "          , receptiondate" & _
'                   "          , vehicle" & _
'                   "          , own" & _
'                   "          , harvestingtranporttypecode" & _
'                   "          , agreementcode" & _
'                   "          , contractorcode" & _
'                   "          , contractor_api.get_contractorname(contractorcode, fcba)" & _
'                   "          , lineno"
					 
txtsql = txtsql & " order by spbno, receptiondate, fieldcode"

if len(trim(txtSelect)) > 0 then
   txtSelect = "Selected  by : " & txtSelect
else
   txtSelect = "Selected by All Data"
end if

' response.Write txtsql
'response.Flush()
%>
<body>
		<div id="wrapper">
        <!--#include file="print_excel.asp"-->
         <table border="0" width="100%" id="product-table" class="xthf-site">
         <thead>
         <tr>
         <th colspan="17" style=" background-color:#FFFFFF; color:#000000; border:none">
			<div id="header-wrapper">
				<div id="head_judul" style="float:left; text-align:left;">

					<h4>PT. <%=CompanyName%></h4>
					<h5><%=fcbaname%></h5>
				</div>
				<div id="subheader" class="centered" style="width:50%; margin:auto;">
					<h2>HARVESTING</h2>
				</div>
			</div>
	</th>
			</tr>
						<!--<div class="printhide">-->
						<tr>
							<th colspan="19" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></th>
				</tr>
				<!--</div>-->
				<%	
				Set rsCat = connStr.Execute(txtSQL)
				if Not rsCat.eof then
					nCountofField=rsCat.fields.count-1
				end if
				NO=1%>
				<tr>
					<!--<th>No</th>-->
					<th>Reception Date</th>
					<th>Employee Code</th>
					<th>Employee Name</th>
					<th>SPB No</th>
					<th>Harvest Date</th>
					<th>Field Code</th>
					<th>Crop Code</th>
					<th>Product Code</th>					
					<th>Vehicle</th>
					<th>Own</th>
					<th>Bunch</th>
					<th>Brondolan</th>
					<th>BJR Census</th>
					<th>Estate Weight</th>
					<th>Transport Type</th>
					<th>Contract Number</th>
				</tr>
                 </thead>
                 <tbody>
				<%

				Dim v_Quantity, nEstateWeight, nMillweight, nMillweightbruto
				v_bunch=0
				v_Bucket=0
				v_bunch_ESTATEWEIGHT=0
				nMillweight=0
				nMillweightbruto=0
				While Not rsCat.EOF%>

				<tr>
					<!--<td align="left" style="<%=cmdstyleLeft%>"><font face="Verdana" size="1"><%=NO%></td>-->
					<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("RECEPTIONDATE")%></td>
					<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("driver")%></td>
					<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("drivername")%></td>
					<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("spbno")%></td>
					<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("harvestdate")%></td>
					<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("FIELDCODE")%></td>
					<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("CROPCODE")%></td>
					<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("PRODUCTCODE")%></td>					
					<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("VEHICLE")%> : <%=rscat("fcname")%></td>
					<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("OWN")%></td>
					<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(rscat("Bunch"))), 0)%></td>
					<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(rscat("brondolan"))), 0)%></td>
					<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(rscat("pressemester_abw"))), 2)%></td>
					<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(rscat("bunch_ESTATEWEIGHT"))), 0)%></td>
					<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("Transporttype")%></td>
					<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("agreementcode")%></td>
					
				</tr>
				<%
				v_bunch=v_bunch + cdbl(to_nol(rsCat("Bunch")))
				v_brondolan=v_brondolan + cdbl(to_nol(rsCat("brondolan")))
				v_bunch_ESTATEWEIGHT=v_bunch_ESTATEWEIGHT + cdbl(to_nol(rsCat("bunch_ESTATEWEIGHT")))
				' nEstateWeight=nEstateWeight+trim(cdbl(rsCat("ESTATE_WEIGHT")))
				' nMillweightbruto=nMillweightbruto+trim(cdbl(rsCat("MILL_WEIGHT_BRUTO")))
				' nMillweight=nMillweight+trim(cdbl(rsCat("MILL_WEIGHT")))
				'response.write cdbl(trim(rsCat("MTHOPENBALANCE")))+1
				rsCat.MoveNext
				'response.Flush()
				no=no+1
				Wend
				rsCat.close()
				rsCat=null%>
				<input  type="hidden" value="<%=txtTable%>" name="txtTable">
				<input  type="hidden" value="<%=txtModule%>" name="txtModule">
				</form>
				<tr bgcolor="#F4F4A0">

					<td align="Center" colspan="10"> <strong><i><font face="Verdana" size="1">T o t a l</i></strong></td>
					<td align="Right"> <strong><i> <font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(v_bunch)), 0)%></i></strong></td>
					<td align="Right"> <strong><i> <font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(v_brondolan)), 0)%></i></strong></td>
					<td align="Right"> <strong><i> <font face="Verdana" size="1"></i></strong></td>
					<td align="Right"> <strong><i> <font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(v_bunch_ESTATEWEIGHT)), 0)%></i></strong></td>
					<td align="Center" colspan="5"> <strong><i><font face="Verdana" size="1">&nbsp;</i></strong></td>
				</tr>
                </tbody>
			</table>
<!--#include file="Bottom_Report.asp"-->
		</div>
	</body>
</html>