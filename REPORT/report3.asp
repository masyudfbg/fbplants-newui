<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
 <%
Dim v_txtYear, v_txtMonth, v_txtGangCode, v_txtGrade
Dim ScriptName
Dim rsErptoSap, strErptoSap, rsMandor, strMandor, NOMOR
Dim NO, cClass, nday

v_txtYear		=	Request("txtYear")
v_txtMonth		=	Request("txtMonth")
ScriptName		=	Request.ServerVariables("SCRIPT_NAME")

strErptoSap		= 	" select fcba, general_upload_type,wbnumber, line, material_code, uom, transporter, to_char(delivery_date,'dd.mm.yyyy') delivery_date, " & _
					" to_char(posting_date,'dd.mm.yyyy') posting_date, gi_plant, " & _
					" gi_store, gi_batch, gi_wb_code, gi_time_in, gi_time_out, mill_gross_qty, mill_tare_qty, gr_bill_of_lading , " & _
					" vehicle_no, gr_plant, gr_store, mill_net_qty, transport_type, transaction_type, gr_discharge " & _
					" from ERP_REPORT_SAP_SD02  " & _
					" WHERE FCBA = '" & Session("FCBA") & "' "
										
   if  LEN(TRIM(v_txtYear)) > 0 then
	   strErptoSap	=	strErptoSap & " AND to_char(delivery_date,'YYYY') = " & v_txtYear
   end if
   if  LEN(TRIM(v_txtMonth)) > 0 then
	   strErptoSap	=	strErptoSap & " AND to_char(delivery_date,'MM') = " & v_txtMonth
   end if
   
   strErptoSap	=	strErptoSap & " order by delivery_date "
%>
<html>

<head>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>iPla'S Reporting - PRODUCT SALES</title>
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
</head>

<body>
	<div id="wrapper">
	<!--#include file="../print_excel.asp"-->
		<div id="header-wrapper">
			<!--#include file="menutop.asp"-->
            <table id="form-table" class="centered-margin">
	<form action="<%=ScriptName%>" method="post" name="frmEntry">
	<tr>
		<td><font face="verdana" size="1">YEAR</font></td>
<% 
	Dim mon, Yr, MinYr, MaxYr, i, MM
    mon = Month(Session("v_Costbook"))
    If len(v_txtMonth)=0 then v_txtMonth = mon

    Yr = Year(Session("v_Costbook"))
    MinYr = 2015
    MaxYr = Yr  + 2
    If len(v_txtYear)= 0 then v_txtYear = Yr
%>
		<td width="4">
			<select name="txtYear">
<%
    For i = MinYr to MaxYr
%>
        <Option Value="<%=i%>" <%If cint(i) = cint(v_txtYear) then%> Selected <%End If%>><%=i%></option>
<%
    Next
%>
			</select>
		</td>
		<td><font face="verdana" size="1">MONTH</font></td>
		<td width="2">
			<select name="txtMonth">
<%
    For i = 1 to 12
%>
        <Option Value="<%=i%>" <%If cint(i) = cint(v_txtMonth) then%> Selected <%End If%>><%=Bulan(i)%></option>
<%
    Next
%>
			</select>
		</td>
		<td><input type="submit" value="View" /></td>
	</tr>
	</form>
</table>
</div>
<table border="0" width="100%" id="product-table" class="xthf-site">
<thead>
	<tr>
    <th colspan="20" style="background-color:#FFFFFF; border:none;">  
	<div id="header-wrapper">
		<div id="head_judul">
				<h4>PT. <%=Session("COMPANY")%></h4>
				<h5><%=Session("BA")%></h5>
			</div>
		<div id="subheader" class="centered" style="width:50%; margin:auto;">
			<h3>INTEGRASI IPLAS - SAP</h3>
			<h4>&nbsp;</h4>
			<h2>PRODUCT SALES (SD / SALES & DISTRIBUTION)</h2>
			<br/>
			<h4>PERIODE&nbsp;:&nbsp;<%=textMonth(v_txtMonth)%>&nbsp;<%=v_txtYear%></h4>
		</div>
	</div>
 </th>
  </tr>
  
	<tr>
	    <th  style="color:#FFFFFF" >GENERAL UPLOAD TYPE</th>
		<th  style="color:#FFFFFF" >WEIGHBRIDGE NO (MILL)</th>
		<th  style="color:#FFFFFF" >STO NO</th>
		<th  style="color:#FFFFFF" >LINE</th>
		<th  style="color:#FFFFFF" >MATERIAL</th>
		<th  style="color:#FFFFFF" >UOM</th>
		<th  style="color:#FFFFFF" >TRANSPORTER</th>
		<th  style="color:#FFFFFF" >DELIVERY DATE</th>
		<th  style="color:#FFFFFF" >DELIVERY POSTING DATE</th>
		<th  style="color:#FFFFFF" >GI PLANT</th>
		<th  style="color:#FFFFFF" >GI STOR</th>
		<th  style="color:#FFFFFF" >GI BATCH</th>
		<th  style="color:#FFFFFF" >GI WB_CODE</th>
		<th  style="color:#FFFFFF" >GI TIME IN</th>
		<th  style="color:#FFFFFF" >GI TIME OUT</th>
		<th  style="color:#FFFFFF" >MILL GROSS QTY</th>
		<th  style="color:#FFFFFF" >MILL TARE QTY</th>
		<th  style="color:#FFFFFF" >WEIGHBRIDGE NO</th>
		<th  style="color:#FFFFFF" >RECEIVE DATE</th>
		<th  style="color:#FFFFFF" >RECEIVE POSTING DATE</th>
		<th  style="color:#FFFFFF" >GR BILL OF LEADING</th>
		<th  style="color:#FFFFFF" >VEHICLE NO</th>
		<th  style="color:#FFFFFF" >GR PLANT</th>
		<th  style="color:#FFFFFF" >GR STORE</th>
		<th  style="color:#FFFFFF" >MILL NETT QTY</th>
		<th  style="color:#FFFFFF" >REFINERY NETT QTY</th>
        <th  style="color:#FFFFFF" >BATCH(DO SAP NO.)</th>
        <th  style="color:#FFFFFF" >GR TIME IN</th>
		<th  style="color:#FFFFFF" >GR TIME OUT</th>
		<th  style="color:#FFFFFF" >REFINERY GROSS QTY</th>
		<th  style="color:#FFFFFF" >REFINERY TARE QTY</th>
		<th  style="color:#FFFFFF" ></th>
		<th  style="color:#FFFFFF" >REFINERY FFA</th>
		<th  style="color:#FFFFFF" >REFINERY M&I</th>
		<th  style="color:#FFFFFF" >REFINERY D&S</th>
		<th  style="color:#FFFFFF" >REFINERY YIELD</th>
		<th  style="color:#FFFFFF" >MILL FFA</th>
		<th  style="color:#FFFFFF" >MILL M&I</th>
		<th  style="color:#FFFFFF" >MILL D&S</th>
		<th  style="color:#FFFFFF" >TRANSPORT TYPE</th>
		<th  style="color:#FFFFFF" >TRANSACTION TYPE</th>
		<th  style="color:#FFFFFF" ></th>
		<th  style="color:#FFFFFF" ></th>
		<th  style="color:#FFFFFF" ></th>
		<th  style="color:#FFFFFF" ></th>
		<th  style="color:#FFFFFF" ></th>
		<th  style="color:#FFFFFF" >GR DISCHARGE</th>
	</tr>
  </thead>
<tbody>
<%
Set rsErptoSap = Server.CreateObject("ADODB.Recordset")
rsErptoSap.Open strErptoSap, connstr, adOpenStatic
'response.write strErptoSap
if rsErptoSap.eof then
%>
	<tr><td colspan="50">NO DATA AVAILABLE</td></tr>
<%
else
	NO = 1

	while not rsErptoSap.eof
		if no  mod 2 = 1 then
			cClass = "odd"
		else
			cClass ="default"
		end if

' " select fcba, general_upload_type, line, material_code, uom, transporter, to_char(delivery_date,'dd.mm.yyyy') delivery_date, " & _
					' " to_char(posting_date,'dd.mm.yyyy') posting_date, gi_plant, " & _
					' " gi_store, gi_batch, gi_wb_code, gi_time_in, gi_time_out, mill_gross_qty, mill_tare_qty, gr_bill_of_lading, " & _
					' " vehicle_no, gr_plant, gr_store, mill_net_qty, transport_type, transaction_type, gr_discharge " & _
		
%>
	<tr class="<%=cClass%>">
		<td align="center"><%=rsErptoSap("general_upload_type")%></td>
		<td align="right"><%=FbpNum2(rsErptoSap("wbnumber"))%></td>
		<td align="center"></td>
		<td align="center"><%=rsErptoSap("line")%></td>
		<td align="center"><%=rsErptoSap("material_code")%></td>
		<td align="center"><%=rsErptoSap("uom")%></td>
		<td align="center"><%=rsErptoSap("transporter")%></td>
		<td align="center"><%=rsErptoSap("delivery_date")%></td>
		<td align="center"><%=rsErptoSap("posting_date")%></td>
		
		<td align="center"><%=rsErptoSap("gi_plant")%></td>
		<td align="center"><%=rsErptoSap("gi_store")%></td>
		
		<td align="center"><%=rsErptoSap("gi_batch")%></td>
		<td align="center"><%=rsErptoSap("gi_wb_code")%></td>
		<td align="center"><%=rsErptoSap("gi_time_in")%></td>
		<td align="center"><%=rsErptoSap("gi_time_out")%></td>		
		<td align="right"><%=FbpNum2(rsErptoSap("mill_gross_qty"))%></td>
		<td align="right"><%=FbpNum2(rsErptoSap("mill_tare_qty"))%></td>
        <td align="right"></td>
		<td align="center"></td>
		<td align="center"></td>
		<td align="center"><%=rsErptoSap("gr_bill_of_lading")%></td>
		<td align="center"><%=rsErptoSap("vehicle_no")%></td>
		<td align="center"><%=rsErptoSap("gr_plant")%></td>
		<td align="center"><%=rsErptoSap("gr_store")%></td>
		<td align="right"><%=FbpNum2(rsErptoSap("mill_net_qty"))%></td>
		<td align="right"></td>
		<td align="center"></td>
		<td align="center"></td>
		<td align="right"></td>
		<td align="center"></td>
		<td align="center"></td>
		<td align="center"></td>
		<td align="right"></td>
		<td align="center"></td>
		<td align="center"></td>
		<td align="center"></td>
		<td align="center"></td>
		<td align="center"></td>
		<td align="center"></td>

		<td align="center"><%=rsErptoSap("transport_type")%></td>
		<td align="center"><%=rsErptoSap("transaction_type")%></td>
		<td align="center"></td>
		<td align="center"></td>
		<td align="center"></td>
		<td align="center"></td>
		<td align="center"></td>
		<td align="center"><%=rsErptoSap("gr_discharge")%></td>
	</tr>	
<%
	NO = NO + 1
	rsErptoSap.MoveNext
Wend
end if
rsErptoSap.Close
set rsErptoSap = Nothing
%>
	
</tbody>
</table>
  
<table width="90%" class="data">
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
		<td width="50%" align="center"><font face="Verdana" size="2"><strong><u><%=Session("FCMANAGER")%></u></strong></td>
		<td width="50%" align="center"><font face="Verdana" size="2"><strong><u><%=Session("KTU")%></u></strong></td>
	</tr>
	<tr> 
		<td width="50%" align="center"><font face="Verdana" size="1"><strong><i>MANAGER</i></strong></td>
		<td width="50%" align="center"><font face="Verdana" size="1"><strong><i>KTU</i></strong></td>
	</tr>
</table>
    </body>
    </html>
