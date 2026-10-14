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

strErptoSap		= 	" select t.fcba,t.post_date,t.spb_date,t.received_date,t.time_received,t.ticket_no,t.customer_code,t.customer_desc,t.location, " & _
					" t.batch,t.weight_delivery,t.first_weight,t.second_weight,t.first_minus_second,t.total_sortasi,t.after_grading,t.uom,t.vehicleno, " & _
					" t.drivername,t.total_tandan,t.sap_type_reference from ERP_REPORT_SAP_SD01 t " & _
					" WHERE t.FCBA = '" & Session("FCBA") & "'  "
										
  if  LEN(TRIM(v_txtYear)) > 0 then
	  strErptoSap	=	strErptoSap & " AND substr(t.post_date,7,4) = " & v_txtYear
  end if
  if  LEN(TRIM(v_txtMonth)) > 0 then
	  strErptoSap	=	strErptoSap & " AND substr(t.post_date,4,2) = " & v_txtMonth
  end if
strErptoSap	=	strErptoSap & " order by ticket_no "
%>
<html>

<head>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>iPla'S Reporting - FFB ESTATE TO OWN POM </title>
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
    <th colspan="24" style="background-color:#FFFFFF; border:none;">  
	<div id="header-wrapper">
		<div id="head_judul">
				<h4>PT. <%=Session("COMPANY")%></h4>
				<h5><%=Session("BA")%></h5>
			</div>
		<div id="subheader" class="centered">
			<h4>INTEGRASI IPLAS - SAP</h4>
            <br/>
           
			<h2>FFB OUTSIDE TO OWN POM (SD / SALES & DISTRIBUTION)</h2>
			<br/>
			 <h4>PERIODE&nbsp;:&nbsp;<%=textMonth(v_txtMonth)%>&nbsp;<%=v_txtYear%></h4>
            <br/>
		</div>
	</div>

  </th>
  </tr>
  
	<tr>
	    <th  style="color:#FFFFFF" >POST DATE</th>
		<th  style="color:#FFFFFF" >TGL SPB</th>
		<th  style="color:#FFFFFF" >TGL TERIMA</th>
		<th  style="color:#FFFFFF" >JAM TERIMA</th>
		<th  style="color:#FFFFFF" >TICKET NUMBER</th>
		<th  style="color:#FFFFFF" >ESTATE</th>
		<th  style="color:#FFFFFF" ></th>
		<th  style="color:#FFFFFF" >LOKASI BONGKAR MUAT</th>
		<th  style="color:#FFFFFF" >BATCH</th>
		<th  style="color:#FFFFFF" >QTY KIRIM</th>
		<th  style="color:#FFFFFF" >QTY TIMBANG PERTAMA</th>
		<th  style="color:#FFFFFF" >QTY TIMBANG KEDUA</th>
		<th  style="color:#FFFFFF" >NETTO</th>
		<th  style="color:#FFFFFF" >QTY SETELAH SORTASI</th>
		<th  style="color:#FFFFFF" >UOM</th>
		<th  style="color:#FFFFFF" >LECENSE PLATE</th>
		<th  style="color:#FFFFFF" >DRIVER NAME</th>
		<th  style="color:#FFFFFF" >SPB NO.</th>
		<th  style="color:#FFFFFF" >BJR</th>
		<th  style="color:#FFFFFF" >R.CPO</th>
		<th  style="color:#FFFFFF" >TOTAL TANDAN</th>
		<th  style="color:#FFFFFF" >FRO%00</th>
		<th  style="color:#FFFFFF" >FR5</th>
		<th  style="color:#FFFFFF" >TANDAN KOSONG</th>
		<th  style="color:#FFFFFF" >TANGKAI PANJANG</th>
        <th  style="color:#FFFFFF" >AIR</th>
        <th  style="color:#FFFFFF" >SAMPAH</th>
		<th  style="color:#FFFFFF" >PASIR</th>
		<th  style="color:#FFFFFF" >LAIN-LAIN</th>
		 <th  style="color:#FFFFFF" >DURA</th>
		  <th  style="color:#FFFFFF" ></th>
		   <th  style="color:#FFFFFF" >TYPE</th>
		  <th  style="color:#FFFFFF" >NEGO NO</th>
		 <th  style="color:#FFFFFF" >STO NO</th>
		 <th  style="color:#FFFFFF" >STO ITEM</th>
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
		
		
%>
	<tr class="<%=cClass%>">
		<td align="center"><%=rsErptoSap("post_date")%></td>
		<td align="center"><%=rsErptoSap("spb_date")%></td>
		<td align="center"><%=rsErptoSap("received_date")%></td>
		<td align="center"><%=rsErptoSap("time_received")%></td>
		<td align="center"><%=rsErptoSap("ticket_no")%></td>
		<td align="center"><%=rsErptoSap("customer_desc")%></td>
		<td align="center"></td>
		<td align="center"><%=rsErptoSap("location")%></td>
		<td align="center"><%=rsErptoSap("batch")%></td>
		<td align="right"><%=rsErptoSap("weight_delivery")%></td>
		<td align="right"><%=rsErptoSap("first_weight")%></td>
		<td align="right"><%=rsErptoSap("second_weight")%></td>
		<td align="right"><%=rsErptoSap("first_minus_second")%></td>
		<td align="right"><%=rsErptoSap("after_grading")%></td>
		<td align="center"><%=rsErptoSap("uom")%></td>
		<td align="center"><%=rsErptoSap("vehicleno")%></td>
		<td align="center"><%=rsErptoSap("drivername")%></td>
		<td align="right"></td>
		<td align="center"></td>
		<td align="center"></td>
		<td align="right"><%=rsErptoSap("total_tandan")%></td>
		<td align="center"></td>
		<td align="right"></td>
		<td align="center"></td>
		<td align="center"></td>
		<td align="center"></td>
		<td align="center"></td>
		<td align="center"></td>
		<td align="center"><%=rsErptoSap("total_sortasi")%></td>
		<td align="center"></td>
		<td align="center"></td>
		<td align="center"><%=rsErptoSap("sap_type_reference")%></td>
		<td align="center"></td>
		<td align="center"></td>
		<td align="center"></td>
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
