<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->

<%
Dim v_txtDateFrom, v_txtDateFrom2, v_txtDateTo2, v_txtDateTo,v_txtSupplier
Dim ScriptName, NO, cClass
Dim nBRUTO, nSORTASI, nNETTO, selectifkebun, selectifpabrik, v_txtfcba2

 if Trim(Request("txtDateFrom")) <> "" then
    v_txtDateFrom = Trim(Request("txtDateFrom"))
 else
   v_txtDateFrom = textDate(date)
 end if
if Trim(Request("txtDateTo")) <> "" then
    v_txtDateTo = Trim(Request("txtDateTo"))
 else
    v_txtDateTo = textDate(date)
 end if
 
 'if Trim(Request("txtDateFrom2")) <> "" then
 '   v_txtDateFrom2 = Trim(Request("txtDateFrom2"))
 'else
 '  v_txtDateFrom2 = textDate(date)
 'end if
if Trim(Request("txtDateTo2")) <> "" then
    v_txtDateTo2 = Trim(Request("txtDateTo2"))
 else
    v_txtDateTo2 = textDate(date)
 end if
 

v_txtfcba 			  = Request("txtfcba")
v_txtfcba2 			  = Request("txtfcba2")

if v_txtfcba = "" then
selectifkebun = ""
else
selectifkebun = " and field_estate_estatecode = '"&v_txtfcba&"'" 
end if

if v_txtfcba2 = "" then

selectifpabrik = ""
else
selectifpabrik = " and FCBA = '"&v_txtfcba2&"'" 
end if

if v_txtfcba <> "" and v_txtfcba2 <> "" then
selectifkebun  = ""
selectifpabrik = ""
end if
 
ScriptName	=	Request.ServerVariables("SCRIPT_NAME")
Dim rsTerimaTBS, strTerimaTBS
strTerimaTBS ="SELECT FCBA PMKS, CHITNO, VEHICLENO, DATEIN, TIMEIN, WEIGHTIN, DATEOUT, TIMEOUT, POSTINGDATE,          field_estate_estatecode EstateCode, spbno, drivername, division afdeling, field_fieldcode fieldcode, weighbridgeticket_api.getListTTWB(field_fieldcode,field_estate_estatecode,datein) Plantingyear, gradingrecord_api.GET_BUNCHES(CHITNO,fcba) JJG,weightin, weightout, weightin - weightout bruto, gradingrecord_api.GET_SORTASE(CHITNO,fcba) Sortase FROM weighbridgeticket a WHERE PRODUCT_PRODUCTCODE = 'FFB'" &_ 
 selectifkebun & selectifpabrik
					
if  len(trim(v_txtDateFrom2)) > 0 then
	 strTerimaTBS = strTerimaTBS & " and POSTINGDATE = to_Date('" & v_txtDateFrom2 & "','MM/DD/YYYY')"
else
     strTerimaTBS = strTerimaTBS & " and datein between to_Date('" & v_txtDateFrom & "','MM/DD/YYYY') and to_Date('" & v_txtDateFrom & "','MM/DD/YYYY')"
end if

'response.write strTerimaTBS
Set rsTerimaTBS = Server.CreateObject("ADODB.Recordset")
rsTerimaTBS.Open strTerimaTBS, connstr, adOpenStatic


%>
<html>

<head>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>ePlaS Reporting - Penerimaan TBS PMKS</title>
<link rel="stylesheet" href="/css/jquery-ui.css" />
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
    <script src="/js/jquery-1.9.1.js"></script>
	<script src="/js/jquery-ui.js"></script>
     <script>
	$(function() {
	$( ".datepicker" ).datepicker();
	});
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
	<!--#include file="menutop.asp"-->
	<!--#include file="../print_excel.asp"-->
	<form action="<%=ScriptName%>" method="post" name="frmEntry">
	<div id="header-wrapper" style=" ">
	<!--#include virtual="/REPORT/FormSelectFcbaComparasi.html"-->
    <table id="form-table" class="centered-margin">

	<tr>		
		<td><font face="verdana" size="1">From Date in</font></td>
		<td align="center">
		<% if v_txtDateFrom <> "" and v_txtDateFrom = v_txtDateFrom then %>
			<input class="datepicker" type="text" maxlength="10" name="txtDateFrom" size="10" value="<%=v_txtDateFrom%>" title="Type Date mm/dd/yyyy" >
		<% else %>
			<input class="datepicker" type="text" maxlength="10" name="txtDateFrom" size="10" value="<%=textDate(Date)%>" title="Type Date mm/dd/yyyy" >
		<% end if%>
	    </td>
		<td><font face="verdana" size="1">To Date in</font></td>
		<td align="center">
		<% if v_txtDateTo <> "" and v_txtDateTo = v_txtDateTo then %>
			<input class="datepicker" type="text" maxlength="10" name="txtDateTo" size="10" value="<%=v_txtDateTo%>" title="Type Date mm/dd/yyyy" >
		<% else %>
			<input class="datepicker" type="text" maxlength="10" name="txtDateTo" size="10" value="<%=textDate(Date)%>" title="Type Date mm/dd/yyyy" >
		<% end if%>
			
		</td>
				<td><font face="verdana" size="1">Posting Date</font></td>
		<td align="center">
		<% if v_txtDateFrom2 <> "" and v_txtDateFrom2 = v_txtDateFrom2 then %>
			<input class="datepicker" type="text" maxlength="10" name="txtDateFrom2" size="10" value="<%=v_txtDateFrom2%>" title="Type Date mm/dd/yyyy" >
		<% else %>
			<input class="datepicker" type="text" maxlength="10" name="txtDateFrom2" size="10" value="" title="Type Date mm/dd/yyyy" >
		<% end if%>
		<td><input type="submit" value="View" /></td>
		</tr>
		
	</tr>
	</table>
	</div>
	</form>

<table border="0" width="100%" id="product-table" class="xthf-site">
<thead>
	<tr>
    <th colspan="50" style="background-color:#FFFFFF; border:none;">  
	<div id="header-wrapper">
		<div id="head_judul" style="float:left; text-align:left;">
		<h4>PT. <%=Session("COMPANY")%></h4>
		<h5><%=Session("BA")%></h5>
	    </div>
		<div id="subheader" class="centered"  style="width:50%; margin:auto;">
		<h2>LAPORAN PENERIMAAN TBS PMKS</h3>
		<br/>
		<% if (v_txtdatefrom <> "" or v_txtdateto <> "") and v_txtdatefrom = v_txtdateto then%>
			<h5>TANGGAL: <%=v_txtdateto%></h5>
		<% else %>
			<h5>TANGGAL: <%=v_txtdatefrom%> - <%=v_txtdateto%></h5>
		<% end if %>
		</div>
	</div>
 	</th>
   </tr>
	<tr>
	    <th style="color:#FFFFFF">NO</th>
		<th style="color:#FFFFFF">PMKS</th>
		<th style="color:#FFFFFF">CHIT NO</th>
		<th style="color:#FFFFFF">VEHICLE NO</th>
		<th style="color:#FFFFFF">DATE IN</th>
		<th style="color:#FFFFFF">TIME IN</th>
		<th style="color:#FFFFFF">DATE OUT</th>
		<th style="color:#FFFFFF">TIME OUT</th>
		<th style="color:#FFFFFF">POSTING DATE</th>
		<th style="color:#FFFFFF">ESTATE CODE</th>
		<th style="color:#FFFFFF">SPB NO</th>
		<th style="color:#FFFFFF">DRIVER NAME</th>
		<th style="color:#FFFFFF">AFDELING</th>
		<th style="color:#FFFFFF">FIELD CODE</th>
		<th style="color:#FFFFFF">PLANTING YEAR</th>
		<th style="color:#FFFFFF">JUMLAH JANJANG</th>
		<th style="color:#FFFFFF">WEIGHT IN</th>
		<th style="color:#FFFFFF">WEIGHT OUT</th>
		<th style="color:#FFFFFF">BRUTO</th>
		<th style="color:#FFFFFF">SORTASE</th>
		<th style="color:#FFFFFF">NETTO</th>
	</tr>
    </thead>
	<tbody>
<%
dim GrandTOTAL(3)
if rsTerimaTBS.bof and rsTerimaTBS.eof then
%>
	<tr><td colspan="22">NO DATA AVAILABLE</td></tr>
<%else
		NO = 1
		Do while not rsTerimaTBS.eof
			if NO  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
%>
	<tr class="<%=cClass%>">
		<td align="center"><%=NO%></td>
		<td align="left"><%=rsTerimaTBS("PMKS")%></td>
		<td align="left"><%=rsTerimaTBS("chitno")%></td>
		<td align="left"><%=rsTerimaTBS("vehicleno")%></td>
		<td align="left"><%=rsTerimaTBS("datein")%></td>
		<td align="left"><%=rsTerimaTBS("timein")%></td>
		<td align="left"><%=rsTerimaTBS("dateout")%></td>
		<td align="left"><%=rsTerimaTBS("timeout")%></td>
		<td align="left"><%=rsTerimaTBS("postingdate")%></td>
		<td align="left"><%=rsTerimaTBS("EstateCode")%></td>
		<td align="left"><%=rsTerimaTBS("spbno")%></td>
		<td align="left"><%=rsTerimaTBS("drivername")%></td>
		<td align="left"><%=rsTerimaTBS("afdeling")%></td>
		<td align="left"><%=rsTerimaTBS("fieldcode")%></td>
		<td align="left"><%=rsTerimaTBS("Plantingyear")%></td>
		<td align="Right"><%=FormatNumber(rsTerimaTBS("JJG"),0)%></td>
		<td align="Right"><%=FormatNumber(rsTerimaTBS("weightin"),0)%></td>
		<td align="Right"><%=FormatNumber(rsTerimaTBS("weightout"),0)%></td>
		<td align="Right"><%=FormatNumber(rsTerimaTBS("bruto"),0)%></td>
		<td align="Right"><%=FormatNumber(rsTerimaTBS("Sortase"),0)%></td>
		<td align="Right"><%= FormatNumber(cdbl(rsTerimaTBS("bruto")) - cdbl (rsTerimaTBS("sortase")),0)%></td>
	</tr>
<%
            NO = NO + 1
			GrandTOTAL(1) =  GrandTOTAL(1) + cdbl(rsTerimaTBS("bruto"))
			GrandTOTAL(2) =  GrandTOTAL(2) + cdbl(rsTerimaTBS("Sortase"))
			GrandTOTAL(3) =  GrandTOTAL(3) + cdbl(rsTerimaTBS("bruto")) - cdbl(rsTerimaTBS("sortase"))
			
			rsTerimaTBS.MoveNext
			

            		
		loop
end if

rsTerimaTBS.Close
set rsTerimaTBS = Nothing
%>
         <tr bgcolor='#fff600'>
		              <td align="Right" colspan="18"><strong> Total </strong></td>
		              <td align="Right"><strong><%=FormatNumber(GrandTOTAL(1),0)%></strong></td>
                      <td align="Right"><strong><%=FormatNumber(GrandTOTAL(2),0)%></strong></td>
					  <td align="Right"><strong><%=FormatNumber(GrandTOTAL(3),0)%></strong></td>
		 </tr>	
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
