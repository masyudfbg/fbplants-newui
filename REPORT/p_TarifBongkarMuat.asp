<%@Language=VBScript%>
<% ValidateSession %>
<!--#include virtual="/conndB1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
Dim v_txtFcba, v_txtYear, v_txtMonth,v_txtDay, v_txtAfdeling, v_txtBlok, v_txtColor
Dim v_txtDateFrom, v_txtDateTo, v_txtSupplier
Dim ScriptName
Dim rsPusingan, strPusingan, strsql
Dim NO, cClass, nday,ColorID
Dim rsDefaultBa, strDefaultBa, varDefaultBa
Dim rsBaInfo, strBaInfo, varFcbaName,varCompanyCode, varCompanyName
Dim oRS, oSQL

cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"

v_txtDateFrom = Request("txtDateFrom")
v_txtDateTo = Request("txtDateTo")

If "" & v_txtDateFrom = "" Then v_txtDateFrom = TextDate(date)
If "" & v_txtDateTo = "" Then v_txtDateTo = TextDate(date)


v_txtSupplier				=	Request("txtSupplier")	'' Supplier
v_txtFcba					=	Request("txtFcba") 		'' Business Unit

''Set Default FCBA ( Business Unit Kebun dan Pabrik )
strDefaultBa = " SELECT" & _
				"     FCCODE as FCBA" & _
				" FROM BUSINESSUNIT BU" & _
				" INNER JOIN USER_FCBA_TAB UFT ON (BU.FCCODE = UFT.FCBA) " & _
				" WHERE" & _
				"     ISDEFAULT = 'TRUE'" & _
	    		"     AND FCUSERCODE = '" & Session("CodeName") & "'"

''response.write strDefaultBa
Set rsDefaultBa = Server.CreateObject("ADODB.Recordset")
rsDefaultBa.Open strDefaultBa, connstr, adOpenStatic

If Not rsDefaultBa.EOF Then
	varDefaultBa = rsDefaultBa("FCBA")
Else
	varDefaultBa = ""
End If
rsDefaultBa.Close
set rsDefaultBa = Nothing

If (v_txtFcba = "" Or Len(Trim(v_txtFcba)) = 0) Then
	v_txtFcba = varDefaultBa
End If



ScriptName	=	Request.ServerVariables("SCRIPT_NAME")
<!--Additional -->

strPusingan	=   " SELECT " & _
				" w.postingdate as Tanggal,  w.chitno as No_Tiket,  w.timeout as Time_Out,  w.cust_customercode, " & _
				" customer_api.get_description(w.cust_customercode,w.fcba) as customername," & _
				" '' as UOM, '' as Tarif, w.vhtype as Jenis_Vehicle, " & _
				" w.vehicleno as No_Kendaraan,  w.spbno, w.drivername as Driver, " & _
				" w.weightin,  w.weightout, " & _
				" w.weightin - w.weightout as Bruto, " & _
				" (w.weightin - w.weightout) - (w.weightin - w.weightout - gradingrecord_api.GET_SORTASE(w.chitno, w.fcba)) as Tarra, " & _
				" w.weightin - w.weightout - gradingrecord_api.GET_SORTASE(w.chitno, w.fcba) as netto, " & _
				" case w.vhtype  " & _
				"   when 'DT' then w.tarif " & _
				"   when 'TR' then w.tarif *  ( w.weightin - w.weightout) " & _
				"   else 0 " & _
				" end  total_tarif " & _
				" FROM weighbridgeticket w " & _
				" WHERE 1=1" & _
				" AND PRODUCT_PRODUCTCODE ='FFB' " & _
				" AND FCBA = '" & varDefaultBa & "'" & _
				" AND postingdate BETWEEN TO_DATE('" & v_txtDateFrom & "', 'mm/dd/yyyy') AND TO_DATE('" & v_txtDateTo & "', 'mm/dd/yyyy')"

			''Filter Division / Afdeling
			If v_txtSupplier <> "ALL" And Len(Trim(v_txtSupplier)) > 0 Then
				strPusingan	= strPusingan	& "  AND cust_customercode = '" & v_txtSupplier & "'"
			Else
				strPusingan	= strPusingan	& ""
			End If
	
			strPusingan	= strPusingan	& " ORDER BY cust_customercode, postingdate "	 
		
''response.write strpusingan
									
strBaInfo = " SELECT" & _
			"     FCNAME as FCBANAME" & _
			"    , fccompanycode  , FCCOMPANYNAME as COMPANYNAME" & _
			" FROM BUSINESSUNIT BU" & _
			" WHERE" & _
			"     FCCODE = '" & v_txtFcba & "'"

Set rsBaInfo = Server.CreateObject("ADODB.Recordset")
rsBaInfo.Open strBaInfo, connstr, adOpenStatic

If Not rsBaInfo.EOF Then
	varFcbaName = rsBaInfo("FCBANAME")
	varCompanyName = rsBaInfo("COMPANYNAME")
	varCompanyCode = rsBaInfo("fccompanycode")
Else
	varFcbaName = ""
	varCompanyName = ""
	varCompanyCode = ""
End If
rsBaInfo.Close
set rsBaInfo = Nothing

''response.write strPusingan
%>

<%
			Dim strTitle, rsTitle
			Dim titleGroupPlantation, titlePlantation, titleGroupRegion, titleRegion, titleCompany, titleSite
			
	        strTitle = " select PARAMETERSYSDETAIL_API.GETPARAM1 ('REPORTGROUP','GROUPPLANTATION') titleGP, " & _
							    " PARAMETERSYSDETAIL_API.GETPARAM1 ('REPORTGROUP','PLANTATION') titleP , " & _
								" PARAMETERSYSDETAIL_API.GETPARAM1 ('REPORTGROUP','GROUPREGION') titleGR, " & _
								" PARAMETERSYSDETAIL_API.GETPARAM1 ('REPORTGROUP','REGION') titleR, " & _
								" PARAMETERSYSDETAIL_API.GETPARAM1 ('REPORTGROUP','COMPANY') titleC, " & _
								" PARAMETERSYSDETAIL_API.GETPARAM1 ('REPORTGROUP','SITE') titleS from dual"
								
			Set rsTitle = Server.CreateObject("ADODB.Recordset")
			'response.write rsTitle
			rsTitle.Open strTitle, connstr, 3,3	

			if not rsTitle.EOF then
				no = 0
				titleGroupPlantation = rsTitle ("titleGP")			
				titlePlantation = rsTitle ("titleP")
				titleGroupRegion = rsTitle ("titleGR")
				titleRegion = rsTitle ("titleR")
				titleCompany = rsTitle ("titleC")
				titleSite = rsTitle ("titleS")
			else
				titleGroupPlantation = "Group Plantation"			
				titlePlantation = "Plantation"
				titleGroupRegion = "Group Region"
				titleRegion = "Region"
				titleCompany = "Company"
				titleSite = "Site"
			end if
			rsTitle.Close
			set rsTitle = Nothing
			
%>
<html>

<head>
<title>ePlaS Reporting - Tarif Bongkar Muat</title>
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
	$( ".datepicker" ).datepicker({
		changeMonth: true,
		changeYear: true
	});
});
</script>
<!--
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
-->

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
						<!--<td><font face="verdana" size="1">TAHUN</font></td>
						<% 
							Dim mon, Yr, MinYr, MaxYr, i, MM
							mon = Month(Session("v_Costbook"))
							If len(v_txtMonth)=0 then v_txtMonth=mon

							Yr = Year(Session("v_Costbook"))
							MinYr = Yr - 2
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
							<td><font face="verdana" size="1">BULAN</font></td>
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
							</td>-->
							
							<!-- From Date -->
							<td>
								<font face="verdana" size="1">FROM DATE</font>
							</td>
							<td width="2">
								<input class="datepicker" type="text" name="txtDateFrom" value="<%=v_txtDateFrom%>" />
							</td>
							
							<!-- To Date -->
							<td>
								<font face="verdana" size="1">TO DATE</font>
							</td>
							<td width="2">
								<input class="datepicker" type="text" name="txtDateTo" value="<%=v_txtDateTo%>" />
							</td>
							
							<!-- Supplier -->
							<td>
								<font face="verdana" size="1">SUPPLIER</font>
							</td>
							<td>
								<select name="txtSupplier" >
									<option value="ALL"<%If v_txtSupplier = "ALL" then%> Selected <%End If%>>ALL</option>
									<%
									Dim rsSupplier, strSupplier
									strSupplier	=	" SELECT DISTINCT CUSTOMERCODE, DESCRIPTION FROM CUSTOMER" & _
													" WHERE FCBA = '" & v_txtfcba & "'" & _
													" ORDER BY CUSTOMERCODE"
									
									Set rsSupplier = Server.CreateObject("ADODB.Recordset")
									rsSupplier.Open strSupplier, connstr, adOpenStatic 
									
									if not rsSupplier.eof then
									Do While Not rsSupplier.eof
									if rsSupplier("CUSTOMERCODE") = v_txtSupplier then%>
									<option value="<%=rsSupplier("CUSTOMERCODE")%>" Selected><%=rsSupplier("DESCRIPTION")%></option> 
									<%Else%>
									<option value="<%=rsSupplier("CUSTOMERCODE")%>"><%=rsSupplier("DESCRIPTION")%></option>  
									<%End if
										rsSupplier.movenext
									loop
									end if
									rsSupplier.Close
									set rsSupplier = Nothing
									%>
								</select>
							</td>
						
							<!-- View -->
							<td>
								<input type="submit" value="View" />
							</td>
						</tr>				
					</table>
				</div>
			</form>
		</div>
		
				
<table border="0" width="100%" id="product-table" class="xthf-site">
	<thead style="border-left:1px #000000 solid;">
	<tr>
         <th colspan="<%=12%>" style="background:#FFFFFF; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
		 <div id="header-wrapper">
			<div id="head_judul">
				<h4><%If Len(Trim(varCompanyName)) > 0 Then%>PT. <%End If%><%=varCompanyName%></h4>
				<h5><%=varFcbaName%></h5>
			</div>
			<div id="subheader" class="centered">
				<h2>TARIF BONGKAR MUAT</h3>
				<br/>

				<!-- 
				<h4>PERIODE&nbsp;:&nbsp;<%=Bulan(v_txtMonth)%>&nbsp;<%=v_txtYear%></h4>
				<h4>Dari Tanggal&nbsp;:&nbsp;<%=v_txtDay%>&nbsp;<%=Bulan(v_txtMonth)%>&nbsp;&nbsp;<%=v_txtYear%></h4>
				-->
				<h4>Dari Tanggal&nbsp;:&nbsp;<%=v_txtDateFrom%>&nbsp; s/d Tanggal&nbsp;:&nbsp;<%=v_txtDateTo%></h4>
				<br/>
			</div>
          </div>  
		 </th>
	</tr>
<% 
' dim selec_day
	' if len(trim(DAY(v_txtDate))) = 0 then
		' selec_day = GetLastDay(v_txtMonth, v_txtYear)
	' else
		' selec_day = DAY(v_txtDate)
	' end if
%>
	<tr>
		<th style="background-color:#00A300;color:#fff;">Tanggal</th>
		<th style="background-color:#00A300;color:#fff;">No.Tiket</th>
		<th style="background-color:#00A300;color:#fff;">Time Out</th>
		<th style="background-color:#00A300;color:#fff;">Supplier</th>
		<!--
		<th style="background-color:#00A300;color:#fff;">UOM</th>
		<th style="background-color:#00A300;color:#fff;">Tarif</th>
		-->
		<th style="background-color:#00A300;color:#fff;">Jenis Vehicle</th>
		<th style="background-color:#00A300;color:#fff;">No.Kendaraan</th>
		<th style="background-color:#00A300;color:#fff;">No.SPBS</th>
		<th style="background-color:#00A300;color:#fff;">Driver</th>
		<th style="background-color:#00A300;color:#fff;">Bruto</th>
		<th style="background-color:#00A300;color:#fff;">Tarra</th>
		<th style="background-color:#00A300;color:#fff;">Netto</th>
		<th style="background-color:#00A300;color:#fff;">Total Tarif</th>
		
    </thead>
	<tbody>
<%
''response.write strPusingan

Set rsPusingan = Server.CreateObject("ADODB.Recordset")
rsPusingan.Open strPusingan, connstr, adOpenStatic

Dim xNum

Dim Total_Ha : Total_Ha = 0
Dim Total_LastInterval : Total_LastInterval = 0
Dim DataIsPersen : DataIsPersen = 0
Dim CountData : CountData = 0


if rsPusingan.bof and rsPusingan.eof then

%>
	<tr>
		<td colspan="14">NO DATA AVAILABLE</td>
	</tr>
<%end if


		Dim n_Bruto, n_Tarra, n_Netto, n_TotalTariff
			n_Bruto=0
			n_Tarra=0
			n_Netto=0
			n_TotalTariff=0
				
		while not rsPusingan.eof
			CountData = CountData + 1
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
%>

			<tr class="<%=cClass%>">
				<td align="center"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsPusingan("Tanggal")%></td>
				<td align="center"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsPusingan("No_Tiket")%></td>
				<td align="center"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsPusingan("Time_Out")%></td>
				<td align="center"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsPusingan("customername")%></td>
				<!--
				<td align="center"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsPusingan("UOM")%></td>
				<td align="center"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsPusingan("Tarif")%></td>
				-->
				<td align="center"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsPusingan("Jenis_Vehicle")%></td>
				<td align="center"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsPusingan("No_Kendaraan")%></td>
				<td align="center"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsPusingan("spbno")%></td>
				<td align="center"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsPusingan("Driver")%></td>
				<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsPusingan("Bruto"),0)%></td>
				<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsPusingan("Tarra"),0)%></td>
				<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsPusingan("netto"),0)%></td>
				<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsPusingan("total_tarif"),0)%></td>
						
			</tr>

<%
			n_Bruto=n_Bruto+trim(cdbl(rsPusingan("Bruto")))
			n_Tarra=n_Tarra+trim(cdbl(rsPusingan("Tarra")))
			n_Netto=n_Netto+trim(cdbl(rsPusingan("netto")))
			n_TotalTariff=n_TotalTariff+trim(cdbl(rsPusingan("total_tarif")))
				
			rsPusingan.MoveNext
		Wend
rsPusingan.Close
set rsPusingan = Nothing
'response.write strPusingan
'response.write  Total_LastInterval
'response.write  Total_Ha 
%>

		<tr bgcolor="#F4F4A0">
			<td align="Right" colspan="8"> <strong><i><font face="Verdana" size="1">Grand Total</i></strong></td>
			<td align="Right"> <strong><i> <font face="Verdana" size="1"><%=FormatNumber(n_Bruto, 0)%></i></strong></td>
			<td align="Right"> <strong><i> <font face="Verdana" size="1"><%=FormatNumber(n_Tarra, 0)%></i></strong></td>
			<td align="Right"> <strong><i> <font face="Verdana" size="1"><%=FormatNumber(n_Netto, 0)%></i></strong></td>
			<td align="Right"> <strong><i> <font face="Verdana" size="1"><%=FormatNumber(n_TotalTariff, 0)%></i></strong></td>
		</tr>

</tbody>
</table>

   <table width="90%" class="data">
		<br/>
	   <tr>
 		<td width="30%" align="center"><font face="Verdana" size="1">Prepared by,</font></td>
		<td width="30%" align="center"><font face="Verdana" size="1">Checked by,</font></td>
		<td width="30%" align="center"><font face="Verdana" size="1">AcKnowledge by,</font></td>	
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
	   <%
		DIM rsttd,ttd
		
		ttd = "SELECT FCMANAGER,FCKTU FROM BusinessUnit WHERE FCCODE = '" & v_txtFcba & "'"
		
	    Set rsttd = Server.CreateObject("ADODB.Recordset")
		rsttd.Open ttd, connstr, adOpenStatic

		If Not rsttd.EOF Then
	   %>
	   <tr> 
	   <!--
		<td width="30%" align="center"><font face="Verdana" size="2"><strong><u><%=rsttd("FCKTU")%></u></strong></td>
		<td width="30%" align="center"><font face="Verdana" size="2"><strong><u><%=rsttd("FCMANAGER")%></u></strong></td>
		<td width="30%" align="center"><font face="Verdana" size="2"><strong><u><%=rsttd("FCMANAGER")%></u></strong></td>
		-->
	   </tr>
	<%  End If
%> 
	   	<tr> 
			<td width="30%" align="center"><font face="Verdana" size="1"><strong><i>Krani Timbangan</i></strong></td>
			<td width="30%" align="center"><font face="Verdana" size="1"><strong><i>Manager</i></strong></td>
			<td width="30%" align="center"><font face="Verdana" size="1"><strong><i>General Manager</i></strong></td>
	    </tr>
   </table>
   </body>
   </html>
