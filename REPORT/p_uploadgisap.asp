<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/conndB1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
Dim v_txtFcba, v_txtYear, v_txtMonth,v_txtDay,v_txtDate, v_txtAfdeling, v_txtBlok, v_txtColor
Dim ScriptName
Dim rsuploadgisap, struploadgisap, strsql
Dim NO, cClass, nday,ColorID
Dim rsDefaultBa, strDefaultBa, varDefaultBa
Dim rsBaInfo, strBaInfo, varFcbaName, varCompanyName
Dim oRS, oSQL
Dim v_txtdatefrom, v_txtdateto, v_txtpocode

v_txtdatefrom = Request("txtdatefrom")
v_txtdateto = Request("txtdateto")

If Len(Trim(v_txtdatefrom)) = 0 Then v_txtdatefrom = textdate(date)
If Len(Trim(v_txtdateto)) = 0 Then v_txtdateto = textdate(date)

v_txtMonth		=	month(v_txtDate)
v_txtDay		=	DAY(v_txtDate)
v_txtYear		=	year(v_txtDate)
	
'v_txtYear		=	Request("txtYear")
'v_txtMonth		=	Request("txtMonth")
v_txtAfdeling	=	Request("txtAfdeling")
v_txtBlok		=	Request("txtBlok")
v_txtFcba		=	Request("txtFcba")
v_txtColor		=	Request("txtColor")
v_txtpocode		=	Request("txtpocode")

Dim v_txtPersen : v_txtPersen = Request("txtPersen")
If Len(trim(v_txtPersen)) = 0 Then v_txtPersen = 0


strDefaultBa = " SELECT" & _
				"     FCCODE as FCBA" & _
				" FROM BUSINESSUNIT BU" & _
				" INNER JOIN USER_FCBA_TAB UFT ON (BU.FCCODE = UFT.FCBA) " & _
				" WHERE" & _
				"     ISDEFAULT = 'TRUE'" & _
				"     AND FCUSERCODE = '" & Session("CodeName") & "'"

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

If v_txtYear = "" Then
	v_txtYear = Year(Session("v_Costbook"))
End If

If v_txtMonth = "" Then
	v_txtMonth = Month(Session("v_Costbook"))
End If

ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

struploadgisap	=	" select documentdate" & _
					"        , postingdate" & _
					"        , materialslip" & _
					"        , headertext" & _
					"        , movementtype" & _
					"        , pla" & _
					"        , storloc" & _
					"        , xfull" & _
					"        , version1" & _
					"        , materialnumber" & _
					"		 , stock_api.get_stockName(materialnumber,fcba) MatrialName " & _
					"        , nvl(sum(nvl(quantityinunit, 0)), 0) quantityinunit" & _
					"        , uom" & _
					"        , plant" & _					
					"        , storagelocation" & _
					"        , batchnumber" & _
					"        , glacc" & _
					"        , fcba" & _
					"        , description" & _
					"        , version2" & _
					"        , costcenter" & _
					" from (" & _
					"     select to_char(rn.rndate,'dd.mm.yyyy') documentdate" & _
					"            , to_char(rn.rndate,'dd.mm.yyyy') postingdate" & _
					"            , rn.pocode materialslip" & _
					"            , rn.pocode headertext" & _
					"            , '201' movementtype" & _
					"            , supplier_api.get_mapping_sap(rn.suppliercode, rnd.fcba) pla" & _
					"            , purchase_report_api.getstoragelocation(rnd.fcba) storloc" & _
					"            , 'X' xfull" & _
					"            , 'X' version1" & _
					"            , rnd.purchaseitemcode materialnumber" & _
					"            , rnd.quantity quantityinunit" & _
					"            , null uom" & _
					"            , supplier_api.get_mapping_sap(rn.suppliercode, rnd.fcba) plant" & _
					"            , purchase_report_api.getstoragelocation(rnd.fcba) storagelocation" & _
					"            , substr(rn.pocode,1,10) batchnumber" & _
					"            , '9000002' glacc" & _
					"            , rnd.fcba fcba" & _
					"            , rn.pocode description" & _
					"            , 'X' version2" & _
					"            , erp_report_api.getcompany(rnd.fcba)||'02TC8999' costcenter" & _					
					"     from receivenotedetail rnd" & _
					"          , receivenote rn" & _
					"     where rnd.fcba = rn.fcba" & _
					"           and rnd.receivenotecode = rn.receivenotecode" & _
					"           and supplier_api.get_isincludegisap(rn.suppliercode, rnd.fcba) = 'TRUE'"
If  len(trim(v_txtFcba)) > 0 Then
	If v_txtFcba = "ALL" Then
		struploadgisap = struploadgisap & " "
	Else
		struploadgisap = struploadgisap & "       and rnd.fcba = '" & v_txtFcba & "'"
	End If
End If
If Len(trim(v_txtdatefrom)) > 0 then
	struploadgisap = struploadgisap & "           and rn.rndate >= to_date('" & v_txtdatefrom & "', 'mm/dd/yyyy')"
End If

If Len(trim(v_txtdateto)) > 0 then
	struploadgisap = struploadgisap & "           and rn.rndate <= to_date('" & v_txtdateto & "', 'mm/dd/yyyy')"
End If

If  len(trim(v_txtpocode)) > 0 Then
	struploadgisap = struploadgisap & "           and pocode = '" & v_txtpocode & "'"
End If

	struploadgisap = struploadgisap & "     )" & _
									  " group by documentdate" & _
									  "          , postingdate" & _
									  "          , materialslip" & _
									  "          , headertext" & _
									  "          , movementtype" & _
									  "          , pla" & _
									  "          , storloc" & _
									  "          , xfull" & _
									  "          , version1" & _
									  "          , materialnumber" & _
									  "          , uom" & _
									  "          , plant" & _					
									  "          , storagelocation" & _
									  "          , batchnumber" & _
									  "          , glacc" & _
									  "          , fcba" & _
									  "          , description" & _
									  "          , version2" & _
									  "          , costcenter" & _
									  " order by fcba" & _
									  "          , documentdate" & _
									  "          , postingdate" & _
									  "          , description"

strBaInfo = " SELECT" & _
			"     FCNAME as FCBANAME" & _
			"     , FCCOMPANYNAME as COMPANYNAME" & _
			" FROM BUSINESSUNIT BU" & _
			" WHERE" & _
			"     FCCODE = '" & v_txtFcba & "'"

Set rsBaInfo = Server.CreateObject("ADODB.Recordset")
rsBaInfo.Open strBaInfo, connstr, adOpenStatic

If Not rsBaInfo.EOF Then
	varFcbaName = rsBaInfo("FCBANAME")
	varCompanyName = rsBaInfo("COMPANYNAME")
Else
	varFcbaName = ""
	varCompanyName = ""
End If
rsBaInfo.Close
set rsBaInfo = Nothing

'response.write struploadgisap
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
<title>ePlaS Reporting - Upload GI SAP</title>
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
							<td>
								<font face="verdana" size="1">FROM</font>
							</td>
							<td width="2">
								<input class="datepicker" type="text" name="txtdatefrom" value="<%=v_txtdatefrom%>" />
							</td>
							<td>
								<font face="verdana" size="1">TO</font>
							</td>
							<td width="2">
								<input class="datepicker" type="text" name="txtdateto" value="<%=v_txtdateto%>" />
							</td>
							<td><font face="verdana" size="1">PO NO</font></td>
							<td align="center">
								<input type="text" maxlength="50" name="txtpocode" size="50" value="<%=v_txtpocode%>" title="Type PO NO" >
							</td>
							<td>
								<font face="verdana" size="1"><%=titleSite%></font>
							</td>
							<td>
								<select name="txtFcba" >
									<option value="ALL">--ALL--</option>
									<%
									Dim rsFcba, strFcba
									strFcba	=	" SELECT FCCODE, ORDERDATA FROM BUSINESSUNIT" & _
												" WHERE USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',fccode) IS NOT NULL" & _												
												" ORDER BY ORDERDATA"
									
									Set rsFcba = Server.CreateObject("ADODB.Recordset")
									rsFcba.Open strFcba, connstr, adOpenStatic
									
										if not rsFcba.eof then
										Do While Not rsFcba.eof
										%>
										<option value="<%=rsFcba("FCCODE")%>" <%if v_txtFcba = rsFcba("FCCODE") then%> Selected <%End If%>><%=rsFcba("FCCODE")%></option>
										<%
											rsFcba.movenext
										loop
									end if
									rsFcba.Close
									set rsFcba = Nothing
									%>
								</select>
							</td>
							<td><input type="submit" value="View" /></td>
						</tr>				
					</table>
				</div>
			</form>
		</div>
<table border="0" width="100%" id="product-table" class="xthf-site">
	<thead style="border-left:1px #000000 solid;">
	<tr>
         <th colspan="<%=21%>" style="background:#FFFFFF; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
		 <div id="header-wrapper">
			<div id="head_judul">
				<h4><%If Len(Trim(varCompanyName)) > 0 Then%>PT. <%End If%><%=varCompanyName%></h4>
				<h5><%=varFcbaName%></h5>
			</div>
			<div id="subheader" class="centered">
				<h2>UPLOAD GI SAP</h3>
				<br/>
				<%If v_txtdateto = v_txtdatefrom then%>
					<h4>
						Tanggal&nbsp;:&nbsp;<%=Day(v_txtdatefrom)%>&nbsp;<%=Bulan(Month(v_txtdatefrom))%>&nbsp;<%=Year(v_txtdatefrom)%>
					</h4>
				<%Else%>
					<h4>
						Tanggal&nbsp;:&nbsp;<%=Day(v_txtdatefrom)%>&nbsp;<%=Bulan(Month(v_txtdatefrom))%>&nbsp;<%=Year(v_txtdatefrom)%>
						&nbsp;s/d&nbsp;Tanggal&nbsp;:&nbsp;<%=Day(v_txtdateto)%>&nbsp;<%=Bulan(Month(v_txtdateto))%>&nbsp;<%=Year(v_txtdateto)%>
					</h4>
				<%End If%>
				
			</div>
          </div>  
		 </th>
	</tr>
	
	<tr>
		<th style="background-color:#00A300;color:#fff;">Doc Date</th>
		<th style="background-color:#00A300;color:#fff;">Posting Date</th>
		<th style="background-color:#00A300;color:#fff;">Material Slip</th>
		<th style="background-color:#00A300;color:#fff;">Header Text</th>
		<th style="background-color:#00A300;color:#fff;">Mov Type</th>
		<th style="background-color:#00A300;color:#fff;">Plant</th>
		<th style="background-color:#00A300;color:#fff;">Stor Loc</th>
		<th style="background-color:#00A300;color:#fff;">XFULL</th>
		<th style="background-color:#00A300;color:#fff;">Version 1</th>
		<th style="background-color:#00A300;color:#fff;">Material No</th>
		<th style="background-color:#00A300;color:#fff;">Matrial Name</th>
		<th style="background-color:#00A300;color:#fff;">Qty</th>
		<th style="background-color:#00A300;color:#fff;">Uom</th>
		<th style="background-color:#00A300;color:#fff;">Plant</th>
		<th style="background-color:#00A300;color:#fff;">Storage Location</th>
		<th style="background-color:#00A300;color:#fff;">Batch</th>
		<th style="background-color:#00A300;color:#fff;">Gl acc</th>
		<th style="background-color:#00A300;color:#fff;">Good Receipt</th>
		<th style="background-color:#00A300;color:#fff;">Description</th>
		<th style="background-color:#00A300;color:#fff;">Version 2</th>
		<th style="background-color:#00A300;color:#fff;">Cost Center</th>
	</tr>
    </thead>
	<tbody>
<%
Set rsuploadgisap = Server.CreateObject("ADODB.Recordset")
rsuploadgisap.Open struploadgisap, connstr, adOpenStatic

if rsuploadgisap.bof and rsuploadgisap.eof then

%>
	<tr>
		<td colspan="<%=20%>">NO DATA AVAILABLE</td>
	</tr>
<%end if


		while not rsuploadgisap.eof
		
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
%>

			<tr class="<%=cClass%>">
				<td align="left"><%=rsuploadgisap("documentdate")%></td>
				<td align="left"><%=rsuploadgisap("postingdate")%></td>
				<td align="left"><%=rsuploadgisap("materialslip")%></td>
				<td align="left"><%=rsuploadgisap("headertext")%></td>
				<td align="right"><%=FbpNum2(rsuploadgisap("movementtype"))%></td>
				<td align="left"><%=rsuploadgisap("pla")%></td>
				<td align="left"><%=rsuploadgisap("storloc")%></td>
				<td align="left"><%=rsuploadgisap("xfull")%></td>
				<td align="left"><%=rsuploadgisap("version1")%></td>
				<td align="left"><%=rsuploadgisap("materialnumber")%></td>
				<td align="left"><%=rsuploadgisap("MatrialName")%></td>
				<td align="right"><%=FbpNum2(rsuploadgisap("quantityinunit"))%></td>
				<td align="left"><%=rsuploadgisap("uom")%></td>
				<td align="left"><%=rsuploadgisap("plant")%></td>
				<td align="left"><%=rsuploadgisap("storagelocation")%></td>
				<td align="left"><%=rsuploadgisap("batchnumber")%></td>
				<td align="right"><%=FbpNum2(rsuploadgisap("glacc"))%></td>
				<td align="left"><%=rsuploadgisap("fcba")%></td>
				<td align="left"><%=rsuploadgisap("description")%></td>
				<td align="left"><%=rsuploadgisap("version2")%></td>
				<td align="left"><%=rsuploadgisap("costcenter")%></td>
			</tr>

<%
			rsuploadgisap.MoveNext
		Wend
rsuploadgisap.Close
set rsuploadgisap = Nothing
%>

</tbody>
</table>

   </body>
   </html>
