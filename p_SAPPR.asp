<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/ConndB1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%
Dim ScriptName, v_txtDateFrom, v_txtDateTo, v_txtprcode
Dim rsDefaultBa, strDefaultBa, varDefaultBa, v_txtFcba
Dim rsBaInfo, strBaInfo, varFcbaName, varCompanyName
Dim no

v_txtDateFrom = Request("txtDateFrom")
v_txtDateTo = Request("txtDateTo")
v_txtFcba = Request("txtFcba")
v_txtprcode = Request("txtprcode")

If Len(Trim(v_txtDateFrom)) = 0 Then v_txtDateFrom = textDate(Date)
If Len(Trim(v_txtDateTo)) = 0 Then v_txtDateTo = textDate(Date)

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

ScriptName = Request.ServerVariables("SCRIPT_NAME")

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
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>ePlaS Reporting - Upload PR to SAP</title>
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
<style type="text/css">
@media print
{
.printhide{display:none;}
}
</style>
</head>
<body>
	<div id="wrapper">
		<div id="header-wrapper">
			<div class="printhide">
				<!--#include file="../print_excel.asp"-->
				<!--#include file="menutop.asp"-->
			</div>
			<form name="frmEntry" method="post" action="<%=ScriptName%>">
			<div class="printhide">
				<table id="form-table" class="centered-margin">
					<tr>
						<td><font face="verdana" size="1">Dari</font></td>
						<td align="center">
						<% If v_txtDateFrom <> "" And v_txtDateFrom = v_txtDateFrom Then %>
							<input class="datepicker" type="text" maxlength="10" name="txtDateFrom" size="10" value="<%=v_txtDateFrom%>" title="Type Date mm/dd/yyyy" >
						<% Else %>
							<input class="datepicker" type="text" maxlength="10" name="txtDateFrom" size="10" value="<%=textDate(Date)%>" title="Type Date mm/dd/yyyy" >
						<% End If%>
						</td>
						<td><font face="verdana" size="1">Sampai</font></td>
						<td align="center">
						<% If v_txtDateTo <> "" And v_txtDateTo = v_txtDateTo Then %>
							<input class="datepicker" type="text" maxlength="10" name="txtDateTo" size="10" value="<%=v_txtDateTo%>" title="Type Date mm/dd/yyyy" >
						<% Else %>
							<input class="datepicker" type="text" maxlength="10" name="txtDateTo" size="10" value="<%=textDate(Date)%>" title="Type Date mm/dd/yyyy" >
						<% End If%>
						</td>
						<td><font face="verdana" size="1">PR NO</font></td>
						<td align="center">
							<input type="text" maxlength="50" name="txtprcode" size="50" value="<%=v_txtprcode%>" title="Type PR NO" >
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
			<thead>
				<tr>
					<th colspan="39" style="background-color:#FFFFFF; border:none;">
						<div id="header-wrapper">
							<div id="head_judul" style="margin-bottom:-30px; position:static; text-align:left;">
								<h4><%If Len(Trim(varCompanyName)) > 0 Then%>PT. <%End If%><%=varCompanyName%></h4>
								<h5><%=varFcbaName%></h5>
							</div>
							<div id="subheader" class="centered">
								<h2>LAPORAN PR SAP</h2>			 
							<% If (v_txtdatefrom <> "" Or v_txtdateto <> "") And v_txtdatefrom = v_txtdateto Then%>
								<h5>TANGGAL: <%=v_txtdateto%></h5>
							<% Else %>
								<h5>TANGGAL: <%=v_txtdatefrom%> - <%=v_txtdateto%></h5>
							<% End If %>	
							</div>
						</div>
					</th>
				</tr>
				<tr>
					<th style="color:#fff;">NO PR iPlaS</th>		
					<th style="color:#fff;">DOCUMENT TYPE</th>
					<th style="color:#fff;">L</th>
					<th style="color:#fff;">ITEMNUMBER</th>
					<th style="color:#fff;">MATERIAL CODE</th>
					<th style="color:#fff;">PR QUANTITY</th>
					<th style="color:#fff;">MATERIAL UOM</th>
					<th style="color:#fff;">&nbsp;</th>
					<th style="color:#fff;">PLANT</th>
					<th style="color:#fff;">STORAGE LOCATION</th>
					<th style="color:#fff;">MATERIAL DESCRIPTION</th>
					<th style="color:#fff;">DELIVERY DATE</th>
					<th style="color:#fff;">REQUISITIONER</th>
					<th style="color:#fff;">PURCHASING GROUP</th>
					<th style="color:#fff;">PR DATE</th>
					<th style="color:#fff;">VALUATION PRICE</th>
					<th style="color:#fff;">CURRENCY</th>
					<th style="color:#fff;">PRICE UNIT</th>
					<th style="color:#fff;">&nbsp;</th>
					<th style="color:#fff;">PURCHASING ORGANIZATION</th>
					<th style="color:#fff;">&nbsp;</th>
					<th style="color:#fff;">ITEM TEXT 1</th>
					<th style="color:#fff;">ITEM TEXT 2</th>
					<th style="color:#fff;">ITEM TEXT 3</th>
					<th style="color:#fff;">ITEM TEXT 4</th>
					<th style="color:#fff;">ITEM TEXT 5</th>
					<th style="color:#fff;">ITEM NOTE 1</th>
					<th style="color:#fff;">ITEM NOTE 2</th>
					<th style="color:#fff;">ITEM NOTE 3</th>
					<th style="color:#fff;">ITEM NOTE 4</th>
					<th style="color:#fff;">ITEM NOTE 5</th>
					<th style="color:#fff;">ITEM DELIVERY 1</th>
					<th style="color:#fff;">ITEM DELIVERY 2</th>
					<th style="color:#fff;">ITEM DELIVERY 3</th>
					<th style="color:#fff;">ITEM DELIVERY 4</th>
					<th style="color:#fff;">ITEM DELIVERY 5</th>
					<th style="color:#fff;">ITEM DELIVERY 6</th>
					<th style="color:#fff;">NO AFCE</th>
					<th style="color:#fff;">PROJECTCODE</th>
				</tr>
			</thead>
			<tbody>
<%
	Dim rsSAPPR, strSAPPR , noline
	
	strSAPPR	=	" select purchase_report_api.getdocumenttype (a.fcba) document_type" & _
					"        , b.thenumber" & _
					"        , a.lineid" & _
					"        , a.itemcode" & _
					"        , a.approvedqty" & _
					"        , a.unitofmeasurement" & _
					"        , purchase_report_api.GetPlantReport (a.fcba, b.managedid) plant" & _
					"        , purchase_report_api.getstoragelocation (a.fcba) storage_location" & _
					"        , to_char (a.wanteddeliverydate, 'dd.mm.yyyy') delivery_date" & _
					"        , b.requestorid" & _
					"        , b.purchasegroupid" & _
					"        , to_char (b.prdate, 'dd.mm.yyyy') prdate" & _
					"        , purchase_report_api.getpurchasingorganization (a.fcba, b.managedid) purchasing_organization" & _
					"        ,stock_api.get_stockName(a.itemcode,a.fcba) Matrial_desc" & _
					"        , a.description" & _
					"        , a.prcode" & _
					" from purchaserequestdetails a," & _
					"      (" & _
					"         Select rownum thenumber" & _
					"                , prcode" & _
					"                , fcba" & _
					"                , prdate" & _
					"                , managedid" & _
					"                , requestorid" & _
					"                , purchasegroupid" & _
					"                , rowstateid" & _
					"         from (" & _
					"             select prcode" & _
					"                    , fcba" & _
					"                    , prdate" & _
					"                    , managedid" & _
					"                    , requestorid" & _
					"                    , purchasegroupid" & _
					"                    , rowstateid" & _
					"             from purchaserequest" & _
					"             where upper(rowstateid) = upper('Approved')" & _
					"                   and upper(managedid) <> upper('LOCAL')"
If  len(trim(v_txtFcba)) > 0 Then
	If v_txtFcba = "ALL" Then
		strSAPPR = strSAPPR & " "
	Else
		strSAPPR = strSAPPR & "             and fcba = '" & v_txtFcba & "'"
	End If
End If
If  len(trim(v_txtDateFrom)) > 0 Then
	strSAPPR = strSAPPR & "             and prdate >= to_date('" & v_txtDateFrom & "', 'mm/dd/yyyy')"
End If
If  len(trim(v_txtDateTo)) > 0 Then
	strSAPPR = strSAPPR & "             and prdate <= to_date('" & v_txtDateTo & "', 'mm/dd/yyyy')"
End If
If  len(trim(v_txtprcode)) > 0 Then
	strSAPPR = strSAPPR & "             and prcode = '" & v_txtprcode & "'"
End If

strSAPPR = strSAPPR & "             order by prcode" & _
					  "         )" & _
					  "      ) b" & _
					  " where a.prcode = b.prcode" & _
					  "       and a.fcba = b.fcba" & _
					  "       and a.approvedqty <> 0"

	Set rsSAPPR = Server.CreateObject("ADODB.Recordset")
	'response.write strSAPPR
	rsSAPPR.Open strSAPPR, connstr, adOpenStatic

	If rsSAPPR.BOF And rsSAPPR.EOF Then
%>
				<tr>
					<td align="left" colspan="39"><strong>DATA TIDAK DITEMUKAN</strong></td>
				</tr>	
<%
	ElseIf Not rsSAPPR.EOF then		
		While Not rsSAPPR.EOF
%>
				<tr>
					<td align="left"><%=rsSAPPR("prcode")%></td>
					<td align="left"><%=rsSAPPR("document_type")%></td>
					<td align="right"><%=FbpNum2(rsSAPPR("theNumber"))%></td>
					<td align="right"><%=rsSAPPR("lineid")%></td>
					<td align="left"><%=rsSAPPR("itemcode")%></td>
					<td align="right"><%=FbpNum2(rsSAPPR("approvedqty"))%></td>
					<td align="left"><%=rsSAPPR("unitofmeasurement")%></td>
					<td align="left">&nbsp;</td>
					<td align="left"><%=rsSAPPR("Plant")%></td>
					<td align="left"><%=rsSAPPR("Storage_Location")%></td>
					<td align="left"><%=rsSAPPR("Matrial_desc")%></td>
					<td align="right"><%=FbpNum2(rsSAPPR("delivery_date"))%></td>
					<td align="left"><%=rsSAPPR("REQUESTORID")%></td>
					<td align="left"><%=rsSAPPR("PURCHASEGROUPID")%></td>
					<td align="right"><%=FbpNum2(rsSAPPR("prdate"))%></td>
					<td align="left">1</td>
					<td align="left">IDR</td>
					<td align="left">1</td>
					<td align="left">&nbsp;</td>
					<td align="left"><%=rsSAPPR("Purchasing_Organization")%></td>
					<td align="left">&nbsp;</td>
					<td align="left">&nbsp;</td>
					<td align="left">&nbsp;</td>
					<td align="left">&nbsp;</td>
					<td align="left">&nbsp;</td>
					<td align="left">&nbsp;</td>
					<td align="left"><%=rsSAPPR("prcode")%></td>
					<td align="left"><%=rsSAPPR("description")%></td>
					<td align="left">&nbsp;</td>
					<td align="left">&nbsp;</td>
					<td align="left">&nbsp;</td>
					<td align="left">&nbsp;</td>
					<td align="left">&nbsp;</td>
					<td align="left">&nbsp;</td>
					<td align="left">&nbsp;</td>
					<td align="left">&nbsp;</td>
					<td align="left">&nbsp;</td>
					<td align="left">&nbsp;</td>
					<td align="left">&nbsp;</td>
				</tr>
<%
			rsSAPPR.MoveNext
		Wend
	End If
rsSAPPR.Close
Set rsSAPPR = Nothing
%>
			</tbody>
		</table>
	</div>
</body>
</html>