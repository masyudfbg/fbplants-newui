<% ValidateSession %>
<!--#include virtual="/conndB1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
maxCol = 0

txtYear = Request("txtYear")
txtMonth = Request("txtMonth")

txttdate = Request("txttdate")

If "" & txtYear = "" Then txtYear = Year(Session("v_Costbook"))
If "" & txtMonth = "" Then txtMonth = Month(Session("v_Costbook"))
If "" & txttdate = "" Then txttdate = textDate(date)

txttdate = txtMonth & "" & txtYear

add_col_bud_prod_qty = ""
d_span = 28
	
v_txtFcba = Request("txtFcba")
v_txtItemCode = Request("txtItemCode")

  
    if  v_txtfcba <> "--ALL--"  and v_txtfcba <> "" then
       whereCons = whereCons & " And fcba = '" & v_txtfcba & "' "
	end if
    ' selectCons = selectCons & ",fcba"
    ' orderCons  = orderCons & ",BUSINESSUNIT_API.get_OrderData(fcba)"
	q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & v_txtfcba & "' "
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

	if  v_txtItemCode <> "--ALL--"  and v_txtItemCode <> "" then
    	whereCons = whereCons & "AND INSTR('"& v_txtItemCode &"', SUBSTR(itemcode,1,6) ) <> 0 "
	end if

strDefaultBa = " SELECT" & _
				"     FCCODE as FCBA" & _
				" FROM BUSINESSUNIT BU" & _
				" INNER JOIN USER_FCBA_TAB UFT ON (BU.FCCODE = UFT.FCBA) " & _
				" WHERE" & _
				"     ISDEFAULT = 'TRUE'" & _
				"     AND FCTYPE = 'E'" & _
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
ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

 


strBaInfo = " SELECT" & _
			"     FCNAME as FCBANAME" & _
			"    , fccompanycode  , FCCOMPANYNAME as COMPANYNAME" & _
			" FROM BUSINESSUNIT BU" & _
			" WHERE" & _
			"     FCCODE = '" & v_txtFcba & "'"
' response.write strBaInfo
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
<title>ePlaS Reporting - Budget Purchase</title>
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

th.rotate-270 {
  height: 200px; /* 1 */
  white-space: nowrap;
  min-width:40px;
  max-width:40px;
  overflow:hidden;
}
th.rotate-270 > div {
	transform: translate(8px /* 2 */ , 80px /* 3 */ ) rotate(-90deg);
	-webkit-transform: translate(8px /* 2 */ , 80px /* 3 */ ) rotate(-90deg);
	-ms-transform: translate(8px /* 2 */ , 80px /* 3 */ ) rotate(-90deg);
    width:15px; /* 4 */
	font-size:10px;
	font-weight:normal;
}

/*
1. Mengatur tinggi table head.
2. mengatur posisi X div
3. mengatur posisi Y div.
4. mengatur lebarnya DIV yang akan mempengaruhi Lebarnya Table Head.
*/
</style>


<link href="/plugins/chosen/chosen.min.css" rel="stylesheet" />
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
							<th>SITE</th>
							<th>TAHUN</th>
							<th>KODE MATERIAL</th>
						</tr>
						<tr>
							<td>
								<select name="txtFcba" id="vFcba" onchange="javascript:filtering()">
									<% 
									Set rsFcba = Server.CreateObject("ADODB.Recordset")
									rsFcba.CursorLocation = adUseClient
									strsql =    " SELECT '--ALL--' as fcba from DUAL UNION " & _  
									            " select fccode fcba from businessunit " & _
												" where " & _
												" USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',fccode) IS NOT NULL " 
									  
									
									rsFcba.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
									if not rsFcba.eof then
										DO while not rsFcba.eof
											if v_txtFcba = rsFcba("fcba") then
											%>			
												<option value="<%=rsFcba("fcba")%>" Selected><%=rsFcba("fcba")%></option>
											<% Else %>	
												<option value="<%=rsFcba("fcba")%>"><%=rsFcba("fcba")%></option>
											<% End If
											
											rsFcba.movenext
										loop
									End if
									rsFcba.Close
									set rsFcba = Nothing
									%>
								</select>
							</td> 
								<% 
									Dim mon, Yr, MinYr, MaxYr, i, MM

									Yr = txtYear
									MinYr = 2015
									MaxYr = Yr  + 2
								%>
										<td width="4">
											<select name="txtYear">
								<%
									For i = MinYr to MaxYr
								%>
										<Option Value="<%=i%>" <%If cint(i) = cint(txtYear) then%> Selected <%End If%>><%=i%></option>
								<%
									Next
								%>
								</select>
							<td>
								<select name="txtItemCode"  multiple class="chosen-select">
									<option>--ALL--</option>
									<% 
									Set rsFcba = Server.CreateObject("ADODB.Recordset")
									rsFcba.CursorLocation = adUseClient
									strsql =    " SELECT DISTINCT fccode, fcName FROM bankstock ORDER BY fccode" 
									  
									
									rsFcba.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
									if not rsFcba.eof then
										DO while not rsFcba.eof
											if v_txtItemCode = rsFcba("Fccode") then
											%>			
												<option value="<%=rsFcba("fccode")%>" Selected><%=rsFcba("fccode")%> : <%=rsFcba("fcName")%></option>
											<% Else %>	
												<option value="<%=rsFcba("fccode")%>"><%=rsFcba("fccode")%> : <%=rsFcba("fcName")%></option>
											<% End If
											
											rsFcba.movenext
										loop
									End if
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
<table border="0" width="100%" id="product-table" class="xthf-site" >
	<thead style="border-left:1px #000000 solid;">
	<tr>
         <th colspan="<%=d_span%>" style="background:#FFFFFF; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
		 <div id="header-wrapper">
			
			<div id="subheader" class="centered">
				<h2>BUDGET PURCHASE REPORT</h2>
				<br/>
				<h4>PERIODE&nbsp;:&nbsp;<%=Bulan(txtMonth)%>&nbsp;<%=txtYear%></h4>
			</div>
          </div>  
		 </th>
	</tr>
	<tr>
		<th rowspan="2">Budget Year</th>
		<th rowspan="2">Stock Code</th>
		<th rowspan="2">Stock Name</th>
		<th rowspan="2">UOM</th>
		<th rowspan="2">Budget Quantity</th>
		<th colspan="2">Purchase Request</th>
		<th colspan="5">Purchase Order</th>
	</tr> 
	<tr>
		<th>PR Quantity</th>
		<th>Ending Balance PR Quantity</th>

		<th>PO Quantity</th>
		<th>PO Amount</th>
		<th>Ending Balance PO Quantity</th>
		<th>Budget Amount</th>
		<th>Ending Balance PO Amount</th>
	</tr>

    </thead>
	<tbody>
		<%
		
		v_budget_qty = 0
		v_pr_qty = 0
		v_ending_balance_qty_pr = 0 
		v_po_qty = 0
		v_po_amount = 0
		v_ending_balance_qty_po = 0
		v_budget_amount = 0
		v_ending_balance_amount_po = 0
		 

		strPusingan	=  	" SELECT btp.fcba," & _
						"        btp.budgetyear," & _
						"        btp.stockcode," & _
						"        st.fcname AS stockname," & _
						"        st.unitofmeasurement," & _
						"        btp.quantity AS budget_qty," & _
						"        pr.pr_qty," & _
						"        po.po_qty," & _
						"        (btp.quantity - pr.pr_qty) AS ending_balance_qty_pr," & _
						"        (btp.quantity - po.po_qty) AS ending_balance_qty_po," & _
						"        btp.amount AS budget_amount," & _
						"        po.po_amount," & _
						"        (btp.amount - po.po_amount) AS ending_balance_amount_po" & _
						"   FROM budget_transaction_purchase btp" & _
						"        INNER JOIN" & _
						"        (  SELECT a.fcba," & _
						"                  TO_CHAR (b.prdate, 'yyyy') AS pryear," & _
						"                  a.itemcode," & _
						"                  SUM (a.ApprovedQTY) AS pr_qty" & _
						"             FROM PURCHASEREQUESTdetails a" & _
						"                  INNER JOIN purchaserequest b" & _
						"                     ON (a.fcba = b.fcba AND a.prcode = b.prcode)" & _
						"            WHERE     b.rowstateid = 'Approved'" & _
						"                  AND TO_CHAR (b.PRDATE, 'YYYY') = '" & txtYear & "'" & _
						"                  AND a.fcba = '" & v_txtfcba & "'" & _ 
						"         GROUP BY a.fcba, TO_CHAR (b.prdate, 'yyyy'), a.itemcode) pr" & _
						"           ON (    pr.fcba = btp.fcba" & _
						"               AND pr.pryear = btp.budgetyear" & _
						"               AND pr.itemcode = btp.stockcode)" & _
						"        INNER JOIN" & _
						"        (  SELECT a.fcba," & _
						"                  TO_CHAR (b.podate, 'yyyy') AS poyear," & _
						"                  a.itemcode," & _
						"                  SUM (a.quantity) AS po_qty," & _
						"                  SUM (a.amount) AS po_amount" & _
						"             FROM purchaseorderdetails a" & _
						"                  INNER JOIN purchaseorder b" & _
						"                     ON (a.fcba = b.fcba AND a.pocode = b.pocode)" & _
						"            WHERE     b.rowstateid = 'Ordered'" & _
						"                  AND TO_CHAR (b.PoDATE, 'YYYY') = '" & txtYear & "'" & _
						"                  AND a.fcba = '" & v_txtfcba & "'" & _ 
						"         GROUP BY a.fcba, TO_CHAR (b.podate, 'yyyy'), a.itemcode) po" & _
						"           ON (    po.fcba = btp.fcba" & _
						"               AND po.poyear = btp.budgetyear" & _
						"               AND po.itemcode = btp.stockcode)" & _
						"        INNER JOIN stock st" & _
						"           ON (st.fcba = btp.fcba AND st.fccode = btp.stockcode)" 


		' response.write strPusingan
		Set qORS = Server.CreateObject("ADODB.Recordset")
		qORS.Open strPusingan, connstr, adOpenStatic
		 

		if Not qORS.eof then
			While Not qORS.EOF
		%>
		<tr>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=qORS("budgetyear")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=qORS("stockcode")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=qORS("stockname")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=qORS("unitofmeasurement")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(cdbl(to_nol(qORS("budget_qty"))),0)%></td> 
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(cdbl(to_nol(qORS("pr_qty"))),0)%></td> 
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(cdbl(to_nol(qORS("ending_balance_qty_pr"))),0)%></td> 
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(cdbl(to_nol(qORS("po_qty"))),0)%></td> 
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(cdbl(to_nol(qORS("po_amount"))),0)%></td> 
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(cdbl(to_nol(qORS("ending_balance_qty_po"))),0)%></td> 
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(cdbl(to_nol(qORS("budget_amount"))),0)%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(cdbl(to_nol(qORS("ending_balance_amount_po"))),0)%></td> 
		</tr>
	<% 
		v_budget_qty =  v_budget_qty +  cdbl(to_nol(qORS("budget_qty")))
		v_pr_qty =  v_pr_qty +  cdbl(to_nol(qORS("pr_qty")))
		v_ending_balance_qty_pr = v_ending_balance_qty_pr +  cdbl(to_nol(qORS("ending_balance_qty_pr"))) 
		v_po_qty =  v_po_qty +  cdbl(to_nol(qORS("po_qty")))
		v_po_amount =  v_po_amount +  cdbl(to_nol(qORS("po_amount")))
		v_ending_balance_qty_po =  v_ending_balance_qty_po +  cdbl(to_nol(qORS("ending_balance_qty_po")))
		v_budget_amount =  v_budget_amount +  cdbl(to_nol(qORS("budget_amount")))
		v_ending_balance_amount_po = v_ending_balance_amount_po +  cdbl(to_nol(qORS("ending_balance_amount_po"))) 


		qORS.MoveNext
		Wend
	Else
		response.write "<tr><td colspan='26'>No Data Available</td></tr>"
	End If
	qORS.close()
	qORS=null%>
	<input  type="hidden" value="<%=txtTable%>" name="txtTable">
	<input  type="hidden" value="<%=txtModule%>" name="txtModule">
		<tr bgcolor="#F4F4A0">
			<td align="center" colspan="4"><strong>TOTAL</strong></td>
			<td align="right"><strong><%=FormatNumber(cdbl(to_nol(v_budget_qty)),0)%></strong></td>
			<td align="right"><strong><%=FormatNumber(cdbl(to_nol(v_pr_qty)),0)%></strong></td>
			<td align="right"><strong><%=FormatNumber(cdbl(to_nol(v_ending_balance_qty_pr)),0)%></strong></td>
			<td align="right"><strong><%=FormatNumber(cdbl(to_nol(v_po_qty)),0)%></strong></td>
			<td align="right"><strong><%=FormatNumber(cdbl(to_nol(v_po_amount)),0)%></strong></td>
			<td align="right"><strong><%=FormatNumber(cdbl(to_nol(v_ending_balance_qty_po)),0)%></strong></td> 
			<td align="right"><strong><%=FormatNumber(cdbl(to_nol(v_budget_amount)),0)%></strong></td> 
			<td align="right"><strong><%=FormatNumber(cdbl(to_nol(v_ending_balance_amount_po)),0)%></strong></td> 
		</tr>
	</tbody>
</table>


<script src="/plugins/chosen/chosen.jquery.min.js"></script>
	<script type="text/javascript">	
		var config = {
		  '.chosen-select'           : {},
		  '.chosen-select-deselect'  : {allow_single_deselect:true},
		  '.chosen-select-no-single' : {disable_search_threshold:10},
		  '.chosen-select-no-results': {no_results_text:'Oops, nothing found!'},
		  '.chosen-select-width'     : {width:"95%"}
		}
		for (var selector in config) {
		  $(selector).chosen(config[selector]);
		}
		
	</script>
	</body>
</html>