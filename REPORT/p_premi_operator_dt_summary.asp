<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
	Dim v_txtYear, v_txtMonth
	Dim ScriptName
	Dim rsGrnInv, strGrnInv
	Dim v_txtPOCode
	Dim NO, cClass
	
	ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

	v_txtFcba = Request("txtFcba")
	v_txtMonth = Request("txtMonth")
	v_txtYear = Request("txtYear")
	v_txtsectionname = Request("txtsectionname")
	v_txtempcode = Request("txtempcode")

	txttdate = txtMonth & "" & txtYear
	If "" & v_txtYear = "" Then v_txtYear = Year(Session("v_Costbook"))
	If "" & v_txtMonth = "" Then v_txtMonth = Month(Session("v_Costbook"))
	If "" & txttdate = "" Then txttdate = textDate(date)
	varDefaultBa = Session("fcBA")

	v_txtdatefrom = Request("txtdatefrom")
	v_txtdateuntil = Request("txtdateuntil")
	If (v_txtdatefrom = "" Or Len(Trim(v_txtdatefrom)) = 0) Then v_txtdatefrom = textdate(date)
	If (v_txtdateuntil = "" Or Len(Trim(v_txtdateuntil)) = 0) Then v_txtdateuntil = textdate(date)

	If (v_txtFcba = "" Or Len(Trim(v_txtFcba)) = 0) Then
		v_txtFcba = varDefaultBa
	End If	
	nSpanTotal = 0
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

	q = "SELECT type_2 FROM allowded WHERE fccode = 'A62' AND fcba = '"& v_txtfcba &"'"

	' response.Write q
	Set rsCallCombo = Server.CreateObject("ADODB.Recordset")
	rsCallCombo.Open q, connstr
	 
	If not rsCallCombo.eof then
	   	type_report = rsCallCombo("type_2")
	End if 
	
	If type_report = "MIN" Then

		strGrnInv	=	" SELECT FCBA," & _
						" 		 PERIOD_MONTH," & _
						"		 PERIOD_YEAR," & _
						"        OPERATORCODE," & _
						"        OPERATORNAME," & _
						"        TOTAL_PREMI_TONASE_MIN_BASIS," & _ 
						"        HK_EFEKTIF_BRUTO_MIN_BASIS," & _
						"        POTONGAN_HKE," & _
						"        HK_EFEKTIF_NETTO_MIN_BASIS," & _
						"        PREMI_INSENTIF_MIN_BASIS," & _ 
						"        TOTAL_PREMI_OPERATOR_MIN_BASIS" & _
						"  FROM SUMMARY_PREMI_OPR_DT_MINBASIS" & _
						"  WHERE  FCBA = '"& v_txtfcba &"'" 
	Else 
		strGrnInv	=	" SELECT FCBA," & _
						" 		 PERIOD_MONTH," & _
						"		 PERIOD_YEAR," & _
						"        OPERATORCODE," & _
						"        OPERATORNAME," & _ 
						"        TOTAL_PREMI_TONASE_AVG_BASIS," & _ 
						"        HK_EFEKTIF_BRUTO_AVG_BASIS," & _
						"        POTONGAN_HKE," & _ 
						"        HK_EFEKTIF_NETTO_AVG_BASIS," & _
						"        PREMI_INSENTIF_AVG_BASIS," & _
						"        TOTAL_PREMI_OPERATOR_AVG_BASIS" & _
						"  FROM SUMMARY_PREMI_OPR_DT_AVGBASIS" & _
						"  WHERE  FCBA = '"& v_txtfcba &"'" 
	End If 
 
		
		if  len(trim(v_txtYear)) > 0 then
		    strGrnInv	=	strGrnInv & " AND PERIOD_YEAR = '"& v_txtYear &"'" 
		    txtSelect = txtSelect & " Year " & v_txtYear
		end if 
		if  len(trim(v_txtMonth)) > 0 then
		    strGrnInv	=	strGrnInv & " AND PERIOD_MONTH = '"& v_txtMonth &"'" 
		    txtSelect = txtSelect & " Month " & v_txtMonth
		end if 
		if  len(trim(v_txtempcode)) > 0 then
		    strGrnInv	=	strGrnInv & " AND operatorcode = '"& v_txtempcode &"'" 
		    txtSelect = txtSelect & " Employee Code " & v_txtempcode
		end if 
									
		strGrnInv	=	strGrnInv & " ORDER BY fcba," & _
									"          operatorcode," & _
									"		   period_year," & _
									"		   period_month"

		' response.write strGrnInv	
		Set rsGrnInv = Server.CreateObject("ADODB.Recordset")
		rsGrnInv.Open strGrnInv, connstr, adOpenStatic  
%>


<html>

<head>
<title>ePlaS Reporting - Premi Operator DT Summary</title>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<link rel="stylesheet" href="/css/jquery-ui.css" />
<script src="/js/jquery-1.9.1.js"></script>
<script src="/js/jquery-ui.js"></script>
<link href="/plugins/tabcontent/tabcontent.css" rel="stylesheet" type="text/css" />
<link href="/plugins/chosen/chosen.min.css" rel="stylesheet" />

<link rel="stylesheet" href="/plugins/responsivegridsystem/col.css" media="all">
<link rel="stylesheet" href="/plugins/responsivegridsystem/12cols.css" media="all">
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
							<font face="verdana" size="1">KEBUN</font>
						</td>
						<td>
							<select name="txtFcba" id="vFcba" onchange="javascript:filtering()">
							<%
							Dim rsFcba
							Set rsFcba = Server.CreateObject("ADODB.Recordset")
							rsFcba.CursorLocation = adUseClient
							strsql =   " select fccode fcba from businessunit " & _
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
						<td>
							<font face="verdana" size="1">TAHUN</font>
						</td>
						<% 
							Dim mon, Yr, MinYr, MaxYr, i, MM
							mon = Month(Session("v_Costbook"))
							If len(v_txtMonth)=0 then v_txtMonth=mon
						
							Yr = Year(Session("v_Costbook"))
							MinYr = 2015
							MaxYr = Yr  + 2
							If len(v_txtYear)= 0 then v_txtYear = Yr
						%>
						<td width="4">
							<select name="txtYear">
							<%For i = MinYr to MaxYr%>
								<Option Value="<%=i%>" <%If cint(i) = cint(v_txtYear) then%> Selected <%End If%>><%=i%></option>
							<%Next%>
							</select>
						</td>
						<td>
							<font face="verdana" size="1">BULAN</font>
						</td>
						<td width="2">
							<select name="txtMonth">
							<%For i = 1 to 12%>
								<Option Value="<%=i%>" <%If cint(i) = cint(v_txtMonth) then%> Selected <%End If%>><%=Bulan(i)%></option>
							<%Next%>
							</select>
						</td>
						<td>
							<font face="verdana" size="1">EMPLOYEE</font>
						</td>
						<td>
							<select name="txtempcode" class="chosen-select">
								<option value="">-- ALL --</option>
									<% 
									strsectionname	=	"select fccode, fcname from employee where fcba = '"& v_txtFcba &"' and dateterminate is null order by fccode"
									
									Set rssectionname = Server.CreateObject("ADODB.Recordset")
									rssectionname.Open strsectionname, connstr, adOpenStatic 
									
									if not rssectionname.eof then
									Do While Not rssectionname.eof
									if rssectionname("fccode") = v_txtempcode then%>
											<option value="<%=rssectionname("fccode")%>" Selected><%=rssectionname("fccode")%> : <%=rssectionname("fcname")%></option> 
										 <%Else%>
											<option value="<%=rssectionname("fccode")%>"><%=rssectionname("fccode")%> : <%=rssectionname("fcname")%></option>  
										<%End if
										rssectionname.movenext
									loop
									end if
									rssectionname.Close
									set rssectionname = Nothing
									%>
							</select>
						</td> 
						<td>
							<input type="submit" value="View" />
						</td>
					</tr>
				</table>
			</div>
			</form>
			
			</div>
<table border="0" width="100%" id="product-table" class="xthf-site">
	<thead>
		<tr>
    		<th colspan="11" style="background-color:#FFFFFF; border:none;">  
				<div id="header-wrapper">
					<div id="head_judul" style="float:left; text-align:left;">
						<h4>PT. <%=CompanyName%></h4>
						<h5><%=fcbaname%></h5>
					</div>
					<div id="subheader" class="centered" tyle="width:50%; margin:auto;">
						<h2>LAPORAN PREMI OPERATOR DT SUMMARY</h2>
		                <br/>
						<h4>PERIODE: Year <%=txtyear%> Month <%=txtMonth%> </h4>
					</div> 
				</div>
			</th>
		</tr>
		<tr> 
			<th style="background-color:#00A300;color:#fff;">OPERATOR CODE</th> 
			<th style="background-color:#00A300;color:#fff;">OPERATOR NAME</th>
			<th style="background-color:#00A300;color:#fff;">PERIOD YEAR</th> 
			<th style="background-color:#00A300;color:#fff;">PERIOD MONTH</th>
			<th style="background-color:#00A300;color:#fff;">TOTAL PREMI TONASE BASIS</th>
			<th style="background-color:#00A300;color:#fff;">HK EFEKTIF BRUTO BASIS</th>
			<th style="background-color:#00A300;color:#fff;">POTONGAN HKE</th>
			<th style="background-color:#00A300;color:#fff;">HK EFEKTIF NETTO BASIS</th>
			<th style="background-color:#00A300;color:#fff;">PREMI INSENTIF BASIS</th>
			<th style="background-color:#00A300;color:#fff;">TOTAL PREMI OPERATOR BASIS</th>		
		</tr>   
		    </thead>
	<tbody> 
	<% If type_report = "MIN" Then %>
		<%if rsGrnInv.BOF and rsGrnInv.EOF then%>
			<tr>
				<td colspan="11" align="left">NO DATA AVAILABLE</td>
			</tr>
		<%end if
				NO = 1
				rows = 1
				while not rsGrnInv.eof
					if no  mod 2 = 1 then
						cClass = "odd"
					else
						cClass = "default"
					end if
		%>
				<tr class="<%=cClass%>"> 
					<td align="center"><%=rsGrnInv("operatorcode")%></td>
					<td align="center"><%=rsGrnInv("operatorname")%></td>
					<td align="center"><%=rsGrnInv("PERIOD_YEAR")%></td>
					<td align="center"><%=rsGrnInv("PERIOD_MONTH")%></td>
					<td align="right"><%=FormatNumber(cdbl(to_nol(rsGrnInv("TOTAL_PREMI_TONASE_MIN_BASIS"))),2)%></td>
					<td align="right"><%=FormatNumber(cdbl(to_nol(rsGrnInv("HK_EFEKTIF_BRUTO_MIN_BASIS"))),2)%></td>
					<td align="right"><%=FormatNumber(cdbl(to_nol(rsGrnInv("POTONGAN_HKE"))),2)%></td> 
					<td align="right"><%=FormatNumber(cdbl(to_nol(rsGrnInv("HK_EFEKTIF_NETTO_MIN_BASIS"))),2)%></td>
					<td align="right"><%=FormatNumber(cdbl(to_nol(rsGrnInv("PREMI_INSENTIF_MIN_BASIS"))),2)%></td> 
					<td align="right"><%=FormatNumber(cdbl(to_nol(rsGrnInv("TOTAL_PREMI_OPERATOR_MIN_BASIS"))),2)%></td>  
				</tr>
			<%	
				nTOTAL_PREMI_TONASE_MIN_BASIS = nTOTAL_PREMI_TONASE_MIN_BASIS + cdbl(to_nol(rsGrnInv("TOTAL_PREMI_TONASE_MIN_BASIS")))
				nHK_EFEKTIF_BRUTO_MIN_BASIS = nHK_EFEKTIF_BRUTO_MIN_BASIS + cdbl(to_nol(rsGrnInv("HK_EFEKTIF_BRUTO_MIN_BASIS")))	
				nPOTONGAN_HKE = nPOTONGAN_HKE + cdbl(to_nol(rsGrnInv("POTONGAN_HKE")))
				nHK_EFEKTIF_NETTO_MIN_BASIS = nHK_EFEKTIF_NETTO_MIN_BASIS + cdbl(to_nol(rsGrnInv("HK_EFEKTIF_NETTO_MIN_BASIS")))
				nPREMI_INSENTIF_MIN_BASIS = nPREMI_INSENTIF_MIN_BASIS + cdbl(to_nol(rsGrnInv("PREMI_INSENTIF_MIN_BASIS")))	 
				nTOTAL_PREMI_OPERATOR_MIN_BASIS = nTOTAL_PREMI_OPERATOR_MIN_BASIS + cdbl(to_nol(rsGrnInv("TOTAL_PREMI_OPERATOR_MIN_BASIS"))) 
				rsGrnInv.MoveNext			
				Wend
				rsGrnInv.Close
				set rsGrnInv = Nothing
			%> 
				<tr bgcolor="#F4F4A0">	 
					<td align="center" colspan="4"><b> TOTAL</b></td>
					<td align="right"><b><%=formatnumber(nTOTAL_PREMI_TONASE_MIN_BASIS,2)%></td>
					<td align="right"><b><%=formatnumber(nHK_EFEKTIF_BRUTO_MIN_BASIS,2)%></td>
					<td align="right"><b><%=formatnumber(nPOTONGAN_HKE,2)%></td>
					<td align="right"><b><%=formatnumber(nHK_EFEKTIF_NETTO_MIN_BASIS,2)%></td>
					<td align="right"><b><%=formatnumber(nPREMI_INSENTIF_MIN_BASIS,2)%></td>
					<td align="right"><b><%=formatnumber(nTOTAL_PREMI_OPERATOR_MIN_BASIS,2)%></td> 
					<!--<td align="right"></td>-->
				</tr> 
	<%Else%>

		<%if rsGrnInv.BOF and rsGrnInv.EOF then%>
			<tr>
				<td colspan="11" align="left">NO DATA AVAILABLE</td>
			</tr>
		<%end if
				NO = 1
				rows = 1
				while not rsGrnInv.eof
					if no  mod 2 = 1 then
						cClass = "odd"
					else
						cClass = "default"
					end if
		%>
				<tr class="<%=cClass%>"> 
					<td align="center"><%=rsGrnInv("operatorcode")%></td>
					<td align="center"><%=rsGrnInv("operatorname")%></td>
					<td align="center"><%=rsGrnInv("PERIOD_YEAR")%></td>
					<td align="center"><%=rsGrnInv("PERIOD_MONTH")%></td>
					<td align="right"><%=FormatNumber(cdbl(to_nol(rsGrnInv("TOTAL_PREMI_TONASE_AVG_BASIS"))),2)%></td>
					<td align="right"><%=FormatNumber(cdbl(to_nol(rsGrnInv("HK_EFEKTIF_BRUTO_AVG_BASIS"))),2)%></td>
					<td align="right"><%=FormatNumber(cdbl(to_nol(rsGrnInv("POTONGAN_HKE"))),2)%></td> 
					<td align="right"><%=FormatNumber(cdbl(to_nol(rsGrnInv("HK_EFEKTIF_NETTO_AVG_BASIS"))),2)%></td>
					<td align="right"><%=FormatNumber(cdbl(to_nol(rsGrnInv("PREMI_INSENTIF_AVG_BASIS"))),2)%></td> 
					<td align="right"><%=FormatNumber(cdbl(to_nol(rsGrnInv("TOTAL_PREMI_OPERATOR_AVG_BASIS"))),2)%></td>  
				</tr>
			<%	
				nTOTAL_PREMI_TONASE_AVG_BASIS = nTOTAL_PREMI_TONASE_AVG_BASIS + cdbl(to_nol(rsGrnInv("TOTAL_PREMI_TONASE_AVG_BASIS")))
				nHK_EFEKTIF_BRUTO_AVG_BASIS = nHK_EFEKTIF_BRUTO_AVG_BASIS + cdbl(to_nol(rsGrnInv("HK_EFEKTIF_BRUTO_AVG_BASIS")))	
				nPOTONGAN_HKE = nPOTONGAN_HKE + cdbl(to_nol(rsGrnInv("POTONGAN_HKE")))
				nHK_EFEKTIF_NETTO_AVG_BASIS = nHK_EFEKTIF_NETTO_AVG_BASIS + cdbl(to_nol(rsGrnInv("HK_EFEKTIF_NETTO_AVG_BASIS")))
				nPREMI_INSENTIF_AVG_BASIS = nPREMI_INSENTIF_AVG_BASIS + cdbl(to_nol(rsGrnInv("PREMI_INSENTIF_AVG_BASIS")))	 
				nTOTAL_PREMI_OPERATOR_AVG_BASIS = nTOTAL_PREMI_OPERATOR_AVG_BASIS + cdbl(to_nol(rsGrnInv("TOTAL_PREMI_OPERATOR_AVG_BASIS"))) 
				rsGrnInv.MoveNext			
				Wend
				rsGrnInv.Close
				set rsGrnInv = Nothing
			%> 
				<tr bgcolor="#F4F4A0">	 
					<td align="center" colspan="4"><b> TOTAL</b></td>
					<td align="right"><b><%=formatnumber(nTOTAL_PREMI_TONASE_AVG_BASIS,2)%></td>
					<td align="right"><b><%=formatnumber(nHK_EFEKTIF_BRUTO_AVG_BASIS,2)%></td>
					<td align="right"><b><%=formatnumber(nPOTONGAN_HKE,2)%></td>
					<td align="right"><b><%=formatnumber(nHK_EFEKTIF_NETTO_AVG_BASIS,2)%></td>
					<td align="right"><b><%=formatnumber(nPREMI_INSENTIF_AVG_BASIS,2)%></td>
					<td align="right"><b><%=formatnumber(nTOTAL_PREMI_OPERATOR_AVG_BASIS,2)%></td> 
					<!--<td align="right"></td>-->
				</tr> 
	<%End If%>
	</tbody>
</table>
<script src="/js/jquery.inputmask.bundle.min.js"></script>
<script src="/js/inputmask.binding.js"></script>
<script src="/plugins/chosen/chosen.jquery.min.js"></script>
<script type="text/javascript">	
	// $(".chosen-select").empty();
	
	var config = {
	  '.chosen-select'           : {},
	  '.chosen-select-deselect'  : {allow_single_deselect:true},
	  '.chosen-select-no-single' : {disable_search_threshold:10},
	  '.chosen-select-no-results': {no_results_text:'Oops, nothing found!'},
	  '.chosen-select-width'     : {width:"45%"}
	}
	for (var selector in config) {
	  $(selector).chosen(config[selector]);
	}
	// $(".chosen-select").trigger("chosen:updated");
	
</script>