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
	If "" & txtYear = "" Then txtYear = Year(Session("v_Costbook"))
	If "" & txtMonth = "" Then txtMonth = Month(Session("v_Costbook"))
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

	q = "SELECT type_2 FROM allowded WHERE fccode = 'A32B' AND fcba = '"& v_txtfcba &"'"

	' response.Write q
	Set rsCallCombo = Server.CreateObject("ADODB.Recordset")
	rsCallCombo.Open q, connstr
	 
	If not rsCallCombo.eof then
	   	type_report = rsCallCombo("type_2")
	End if
	
	If 	type_report = "MIN" then 
		strGrnInv	=	" SELECT fcba," & _
						"        transdate," & _
						"        attendance," & _
						"        operatorcode," & _
						"        operatorname," & _
						"        kg_tbs," & _
						"        kg_brondolan," & _
						"        kg_pupuk," & _
						"        kg_lain_lain," & _
						"        total_kg," & _
						"        basis," & _
						"        kg_premi_lebih_basis," & _
						"        total_premi," & _
						"        hk_efektif" & _
						"   FROM PREMI_OPR_DT_BYMINIMALBASIS" & _
						"  WHERE     fcba = '"& v_txtfcba &"'" 

		if  len(trim(v_txtdatefrom)) > 0 then
			strGrnInv = strGrnInv & " AND transdate >= TO_DATE('" & v_txtdatefrom & "', 'mm/dd/yyyy')"
			txtSelect = " From " & FormatDateTime(v_txtdatefrom,1)
		end if
		if  len(trim(v_txtdateuntil)) > 0 then
			strGrnInv = strGrnInv & " AND transdate <= TO_DATE('" & v_txtdateuntil & "', 'mm/dd/yyyy')"
			txtSelect = txtSelect & " To " & FormatDateTime(v_txtdateuntil,1)
		end if  
		if  len(trim(v_txtempcode)) > 0 then
		    strGrnInv	=	strGrnInv & " AND operatorcode = '"& v_txtempcode &"'" 
		    txtSelect = txtSelect & " Employee Code " & v_txtempcode
		end if 
									
		strGrnInv	=	strGrnInv & " ORDER BY fcba," & _
									"          operatorcode," & _ 
									"          transdate" 

		' response.write strGrnInv	
		Set rsGrnInv = Server.CreateObject("ADODB.Recordset")
		rsGrnInv.Open strGrnInv, connstr, adOpenStatic 
	Else
		strGrnInv	=	" SELECT fcba," & _
						"        transdate," & _
						"        attendance," & _
						"        operatorcode," & _
						"        operatorname," & _
						"        kg_tbs," & _
						"        kg_brondolan," & _
						"        kg_pupuk," & _
						"        kg_lain_lain," & _
						"        total_kg," & _
						"        basis," & _
						"        kg_premi_lebih_basis," & _
						"        total_premi," & _
						"        hk_efektif" & _
						"   FROM PREMI_OPR_DT_BYAVERAGEBASIS" & _
						"  WHERE     fcba = '"& v_txtfcba &"'" 
						
		if  len(trim(v_txtdatefrom)) > 0 then
			strGrnInv = strGrnInv & " AND transdate >= TO_DATE('" & v_txtdatefrom & "', 'mm/dd/yyyy')"
			txtSelect = " From " & FormatDateTime(v_txtdatefrom,1)
		end if
		if  len(trim(v_txtdateuntil)) > 0 then
			strGrnInv = strGrnInv & " AND transdate <= TO_DATE('" & v_txtdateuntil & "', 'mm/dd/yyyy')"
			txtSelect = txtSelect & " To " & FormatDateTime(v_txtdateuntil,1)
		end if 
		if  len(trim(v_txtempcode)) > 0 then
		    strGrnInv	=	strGrnInv & " AND operatorcode = '"& v_txtempcode &"'" 
		    txtSelect = txtSelect & " Employee Code " & v_txtempcode
		end if 
									
		strGrnInv	=	strGrnInv & " ORDER BY fcba," & _
									"          operatorcode," & _ 
									"          transdate"

		' response.write strGrnInv	
		Set rsGrnInv = Server.CreateObject("ADODB.Recordset")
		rsGrnInv.Open strGrnInv, connstr, adOpenStatic 
	End If
%>


<html>

<head>
<title>ePlaS Reporting - Premi Operator DT</title>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<link rel="stylesheet" href="/css/jquery-ui.css" />
<script src="/js/jquery-1.9.1.js"></script>
<script src="/js/jquery-ui.js"></script>
<script language="JavaScript" src='/js/x.js'></script>
<script language="JavaScript" src='/js/xtableheaderfixed.js'></script>
<link href="/plugins/tabcontent/tabcontent.css" rel="stylesheet" type="text/css" />
<link href="/plugins/chosen/chosen.min.css" rel="stylesheet" />

<link rel="stylesheet" href="/plugins/responsivegridsystem/col.css" media="all">
<link rel="stylesheet" href="/plugins/responsivegridsystem/12cols.css" media="all">
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
							<font face="verdana" size="1">DARI</font>
						</td>
						<td width="4">
							<input class="datepicker" type="text" maxlength="10" name="txtdatefrom" size="10" value="<%=v_txtdatefrom%>">
						</td>
						<td>
							<font face="verdana" size="1">SAMPAI</font>
						</td>
						<td width="2">
							<input class="datepicker" type="text" maxlength="10" name="txtdateuntil" size="10" value="<%=v_txtdateuntil%>">
						</td> 
						<td>
							<font face="verdana" size="1">EMPLOYEE</font>
						</td>
						<td>
							<select name="txtempcode" class="chosen-select" >
								<option value="">--ALL--</option>
									<% 
									strsectionname	=	" select fccode, fcname from employee where fcba = '"& v_txtfcba &"' order by fcCode"
									
									Set rssectionname = Server.CreateObject("ADODB.Recordset")
									rssectionname.Open strsectionname, connstr, adOpenStatic 
									
									if not rssectionname.eof then
									Do While Not rssectionname.eof
									if rssectionname("fccode") = v_txtempcode then%>
											<option value="<%=rssectionname("fccode")%>" Selected><%=rssectionname("fccode")%> : <%=rssectionname("fcname")%></option> 
										 <%Else%>
											<option value="<%=rssectionname("fccode")%>"><%=rssectionname("fccode")%> : <%=rssectionname("fcname")%> </option>  
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
    		<th colspan="13" style="background-color:#FFFFFF; border:none;">  
				<div id="header-wrapper">
					<div id="head_judul" style="float:left; text-align:left;">
						<h4>PT. <%=CompanyName%></h4>
						<h5><%=fcbaname%></h5>
					</div>
					<div id="subheader" class="centered" tyle="width:50%; margin:auto;">
						<h2>LAPORAN PREMI OPERATOR DT</h2>
		                <br/>
						<h4>Selected By :<%=txtSelect%></h4>
					</div> 
				</div>
			</th>
		</tr>
		<tr>
			<th style="background-color:#00A300;color:#fff;">OPERATOR CODE</th> 
			<th style="background-color:#00A300;color:#fff;">OPERATOR NAME</th>
			<th style="background-color:#00A300;color:#fff;">DATE</th>
			<th style="background-color:#00A300;color:#fff;">ATTENDANCE</th>
			<th style="background-color:#00A300;color:#fff;">KG TBS</th>
			<th style="background-color:#00A300;color:#fff;">KG BRONDOLAN</th>
			<th style="background-color:#00A300;color:#fff;">KG PUPUK</th>
			<th style="background-color:#00A300;color:#fff;">KG LAIN - LAIN</th>
			<th style="background-color:#00A300;color:#fff;">TOTAL KG</th>		
			<th style="background-color:#00A300;color:#fff;">BASIS</th>		
			<th style="background-color:#00A300;color:#fff;">KG PREMI LEBIH BASIS</th>		
			<th style="background-color:#00A300;color:#fff;">TOTAL PREMI</th>		
			<th style="background-color:#00A300;color:#fff;">HK EFEKTIF</th>						
		</tr>   
		    </thead>
	<tbody>
<%if rsGrnInv.BOF and rsGrnInv.EOF then%>
		<tr>
			<td colspan="13" align="left">NO DATA AVAILABLE</td>
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
			<% if rows = 1 then
					sqlrow = "  select count(operatorcode) rowsd  from ( " & strGrnInv & " ) where operatorcode = '" & rsGrnInv("operatorcode") & "'"
					' response.write sqlrow
					Set rsrow = Server.CreateObject("ADODB.Recordset")
					rsrow.Open sqlrow, connstr, adOpenStatic
					if not rsrow.EOF then
						temprows = CDbl(rsrow("rowsd"))
					else
						temprows = 0
					end if
				end if 
			%>
			<%if rows = 1 then%>
				<td align="center" Rowspan="<%=0+temprows%>"><%=rsGrnInv("operatorcode")%></td>
				<td align="center" Rowspan="<%=0+temprows%>"><%=rsGrnInv("operatorname")%></td>
			<%end if%>
			<td align="center"><%=rsGrnInv("transdate")%></td>
			<td align="center"><%=rsGrnInv("attendance")%></td>
			<td align="right"><%=FormatNumber(cdbl(to_nol(rsGrnInv("kg_tbs"))),0)%></td>
			<td align="right"><%=FormatNumber(cdbl(to_nol(rsGrnInv("kg_brondolan"))),0)%></td>
			<td align="right"><%=FormatNumber(cdbl(to_nol(rsGrnInv("kg_pupuk"))),0)%></td> 
			<td align="right"><%=FormatNumber(cdbl(to_nol(rsGrnInv("kg_lain_lain"))),0)%></td>
			<td align="right"><%=FormatNumber(cdbl(to_nol(rsGrnInv("total_kg"))),0)%></td> 
			<td align="right"><%=FormatNumber(cdbl(to_nol(rsGrnInv("basis"))),0)%></td> 
			<td align="right"><%=FormatNumber(cdbl(to_nol(rsGrnInv("kg_premi_lebih_basis"))),0)%></td>
			<td align="right"><%=FormatNumber(cdbl(to_nol(rsGrnInv("total_premi"))),0)%></td> 
			<td align="right"><%=FormatNumber(cdbl(to_nol(rsGrnInv("hk_efektif"))),0)%></td>
		</tr>
<%	if rows = temprows then
			nkg_tbs = nkg_tbs + cdbl(to_nol(rsGrnInv("kg_tbs")))
			nkg_brondolan = nkg_brondolan + cdbl(to_nol(rsGrnInv("kg_brondolan")))	
			nkg_pupuk = nkg_pupuk + cdbl(to_nol(rsGrnInv("kg_pupuk")))
			nkg_lain_lain = nkg_lain_lain + cdbl(to_nol(rsGrnInv("kg_lain_lain")))
			ntotal_kg = ntotal_kg + cdbl(to_nol(rsGrnInv("total_kg")))	 
			nbasis = nbasis + cdbl(to_nol(rsGrnInv("basis")))
			nkg_premi_lebih_basis = nkg_premi_lebih_basis + cdbl(to_nol(rsGrnInv("kg_premi_lebih_basis")))
			ntotal_premi = ntotal_premi + cdbl(to_nol(rsGrnInv("total_premi")))
			nhk_efektif = nhk_efektif + cdbl(to_nol(rsGrnInv("hk_efektif"))) 
			rows = 1
			
	%>
		<tr bgcolor="#F4F4A0"> 
			<td align="center" colspan="4"><font style="font-weight: bold;">TOTAL</font></td>				
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(nkg_tbs)),0)%></font></td> 
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(nkg_brondolan)),0)%></font></td>  
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(nkg_pupuk)),0)%></font></td> 
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(nkg_lain_lain)),0)%></font></td> 
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(ntotal_kg)),0)%></font></td> 
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(nbasis)),0)%></font></td> 
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(nkg_premi_lebih_basis)),0)%></font></td> 
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(ntotal_premi)),0)%></font></td>  
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(nhk_efektif)),0)%></font></td>   
		</tr>
		<tr>
			<td align="center" colspan="<%= 13+ nSpanTotal %>" ></td>
		</tr>
		<%
			nkg_tbs = cdbl("0")
			nkg_brondolan = cdbl("0")
			nkg_pupuk = cdbl("0")
			nkg_lain_lain = cdbl("0")
			ntotal_kg = cdbl("0") 
			nbasis = cdbl("0")
			nkg_premi_lebih_basis = cdbl("0") 
			ntotal_premi = cdbl("0") 
			nhk_efektif = cdbl("0")  
		else
		' cari sub total
			nkg_tbs = nkg_tbs + cdbl(to_nol(rsGrnInv("kg_tbs")))
			nkg_brondolan = nkg_brondolan + cdbl(to_nol(rsGrnInv("kg_brondolan")))	
			nkg_pupuk = nkg_pupuk + cdbl(to_nol(rsGrnInv("kg_pupuk")))
			nkg_lain_lain = nkg_lain_lain + cdbl(to_nol(rsGrnInv("kg_lain_lain")))
			ntotal_kg = ntotal_kg + cdbl(to_nol(rsGrnInv("total_kg")))
			nbasis = nbasis + cdbl(to_nol(rsGrnInv("basis")))
			nkg_premi_lebih_basis = nkg_premi_lebih_basis + cdbl(to_nol(rsGrnInv("kg_premi_lebih_basis")))
			ntotal_premi = ntotal_premi + cdbl(to_nol(rsGrnInv("total_premi")))
			nhk_efektif = nhk_efektif + cdbl(to_nol(rsGrnInv("hk_efektif")))
			rows = rows + 1
		End if
							  
			' cari grand total 
				gt_kg_tbs = cdbl(gt_kg_tbs) + cdbl(rsGrnInv("kg_tbs"))
				gt_kg_brondolan = cdbl(gt_kg_brondolan) + cdbl(rsGrnInv("kg_brondolan"))
				gt_kg_pupuk = cdbl(gt_kg_pupuk) + cdbl(rsGrnInv("kg_pupuk"))
				gt_kg_lain_lain = cdbl(gt_kg_lain_lain) + cdbl(rsGrnInv("kg_lain_lain"))
				gt_total_kg = cdbl(gt_total_kg) + cdbl(rsGrnInv("total_kg")) 
				gt_basis = cdbl(gt_basis) + cdbl(rsGrnInv("basis"))
				gt_kg_premi_lebih_basis = cdbl(gt_kg_premi_lebih_basis) + cdbl(rsGrnInv("kg_premi_lebih_basis"))
				gt_total_premi = cdbl(gt_total_premi) + cdbl(rsGrnInv("total_premi")) 
				gt_hk_efektif = cdbl(gt_hk_efektif) + cdbl(rsGrnInv("hk_efektif"))  

			rsGrnInv.MoveNext			
			Wend
		rsGrnInv.Close
		set rsGrnInv = Nothing
	%> 
		<tr bgcolor="#FFFF00">	 
			<td align="center" colspan="4"><b>GRAND TOTAL</b></td>
			<td align="right"><b><%=formatnumber(gt_kg_tbs,0)%></td>
			<td align="right"><b><%=formatnumber(gt_kg_brondolan,0)%></td>
			<td align="right"><b><%=formatnumber(gt_kg_pupuk,0)%></td>
			<td align="right"><b><%=formatnumber(gt_kg_lain_lain,0)%></td>
			<td align="right"><b><%=formatnumber(gt_total_kg,0)%></td>
			<td align="right"><b><%=formatnumber(gt_basis,0)%></td>
			<td align="right"><b><%=formatnumber(gt_kg_premi_lebih_basis,0)%></td>
			<td align="right"><b><%=formatnumber(gt_total_premi,0)%></td>
			<td align="right"><b><%=formatnumber(gt_hk_efektif,0)%></td> 
			<!--<td align="right"></td>-->
		</tr>
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