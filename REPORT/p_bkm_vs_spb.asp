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
		
       if  v_txtfcba <> "" then
	       whereCons = whereCons & "  AND fp.fcba = '" & v_txtfcba & "'"
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
	
	strGrnInv		=	"   SELECT spb.fcba," & _
						"          spb.harvestdate," & _
						"          iplasprod.field.division," & _
						"          spb.fieldcode," & _
						"          NVL (spb.bunch_spb, 0) AS bunch_spb," & _
						"          NVL (bkm.bunch_bkm, 0) AS bunch_bkm," & _
						"          NVL ( (spb.bunch_spb - bkm.bunch_bkm), 0) AS var_bunch," & _
						"          NVL (spb.bucket_spb, 0) AS bucket_spb," & _
						"          NVL (bkm.bucket_bkm, 0) AS bucket_bkm," & _
						"          NVL ( (spb.bucket_spb - bkm.bucket_bkm), 0) AS var_bucket" & _
						"     FROM (  SELECT iplasprod.harvestingspb.fcba," & _
						"                    iplasprod.harvestingspb.harvestdate," & _
						"                    iplasprod.harvestingspb.fieldcode," & _
						"                    NVL (SUM (iplasprod.harvestingspb.bunch), 0) AS bunch_spb," & _
						"                    NVL (SUM (iplasprod.harvestingspb.bucket), 0) AS bucket_spb" & _
						"               FROM iplasprod.harvestingspb" & _
						"           GROUP BY iplasprod.harvestingspb.fcba," & _
						"                    iplasprod.harvestingspb.harvestdate," & _
						"                    iplasprod.harvestingspb.fieldcode) spb" & _
						"          FULL JOIN" & _
						"          (  SELECT iplasprod.attendance_gad.fcba," & _
						"                    iplasprod.attendance_gad.fddate," & _
						"                    iplasprod.attendance_gad.locationcode," & _
						"                    NVL (SUM (iplasprod.attendance_gad.unit), 0) AS bunch_bkm," & _
						"                    NVL (SUM (iplasprod.attendance_gad.bucket), 0) AS bucket_bkm" & _
						"               FROM iplasprod.attendance_gad" & _
						"              WHERE     iplasprod.attendance_gad.jobcode IN" & _
						"                           (select fccode from parameterdetail where parcode ='HARDCODE' and fcba = '"& v_txtFcba &"' and parheadcode = 'JOBCODE_PANEN')" & _
						"                    AND iplasprod.attendance_gad.attendance IN ('KJ'," & _
						"                                                                'WH'," & _
						"                                                                'WS'," & _
						"                                                                'KK')" & _
						"           GROUP BY iplasprod.attendance_gad.fcba," & _
						"                    iplasprod.attendance_gad.fddate," & _
						"                    iplasprod.attendance_gad.locationcode) bkm" & _
						"             ON (    spb.fcba = bkm.fcba" & _
						"                 AND spb.harvestdate = bkm.fddate" & _
						"                 AND spb.fieldcode = bkm.locationcode)" & _
						"          INNER JOIN iplasprod.field" & _
						"             ON (    bkm.fcba = iplasprod.field.fcba" & _
						"                 AND bkm.locationcode = iplasprod.field.fccode)" & _
						"    WHERE     spb.fcba = '"& v_txtFcba &"'"  
   	
   	if  len(trim(v_txtdatefrom)) > 0 then
			strGrnInv = strGrnInv & " AND spb.harvestdate >= to_date('" & v_txtdatefrom & "','mm/dd/yyyy')"
			txtSelect = " From " & FormatDateTime(v_txtdatefrom,1)
	end if
	if  len(trim(v_txtdateuntil)) > 0 then
		strGrnInv = strGrnInv & " AND spb.harvestdate <= to_date('" & v_txtdateuntil & "','mm/dd/yyyy')"
		txtSelect = txtSelect & " To " & FormatDateTime(v_txtdateuntil,1)
	end if 
	if  len(trim(v_txtempcode)) > 0 then
		strGrnInv = strGrnInv & " AND employeecode ='"& v_txtempcode &"'"
	end if 
	if  len(trim(v_txtsectionname)) > 0 then
		strGrnInv = strGrnInv & " AND iplasprod.field.division ='"& v_txtsectionname &"'"
	end if 


	' if  cdbl(to_nol(compare1)) = 1 then
	'     strGrnInv	=	strGrnInv & " and attendance_gad.gad_employeecode != vehicleactivity.vh_employeecode"
	' end if
	' if  cdbl(to_nol(compare1)) = 2 then
	'     strGrnInv	=	strGrnInv & " and attendance_gad.gad_employeecode != harvestingspb.harvesting_employeecode" 
	' end if
	' if  cdbl(to_nol(compare1)) = 3 then
	'     strGrnInv	=	strGrnInv & " and vehicleactivity.vh_employeecode != harvestingspb.harvesting_employeecode" 
	' end if

	' strGrnInv	=	strGrnInv &	" GROUP BY" & _
	' 							"     FCBA" & _
	' 							"     , NOPO" & _
	' 							"     , STOCKCODE"
								
	strGrnInv	=	strGrnInv & " ORDER BY spb.fcba," & _
								"          spb.harvestdate," & _
								"          iplasprod.field.division," & _
								"          spb.fieldcode" 

	' response.write strGrnInv	
	Set rsGrnInv = Server.CreateObject("ADODB.Recordset")
	rsGrnInv.Open strGrnInv, connstr, adOpenStatic 
	
%>


<html>

<head>
<title>ePlaS Reporting - BKM VS SPB HARVESTING</title>
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
						<td align="center">
							<font face="verdana" size="1" style="font-weight: bold;">KEBUN</font>
						</td>
						<td align="center">
							<font face="verdana" size="1" style="font-weight: bold;">DARI</font>
						</td>
						<td align="center">
							<font face="verdana" size="1" style="font-weight: bold;">SAMPAI</font>
						</td>
						<td align="center">
							<font face="verdana" size="1" style="font-weight: bold;">DIVISION</font>
						</td>  
					</tr>
					<tr>	 
						<td>
							<select name="txtFcba" id="vFcba" onchange="javascript:filtering()">
							<option></option>
							<%
							Set rsFcba = Server.CreateObject("ADODB.Recordset")
							rsFcba.CursorLocation = adUseClient
							strsql =   " select fccode fcba from businessunit " & _
										" where " & _
										" USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',fccode) IS NOT NULL " 
							  
							
							rsFcba.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
							if not rsFcba.eof then
								DO while not rsFcba.eof
									if v_txtFcba  = rsFcba("fcba") then
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
						<td width="4">
							<input class="datepicker" type="text" maxlength="10" name="txtdatefrom" size="10" value="<%=v_txtdatefrom%>">
						</td> 
						<td width="2">
							<input class="datepicker" type="text" maxlength="10" name="txtdateuntil" size="10" value="<%=v_txtdateuntil%>">
						</td> 
						<td>
							<select name="txtsectionname" >
								<option value="">-- ALL --</option>
									<%
									Dim rssectionname, strsectionname
									strsectionname	=	"select fccode, fcname from sectionname where fccode in ('AFD-01','AFD-02','AFD-03','AFD-04','AFD-05','AFD-06','AFD-07','AFD-08','AFD-09','AFD-10','AFD-11','AFD-12','AFD-13') order by fccode"
									
									Set rssectionname = Server.CreateObject("ADODB.Recordset")
									rssectionname.Open strsectionname, connstr, adOpenStatic 
									
									if not rssectionname.eof then
									Do While Not rssectionname.eof
									if rssectionname("fccode") = v_txtsectionname then%>
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
    <th colspan="9" style="background-color:#FFFFFF; border:none;">  
	<div id="header-wrapper">
		<div id="head_judul" style="float:left; text-align:left;">
				<h4>PT. <%=CompanyName%></h4>
				<h5><%=fcbaname%></h5>
			</div>
			
			<div id="subheader" class="centered" tyle="width:50%; margin:auto;">
				<h2>BKM VS SPB HARVESTING</h2> 
				<h4>Seleted By : <%=txtselect%></h4>
                <br/>
			</div>
		</div>
</th>
		</tr>
		<tr>
			<th style="background-color:#00A300;color:#fff;" rowspan="2">Date</th>
			<th style="background-color:#00A300;color:#fff;" rowspan="2">Division</th>
			<th style="background-color:#00A300;color:#fff;" rowspan="2">Field Code</th>
			<th style="background-color:#00A300;color:#fff;" colspan="3">Bunch</th>
			<th style="background-color:#00A300;color:#fff;" colspan="3">Brondolan</th>
		</tr>
		<tr>
			<th style="background-color:#00A300;color:#fff;">Bunch SPB</th>
			<th style="background-color:#00A300;color:#fff;">Bunch BKM</th>
			<th style="background-color:#00A300;color:#fff;">Var Bunch</th>

			<th style="background-color:#00A300;color:#fff;">Brondolan SPB</th>
			<th style="background-color:#00A300;color:#fff;">Brondolan BKM</th>
			<th style="background-color:#00A300;color:#fff;">Var Brondolan</th>
		</tr>
    </thead>
	<tbody>
<%if rsGrnInv.BOF and rsGrnInv.EOF then%>
		<tr>
			<td colspan="9" align="left">NO DATA AVAILABLE</td>
		</tr>
<%end if
		NO = 1
		v_bunch_spb = 0
		v_bunch_bkm = 0
		v_var_bunch = 0 
		v_bucket_spb = 0
		v_bucket_bkm = 0
		v_var_bucket = 0
		while not rsGrnInv.eof
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
%>
		<tr class="<%=cClass%>">
			<td align="center" rowspan="2"><%=rsGrnInv("harvestdate")%></td>
			<td align="center" rowspan="2"><%=rsGrnInv("division")%></td>
			<td align="center" rowspan="2"><%=rsGrnInv("fieldcode")%></td> 
		</tr>
		<tr class="<%=cClass%>">
			<td align="center" ><%=rsGrnInv("bunch_spb")%></td>
			<td align="center" ><%=rsGrnInv("bunch_bkm")%></td>
			<td align="center" ><%=rsGrnInv("var_bunch")%></td>

			<td align="center" ><%=rsGrnInv("bucket_spb")%></td>
			<td align="center" ><%=rsGrnInv("bucket_bkm")%></td>
			<td align="center" ><%=rsGrnInv("var_bucket")%></td>
		</tr>

<%			
			v_bunch_spb = v_bunch_spb + cdbl(to_nol(rsGrnInv("bunch_spb"))) 
			v_bunch_bkm = v_bunch_bkm + cdbl(to_nol(rsGrnInv("bunch_bkm"))) 
			v_var_bunch = v_var_bunch + cdbl(to_nol(rsGrnInv("var_bunch"))) 
			v_bucket_spb = v_bucket_spb + cdbl(to_nol(rsGrnInv("bucket_spb"))) 
			v_bucket_bkm = v_bucket_bkm + cdbl(to_nol(rsGrnInv("bucket_bkm"))) 
			v_var_bucket = v_var_bucket + cdbl(to_nol(rsGrnInv("var_bucket")))  

			rsGrnInv.MoveNext
		Wend
	rsGrnInv.Close
	set rsGrnInv = Nothing
%>
		<tr bgcolor="#F4F4A0"> 
			<td align="center" colspan="3"><font style="font-weight: bold;">TOTAL</font></td>				
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_bunch_spb)),0)%></font></td>
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_bunch_bkm)),0)%></font></td>
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_var_bunch)),0)%></font></td>
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_bucket_spb)),0)%></font></td> 
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_bucket_bkm)),0)%></font></td> 
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_var_bucket)),0)%></font></td> 
		</tr>
	</tbody>
</table>

<!--#include file="../Bottom_Report.asp"-->
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