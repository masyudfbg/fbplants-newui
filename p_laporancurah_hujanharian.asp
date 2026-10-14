<%@Language=VBScript%>
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

txttdate = txtMonth & "/01/" & txtYear

add_col_bud_prod_qty = ""
d_span = 33
	
v_txtFcba = Request("txtFcba")

strDefaultBa = " SELECT fccode AS fcba" & _
			   " FROM businessunit bu" & _
			   " INNER JOIN user_fcba_tab uft ON (bu.fccode = uft.fcba) " & _
			   " WHERE isactive = 'TRUE'" & _
			   "     AND fcusercode = '" & Session("CodeName") & "'"

Set rsDefaultBa = Server.CreateObject("ADODB.Recordset")
rsDefaultBa.Open strDefaultBa, connstr, adOpenStatic

If Not rsDefaultBa.EOF Then
	listba=0
	do while not rsDefaultBa.EOF
		listba = listba +1
		if listba = 1 then varDefaultBa = rsDefaultBa("FCBA")
		rsDefaultBa.movenext
	loop
Else
	varDefaultBa = ""
End If
rsDefaultBa.Close
set rsDefaultBa = Nothing

If "" & v_txtFcba = "" Then v_txtFcba = varDefaultBa

ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

arr_field_eof = true
strPusingan	= " SELECT fccode, company_nik" & _
              "        , fcname" & _
              "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 1, attendance, NULL)) day_1" & _
              "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 2, attendance, NULL)) day_2" & _
              "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 3, attendance, NULL)) day_3" & _
              "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 4, attendance, NULL)) day_4" & _
              "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 5, attendance, NULL)) day_5" & _
              "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 6, attendance, NULL)) day_6" & _
              "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 7, attendance, NULL)) day_7" & _
              "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 8, attendance, NULL)) day_8" & _
              "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 9, attendance, NULL)) day_9" & _
              "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 10, attendance, NULL)) day_10" & _
              "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 11, attendance, NULL)) day_11" & _
              "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 12, attendance, NULL)) day_12" & _
              "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 13, attendance, NULL)) day_13" & _
              "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 14, attendance, NULL)) day_14" & _
              "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 15, attendance, NULL)) day_15" & _
              "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 16, attendance, NULL)) day_16" & _
              "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 17, attendance, NULL)) day_17" & _
              "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 18, attendance, NULL)) day_18" & _
              "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 19, attendance, NULL)) day_19" & _
              "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 20, attendance, NULL)) day_20" & _
              "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 21, attendance, NULL)) day_21" & _
              "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 22, attendance, NULL)) day_22" & _
              "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 23, attendance, NULL)) day_23" & _
              "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 24, attendance, NULL)) day_24" & _
              "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 25, attendance, NULL)) day_25" & _
              "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 26, attendance, NULL)) day_26" & _
              "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 27, attendance, NULL)) day_27" & _
              "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 28, attendance, NULL)) day_28" & _
              "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 29, attendance, NULL)) day_29" & _
              "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 30, attendance, NULL)) day_30" & _
              "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 31, attendance, NULL)) day_31" & _
              " FROM (" & _
              " SELECT em.fcba, em.fccode, em.company_nik, em.fcname, em.tdate, CASE WHEN uf.timein IS NULL THEN '' ELSE 'KJ' END attendance " & _
              " FROM (" & _
              "     SELECT em.fcba, em.fccode, em.company_nik, em.fcname, dd.tdate" & _
              "     FROM iplasprod.employee em," & _
              "          ( SELECT TO_DATE ('"&txttdate&"', 'MM/DD/YYYY') +ROWNUM -1 tdate" & _
              "            FROM ALL_OBJECTS" & _
              "            WHERE ROWNUM <= LAST_DAY(TO_DATE ('"&txttdate&"', 'MM/DD/YYYY')) - TO_DATE ('"&txttdate&"', 'MM/DD/YYYY') + 1) dd" & _
              "     WHERE em.fcba = '"&v_txtFcba&"'" & _
              "           AND em.presentgrade = 'STF'" & _
              "           AND em.company_nik IS NOT NULL" & _
              "           AND em.dateterminate IS NULL" & _
              " ) em" & _
              " LEFT JOIN userfinger.attendance_tab uf ON (em.fcba = uf.fcbatobe AND em.fccode = uf.empnotobe and em.tdate = uf.tdate)" & _
              " WHERE em.fcba = '"&v_txtFcba&"'" & _
              " ) GROUP BY fccode, company_nik, fcname"

'response.write strPusingan
Set qORS = Server.CreateObject("ADODB.Recordset")
qORS.Open strPusingan, connstr, adOpenStatic

If Not qORS.EOF Then
	arr_field = qORS.getrows()
	arr_field_eof = false
End If
qORS.Close
set qORS = Nothing


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
<title>Absensi Staff - iPlaS Reporting</title>
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
						<td>
								<font face="verdana" size="1"><%=titleSite%></font>
							</td>
							<td>
								<select name="txtFcba" >
									<%
									Dim rsFcba, strFcba
									strFcba	=	" SELECT FCCODE, ORDERDATA FROM BUSINESSUNIT" & _
												" WHERE USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',fccode) IS NOT NULL" & _
												"       AND isactive = 'TRUE'" & _
												" ORDER BY ORDERDATA"
									
									Set rsFcba = Server.CreateObject("ADODB.Recordset")
									rsFcba.Open strFcba, connstr, adOpenStatic
									
									If rsFcba.Bof Or rsFcba.eof Then
									%>
									<option value=""></option>
									<%
									End If
									
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
							
						<td><font face="verdana" size="1">TAHUN</font></td>
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
							</td>
							<td><font face="verdana" size="1">BULAN</font></td>
							<td width="2">
								<select name="txtMonth">
					<%
						For i = 1 to 12
					%>
							<Option Value="<%=i%>" <%If cint(i) = cint(txtMonth) then%> Selected <%End If%>><%=Bulan(i)%></option>
					<%
						Next
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
			<div id="head_judul">
				<h4><%If Len(Trim(varCompanyName)) > 0 Then%>PT. <%End If%><%=varCompanyName%></h4>
				<h5><%=varFcbaName%></h5>
			</div>
			<div id="subheader" class="centered">
				<h2>ABSENSI STAFF</h2>
				<br/>
				<h4>PERIODE&nbsp;:&nbsp;<%=Bulan(txtMonth)%>&nbsp;<%=txtYear%></h4>
			</div>
          </div>  
		 </th>
	</tr>
	<tr>
		<th rowspan="2">NIK</th>
		<th rowspan="2">Nama</th>
		<th colspan="31">Tanggal</th>
	</tr>
	<tr>
		<%for xi=1 to 31%>
		<th style="min-width:20px;max-width:20px;"><%=xi%></th>
		<%next%>
	</tr>

    </thead>
	<tbody>
<%
if not arr_field_eof then
	
	for y=0 to ubound(arr_field,2)
			response.write "<tr>"
			response.write "	<td align='left'>"&arr_field(1,y)&"</td>"
			response.write "	<td align='left'>"&arr_field(2,y)&"</td>"
			for xi=3 to 33
				response.write "	<td align='center'>"&arr_field(xi,y)&"</td>"
			next
			response.write "</tr>"
	next
else
	response.write "<tr><td colspan='"&d_span&"'>NO DATA AVAILABLE</td></tr>"
end if
%>
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