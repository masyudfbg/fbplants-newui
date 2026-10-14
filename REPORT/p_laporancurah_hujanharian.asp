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
d_span = 64
	
v_txtFcba = Request("txtFcba")
v_txtDivisi = Request("txtDivisi")

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

if  v_txtDivisi <> "--ALL--"  and v_txtDivisi <> "" then
	whereCons = whereCons & "AND INSTR('"& v_txtDivisi &"', SUBSTR(widivision,1,6) ) <> 0 "
end if

strDefaultBa = " SELECT fccode AS fcba" & _
			   " FROM businessunit bu" & _
			   " INNER JOIN user_fcba_tab uft ON (bu.fccode = uft.fcba) " & _
			   " WHERE isactive = 'TRUE'" & _
			   "     AND fcusercode = '" & Session("CodeName") & "'"
' response.write strDefaultBa

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
strPusingan	= 	 "  SELECT fcba,widivision, " & _
				 "	SUM(DECODE(HARI,'01',TOTAL1,0))CH01," & _
				 "	SUM(DECODE(HARI,'02',TOTAL1,0))CH02," & _
				 "	SUM(DECODE(HARI,'03',TOTAL1,0))CH03," & _
				 "	SUM(DECODE(HARI,'04',TOTAL1,0))CH04," & _
				 "	SUM(DECODE(HARI,'05',TOTAL1,0))CH05," & _
				 "	SUM(DECODE(HARI,'06',TOTAL1,0))CH06," & _
				 "	SUM(DECODE(HARI,'07',TOTAL1,0))CH07," & _
				 "	SUM(DECODE(HARI,'08',TOTAL1,0))CH08," & _
				 "	SUM(DECODE(HARI,'09',TOTAL1,0))CH09," & _
				 "	SUM(DECODE(HARI,'10',TOTAL1,0))CH10," & _
				 "	SUM(DECODE(HARI,'11',TOTAL1,0))CH11," & _
				 "	SUM(DECODE(HARI,'12',TOTAL1,0))CH12," & _
				 "	SUM(DECODE(HARI,'13',TOTAL1,0))CH13," & _
				 "	SUM(DECODE(HARI,'14',TOTAL1,0))CH14," & _
				 "	SUM(DECODE(HARI,'15',TOTAL1,0))CH15," & _
				 "	SUM(DECODE(HARI,'16',TOTAL1,0))CH16," & _
				 "	SUM(DECODE(HARI,'17',TOTAL1,0))CH17," & _
				 "	SUM(DECODE(HARI,'18',TOTAL1,0))CH18," & _
				 "	SUM(DECODE(HARI,'19',TOTAL1,0))CH19," & _
				 "	SUM(DECODE(HARI,'20',TOTAL1,0))CH20," & _
				 "	SUM(DECODE(HARI,'21',TOTAL1,0))CH21," & _
				 "	SUM(DECODE(HARI,'22',TOTAL1,0))CH22," & _
				 "	SUM(DECODE(HARI,'23',TOTAL1,0))CH23," & _
				 "	SUM(DECODE(HARI,'24',TOTAL1,0))CH24," & _
				 "	SUM(DECODE(HARI,'25',TOTAL1,0))CH25," & _
				 "	SUM(DECODE(HARI,'26',TOTAL1,0))CH26," & _
				 "	SUM(DECODE(HARI,'27',TOTAL1,0))CH27," & _
				 "	SUM(DECODE(HARI,'28',TOTAL1,0))CH28," & _
				 "	SUM(DECODE(HARI,'29',TOTAL1,0))CH29," & _
				 "	SUM(DECODE(HARI,'30',TOTAL1,0))CH30," & _
				 "	SUM(DECODE(HARI,'31',TOTAL1,0))CH31," & _
				 "	SUM(DECODE(HARI,'01',TOTAL2,0))HH01," & _
				 "	SUM(DECODE(HARI,'02',TOTAL2,0))HH02," & _
				 " 	SUM(DECODE(HARI,'03',TOTAL2,0))HH03," & _
				 "	SUM(DECODE(HARI,'04',TOTAL2,0))HH04," & _
				 "	SUM(DECODE(HARI,'05',TOTAL2,0))HH05," & _
				 "	SUM(DECODE(HARI,'06',TOTAL2,0))HH06," & _
				 "	SUM(DECODE(HARI,'07',TOTAL2,0))HH07," & _
				 "	SUM(DECODE(HARI,'08',TOTAL2,0))HH08," & _
				 "	SUM(DECODE(HARI,'09',TOTAL2,0))HH09," & _
				 "	SUM(DECODE(HARI,'10',TOTAL2,0))HH10," & _
				 "	SUM(DECODE(HARI,'11',TOTAL2,0))HH11," & _
				 "	SUM(DECODE(HARI,'12',TOTAL2,0))HH12," & _
				 "	SUM(DECODE(HARI,'13',TOTAL2,0))HH13," & _
				 "	SUM(DECODE(HARI,'14',TOTAL2,0))HH14," & _
				 "	SUM(DECODE(HARI,'15',TOTAL2,0))HH15," & _
				 "	SUM(DECODE(HARI,'16',TOTAL2,0))HH16," & _
				 "	SUM(DECODE(HARI,'17',TOTAL2,0))HH17," & _
				 "	SUM(DECODE(HARI,'18',TOTAL2,0))HH18," & _
				 "	SUM(DECODE(HARI,'19',TOTAL2,0))HH19," & _
				 "	SUM(DECODE(HARI,'20',TOTAL2,0))HH20," & _
				 "	SUM(DECODE(HARI,'21',TOTAL2,0))HH21," & _
				 "	SUM(DECODE(HARI,'22',TOTAL2,0))HH22," & _
				 "	SUM(DECODE(HARI,'23',TOTAL2,0))HH23," & _
				 "	SUM(DECODE(HARI,'24',TOTAL2,0))HH24," & _
				 "	SUM(DECODE(HARI,'25',TOTAL2,0))HH25," & _
				 "	SUM(DECODE(HARI,'26',TOTAL2,0))HH26," & _
				 "	SUM(DECODE(HARI,'27',TOTAL2,0))HH27," & _
				 "	SUM(DECODE(HARI,'28',TOTAL2,0))HH28," & _
				 "	SUM(DECODE(HARI,'29',TOTAL2,0))HH29," & _
				 "	SUM(DECODE(HARI,'30',TOTAL2,0))HH30," & _
				 "	SUM(DECODE(HARI,'31',TOTAL2,0))HH31" & _
				 "	FROM (" & _
				 "	select TO_CHAR(widate,'YYYY')AS TAHUN,TO_CHAR(widate,'MM')AS BULAN,TO_CHAR(widate,'DD')AS HARI, fcba,widivision,sum(TOTRNINCHES) as total1,sum(TOTRNDAYS) as TOTAL2" & _
				 "	from iplasprod.weather" & _
				 "	where to_char(widate,'MMYYYY')='"& txttdate &"'  "& whereCons &""& _
				 "	group by TO_CHAR(widate,'YYYY'),TO_CHAR(widate,'MM'),TO_CHAR(widate,'DD'),fcba,widivision" & _
				 "	)" & _
				 "	group by fcba,widivision" & _
				 "	order by fcba,widivision"

' response.write strPusingan
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
<title>ePlaS Reporting - Curah Hujan Harian</title>
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
							<th>KEBUN</th>
							<th>TAHUN</th>
							<th>BULAN</th>
							<th>DIVISI</th>
						</tr>
						<tr>
							<td>
								<select name="txtFcba" id="vFcba" onchange="javascript:filtering()">
								<%
								Dim rsFcba
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
							</td> 
							<td width="2">
								<select name="txtMonth">
					<% 
						For i = 1 to 12
							
							if i <= 9 then 
								iBulan = 0&i
							else
								iBulan = i
							end if
							i	= iBulan
					%>
							<Option Value="<%=i%>" <%If cint(i) = cint(txtMonth) then%> Selected <%End If%>><%=Bulan(i)%></option>
					<%
						Next
					%>
								</select>
							</td> 
							<td>
								<select name="txtDivisi" class="chosen-select" multiple>
									<option>--ALL--</option>
									<% 
									Set rsFcba = Server.CreateObject("ADODB.Recordset")
									rsFcba.CursorLocation = adUseClient
									strsql =    " SELECT DISTINCT WIDIVISION FROM WEATHER WHERE FCBA = '"& v_txtFcba &"' ORDER BY WIDIVISION" 
									  
									
									rsFcba.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
									if not rsFcba.eof then
										DO while not rsFcba.eof
											if cstr(""&v_txtDivisi) = (""&rsFcba("WIDIVISION")) then
											%>			
												<option value="<%=rsFcba("WIDIVISION")%>" Selected><%=rsFcba("WIDIVISION")%></option>
											<% Else %>	
												<option value="<%=rsFcba("WIDIVISION")%>"><%=rsFcba("WIDIVISION")%></option>
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
				<h2>LAPORAN CURAH HUJAN HARIAN</h2>
				<br/>
				<h4>PERIODE&nbsp;:&nbsp;<%=Bulan(txtMonth)%>&nbsp;<%=txtYear%></h4>
			</div>
          </div>  
		 </th>
	</tr>
	 <tr>
					<th rowspan="2">Kebun</th>
					<th rowspan="2">WidDivision</th>
					<th colspan="32">Curah Hujan (mm)</th>
					<th colspan="32">Hari Hujan (Hari)</th>
				</tr>

				<tr>
					<th>  01</th>
					<th>  02</th>
					<th>  03</th>
					<th>  04</th>
					<th>  05</th>
					<th>  06</th>
					<th>  07</th>
					<th>  08</th>
					<th>  09</th>
					<th>  10</th>
					<th>  11</th>
					<th>  12</th>
					<th>  13</th>
					<th>  14</th>
					<th>  15</th>
					<th>  16</th>
					<th>  17</th>
					<th>  18</th>
					<th>  19</th>
					<th>  20</th>
					<th>  21</th>
					<th>  22</th>
					<th>  23</th>
					<th>  24</th>
					<th>  25</th>
					<th>  26</th>
					<th>  27</th>
					<th>  28</th>
					<th>  29</th>
					<th>  30</th>
					<th>  31</th>
					<th>TOTAL</th>
					
					<th>  01</th>
					<th>  02</th>
					<th>  03</th>
					<th>  04</th>
					<th>  05</th>
					<th>  06</th>
					<th>  07</th>
					<th>  08</th>
					<th>  09</th>
					<th>  10</th>
					<th>  11</th>
					<th>  12</th>
					<th>  13</th>
					<th>  14</th>
					<th>  15</th>
					<th>  16</th>
					<th>  17</th>
					<th>  18</th>
					<th>  19</th>
					<th>  20</th>
					<th>  21</th>
					<th>  22</th>
					<th>  23</th>
					<th>  24</th>
					<th>  25</th>
					<th>  26</th>
					<th>  27</th>
					<th>  28</th>
					<th>  29</th>
					<th>  30</th>
					<th>  31</th>
					<th>TOTAL</th>
				</tr>

    </thead>
	<tbody>
		<%
		Set rsCat = connStr.Execute(strPusingan)

		v_CH01 = 0
		v_CH02 = 0
		v_CH03 = 0
		v_CH04 = 0
		v_CH05 = 0
		v_CH06 = 0
		v_CH07 = 0
		v_CH08 = 0
		v_CH09 = 0
		v_CH10 = 0
		v_CH11 = 0
		v_CH12 = 0
		v_CH13 = 0
		v_CH14 = 0
		v_CH15 = 0
		v_CH16 = 0
		v_CH17 = 0
		v_CH18 = 0
		v_CH19 = 0
		v_CH20 = 0
		v_CH21 = 0
		v_CH22 = 0
		v_CH23 = 0
		v_CH24 = 0
		v_CH25 = 0
		v_CH26 = 0
		v_CH27 = 0
		v_CH28 = 0
		v_CH29 = 0
		v_CH30 = 0
		v_CH31 = 0

		v_HH01 = 0
		v_HH02 = 0
		v_HH03 = 0
		v_HH04 = 0
		v_HH05 = 0
		v_HH06 = 0
		v_HH07 = 0
		v_HH08 = 0
		v_HH09 = 0
		v_HH10 = 0
		v_HH11 = 0
		v_HH12 = 0
		v_HH13 = 0
		v_HH14 = 0
		v_HH15 = 0
		v_HH16 = 0
		v_HH17 = 0
		v_HH18 = 0
		v_HH19 = 0
		v_HH20 = 0
		v_HH21 = 0
		v_HH22 = 0
		v_HH23 = 0
		v_HH24 = 0
		v_HH25 = 0
		v_HH26 = 0
		v_HH27 = 0
		v_HH28 = 0
		v_HH29 = 0
		v_HH30 = 0
		v_HH31 = 0

		v_TOTALCH = 0
		v_TOTALHH = 0

		strPusingan	= 	 "  SELECT COUNT(*) TOTAL FROM (SELECT fcba,widivision, " & _
						 "	SUM(DECODE(HARI,'01',TOTAL1,0))CH01," & _
						 "	SUM(DECODE(HARI,'02',TOTAL1,0))CH02," & _
						 "	SUM(DECODE(HARI,'03',TOTAL1,0))CH03," & _
						 "	SUM(DECODE(HARI,'04',TOTAL1,0))CH04," & _
						 "	SUM(DECODE(HARI,'05',TOTAL1,0))CH05," & _
						 "	SUM(DECODE(HARI,'06',TOTAL1,0))CH06," & _
						 "	SUM(DECODE(HARI,'07',TOTAL1,0))CH07," & _
						 "	SUM(DECODE(HARI,'08',TOTAL1,0))CH08," & _
						 "	SUM(DECODE(HARI,'09',TOTAL1,0))CH09," & _
						 "	SUM(DECODE(HARI,'10',TOTAL1,0))CH10," & _
						 "	SUM(DECODE(HARI,'11',TOTAL1,0))CH11," & _
						 "	SUM(DECODE(HARI,'12',TOTAL1,0))CH12," & _
						 "	SUM(DECODE(HARI,'13',TOTAL1,0))CH13," & _
						 "	SUM(DECODE(HARI,'14',TOTAL1,0))CH14," & _
						 "	SUM(DECODE(HARI,'15',TOTAL1,0))CH15," & _
						 "	SUM(DECODE(HARI,'16',TOTAL1,0))CH16," & _
						 "	SUM(DECODE(HARI,'17',TOTAL1,0))CH17," & _
						 "	SUM(DECODE(HARI,'18',TOTAL1,0))CH18," & _
						 "	SUM(DECODE(HARI,'19',TOTAL1,0))CH19," & _
						 "	SUM(DECODE(HARI,'20',TOTAL1,0))CH20," & _
						 "	SUM(DECODE(HARI,'21',TOTAL1,0))CH21," & _
						 "	SUM(DECODE(HARI,'22',TOTAL1,0))CH22," & _
						 "	SUM(DECODE(HARI,'23',TOTAL1,0))CH23," & _
						 "	SUM(DECODE(HARI,'24',TOTAL1,0))CH24," & _
						 "	SUM(DECODE(HARI,'25',TOTAL1,0))CH25," & _
						 "	SUM(DECODE(HARI,'26',TOTAL1,0))CH26," & _
						 "	SUM(DECODE(HARI,'27',TOTAL1,0))CH27," & _
						 "	SUM(DECODE(HARI,'28',TOTAL1,0))CH28," & _
						 "	SUM(DECODE(HARI,'29',TOTAL1,0))CH29," & _
						 "	SUM(DECODE(HARI,'30',TOTAL1,0))CH30," & _
						 "	SUM(DECODE(HARI,'31',TOTAL1,0))CH31," & _
						 "	SUM(DECODE(HARI,'01',TOTAL2,0))HH01," & _
						 "	SUM(DECODE(HARI,'02',TOTAL2,0))HH02," & _
						 " 	SUM(DECODE(HARI,'03',TOTAL2,0))HH03," & _
						 "	SUM(DECODE(HARI,'04',TOTAL2,0))HH04," & _
						 "	SUM(DECODE(HARI,'05',TOTAL2,0))HH05," & _
						 "	SUM(DECODE(HARI,'06',TOTAL2,0))HH06," & _
						 "	SUM(DECODE(HARI,'07',TOTAL2,0))HH07," & _
						 "	SUM(DECODE(HARI,'08',TOTAL2,0))HH08," & _
						 "	SUM(DECODE(HARI,'09',TOTAL2,0))HH09," & _
						 "	SUM(DECODE(HARI,'10',TOTAL2,0))HH10," & _
						 "	SUM(DECODE(HARI,'11',TOTAL2,0))HH11," & _
						 "	SUM(DECODE(HARI,'12',TOTAL2,0))HH12," & _
						 "	SUM(DECODE(HARI,'13',TOTAL2,0))HH13," & _
						 "	SUM(DECODE(HARI,'14',TOTAL2,0))HH14," & _
						 "	SUM(DECODE(HARI,'15',TOTAL2,0))HH15," & _
						 "	SUM(DECODE(HARI,'16',TOTAL2,0))HH16," & _
						 "	SUM(DECODE(HARI,'17',TOTAL2,0))HH17," & _
						 "	SUM(DECODE(HARI,'18',TOTAL2,0))HH18," & _
						 "	SUM(DECODE(HARI,'19',TOTAL2,0))HH19," & _
						 "	SUM(DECODE(HARI,'20',TOTAL2,0))HH20," & _
						 "	SUM(DECODE(HARI,'21',TOTAL2,0))HH21," & _
						 "	SUM(DECODE(HARI,'22',TOTAL2,0))HH22," & _
						 "	SUM(DECODE(HARI,'23',TOTAL2,0))HH23," & _
						 "	SUM(DECODE(HARI,'24',TOTAL2,0))HH24," & _
						 "	SUM(DECODE(HARI,'25',TOTAL2,0))HH25," & _
						 "	SUM(DECODE(HARI,'26',TOTAL2,0))HH26," & _
						 "	SUM(DECODE(HARI,'27',TOTAL2,0))HH27," & _
						 "	SUM(DECODE(HARI,'28',TOTAL2,0))HH28," & _
						 "	SUM(DECODE(HARI,'29',TOTAL2,0))HH29," & _
						 "	SUM(DECODE(HARI,'30',TOTAL2,0))HH30," & _
						 "	SUM(DECODE(HARI,'31',TOTAL2,0))HH31" & _
						 "	FROM (" & _
						 "	select TO_CHAR(widate,'YYYY')AS TAHUN,TO_CHAR(widate,'MM')AS BULAN,TO_CHAR(widate,'DD')AS HARI, fcba,widivision,sum(TOTRNINCHES) as total1,sum(TOTRNDAYS) as TOTAL2" & _
						 "	from iplasprod.weather" & _
						 "	where to_char(widate,'MMYYYY')='"& txttdate &"'  "& whereCons &""& _
						 "	group by TO_CHAR(widate,'YYYY'),TO_CHAR(widate,'MM'),TO_CHAR(widate,'DD'),fcba,widivision" & _
						 "	)" & _
						 "	group by fcba,widivision" & _
						 "	order by fcba,widivision)"

		' response.write strPusingan 

		Set qORS = Server.CreateObject("ADODB.Recordset")
		qORS.Open strPusingan, connstr, adOpenStatic
		If Not qORS.EOF Then
			Total = cdbl(to_nol(qORS("TOTAL")))
		End If
		qORS.Close
		set qORS = Nothing


		if Not rsCat.eof then
			While Not rsCat.EOF
		TOTALCH = cdbl(to_nol(rscat("CH01")))+cdbl(to_nol(rscat("CH02")))+cdbl(to_nol(rscat("CH03")))+cdbl(to_nol(rscat("CH04")))+cdbl(to_nol(rscat("CH05")))+cdbl(to_nol(rscat("CH06")))+cdbl(to_nol(rscat("CH07")))+cdbl(to_nol(rscat("CH08")))+cdbl(to_nol(rscat("CH09")))+cdbl(to_nol(rscat("CH10")))+cdbl(to_nol(rscat("CH11")))+cdbl(to_nol(rscat("CH12")))+cdbl(to_nol(rscat("CH13")))+cdbl(to_nol(rscat("CH14")))+cdbl(to_nol(rscat("CH15")))+cdbl(to_nol(rscat("CH16")))+cdbl(to_nol(rscat("CH17")))+cdbl(to_nol(rscat("CH18")))+cdbl(to_nol(rscat("CH19")))+cdbl(to_nol(rscat("CH20")))+cdbl(to_nol(rscat("CH21")))+cdbl(to_nol(rscat("CH22")))+cdbl(to_nol(rscat("CH23")))+cdbl(to_nol(rscat("CH24")))+cdbl(to_nol(rscat("CH25")))+cdbl(to_nol(rscat("CH26")))+cdbl(to_nol(rscat("CH27")))+cdbl(to_nol(rscat("CH28")))+cdbl(to_nol(rscat("CH29")))+cdbl(to_nol(rscat("CH30")))+cdbl(to_nol(rscat("CH31")))
		
		TOTALHH = cdbl(to_nol(rscat("HH01")))+cdbl(to_nol(rscat("HH02")))+cdbl(to_nol(rscat("HH03")))+cdbl(to_nol(rscat("HH04")))+cdbl(to_nol(rscat("HH05")))+cdbl(to_nol(rscat("HH06")))+cdbl(to_nol(rscat("HH07")))+cdbl(to_nol(rscat("HH08")))+cdbl(to_nol(rscat("HH09")))+cdbl(to_nol(rscat("HH10")))+cdbl(to_nol(rscat("HH11")))+cdbl(to_nol(rscat("HH12")))+cdbl(to_nol(rscat("HH13")))+cdbl(to_nol(rscat("HH14")))+cdbl(to_nol(rscat("HH15")))+cdbl(to_nol(rscat("HH16")))+cdbl(to_nol(rscat("HH17")))+cdbl(to_nol(rscat("HH18")))+cdbl(to_nol(rscat("HH19")))+cdbl(to_nol(rscat("HH20")))+cdbl(to_nol(rscat("HH21")))+cdbl(to_nol(rscat("HH22")))+cdbl(to_nol(rscat("HH23")))+cdbl(to_nol(rscat("HH24")))+cdbl(to_nol(rscat("HH25")))+cdbl(to_nol(rscat("HH26")))+cdbl(to_nol(rscat("HH27")))+cdbl(to_nol(rscat("HH28")))+cdbl(to_nol(rscat("HH29")))+cdbl(to_nol(rscat("HH30")))+cdbl(to_nol(rscat("HH31")))
		%>
		<tr>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("fcba")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("widivision")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("CH01")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("CH02")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("CH03")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("CH04")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("CH05")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("CH06")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("CH07")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("CH08")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("CH09")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("CH10")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("CH11")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("CH12")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("CH13")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("CH14")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("CH15")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("CH16")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("CH17")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("CH18")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("CH19")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("CH20")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("CH21")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("CH22")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("CH23")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("CH24")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("CH25")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("CH26")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("CH27")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("CH28")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("CH29")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("CH30")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("CH31")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=TOTALCH%></td>

			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("HH01")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("HH02")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("HH03")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("HH04")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("HH05")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("HH06")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("HH07")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("HH08")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("HH09")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("HH10")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("HH11")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("HH12")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("HH13")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("HH14")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("HH15")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("HH16")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("HH17")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("HH18")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("HH19")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("HH20")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("HH21")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("HH22")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("HH23")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("HH24")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("HH25")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("HH26")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("HH27")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("HH28")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("HH29")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("HH30")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("HH31")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=TOTALHH%></td>

		</tr>
	<%	
		v_CH01 = v_CH01 +  cdbl(to_nol(rsCat("CH01")))  / Total 
		v_CH02 = v_CH02 +  cdbl(to_nol(rsCat("CH02")))  / Total 
		v_CH03 = v_CH03 +  cdbl(to_nol(rsCat("CH03")))  / Total 
		v_CH04 = v_CH04 +  cdbl(to_nol(rsCat("CH04")))  / Total 
		v_CH05 = v_CH05 +  cdbl(to_nol(rsCat("CH05")))  / Total 
		v_CH06 = v_CH06 +  cdbl(to_nol(rsCat("CH06")))  / Total 
		v_CH07 = v_CH07 +  cdbl(to_nol(rsCat("CH07")))  / Total 
		v_CH08 = v_CH08 +  cdbl(to_nol(rsCat("CH08")))  / Total 
		v_CH09 = v_CH09 +  cdbl(to_nol(rsCat("CH09")))  / Total 
		v_CH10 = v_CH10 +  cdbl(to_nol(rsCat("CH10")))  / Total 
		v_CH11 = v_CH11 +  cdbl(to_nol(rsCat("CH11")))  / Total 
		v_CH12 = v_CH12 +  cdbl(to_nol(rsCat("CH12")))  / Total 
		v_CH13 = v_CH13 +  cdbl(to_nol(rsCat("CH13")))  / Total 
		v_CH14 = v_CH14 +  cdbl(to_nol(rsCat("CH14")))  / Total 
		v_CH15 = v_CH15 +  cdbl(to_nol(rsCat("CH15")))  / Total 
		v_CH16 = v_CH16 +  cdbl(to_nol(rsCat("CH16")))  / Total 
		v_CH17 = v_CH17 +  cdbl(to_nol(rsCat("CH17")))  / Total 
		v_CH18 = v_CH18 +  cdbl(to_nol(rsCat("CH18")))  / Total 
		v_CH19 = v_CH19 +  cdbl(to_nol(rsCat("CH19")))  / Total 
		v_CH20 = v_CH20 +  cdbl(to_nol(rsCat("CH20")))  / Total 
		v_CH21 = v_CH21 +  cdbl(to_nol(rsCat("CH21")))  / Total 
		v_CH22 = v_CH22 +  cdbl(to_nol(rsCat("CH22")))  / Total 
		v_CH23 = v_CH23 +  cdbl(to_nol(rsCat("CH23")))  / Total 
		v_CH24 = v_CH24 +  cdbl(to_nol(rsCat("CH24")))  / Total 
		v_CH25 = v_CH25 +  cdbl(to_nol(rsCat("CH25")))  / Total 
		v_CH26 = v_CH26 +  cdbl(to_nol(rsCat("CH26")))  / Total 
		v_CH27 = v_CH27 +  cdbl(to_nol(rsCat("CH27")))  / Total 
		v_CH28 = v_CH28 +  cdbl(to_nol(rsCat("CH28")))  / Total 
		v_CH29 = v_CH29 +  cdbl(to_nol(rsCat("CH29")))  / Total 
		v_CH30 = v_CH30 +  cdbl(to_nol(rsCat("CH30")))  / Total 
		v_CH31 = v_CH31 +  cdbl(to_nol(rsCat("CH31")))  / Total  
		v_TOTALCH = v_TOTALCH +  cdbl(to_nol(TOTALCH))  / Total  

		v_HH01 = v_HH01 +  cdbl(to_nol(rsCat("CH01")))  / Total 
		v_HH02 = v_HH02 +  cdbl(to_nol(rsCat("CH02")))  / Total 
		v_HH03 = v_HH03 +  cdbl(to_nol(rsCat("CH03")))  / Total 
		v_HH04 = v_HH04 +  cdbl(to_nol(rsCat("CH04")))  / Total 
		v_HH05 = v_HH05 +  cdbl(to_nol(rsCat("CH05")))  / Total 
		v_HH06 = v_HH06 +  cdbl(to_nol(rsCat("CH06")))  / Total 
		v_HH07 = v_HH07 +  cdbl(to_nol(rsCat("CH07")))  / Total 
		v_HH08 = v_HH08 +  cdbl(to_nol(rsCat("CH08")))  / Total 
		v_HH09 = v_HH09 +  cdbl(to_nol(rsCat("CH09")))  / Total 
		v_HH10 = v_HH10 +  cdbl(to_nol(rsCat("CH10")))  / Total 
		v_HH11 = v_HH11 +  cdbl(to_nol(rsCat("CH11")))  / Total 
		v_HH12 = v_HH12 +  cdbl(to_nol(rsCat("CH12")))  / Total 
		v_HH13 = v_HH13 +  cdbl(to_nol(rsCat("CH13")))  / Total 
		v_HH14 = v_HH14 +  cdbl(to_nol(rsCat("CH14")))  / Total 
		v_HH15 = v_HH15 +  cdbl(to_nol(rsCat("CH15")))  / Total 
		v_HH16 = v_HH16 +  cdbl(to_nol(rsCat("CH16")))  / Total 
		v_HH17 = v_HH17 +  cdbl(to_nol(rsCat("CH17")))  / Total 
		v_HH18 = v_HH18 +  cdbl(to_nol(rsCat("CH18")))  / Total 
		v_HH19 = v_HH19 +  cdbl(to_nol(rsCat("CH19")))  / Total 
		v_HH20 = v_HH20 +  cdbl(to_nol(rsCat("CH20")))  / Total 
		v_HH21 = v_HH21 +  cdbl(to_nol(rsCat("CH21")))  / Total 
		v_HH22 = v_HH22 +  cdbl(to_nol(rsCat("CH22")))  / Total 
		v_HH23 = v_HH23 +  cdbl(to_nol(rsCat("CH23")))  / Total 
		v_HH24 = v_HH24 +  cdbl(to_nol(rsCat("CH24")))  / Total 
		v_HH25 = v_HH25 +  cdbl(to_nol(rsCat("CH25")))  / Total 
		v_HH26 = v_HH26 +  cdbl(to_nol(rsCat("CH26")))  / Total 
		v_HH27 = v_HH27 +  cdbl(to_nol(rsCat("CH27")))  / Total 
		v_HH28 = v_HH28 +  cdbl(to_nol(rsCat("CH28")))  / Total 
		v_HH29 = v_HH29 +  cdbl(to_nol(rsCat("CH29")))  / Total 
		v_HH30 = v_HH30 +  cdbl(to_nol(rsCat("CH30")))  / Total 
		v_HH31 = v_HH31 +  cdbl(to_nol(rsCat("CH31")))  / Total 
		v_TOTALHH = v_TOTALHH +  cdbl(to_nol(TOTALHH))  / Total 

		rsCat.MoveNext
		Wend
	Else
		response.write "<tr><td colspan='64'>No Data Available</td></tr>"
	End If
	rsCat.close()
	rsCat=null%>
	<input  type="hidden" value="<%=txtTable%>" name="txtTable">
	<input  type="hidden" value="<%=txtModule%>" name="txtModule">
	</tbody>
	<tr bgcolor="#F4F4A0">
		<td align="center" colspan="2"><strong>AVG</strong></td>
		<td align="right"><strong><%=FormatNumber(v_CH01,2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(v_CH02,2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(v_CH03,2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(v_CH04,2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(v_CH05,2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(v_CH06,2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(v_CH07,2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(v_CH08,2)%></strong></td> 
		<td align="right"><strong><%=FormatNumber(v_CH09,2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(v_CH10,2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(v_CH11,2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(v_CH12,2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(v_CH13,2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(v_CH14,2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(v_CH15,2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(v_CH16,2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(v_CH17,2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(v_CH18,2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(v_CH19,2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(v_CH20,2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(v_CH21,2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(v_CH22,2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(v_CH23,2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(v_CH24,2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(v_CH25,2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(v_CH26,2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(v_CH27,2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(v_CH28,2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(v_CH29,2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(v_CH30,2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(v_CH31,2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(v_TOTALCH,2)%></strong></td>

		<td align="right"><strong><%=FormatNumber(v_HH01,2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(v_HH02,2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(v_HH03,2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(v_HH04,2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(v_HH05,2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(v_HH06,2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(v_HH07,2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(v_HH08,2)%></strong></td> 
		<td align="right"><strong><%=FormatNumber(v_HH09,2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(v_HH10,2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(v_HH11,2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(v_HH12,2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(v_HH13,2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(v_HH14,2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(v_HH15,2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(v_HH16,2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(v_HH17,2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(v_HH18,2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(v_HH19,2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(v_HH20,2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(v_HH21,2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(v_HH22,2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(v_HH23,2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(v_HH24,2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(v_HH25,2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(v_HH26,2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(v_HH27,2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(v_HH28,2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(v_HH29,2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(v_HH30,2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(v_HH31,2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(v_TOTALHH,2)%></strong></td>
	</tr>
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