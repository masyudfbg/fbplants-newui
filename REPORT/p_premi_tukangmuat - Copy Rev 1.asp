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
	
	strGrnInv	=	" SELECT gad.fcba," & _
					"        gad.employeecode," & _
					"        emp.fcname AS employeedesc," & _
					"        gad.fddate," & _
					"        gad.division," & _
					"        gad.janjang," & _
					"        gad.rate_janjang," & _
					"        ROUND (gad.kg_janjang, 0) AS kg_janjang," & _
					"        (gad.rate_janjang * gad.kg_janjang) AS premi_tbs," & _
					"        gad.bucket," & _
					"        gad.rate_brondolan," & _
					"        ROUND (gad.kg_brondolan, 0) AS kg_brondolan," & _
					"        (gad.rate_brondolan * gad.kg_brondolan) AS premi_brd," & _
					"        gad.amount_empallded AS amount_empallded," & _
					"        (  (gad.rate_janjang * gad.kg_janjang)" & _
					"         + (gad.rate_brondolan * gad.kg_brondolan)" & _
					"         + gad.amount_empallded)" & _
					"           AS total_premi_langsir" & _
					"   FROM (  SELECT fcba," & _
					"                  employeecode," & _
					"                  fddate," & _
					"                  division," & _
					"                  SUM (janjang) AS janjang," & _
					"                  MAX (rate_janjang) AS rate_janjang," & _
					"                  SUM (kg_janjang) AS kg_janjang," & _
					"                  SUM (bucket) AS bucket," & _
					"                  MAX (rate_brondolan) AS rate_brondolan," & _
					"                  SUM (kg_brondolan) AS kg_brondolan," & _
					"                  SUM (amount_empallded) AS amount_empallded" & _
					"             FROM (  SELECT iplasprod.attendance_gad.fcba," & _
					"                            iplasprod.attendance_gad.employeecode," & _
					"                            iplasprod.attendance_gad.fddate," & _
					"                            iplasprod.field.division," & _
					"                            SUM (iplasprod.attendance_gad.kg_janjang) AS janjang," & _
					"                            iplasprod.attendance_gad.rate AS rate_janjang," & _
					"                            SUM (iplasprod.attendance_gad.unit) AS kg_janjang," & _
					"                            0 AS bucket," & _
					"                            0 AS rate_brondolan," & _
					"                            0 AS kg_brondolan," & _
					"                            NVL (SUM (iplasprod.attendance_empallded.amount), 0)" & _
					"                               AS amount_empallded" & _
					"                       FROM iplasprod.attendance_gad" & _
					"                            LEFT JOIN iplasprod.attendance_empallded" & _
					"                               ON (    iplasprod.attendance_gad.fcba =" & _
					"                                          iplasprod.attendance_empallded.fcba" & _
					"                                   AND iplasprod.attendance_gad.linenokey =" & _
					"                                          iplasprod.attendance_empallded.linenokey)" & _
					"                            INNER JOIN iplasprod.field" & _
					"                               ON (    iplasprod.attendance_gad.fcba =" & _
					"                                          iplasprod.field.fcba" & _
					"                                   AND iplasprod.attendance_gad.locationcode =" & _
					"                                          iplasprod.field.fccode)" & _
					"                      WHERE     iplasprod.attendance_gad.attendance = 'KB'" & _
					"                            AND iplasprod.attendance_gad.unit > 0" & _
					"                            AND iplasprod.attendance_gad.jobcode IN" & _
					"                                   (SELECT DISTINCT" & _
					"                                           iplasprod.parameterdetail.fccode" & _
					"                                      FROM iplasprod.parameterdetail" & _
					"                                     WHERE iplasprod.parameterdetail.parcode =" & _
					"                                              'JOBCODE_PENGANGKUTAN')" & _
					"                   GROUP BY iplasprod.attendance_gad.fcba," & _
					"                            iplasprod.attendance_gad.employeecode," & _
					"                            iplasprod.attendance_gad.fddate," & _
					"                            iplasprod.field.division," & _
					"                            iplasprod.attendance_gad.rate" & _
					"                   UNION ALL" & _
					"                     SELECT iplasprod.attendance_gad.fcba," & _
					"                            iplasprod.attendance_gad.employeecode," & _
					"                            iplasprod.attendance_gad.fddate," & _
					"                            iplasprod.field.division," & _
					"                            0 AS janjang," & _
					"                            0 AS rate_janjang," & _
					"                            0 AS kg_janjang," & _
					"                            SUM (iplasprod.attendance_gad.bucket) AS bucket," & _
					"                            iplasprod.attendance_gad.rate AS rate_brondolan," & _
					"                            SUM (iplasprod.attendance_gad.kg_brondolan)" & _
					"                               AS kg_brondolan," & _
					"                            NVL (SUM (iplasprod.attendance_empallded.amount), 0)" & _
					"                               AS amount_empallded" & _
					"                       FROM iplasprod.attendance_gad" & _
					"                            LEFT JOIN iplasprod.attendance_empallded" & _
					"                               ON (    iplasprod.attendance_gad.fcba =" & _
					"                                          iplasprod.attendance_empallded.fcba" & _
					"                                   AND iplasprod.attendance_gad.linenokey =" & _
					"                                          iplasprod.attendance_empallded.linenokey)" & _
					"                            INNER JOIN iplasprod.field" & _
					"                               ON (    iplasprod.attendance_gad.fcba =" & _
					"                                          iplasprod.field.fcba" & _
					"                                   AND iplasprod.attendance_gad.locationcode =" & _
					"                                          iplasprod.field.fccode)" & _
					"                      WHERE     iplasprod.attendance_gad.attendance = 'KB'" & _
					"                            AND iplasprod.attendance_gad.kg_brondolan > 0" & _
					"                            AND iplasprod.attendance_gad.jobcode IN" & _
					"                                   (SELECT DISTINCT" & _
					"                                           iplasprod.parameterdetail.fccode" & _
					"                                      FROM iplasprod.parameterdetail" & _
					"                                     WHERE iplasprod.parameterdetail.parcode =" & _
					"                                              'JOBCODE_PENGANGKUTAN')" & _
					"                   GROUP BY iplasprod.attendance_gad.fcba," & _
					"                            iplasprod.attendance_gad.employeecode," & _
					"                            iplasprod.attendance_gad.fddate," & _
					"                            iplasprod.field.division," & _
					"                            iplasprod.attendance_gad.rate)" & _
					"         GROUP BY fcba," & _
					"                  employeecode," & _
					"                  fddate," & _
					"                  division" & _
					"         ORDER BY fcba, employeecode, fddate) gad" & _
					"        INNER JOIN iplasprod.employee emp" & _
					"           ON (gad.fcba = emp.fcba AND gad.employeecode = emp.fccode)" & _
					"  WHERE     gad.fcba = '"& v_txtfcba &"'" & _
					"        AND TO_NUMBER (TO_CHAR (gad.fddate, 'mm')) = '"& v_txtMonth &"'" & _
					"        AND TO_NUMBER (TO_CHAR (gad.fddate, 'yyyy')) = '"& v_txtYear &"'" 		 


	if  len(trim(v_txtsectionname)) > 0 then
	    strGrnInv	=	strGrnInv & " AND gad.division = '"& v_txtsectionname &"'" 
	end if
	if  len(trim(v_txtempcode)) = 2 then
	    strGrnInv	=	strGrnInv & " AND gad.employeecode = '"& v_txtempcode &"'" 
	end if
	' if  cdbl(to_nol(compare1)) = 3 then
	'     strGrnInv	=	strGrnInv & " and vehicleactivity.vh_employeecode != harvestingspb.harvesting_employeecode" 
	' end if

	' strGrnInv	=	strGrnInv &	" GROUP BY" & _
	' 							"     FCBA" & _
	' 							"     , NOPO" & _
	' 							"     , STOCKCODE"
								
	strGrnInv	=	strGrnInv & " ORDER BY gad.fcba," & _
								"          gad.employeecode," & _
								"          gad.division," & _
								"          gad.fddate" 

	' response.write strGrnInv	
	Set rsGrnInv = Server.CreateObject("ADODB.Recordset")
	rsGrnInv.Open strGrnInv, connstr, adOpenStatic 
	
%>


<html>

<head>
<title>ePlaS Reporting - Premi Tukang Muat</title>
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
							<font face="verdana" size="1">DIVISION</font>
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
							<font face="verdana" size="1">EMPLOYEE</font>
						</td>
						<td>
							<select name="txtempcode" >
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
    		<th colspan="14" style="background-color:#FFFFFF; border:none;">  
				<div id="header-wrapper">
					<div id="head_judul" style="float:left; text-align:left;">
						<h4>PT. <%=CompanyName%></h4>
						<h5><%=fcbaname%></h5>
					</div>
					<div id="subheader" class="centered" tyle="width:50%; margin:auto;">
						<h2>LAPORAN PREMI TUKANG MUAT</h2>
		                <br/>
						<h4>PERIODE&nbsp;:&nbsp;<%=textMonth(v_txtMonth)%>&nbsp;<%=v_txtYear%></h4>
					</div>
				</div>
			</th>
		</tr>
		<tr>
			<th style="background-color:#00A300;color:#fff">EMPLOYEE CODE</th>
			<th style="background-color:#00A300;color:#fff">EMPLOYEE NAME</th> 
			<th style="background-color:#00A300;color:#fff;">DATE</th>
			<th style="background-color:#00A300;color:#fff;">DIVISION</th>
			<th style="background-color:#00A300;color:#fff;">JANJANG</th>
			<th style="background-color:#00A300;color:#fff;">RATE JANJANG</th>
			<th style="background-color:#00A300;color:#fff;">KG JANJANG</th>
			<th style="background-color:#00A300;color:#fff;">PREMI TBS</th>		
			<th style="background-color:#00A300;color:#fff;">BUCKET</th>		
			<th style="background-color:#00A300;color:#fff;">RATE BRONDOLAN</th>		
			<th style="background-color:#00A300;color:#fff;">KG BRONDOLAN</th>		
			<th style="background-color:#00A300;color:#fff;">PREMI BRONDOLAN</th>		
			<th style="background-color:#00A300;color:#fff;">AMOUNT EMPALLDED</th>
			<th style="background-color:#00A300;color:#fff;">TOTAL PREMI LANGSIR</th>				
		</tr>   
		    </thead>
	<tbody>
<%if rsGrnInv.BOF and rsGrnInv.EOF then%>
		<tr>
			<td colspan="14" align="left">NO DATA AVAILABLE</td>
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
					sqlrow = "  select count(employeecode) rowsd  from ( " & strGrnInv & " ) where employeecode = '" & rsGrnInv("employeecode") & "'"
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
				<td align="center" Rowspan="<%=temprows+1 %>"><%=rsGrnInv("employeecode")%></td>
				<td align="center" Rowspan="<%=temprows+1 %>"><%=rsGrnInv("employeedesc")%></td>
			<%end if%>
			<td align="center"><%=rsGrnInv("fddate")%></td>
			<td align="center"><%=rsGrnInv("division")%></td> 
			<td align="right"><%=FormatNumber(cdbl(to_nol(rsGrnInv("janjang"))),0)%></td>
			<td align="right"><%=FormatNumber(cdbl(to_nol(rsGrnInv("rate_janjang"))),0)%></td>
			<td align="right"><%=FormatNumber(cdbl(to_nol(rsGrnInv("kg_janjang"))),0)%></td> 
			<td align="right"><%=FormatNumber(cdbl(to_nol(rsGrnInv("premi_tbs"))),0)%></td>
			<td align="right"><%=FormatNumber(cdbl(to_nol(rsGrnInv("bucket"))),0)%></td> 
			<td align="right"><%=FormatNumber(cdbl(to_nol(rsGrnInv("rate_brondolan"))),0)%></td> 
			<td align="right"><%=FormatNumber(cdbl(to_nol(rsGrnInv("kg_brondolan"))),0)%></td>
			<td align="right"><%=FormatNumber(cdbl(to_nol(rsGrnInv("premi_brd"))),0)%></td> 
			<td align="right"><%=FormatNumber(cdbl(to_nol(rsGrnInv("amount_empallded"))),0)%></td>
			<td align="right"><%=FormatNumber(cdbl(to_nol(rsGrnInv("total_premi_langsir"))),0)%></td>
		</tr>
<%	if rows = temprows then
			njanjang = njanjang + cdbl(to_nol(rsGrnInv("janjang")))
			nrate_janjang = nrate_janjang + cdbl(to_nol(rsGrnInv("rate_janjang")))	
			nkg_janjang = nkg_janjang + cdbl(to_nol(rsGrnInv("kg_janjang")))
			npremi_tbs = npremi_tbs + cdbl(to_nol(rsGrnInv("premi_tbs")))
			nbucket = nbucket + cdbl(to_nol(rsGrnInv("bucket")))	

			nrate_brondolan = nrate_brondolan + cdbl(to_nol(rsGrnInv("rate_brondolan")))
			nkg_brondolan = nkg_brondolan + cdbl(to_nol(rsGrnInv("kg_brondolan")))
			npremi_brd = npremi_brd + cdbl(to_nol(rsGrnInv("premi_brd")))
			namount_empallded = namount_empallded + cdbl(to_nol(rsGrnInv("amount_empallded")))
			ntotal_premi_langsir = ntotal_premi_langsir + cdbl(to_nol(rsGrnInv("total_premi_langsir"))) 
			rows = 1
			
	%>
		<tr bgcolor="#F4F4A0"> 
			<td align="center" colspan="2"><font style="font-weight: bold;">TOTAL</font></td>				
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(njanjang)),0)%></font></td> 
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(nrate_janjang)),0)%></font></td>  
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(nkg_janjang)),0)%></font></td> 
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(npremi_tbs)),0)%></font></td> 
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(nbucket)),0)%></font></td> 
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(nrate_brondolan)),0)%></font></td> 
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(nkg_brondolan)),0)%></font></td> 
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(npremi_brd)),0)%></font></td>  
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(namount_empallded)),0)%></font></td>  
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(ntotal_premi_langsir)),0)%></font></td>
		</tr>
		<tr>
			<td align="center" colspan="<%= 14+ nSpanTotal %>" ></td>
		</tr>
		<%
			njanjang = cdbl("0")
			nrate_janjang = cdbl("0")
			nkg_janjang = cdbl("0")
			npremi_tbs = cdbl("0")
			nbucket = cdbl("0") 
			nrate_brondolan = cdbl("0")
			nkg_brondolan = cdbl("0") 
			npremi_brd = cdbl("0") 
			namount_empallded = cdbl("0") 
			ntotal_premi_langsir = cdbl("0")
		else
		' cari sub total
			njanjang = njanjang + cdbl(to_nol(rsGrnInv("janjang")))
			nrate_janjang = nrate_janjang + cdbl(to_nol(rsGrnInv("rate_janjang")))	
			nkg_janjang = nkg_janjang + cdbl(to_nol(rsGrnInv("kg_janjang")))
			npremi_tbs = npremi_tbs + cdbl(to_nol(rsGrnInv("premi_tbs")))
			nbucket = nbucket + cdbl(to_nol(rsGrnInv("bucket")))
			nrate_brondolan = nrate_brondolan + cdbl(to_nol(rsGrnInv("rate_brondolan")))
			nkg_brondolan = nkg_brondolan + cdbl(to_nol(rsGrnInv("kg_brondolan")))
			npremi_brd = npremi_brd + cdbl(to_nol(rsGrnInv("premi_brd")))
			namount_empallded = namount_empallded + cdbl(to_nol(rsGrnInv("amount_empallded")))
			ntotal_premi_langsir = ntotal_premi_langsir + cdbl(to_nol(rsGrnInv("total_premi_langsir")))
			rows = rows + 1
		End if
							  
			' cari grand total 
				gt_janjang = cdbl(gt_janjang) + cdbl(rsGrnInv("janjang"))
				gt_rate_janjang = cdbl(gt_rate_janjang) + cdbl(rsGrnInv("rate_janjang"))
				gt_kg_janjang = cdbl(gt_kg_janjang) + cdbl(rsGrnInv("kg_janjang"))
				gt_premi_tbs = cdbl(gt_premi_tbs) + cdbl(rsGrnInv("premi_tbs"))
				gt_bucket = cdbl(gt_bucket) + cdbl(rsGrnInv("bucket")) 
				gt_rate_brondolan = cdbl(gt_rate_brondolan) + cdbl(rsGrnInv("rate_brondolan"))
				gt_kg_brondolan = cdbl(gt_kg_brondolan) + cdbl(rsGrnInv("kg_brondolan"))
				gt_premi_brd = cdbl(gt_premi_brd) + cdbl(rsGrnInv("premi_brd")) 
				gt_amount_empallded = cdbl(gt_amount_empallded) + cdbl(rsGrnInv("amount_empallded")) 
				gt_total_premi_langsir = cdbl(gt_total_premi_langsir) + cdbl(rsGrnInv("total_premi_langsir")) 

			rsGrnInv.MoveNext			
			Wend
		rsGrnInv.Close
		set rsGrnInv = Nothing
	%>
		<tr bgcolor="#FFFF00">	 
			<td align="center" colspan="4"><b>GRAND TOTAL</b></td>
			<td align="right"><b><%=formatnumber(gt_janjang,0)%></td>
			<td align="right"><b><%=formatnumber(gt_rate_janjang,0)%></td>
			<td align="right"><b><%=formatnumber(gt_kg_janjang,0)%></td>
			<td align="right"><b><%=formatnumber(gt_premi_tbs,0)%></td>
			<td align="right"><b><%=formatnumber(gt_bucket,0)%></td>
			<td align="right"><b><%=formatnumber(gt_rate_brondolan,0)%></td>
			<td align="right"><b><%=formatnumber(gt_kg_brondolan,0)%></td>
			<td align="right"><b><%=formatnumber(gt_premi_brd,0)%></td>
			<td align="right"><b><%=formatnumber(gt_amount_empallded,0)%></td>
			<td align="right"><b><%=formatnumber(gt_total_premi_langsir,0)%></td>
			<!--<td align="right"></td>-->
		</tr>
	</tbody>
</table>