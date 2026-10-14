<% ValidateSession %>
<!--#include virtual="/ConndB1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->

<%
Dim v_txtMonth, v_txtYear
Dim rsDefaultBa, strDefaultBa, varDefaultBa
Dim rscurrentdate, strcurrentdate, varDefaultYear, varDefaultMonth
Dim ScriptName, v_txtsectionname, v_txtGangCode, v_txtStatus, v_txtposting
Dim rsBaInfo, strBaInfo, varFcbaName, varCompanyName, varManagerName, varKTUName, no,v_txtFddateto

v_txtFddate        = Request("txtFddate")
v_txtFddateto      = Request("txtFddateto")

v_txtsectionname   = Request("txtsectionname")
v_txtGangCode	   = Request("txtGangCode")
v_txtStatus		   = Request("txtStatus")
ScriptName		   = Request.ServerVariables("SCRIPT_NAME")
v_txtFcba          = Request("txtFcba")
v_txtposting       = Request("txtposting")

' BA default user bila belum dipilih (sama seperti p_payroll_bkm.asp)
If Len(Trim(v_txtFcba)) = 0 Then
	strDefaultBa = " select fccode as fcba from businessunit bu" & _
	               " inner join user_fcba_tab uft on (bu.fccode = uft.fcba) " & _
	               " where isdefault = 'TRUE' and bu.IsActive= 'TRUE' and fcusercode = '" & Session("CodeName") & "'"
	Set rsDefaultBa = Server.CreateObject("ADODB.Recordset")
	rsDefaultBa.Open strDefaultBa, connstr, adOpenStatic
	If Not rsDefaultBa.EOF Then
		v_txtFcba = rsDefaultBa("fcba")
	Else
		v_txtFcba = Session("fcba")
	End If
	rsDefaultBa.Close
	Set rsDefaultBa = Nothing
End If

'Response.Write v_txtFcba
'Response.end

nSpanTotal = 0
'if v_txtposting = "" then 
'	v_txtposting = "1"
'end if

if v_txtFddate = "" then
	   v_txtFddate = textdate(Now())
end if	

if v_txtFddateto = "" then
	   v_txtFddateto = textdate(Now())
end if	

if len(trim(v_txtStatus)) > 0 then 
	arr_indeks_eof = true
	oSql =  " SELECT '--ALL--' as STATUS FROM DUAL UNION SELECT DISTINCT PRESENTGRADE as STATUS FROM EMPLOYEE WHERE FCBA = '"& v_txtFcba &"' ORDER BY Status ASC"
	set oRs = server.CreateObject("Adodb.recordset")
	oRs.Open oSql, connstr
	if Not oRs.Eof then
		arr_indeks2 = oRs.getrows()
		arr_indeks_eof = false
	End if
	oRs.Close
	Set oRs = Nothing
else
	arr_indeks_eof = true
	oSql =  " SELECT '--ALL--' as STATUS FROM DUAL UNION SELECT DISTINCT PRESENTGRADE as STATUS FROM EMPLOYEE WHERE FCBA = '" & v_txtFcba & "' ORDER BY Status ASC"
	set oRs = server.CreateObject("Adodb.recordset")
	oRs.Open oSql, connstr
	if Not oRs.Eof then
		arr_indeks2 = oRs.getrows()
		arr_indeks_eof = false
	End if
	oRs.Close
	Set oRs = Nothing
end if



strBaInfo = " SELECT" & _
			"     FCNAME as FCBANAME" & _
			"     , FCCOMPANYNAME as COMPANYNAME" & _
			"     , FCMANAGER" & _
			"     , FCKTU" & _
			" FROM BUSINESSUNIT BU" & _
			" WHERE" & _
			"     FCCODE = '" & v_txtFcba & "'"

Set rsBaInfo = Server.CreateObject("ADODB.Recordset")
rsBaInfo.Open strBaInfo, connstr, adOpenStatic

If Not rsBaInfo.EOF Then
	varFcbaName = rsBaInfo("FCBANAME")
	varCompanyName = rsBaInfo("COMPANYNAME")
	varManagerName = rsBaInfo("FCMANAGER")
	varKTUName = rsBaInfo("FCKTU")
Else
	varFcbaName = ""
	varCompanyName = ""
	varManagerName = ""
	varKTUName = ""
End If
rsBaInfo.Close
set rsBaInfo = Nothing

' Added By Sendy 11/19/2018
	' if v_Emptype <>  "--ALL--" and v_Emptype <> "" then
		' whereCons = whereCons & " AND EMP_TYPE = '" & v_Emptype & "'"
	' else
		' nSpanTotal = nSpanTotal - 1  
	' end if
' End

    if  len(trim(v_txtsectionname)) > 0 then
		whereCons = whereCons & " and sectionname = '" & v_txtsectionname & "'"
	end if
	if  len(trim(v_txtGangCode)) > 0 then
		whereCons = whereCons & " and gangcode = '" & v_txtGangCode & "'"
	end if
	if  len(trim(v_txtStatus)) > 0 then
		whereCons = whereCons & " AND INSTR('"& v_txtStatus &"', SUBSTR(presentgrade,1,3) ) <> 0"
	end if

	Dim rsPayroll, strPayroll, noline 
	Dim nGet_Hk, nGet_Hk_kj, nget_mkp1, nBase_Salary, nrapelgaji, nOvertime, nPremi, nTunjangan, nBorongan, nP_Jamsostek, nP_bpjskes, nP_Mk_P1, nP_Panen, nP_Pph21, nP_koperasi, nP_Lain, nTotal
		strPayroll = 	" WITH payroll_daily AS  " & _
						"	(" & _
						"	  SELECT" & _
						"	  a.fcba," & _
						"	  c.postingflg," & _
						"	  BKMDATE FDDATE," & _
						"	  EXTRACT(MONTH FROM BKMDATE) Period_Month," & _
						"	  EXTRACT(YEAR FROM BKMDATE)  Period_YEAR," & _
						"	  a.EMPLOYEECODE," & _
						"	  employee_api.get_employeename(a.fcba,a.EMPLOYEECODE) EMPLOYEENAME," & _
						"	  Nvl(iplasprod.employee_hist_api.get_sectionname ( To_Number(To_Char(BKMDATE,'YYYY')) , To_Number(To_Char(BKMDATE,'MM')),a.fcba, a.EMPLOYEECODE),iplasprod.employee_api.get_sectionname (a.fcba, a.EMPLOYEECODE)) SECTIONNAME," & _
						"	  Nvl(iplasprod.employee_hist_api.get_gangcode ( To_Number(To_Char(BKMDATE,'YYYY')) , To_Number(To_Char(BKMDATE,'MM')),a.fcba, a.EMPLOYEECODE),iplasprod.employee_api.get_gangcode (a.fcba, a.EMPLOYEECODE)) GANGCODE," & _
						"	  Nvl(iplasprod.employee_hist_api.get_presentgrade ( To_Number(To_Char(BKMDATE,'YYYY')) , To_Number(To_Char(BKMDATE,'MM')),a.fcba, a.EMPLOYEECODE),iplasprod.employee_api.get_presentgrade (a.fcba, a.EMPLOYEECODE)) PRESENTGRADE," & _
						"	  SUM(Decode(a.allowdedcode,'A01',b.Mandays,0)) AS HK," & _
						"	  SUM(Decode(d.Allded_Group||d.allded,'GAJIA',a.AMOUNT,0))  AS GAJI," & _
						"	  SUM(Decode(d.Allded_Group||d.allded,'LEMBURA',a.AMOUNT,0))  AS LEMBUR," & _
						"	  SUM(Decode(d.Allded_Group||d.allded,'PREMIA',a.AMOUNT,0))  AS PREMI," & _
						"	  SUM(Decode(d.Allded_Group||d.allded,'BORONGANA',a.AMOUNT,0))  AS BORONGAN," & _
						"	  SUM(Decode(d.Allded_Group||d.allded,'TUNJANGANA',a.AMOUNT,0))  AS TUNJANGAN," & _
						"	  SUM(Decode(d.Allded_Group||d.allded,'THRA',a.AMOUNT,0)) AS THR_BONUS," & _
						"	  SUM(Decode(InStr('GAJID,PREMID',d.Allded_Group||d.allded),0,0,a.AMOUNT))  AS DENDA" & _
						"	  FROM BKMproses a" & _
						"	  INNER JOIN BKMUPAHDTL b ON a.BKMNO = b.BKMNO AND a.fcba=b.fcba AND a.EMPLOYEECODE=b.EMPLOYEECODE AND a.LINENOKEY=b.LINENOKEY" & _
						"	  INNER JOIN BKMHDR c ON a.BKMNO = c.BKMNO AND a.fcba=c.fcba" & _
						"	  INNER JOIN iplasprod.allowded d ON a.allowdedcode = d.fccode AND a.fcba=d.fcba" & _
						"	  WHERE a.fcba='" & v_txtFcba & "'" & _
						"	  AND BKMDATE BETWEEN TO_DATE ('" & v_txtFddate & "', 'MM/DD/YYYY') AND TO_DATE ('" & v_txtFddateto & "', 'MM/DD/YYYY') " & _
						"     AND c.postingflg like '" & v_txtposting & "%'" & _
						"	  GROUP BY  a.fcba,BKMDATE, a.EMPLOYEECODE,c.postingflg " & _
						"	  ORDER BY a.employeecode,bkmdate" & _
						"	)" & _
						" SELECT EMPLOYEECODE,EMPLOYEENAME,SECTIONNAME,GANGCODE,PRESENTGRADE, " & _
						" Sum(HK)HK,Sum(GAJI)GAJI,Sum(LEMBUR)LEMBUR,Sum(PREMI)PREMI,Sum(BORONGAN)BORONGAN,Sum(TUNJANGAN)TUNJANGAN,Sum(THR_BONUS)THR_BONUS,Sum(SubGaji)SubGaji,Sum(POT_DENDA)POT_DENDA,Sum(Sub_Potongan)Sub_Potongan,Sum(Total) Total, " & _
						" Sum(HK_shi)HK_shi,Sum(GAJI_shi)GAJI_shi,Sum(LEMBUR_shi)LEMBUR_shi,Sum(PREMI_shi)PREMI_shi,Sum(BORONGAN_shi)BORONGAN_shi,Sum(TUNJANGAN_shi)TUNJANGAN_shi,Sum(THR_BONUS_shi)THR_BONUS_shi,Sum(SubGaji_shi)SubGaji_shi,Sum(POT_DENDA_shi)POT_DENDA_shi,Sum(Sub_Potongan_shi)Sub_Potongan_shi,Sum(Total_shi) Total_shi " & _
						" FROM" & _
						" ( " & _
						" SELECT  EMPLOYEECODE,EMPLOYEENAME,SECTIONNAME,GANGCODE,PRESENTGRADE, " & _
						" HK,GAJI,LEMBUR,PREMI,BORONGAN,TUNJANGAN,THR_BONUS,Gaji+Lembur+Premi+Borongan+Tunjangan+Thr_Bonus AS SubGaji,DENDA POT_DENDA,DENDA Sub_Potongan, " & _ 
						" Gaji+Lembur+Premi+Borongan+Tunjangan+Thr_Bonus+DENDA Total," & _
						" 0 HK_shi, 0 GAJI_shi, 0 LEMBUR_shi,0 PREMI_shi,0 BORONGAN_shi,0 TUNJANGAN_shi, 0 THR_BONUS_shi,0 SubGaji_shi,0 POT_DENDA_shi, 0 Sub_Potongan_shi,0 Total_shi " & _
						" FROM  " & _
						" payroll_daily " & _
						" WHERE FDDATE=To_Date('" & v_txtFddateto & "','MM/DD/YYYY') " & _
						" UNION " & _
						" SELECT  EMPLOYEECODE,EMPLOYEENAME,SECTIONNAME,GANGCODE,PRESENTGRADE, " & _
						" 0 HK, 0 GAJI, 0 LEMBUR,0 PREMI,0 BORONGAN,0 TUNJANGAN, 0 THR_BONUS,0 SubGaji,0 POT_DENDA,0 Sub_Potongan,0 Total, " & _
						" Sum(HK)HK_shi,Sum(GAJI)GAJI_shi,Sum(LEMBUR)LEMBUR_shi,Sum(PREMI)PREMI_shi,Sum(BORONGAN)BORONGAN_shi,Sum(TUNJANGAN)TUNJANGAN_shi,Sum(THR_BONUS)THR_BONUS_shi," & _ 
						" Sum(GAJI) + Sum(LEMBUR) + Sum(PREMI) + Sum(BORONGAN) + Sum(TUNJANGAN) + Sum(THR_BONUS) AS SubGaji_shi,Sum(DENDA) POT_DENDA_shi,Sum(DENDA) Sub_Potongan_shi, " & _ 
						" Sum(GAJI) + Sum(LEMBUR) + Sum(PREMI) + Sum(BORONGAN) + Sum(TUNJANGAN) + Sum(THR_BONUS) + Sum(DENDA) Total_shi " & _
						" FROM   " & _
						" payroll_daily " & _
						" WHERE FDDATE BETWEEN TO_DATE ('" & v_txtFddate & "', 'MM/DD/YYYY') AND TO_DATE ('" & v_txtFddateto & "', 'MM/DD/YYYY') " & _
						" GROUP BY   EMPLOYEECODE,EMPLOYEENAME,SECTIONNAME,GANGCODE,PRESENTGRADE " & _
						" ) WHERE 1=1 " & whereCons & "" & _
						" GROUP BY " & _
						" EMPLOYEECODE,EMPLOYEENAME,SECTIONNAME,GANGCODE,PRESENTGRADE "

			strPayroll = strPayroll & " order by sectionname  asc, PRESENTGRADE desc, EMPLOYEENAME asc" 
		'	response.write strPayroll
			Set rsPayroll = Server.CreateObject("ADODB.Recordset")
			Dim payrollOK, errPayroll, secCount
			payrollOK = True : errPayroll = ""
			On Error Resume Next
			rsPayroll.Open strPayroll, connstr, adOpenStatic
			If Err.Number <> 0 Then
				payrollOK = False
				errPayroll = Err.Description
				Err.Clear
			End If
			On Error GoTo 0

			' Jumlah baris per section (untuk rowspan & sub total) dihitung sekali di memori.
			' Sebelumnya query di atas dijalankan ulang per section (sangat lambat / timeout).
			Set secCount = Server.CreateObject("Scripting.Dictionary")
			If payrollOK Then
				Do While Not rsPayroll.EOF
					secCount("" & rsPayroll("sectionname")) = secCount("" & rsPayroll("sectionname")) + 1
					rsPayroll.MoveNext
				Loop
				If Not (rsPayroll.BOF And rsPayroll.EOF) Then rsPayroll.MoveFirst
			End If
			
%>
<html>
<head>
	<title>Monitoring - Payroll Harian Karyawan</title>
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
	  }, false
	);
	</script>
	<script>
	$(function() {
	$( ".datepicker" ).datepicker();
	});
	</script>
	
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
@media all {
#footer_note {width:65%;margin:auto;margin-top:20px;}
#footer_note .block { font-size:10px; float:left; width:20%; height:90px; background-color:#fff; margin-left:1%; text-align:center;}
input[type="text"].isian {border:0px;box-shadow:none; padding:0px;font-size: 10px;color:#000;height: 20px;width:100%; text-align:center; font-weight:bold;}
hr{border-top:1px #000000 solid; border-bottom:0px; border-left:0px; border-right:0px;}
}
@media print
{
#footer_note {margin-top:20px;}
#footer_note .block { font-size:10px; float:left; width:20%; height:90px; background-color:#fff; margin-left:1%; text-align:center;}
input[type="text"].isian {border:0px;box-shadow:none; padding:0px;font-size: 10px;color:#000;height: 20px;width:100%; text-align:center; font-weight:bold;}
hr{border-top:1px #000000 solid; border-bottom:0px; border-left:0px; border-right:0px;}
.printhide{display:none;}
@page { 
 size:Letter landscape;
}
}

</style>
<link href="/plugins/chosen/chosen.min.css" rel="stylesheet" />
</head>

<body>
	<div id="wrapper">	
		<div id="header-wrapper">
			<div class="printhide">
			<input type="button" class="tableToPdf" value="Export PDF" onClick="window.print()" style="font-family: 'FontAwesome';margin-top:-10px;">
				<!--#include file="REPORT/menutop.asp"-->	
				<!--#include file="print_excel.asp"-->
			</div>
			<form name="frmAR" method="post" action="<%=ScriptName%>">
				<div class="printhide">
				<!--#include virtual="/REPORT/FormSelectFcbaOnly.html"-->
				<table id="form-table" class="centered-margin">
					<tr>
						<td>
							<font face="verdana" size="1">TANGGAL</font>
						</td>
						<td width="4">
						    	<input class="datepicker" type="text" maxlength="10" name="txtFddate" size="10" value="<%=v_txtFddate%>" title="Type Date mm/dd/yyyy" >
								</td>
						<td>
							<font face="verdana" size="1">TO</font>
						</td>
						<td width="4">
						 		<input class="datepicker" type="text" maxlength="10" name="txtFddateto" size="10" value="<%=v_txtFddateto%>" title="Type Date mm/dd/yyyy" >
						</td>
						<td><font face="verdana" size="1">SECTION NAME</font></td>
						<td><select name="txtsectionname" style="width:270px;">
						<option value=""></option>
						<%
						Dim rssectionname, strsectionname
						strsectionname	=	" select fccode, fcname from sectionname where fccode not in ('AFD-1','AFD-2','AFD-3','AFD-4','AFD-5','AFD-6','AFD-7','AFD-8','AFD-9') order by fccode asc"
						
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
						</select></td>
						
						<td>
						<font face="verdana" size="1">STATUS</font>
						</td>
						<td>
							<select name="txtStatus" class="chosen-select" multiple style="width:150px;">
								<%
									if not arr_indeks_eof then
										for xi=0 to ubound(arr_indeks2,2)
											if INSTR( cstr(""&v_txtStatus), cstr(""&arr_indeks2(0, xi)) ) <> 0 then
												response.write "<option value='"&arr_indeks2(0, xi)&"' selected>"&arr_indeks2(0, xi)&"</option>"
											else
												response.write "<option value='"&arr_indeks2(0, xi)&"'>"&arr_indeks2(0, xi)&"</option>"
											end if
										next
									end if
								%>
							</select> 
						</td>
						
						<td>
						<font face="verdana" size="1">Posting</font>
						</td>
						<td>
							<select name="txtposting" class="chosen-select" style="width:100px;">
								<% if v_txtposting="1" then %>
								<option value="1" Selected >Ya</option>
								<option value="0" >Tidak</option>
								<option value="" >All</option>
								<% end if%>
								<% if v_txtposting="0" then%>
								<option value="1" >Ya</option>
								<option value="0" Selected>Tidak</option>	
								<option value="" >All</option>
								<% end if%>
								<% if v_txtposting="" then%>
								<option value="1" >Ya</option>
								<option value="0" >Tidak</option>	
								<option value="" Selected >All</option>
								<% end if%>
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
				<th colspan="24" style="background-color:#FFFFFF; border:none;">
					<div id="header-wrapper">
					<div id="head_judul" style="margin-bottom:-30px; position:static; text-align:left;">
						<h4>PT. <%=varCompanyName%></h4>
						<h5><%=varFcbaName%></h5>
					</div>
					<div id="subheader" class="centered">
						<h2>MONITOR PAYROLL HARIAN</h2>
						<br/>
						<h4>
							PERIODE S/D &nbsp;:&nbsp;<%=textDate(v_txtFddate)%>  TO <%=textDate(v_txtFddateto)%>
						</h4>
						<%
							Dim rsqGang, qGang, txtqGang,txtAssistqGang,txtPostqGang
							qGang	=	" SELECT FCNAME, ASSISTANTNAME, POSITIONNAME FROM GANG WHERE FCBA = '" & v_txtFcba & "' AND FCCODE = '"& v_txtGangCode & "'"
							
							Set rsqGang = Server.CreateObject("ADODB.Recordset")
							rsqGang.Open qGang, connstr, adOpenStatic
							
							if not rsqGang.eof then
								txtqGang	=	rsqGang("FCNAME")
								txtAssistqGang	=	rsqGang("ASSISTANTNAME")
								txtPostqGang	=	rsqGang("POSITIONNAME")
							end if
							rsqGang.Close
							set rsqGang = Nothing
						%>
						<h4>
						<% If Len(Trim(v_txtsectionname)) > 0 Then %>
								Section Name : <%=PCase(v_txtsectionname)%>
							<%Else%>
								Section Name : All
							<% End If%>
						<% If Len(Trim(v_txtGangCode)) > 0 Then %>
						GANG : <%=v_txtGangCode%> - <%=txtqGang%>
						<%END IF%>
						<% If Len(Trim(v_txtStatus)) > 0  AND Len(Trim(v_txtGangCode)) > 0 Then %>
						& STATUS : <%=v_txtStatus%>
						<% elseIf Len(Trim(v_txtStatus)) > 0 AND Len(Trim(v_txtGangCode)) = 0 Then %>
						STATUS : <%=v_txtStatus%>
						<%END IF%>
						</h4>
						<br/>

					</div>
					</div>
				</th>
			</tr>
		<tr>
			<th rowspan="2" style="color:#fff;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Divisi&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</th>
			<th rowspan="2" style="color:#fff;">NO</th>
			<th rowspan="2" style="color:#fff;">NIK</th>
			<th rowspan="2" style="color:#fff;">Nama Karyawan</th>
			<th rowspan="2" style="color:#fff;">Gang</th>
			<th rowspan="2" style="color:#fff;">Status</th>
			<th colspan="9" style="color:#fff;">Hari ini</th>
			<th colspan="9" style="background-color:#F4F4A0;color:#000;" >S/D Hari ini</th>
		</tr>
		<tr>	
			<th style="color:#fff;">HK</th>
			<th style="color:#fff;">Gaji</th>
			<th style="color:#fff;">Lembur</th>
			<th style="color:#fff;">Premi</th>
			<th style="color:#fff;">Borongan</th>
			<th style="color:#fff;">Total Upah</th>
			<th style="color:#fff;">Denda</th>
			<th style="color:#fff;">Total Potongan</th>
			<th style="color:#fff;">Total</th>
			<th style="background-color:#F4F4A0;color:#000;">HK</th>
			<th style="background-color:#F4F4A0;color:#000;">Gaji</th>
			<th style="background-color:#F4F4A0;color:#000;">Lembur</th>
			<th style="background-color:#F4F4A0;color:#000;">Premi</th>
			<th style="background-color:#F4F4A0;color:#000;">Borongan</th>
			<th style="background-color:#F4F4A0;color:#000;">Total Upah</th>
			<th style="background-color:#F4F4A0;color:#000;">Denda</th>
			<th style="background-color:#F4F4A0;color:#000;">Total Potongan</th>
			<th style="background-color:#F4F4A0;color:#000;">Total</th>
		</tr>
		
	</thead>
	<tbody>
<%if not payrollOK then%>
		<tr>
			<td colspan="24" align="left" style="color:#c0392b;">Data gagal dimuat: <%=Server.HTMLEncode(errPayroll)%></td>
		</tr>
<%elseif rsPayroll.BOF and rsPayroll.EOF then%>
		<tr>
			<td colspan="24" align="left">NO DATA AVAILABLE</td>
		</tr>
<%end if
		if payrollOK then
		NO = 0
		rows = 1
		while not rsPayroll.eof
%>
		<tr class="<%=cClass%>">
			<% if rows = 1 then
					temprows = CDbl(0 + secCount("" & rsPayroll("sectionname")))
				end if 
			no = NO+1
			%>
			<%if rows = 1 then%>
				<td align="center" Rowspan="<%=0+temprows%>"><%=rsPayroll("sectionname")%></td> 
			<%end if%>
			<td align="center"><%=No%></td>
			<td align="left"><%=rsPayroll("EMPLOYEECODE")%></td>
			<td align="left"><%=rsPayroll("EMPLOYEENAME")%></td>
			<td align="center"><%=rsPayroll("GANGCODE")%></td>
			<td align="center"><%=rsPayroll("PRESENTGRADE")%></td>
			<td align="right"><%=formatnumber( cdbl(to_nol(rsPayroll("HK"))), 1)%></td>
			<td align="right"><%=formatnumber( cdbl(to_nol(rsPayroll("GAJI"))),0)%></td>
			<td align="right"><%=formatnumber( cdbl(to_nol(rsPayroll("LEMBUR"))),0)%></td>
			<td align="right"><%=formatnumber( cdbl(to_nol(rsPayroll("PREMI"))),0)%></td>
			<td align="right"><%=formatnumber( cdbl(to_nol(rsPayroll("BORONGAN"))),0)%></td>
			<td align="right"><%=formatnumber( cdbl(to_nol(rsPayroll("SubGaji"))),0)%></td>
			<td align="right"><%=formatnumber( cdbl(to_nol(rsPayroll("POT_DENDA"))),0)%></td>
			<td align="right"><%=formatnumber( cdbl(to_nol(rsPayroll("Sub_Potongan"))),0)%></td>
			<td align="right"><%=formatnumber( cdbl(to_nol(rsPayroll("Total"))),0)%></td>
		
		    <td align="right"><%=formatnumber( cdbl(to_nol(rsPayroll("HK_shi"))), 1)%></td>
			<td align="right"><%=formatnumber( cdbl(to_nol(rsPayroll("GAJI_shi"))),0)%></td>
			<td align="right"><%=formatnumber( cdbl(to_nol(rsPayroll("LEMBUR_shi"))),0)%></td>
			<td align="right"><%=formatnumber( cdbl(to_nol(rsPayroll("PREMI_shi"))),0)%></td>
			<td align="right"><%=formatnumber( cdbl(to_nol(rsPayroll("BORONGAN_shi"))),0)%></td>
			<td align="right"><%=formatnumber( cdbl(to_nol(rsPayroll("SubGaji_shi"))),0)%></td>
			<td align="right"><%=formatnumber( cdbl(to_nol(rsPayroll("POT_DENDA_shi"))),0)%></td>
			<td align="right"><%=formatnumber( cdbl(to_nol(rsPayroll("Sub_Potongan_shi"))),0)%></td>
			<td align="right"><%=formatnumber( cdbl(to_nol(rsPayroll("Total_shi"))),0)%></td>
		</tr>
<%	if rows = temprows then
			noline			= noline		+ 1
			nGet_Hk 		= nGet_Hk 		+ cdbl(to_nol(rsPayroll("HK")))
			nBase_Salary 	= nBase_Salary 	+ cdbl(to_nol(rsPayroll("GAJI")))
			nOvertime 		= nOvertime 	+ cdbl(to_nol(rsPayroll("LEMBUR")))
			nPremi 		    = nPremi 	    + cdbl(to_nol(rsPayroll("PREMI")))
			nBorongan 	    = nBorongan 	+ cdbl(to_nol(rsPayroll("BORONGAN")))
			nTotal 			= nTotal		+ cdbl(to_nol(rsPayroll("SubGaji")))
			nDenda 		    = nDenda	    + cdbl(to_nol(rsPayroll("POT_DENDA")))
			nTotalDenda 	= nTotalDenda 	+ cdbl(to_nol(rsPayroll("Sub_Potongan")))
			nTotalgaji 		= nTotalgaji 	+ cdbl(to_nol(rsPayroll("Total")))
			
			nGet_Hk_shi 		= nGet_Hk_shi 		+ cdbl(to_nol(rsPayroll("HK_shi")))
			nBase_Salary_shi 	= nBase_Salary_shi 	+ cdbl(to_nol(rsPayroll("GAJI_shi")))
			nOvertime_shi 		= nOvertime_shi 	+ cdbl(to_nol(rsPayroll("LEMBUR_shi")))
			nPremi_shi 		    = nPremi_shi 	    + cdbl(to_nol(rsPayroll("PREMI_shi")))
			nBorongan_shi 	    = nBorongan_shi 	+ cdbl(to_nol(rsPayroll("BORONGAN_shi")))
			nTotal_shi 			= nTotal_shi		+ cdbl(to_nol(rsPayroll("SubGaji_shi")))
			nDenda_shi 		    = nDenda_shi	    + cdbl(to_nol(rsPayroll("POT_DENDA_shi")))
			nTotalDenda_shi 	= nTotalDenda_shi 	+ cdbl(to_nol(rsPayroll("Sub_Potongan_shi")))
			nTotalgaji_shi 		= nTotalgaji_shi 	+ cdbl(to_nol(rsPayroll("Total_shi")))
			
			rows = 1
			
	%>
		<tr style="background-color:#F4F4A0;">	 
			<td align="center" colspan="6" ><strong>Sub Total</strong></td>
			<td align="right"><strong><%=formatnumber(nGet_Hk,1)%></strong></td>
			<td align="right"><strong><%=formatnumber(nBase_Salary,0)%></strong></td>
			<td align="right"><strong><%=formatnumber(nOvertime,0)%></strong></td>
			<td align="right"><strong><%=formatnumber(nPremi,0)%></strong></td> 
			<td align="right"><strong><%=formatnumber(nBorongan,0)%></strong></td> 
			<td align="right"><strong><%=formatnumber(nTotal,0)%></strong></td>
			<td align="right"><strong><%=formatnumber(nDenda,0)%></strong></td>
			<td align="right"><strong><%=formatnumber(nTotalDenda,0)%></strong></td>
			<td align="right"><strong><%=formatnumber(nTotalgaji,0)%></strong></td>
			
			<td align="right"><strong><%=formatnumber(nGet_Hk_shi,1)%></strong></td>
			<td align="right"><strong><%=formatnumber(nBase_Salary_shi,0)%></strong></td>
			<td align="right"><strong><%=formatnumber(nOvertime_shi,0)%></strong></td>
			<td align="right"><strong><%=formatnumber(nPremi_shi,0)%></strong></td> 
			<td align="right"><strong><%=formatnumber(nBorongan_shi,0)%></strong></td> 
			<td align="right"><strong><%=formatnumber(nTotal_shi,0)%></strong></td>
			<td align="right"><strong><%=formatnumber(nDenda_shi,0)%></strong></td>
			<td align="right"><strong><%=formatnumber(nTotalDenda_shi,0)%></strong></td>
			<td align="right"><strong><%=formatnumber(nTotalgaji_shi,0)%></strong></td>
			
		</tr>
		<tr>
			<td align="center" colspan="<%= 24+ nSpanTotal %>" ></td>
		</tr>
		<%	
			noline = 1
			
			
			nGet_Hk = cdbl("0")
			nBase_Salary = cdbl("0")
			nOvertime = cdbl("0")
			nPremi = cdbl("0")
			nBorongan = cdbl("0")
			nTotal = cdbl("0")
			nDenda = cdbl("0")
			nTotalDenda = cdbl("0")
			nTotalgaji = cdbl("0")
			
			nGet_Hk_shi = cdbl("0")
			nBase_Salary_shi = cdbl("0")
			nOvertime_shi = cdbl("0")
			nPremi_shi = cdbl("0")
			nBorongan_shi = cdbl("0")
			nTotal_shi= cdbl("0")
			nDenda_shi= cdbl("0")
			nTotalDenda_shi= cdbl("0")
			nTotalgaji_shi= cdbl("0")
			 
		else
		' cari sub total
			noline			= noline		+ 1
			nGet_Hk 		= nGet_Hk 		+ cdbl(to_nol(rsPayroll("HK")))
			nBase_Salary 	= nBase_Salary 	+ cdbl(to_nol(rsPayroll("GAJI")))
			nOvertime 		= nOvertime 	+ cdbl(to_nol(rsPayroll("LEMBUR")))
			nPremi 		    = nPremi 	    + cdbl(to_nol(rsPayroll("PREMI")))
			nBorongan 	    = nBorongan 	+ cdbl(to_nol(rsPayroll("BORONGAN")))
			nTotal 			= nTotal		+ cdbl(to_nol(rsPayroll("SubGaji")))
			nDenda 		    = nDenda	    + cdbl(to_nol(rsPayroll("POT_DENDA")))
			nTotalDenda 	= nTotalDenda 	+ cdbl(to_nol(rsPayroll("Sub_Potongan")))
			nTotalgaji 		= nTotalgaji 	+ cdbl(to_nol(rsPayroll("Total")))
			nGet_Hk_shi 		= nGet_Hk_shi 		+ cdbl(to_nol(rsPayroll("HK_shi")))
			nBase_Salary_shi 	= nBase_Salary_shi 	+ cdbl(to_nol(rsPayroll("GAJI_shi")))
			nOvertime_shi 		= nOvertime_shi 	+ cdbl(to_nol(rsPayroll("LEMBUR_shi")))
			nPremi_shi 		    = nPremi_shi 	    + cdbl(to_nol(rsPayroll("PREMI_shi")))
			nBorongan_shi 	    = nBorongan_shi 	+ cdbl(to_nol(rsPayroll("BORONGAN_shi")))
			nTotal_shi 			= nTotal_shi		+ cdbl(to_nol(rsPayroll("SubGaji_shi")))
			nDenda_shi 		    = nDenda_shi	    + cdbl(to_nol(rsPayroll("POT_DENDA_shi")))
			nTotalDenda_shi 	= nTotalDenda_shi 	+ cdbl(to_nol(rsPayroll("Sub_Potongan_shi")))
			nTotalgaji_shi 		= nTotalgaji_shi 	+ cdbl(to_nol(rsPayroll("Total_shi")))
			rows = rows + 1
		End if
							  
			' cari grand total 
				gt_GET_HK = cdbl(gt_GET_HK) + cdbl(to_nol(rsPayroll("HK")))
				gt_BASE_SALARY = cdbl(gt_BASE_SALARY) + cdbl(to_nol(rsPayroll("GAJI")))
				gt_OVERTIME = cdbl(gt_OVERTIME) + cdbl(to_nol(rsPayroll("LEMBUR"))) 
				gt_PREMI = cdbl(gt_PREMI) + cdbl(to_nol(rsPayroll("PREMI"))) 
				gt_BORONGAN = cdbl(gt_BORONGAN) + cdbl(to_nol(rsPayroll("BORONGAN")))
				gt_TOTAL = cdbl(gt_TOTAL) + cdbl(to_nol(rsPayroll("SubGaji"))) 
				gt_DENDA = cdbl(gt_DENDA) + cdbl(to_nol(rsPayroll("POT_DENDA"))) 
				gt_TOTALDENDA = cdbl(gt_TOTALDENDA) + cdbl(to_nol(rsPayroll("Sub_Potongan")))  
				gt_TOTALGAJI = cdbl(gt_TOTALGAJI) + cdbl(to_nol(rsPayroll("Total")))

				gt_GET_HK_shi = cdbl(gt_GET_HK_shi) + cdbl(to_nol(rsPayroll("HK_shi")))
				gt_BASE_SALARY_shi = cdbl(gt_BASE_SALARY_shi) + cdbl(to_nol(rsPayroll("GAJI_shi")))
				gt_OVERTIME_shi = cdbl(gt_OVERTIME_shi) + cdbl(to_nol(rsPayroll("LEMBUR_shi"))) 
				gt_PREMI_shi = cdbl(gt_PREMI_shi) + cdbl(to_nol(rsPayroll("PREMI_shi"))) 
				gt_BORONGAN_shi = cdbl(gt_BORONGAN_shi) + cdbl(to_nol(rsPayroll("BORONGAN_shi")))
				gt_TOTAL_shi = cdbl(gt_TOTAL_shi) + cdbl(to_nol(rsPayroll("SubGaji_shi"))) 
				gt_DENDA_shi = cdbl(gt_DENDA_shi) + cdbl(to_nol(rsPayroll("POT_DENDA_shi"))) 
				gt_TOTALDENDA_shi = cdbl(gt_TOTALDENDA_shi) + cdbl(to_nol(rsPayroll("Sub_Potongan_shi")))  
				gt_TOTALGAJI_shi = cdbl(gt_TOTALGAJI_shi) + cdbl(to_nol(rsPayroll("Total_shi")))
				
			rsPayroll.MoveNext			
			Wend
		rsPayroll.Close
		end if
		set rsPayroll = Nothing
	%> 
	<tr style="background-color:#27AE60;">	 
		<td align="center" colspan="6" ><strong>Grand Total</strong></td>
		<td align="right"><strong><%=formatnumber(gt_GET_HK,1)%></strong></td>  
		<td align="right"><strong><%=formatnumber(gt_BASE_SALARY,0)%></strong></td> 
		<td align="right"><strong><%=formatnumber(gt_OVERTIME,0)%></strong></td> 
		<td align="right"><strong><%=formatnumber(gt_PREMI,0)%></strong></td>
		<td align="right"><strong><%=formatnumber(gt_BORONGAN,0)%></strong></td>
		<td align="right"><strong><%=formatnumber(gt_TOTAL,0)%></strong></td>
		<td align="right"><strong><%=formatnumber(gt_DENDA,0)%></strong></td>
		<td align="right"><strong><%=formatnumber(gt_TOTALDENDA,0)%></strong></td>
		<td align="right"><strong><%=formatnumber(gt_TOTALGAJI,0)%></strong></td>
		
		<td align="right"><strong><%=formatnumber(gt_GET_HK_shi,1)%></strong></td>  
		<td align="right"><strong><%=formatnumber(gt_BASE_SALARY_shi,0)%></strong></td> 
		<td align="right"><strong><%=formatnumber(gt_OVERTIME_shi,0)%></strong></td> 
		<td align="right"><strong><%=formatnumber(gt_PREMI_shi,0)%></strong></td>
		<td align="right"><strong><%=formatnumber(gt_BORONGAN_shi,0)%></strong></td>
		<td align="right"><strong><%=formatnumber(gt_TOTAL_shi,0)%></strong></td>
		<td align="right"><strong><%=formatnumber(gt_DENDA_shi,0)%></strong></td>
		<td align="right"><strong><%=formatnumber(gt_TOTALDENDA_shi,0)%></strong></td>
		<td align="right"><strong><%=formatnumber(gt_TOTALGAJI_shi,0)%></strong></td>
	</tr>
	
	</tbody>
</table>
<div id="footer_note">	
<div class="block">
</div>
</div>
 <div class="clear"></div>
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