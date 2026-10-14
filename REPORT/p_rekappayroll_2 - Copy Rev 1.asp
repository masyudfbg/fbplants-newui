<% ValidateSession %>
<!--#include virtual="/ConndB1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->

<%
Dim v_txtMonth, v_txtYear
Dim rsDefaultBa, strDefaultBa, varDefaultBa
Dim rscurrentdate, strcurrentdate, varDefaultYear, varDefaultMonth
Dim ScriptName, v_txtsectionname, v_txtGangCode, v_txtStatus
Dim rsBaInfo, strBaInfo, varFcbaName, varCompanyName, varManagerName, varKTUName, no

v_txtFcba = Request("txtFcba")
v_txtYear		= Request("txtYear")
v_txtMonth		= Request("txtMonth")
v_txtsectionname = Request("txtsectionname")
v_txtGangCode	= Request("txtGangCode")
v_txtStatus		= Request("txtStatus")
v_Emptype		= Request("EMP_TYPE")
ScriptName		= Request.ServerVariables("SCRIPT_NAME")

Dim v_paymenttype : v_paymenttype = Request("paymenttype")

if len(trim(v_Emptype)) > 0 then 
	arr_indeks_eof = true
	oSql =  " SELECT '--ALL--' as EMP_TYPE FROM DUAL UNION SELECT DISTINCT FCCODE as EMP_TYPE FROM PARAMETER WHERE PARCODE = 'E0103' AND FCBA = '"& v_txtFcba &"' ORDER BY EMP_TYPE ASC"
	set oRs = server.CreateObject("Adodb.recordset")
	oRs.Open oSql, connstr
	if Not oRs.Eof then
		arr_indeks = oRs.getrows()
		arr_indeks_eof = false
	End if
	oRs.Close
	Set oRs = Nothing
else
	arr_indeks_eof = true
	oSql =  " SELECT '--ALL--' as EMP_TYPE FROM DUAL UNION SELECT DISTINCT FCCODE as EMP_TYPE FROM PARAMETER WHERE PARCODE = 'E0103' AND FCBA = '"& Session("fcba") &"' ORDER BY EMP_TYPE ASC"
	set oRs = server.CreateObject("Adodb.recordset")
	oRs.Open oSql, connstr
	if Not oRs.Eof then
		arr_indeks = oRs.getrows()
		arr_indeks_eof = false
	End if
	oRs.Close
	Set oRs = Nothing
end if

strDefaultBa = " select" & _
               "     fccode as fcba" & _
               " from businessunit bu" & _
               " inner join user_fcba_tab uft on (bu.fccode = uft.fcba) " & _
               " where isdefault = 'TRUE'" & _
               "       and fcusercode = '" & Session("CodeName") & "'"

Set rsDefaultBa = Server.CreateObject("ADODB.Recordset")
rsDefaultBa.Open strDefaultBa, connstr, adOpenStatic

If Not rsDefaultBa.EOF Then
	varDefaultBa = rsDefaultBa("fcba")
Else
	varDefaultBa = Session("fcba")
End If
rsDefaultBa.Close
set rsDefaultBa = Nothing

If (v_txtFcba = "" Or Len(Trim(v_txtFcba)) = 0) Then
	v_txtFcba = varDefaultBa
End If

strcurrentdate = " select currentdate" & _
                 " from periodcontrol" & _                 
                 " where fcba = '" & v_txtFcba & "'"

Set rscurrentdate = Server.CreateObject("ADODB.Recordset")
rscurrentdate.Open strcurrentdate, connstr, adOpenStatic

If Not rscurrentdate.EOF Then
	varDefaultYear	=	Year(rscurrentdate("currentdate"))
	varDefaultMonth	=	Month(rscurrentdate("currentdate"))
Else
	varDefaultYear	=	Year(Session("v_Employee"))
	varDefaultMonth	=	Month(Session("v_Employee"))
End If
rscurrentdate.Close
set rscurrentdate = Nothing

If (v_txtYear = "" Or Len(Trim(v_txtYear)) = 0) Then v_txtYear = varDefaultYear
If (v_txtMonth = "" Or Len(Trim(v_txtMonth)) = 0) Then v_txtMonth = varDefaultMonth


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

' Rev Emp Type By SAG 7 Agustus 2019
	if v_Emptype <>  "--ALL--" and v_Emptype <> "" then
		whereCons = whereCons & "AND INSTR('"& v_Emptype &"', SUBSTR(pr.fccode,1,12) ) <> 0"
	else
		nSpanTotal = nSpanTotal - 1  
	end if
' End Rev Emp Type By SAG 7 Agustus 2019
			
			
%>
<html>
<head>
	<title>ePlaS Reporting - Payroll Karyawan</title>
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
			<input type="button" value="&#xf02f;" onClick="window.print()" style="font-family: 'FontAwesome';margin-top:-10px;">
				<!--#include file="../print_excel.asp"-->
				<!--#include file="menutop.asp"-->
			</div>
			<form name="frmAR" method="post" action="<%=ScriptName%>">
				<div class="printhide">
				<!--#include virtual="/REPORT/FormSelectFcbaOnly.html"-->
				<table id="form-table" class="centered-margin">
					<tr>
						<td>
							<font face="verdana" size="1">TAHUN</font>
						</td>
						<% 
							Dim mon, Yr, MinYr, MaxYr, i, MM
							mon = varDefaultMonth
							If len(v_txtMonth)=0 then v_txtMonth=mon
						
							Yr = varDefaultYear
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
						<td><font face="verdana" size="1">SECTION NAME</font></td>
						<td><select name="txtsectionname" ><option value=""></option>
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
							<font face="verdana" size="1">EMPLOYEE TYPE</font>
						</td>
						<td>
							<select name="EMP_TYPE" class="chosen-select" multiple style="width:270px;">
								<%
									if not arr_indeks_eof then
										for xi=0 to ubound(arr_indeks,2)
											if INSTR( cstr(""&v_Emptype), cstr(arr_indeks(0, xi)) ) <> 0 then
												response.write "<option value='"&arr_indeks(0, xi)&"' selected>"&arr_indeks(0, xi)&"</option>"
											else
												response.write "<option value='"&arr_indeks(0, xi)&"'>"&arr_indeks(0, xi)&"</option>"
											end if
										next
									end if
								%>
							</select>
						</td>
						<td>
							<font face="verdana" size="1">STATUS</font>
						</td>
						<td>
							<select name="txtStatus">
								<option value = "" <%If Len(Trim(v_txtStatus)) = 0 Then %> Selected <%End If%> >--ALL--</option>
								<option value = "KHL" <%If v_txtStatus = "KHL" Then %> Selected <%End If%> >KHL</option>
								<option value = "KTK" <%If v_txtStatus = "KTK" Then %> Selected <%End If%> >KTK</option>
								<option value = "PBI" <%If v_txtStatus = "PBI" Then %> Selected <%End If%> >PBI</option>
								<option value = "KPB" <%If v_txtStatus = "KPB" Then %> Selected <%End If%> >KPB</option>
								<option value = "STF" <%If v_txtStatus = "STF" Then %> Selected <%End If%> >STF</option>
								<option value = "SKU" <%If v_txtStatus = "SKU" Then %> Selected <%End If%> >SKU</option>
							</select>
						</td>
						
						<td width="57"  align="left" ><font face="Verdana" size="1">Payment Type</font></td>
			
						<% 
						Dim rsPayment,q
						q = "select P.FCCODE,P.FCNAME from  PAYMENTTYPE P "
						Set rsPayment = Server.CreateObject("ADODB.Recordset")
						rsPayment.CursorType = 3
						rsPayment.CursorLocation = 3
						rsPayment.Open q, connstr
						if not rsPayment.eof then%>
						<td align="left"> 
							<select name="paymenttype" title="PAYMENTTYPE Transfer Or Non Transfer">         
								<option value=""></option> 
								<%
								Do While Not rsPayment.eof

								if rsPayment("FCCODE")=Request("PAYMENTTYPE") Then %>
								<option value="<%=rsPayment("FCCODE")%>" selected><%=rsPayment("FCNAME")%></option> 
								<%Else%>			
								<option value="<%=rsPayment("FCCODE")%>"><%=rsPayment("FCNAME")%></option> 
								<%End if%>
								<%
								rsPayment.movenext
								loop
								End if
								rsPayment.Close
								set rsPayment = Nothing
								q=""
								%>
							</select> 
							<font color="#FF0000"></font>
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
				<th colspan="22" style="background-color:#FFFFFF; border:none;">
					<div id="header-wrapper">
					<div id="head_judul" style="margin-bottom:-30px; position:static; text-align:left;">
						<h4>PT. <%=varCompanyName%></h4>
						<h5><%=varFcbaName%></h5>
					</div>
					<div id="subheader" class="centered">
						<h2>DAFTAR UPAH KARYAWAN</h2>
						<br/>
						<h4>
							PERIODE&nbsp;:&nbsp;<%=Bulan(v_txtMonth)%>&nbsp;<%=v_txtYear%>
							<% If Len(Trim(Request("PAYMENTTYPE"))) > 0 Then %>
								- Payment Type : <%=PCase(Request("PAYMENTTYPE"))%>
							<%Else%>
								- Payment Type : All
							<% End If%>
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
			<th rowspan="2" style="color:#fff;">NO</th>
			<th rowspan="2" style="color:#fff;">NIK</th>
			<th rowspan="2" style="color:#fff;">Nama Karyawan</th>
			<th rowspan="2" style="color:#fff;">Jenis Kelamin</th>
			<th rowspan="2" style="color:#fff;">Status</th>
			<th rowspan="2" style="color:#fff;">Golongan</th>
			<th rowspan="2" style="color:#fff;">Tipe</th>
			<th rowspan="2" style="color:#fff;">HK</th>
			<th rowspan="2" style="color:#fff;">Gaji</th>
			<th rowspan="2" style="color:#fff;">Lembur</th>
			<th rowspan="2" style="color:#fff;">Premi</th>
			<th rowspan="2" style="color:#fff;">Borongan</th>
			<th rowspan="2" style="color:#fff;">Total Upah</th>
			<th colspan="6" style="color:#fff;">Potongan</th>
			<th rowspan="2" style="color:#fff;">Sisa Upah</th>
			<th rowspan="2" colspan="6" style="color:#fff;" width="500">Tanda Tangan</th>
		</tr>
		<tr>
			<th style="color:#fff;">BPJS Tenaga Kerja</th>
			<th style="color:#fff;">BPJS Kesehatan</th>
			<th style="color:#fff;">PPh21</th>
			<th style="color:#fff;">Pinjaman</th>
			<th style="color:#fff;">Denda</th>	
			<th style="color:#fff;">Total Potongan</th>	
		</tr>
	</thead>
	<tbody>
	<%
	Dim rsPayroll, strPayroll, noline 
	Dim nGet_Hk, nGet_Hk_kj, nget_mkp1, nBase_Salary, nrapelgaji, nOvertime, nPremi, nTunjangan, nBorongan, nP_Jamsostek, nP_bpjskes, nP_Mk_P1, nP_Panen, nP_Pph21, nP_koperasi, nP_Lain, nTotal
			strPayroll ="  SELECT *" & _
						"    FROM (SELECT year," & _
						"                 month," & _
						"                 fcba," & _
						"                 fccode," & _
						"                 emp_type," & _
						"                 gender," & _
						"                 subgrade," & _
						"                 taxstatus," & _
						"                 employee_hist_api.get_fcname (year," & _
						"                                               month," & _
						"                                               fcba," & _
						"                                               fccode)" & _
						"                    fcname," & _
						"                 employee_hist_api.get_gangcode (year," & _
						"                                                 month," & _
						"                                                 fcba," & _
						"                                                 fccode)" & _
						"                    gangcode," & _
						"                 employee_api.get_presentgrade2 (fcba," & _
						"                                                     fccode)" & _
						"                    presentgrade2," & _
						"                 employee_hist_api.get_sectionname (year," & _
						"                                                    month," & _
						"                                                    fcba," & _
						"                                                    fccode)" & _
						"                    sectionname," & _
						"                 employee_hist_api.get_paymenttype (year," & _
						"                                                    month," & _
						"                                                    fcba," & _
						"                                                    fccode)" & _
						"                    paymenttype," & _
						"                 NVL (get_hk, 0) get_hk," & _
						"                 NVL (get_hk_kj, 0) get_hk_kj," & _
						"                 NVL (get_mkp1, 0) get_mkp1," & _
						"                 NVL (base_salary, 0) base_salary," & _
						"                 NVL (overtime, 0) overtime," & _
						"                 NVL (borongan, 0) borongan," & _
						"                 NVL (p_jamsostek, 0) p_jamsostek," & _
						"                 NVL (p_bpjskes, 0) p_bpjskes," & _
						"                 NVL (p_pph21, 0) p_pph21," & _
						"                 NVL (premi, 0) premi," & _
						"                 NVL (totalgaji, 0) totalgaji," & _
						"                 NVL (denda, 0) denda," & _
						"                 NVL (total, 0) total," & _
						"                 NVL (p_pinjaman, 0) pinjaman" & _
						"            FROM (  SELECT TO_NUMBER (SUBSTR (t.yearmonth, 1, 4)) year," & _
						"                           TO_NUMBER (SUBSTR (t.yearmonth, 5, 2)) month," & _
						"                           t.fcba," & _
						"                           t.fccode," & _
						"                           pr.fccode as emp_type," & _
						"						 	emp.gender," & _
						"						 	emp.subgrade," & _
						"						 	emp.taxstatus," & _
						"                           SUM (get_hk) get_hk," & _
						"                           SUM (get_hk_kj) get_hk_kj," & _
						"                           SUM (get_mkp1) get_mkp1," & _
						"                           SUM (t.base_salary) base_salary," & _
						"                           SUM (t.rapelgaji)" & _
						"                           + SUM (t.tunjangan)" & _
						"                           + SUM (t.premi)  premi," & _
						"                           SUM (t.overtime) overtime," & _
						"                           SUM (t.borongan) borongan," & _
						"                           SUM (t.p_jamsostek) p_jamsostek," & _
						"                           SUM (t.p_bpjskes) p_bpjskes," & _
						"                           SUM (t.p_pph21) p_pph21," & _
						"                           SUM (t.p_pinjaman) p_pinjaman," & _
						"                           SUM (t.p_panen) " & _
						"                           + SUM (t.p_mk_p1)" & _
						"                           + SUM (t.p_koperasi)" & _
						"                           + SUM (t.p_lain) denda," & _
						"                             SUM (t.base_salary)" & _
						"                           + SUM (t.overtime)" & _
						"                           + SUM (t.borongan)" & _
						"                           + SUM (t.premi) totalgaji," & _
						"                             SUM (t.base_salary)" & _
						"                           + SUM (t.overtime)" & _
						"                           + SUM (t.premi)" & _
						"                           + SUM (t.rapelgaji)" & _
						"                           + SUM (t.tunjangan)" & _
						"                           + SUM (t.borongan)" & _
						"                           + SUM (t.p_jamsostek)" & _
						"                           + SUM (t.p_bpjskes)" & _
						"                           + SUM (t.p_mk_p1)" & _
						"                           + SUM (t.p_panen)" & _
						"                           + SUM (t.p_pph21)" & _
						"                           + SUM (t.p_koperasi)" & _
						"                           + SUM (t.p_lain) total" & _
						"                      FROM view_payrolldetail t" & _
						"                      INNER JOIN EMPLOYEE emp ON (emp.fccode = t.fccode and emp.fcba = t.fcba)" & _
						"						INNER JOIN PARAMETER pr ON (pr.fcba = emp.fcba and pr.fccode = emp.emp_type and parcode = 'E0103')" & _
						"                     WHERE     TO_NUMBER (SUBSTR (t.yearmonth, 1, 4)) = " & v_txtYear & "" & _
						"                           AND TO_NUMBER (SUBSTR (t.yearmonth, 5, 2)) = " & v_txtMonth & "" & _
						"                     AND t.fcba = '" & v_txtFcba & "'" & whereCons & "" & _
						"                  GROUP BY TO_NUMBER (SUBSTR (t.yearmonth, 1, 4))," & _
						"                           TO_NUMBER (SUBSTR (t.yearmonth, 5, 2))," & _
						"                           t.fcba," & _
						"                           t.fccode," & _
						"						 	emp.gender," & _
						"						 	emp.subgrade," & _
						"						 	pr.fccode," & _
						"						 	emp.taxstatus," & _
						"                           emp.emp_type))" & _
						"   	WHERE fcba ='" & v_txtFcba & "' AND totalgaji >  0"
								
			if  len(trim(v_txtsectionname)) > 0 then
				strPayroll = strPayroll & " and sectionname = '" & v_txtsectionname & "'"
			end if
			if  len(trim(v_txtGangCode)) > 0 then
				strPayroll = strPayroll & " and gangcode = '" & v_txtGangCode & "'"
			end if
			if  len(trim(v_txtStatus)) > 0 then
				strPayroll = strPayroll & " and presentgrade2 = '" & v_txtStatus & "'"
			end if
			if len(trim(request("PAYMENTTYPE"))) > 0 then
				strPayroll = strPayroll & " and paymenttype = '" & v_paymenttype & "'" 
			end if
			
				strPayroll = strPayroll & " order by PRESENTGRADE2, EMP_TYPE, fccode" 
			' response.write strPayroll
			Set rsPayroll = Server.CreateObject("ADODB.Recordset")
			rsPayroll.Open strPayroll, connstr, adOpenStatic
    

	if Not rsPayroll.eof then
		noline = 1
		nGet_Hk = 0
		nGet_Hk_kj = 0
		nBase_Salary = 0
		nOvertime = 0
		nPremi = 0
		nTunjangan = 0
		nBorongan = 0
		nP_Jamsostek = 0
		nP_Mk_P1 = 0
		nP_Panen = 0
		nP_Pph21 = 0
		nP_Lain = 0
		nTotal = 0
		nDenda = 0
		nTotalgaji = 0
		nPinjaman = 0
		nTotalPotongan = 0
		no = 0

		while Not rsPayroll.eof  
		no = no + 1 
		totalpotongan = cdbl(to_nol(rsPayroll("p_jamsostek"))) + cdbl(to_nol(rsPayroll("p_bpjskes"))) + cdbl(to_nol(rsPayroll("p_pph21"))) + cdbl(to_nol(rsPayroll("pinjaman"))) + cdbl(to_nol(rsPayroll("denda"))) 
		totalsisaupah = cdbl(to_nol(rsPayroll("totalgaji"))) - (-1 * cdbl(to_nol(totalpotongan)) )
		%>
		           
		  <tr>
			<td align="center"><%=NO%></td>
			<td align="left"><%=rsPayroll("FCCODE")%></td>
			<td align="left"><%=rsPayroll("FCNAME")%></td>
			<td align="center"><%=rsPayroll("GENDER")%></td>
			<td align="center"><%=rsPayroll("taxstatus")%></td>
			<td align="center"><%=rsPayroll("PRESENTGRADE2")%></td>
			<td align="center"><%=rsPayroll("EMP_TYPE")%></td>
			<td align="right"><%=formatnumber( cdbl(rsPayroll("GET_HK")), 1)%></td>
			<td align="right"><%=formatnumber( cdbl(rsPayroll("BASE_SALARY")),0)%></td>
			<td align="right"><%=formatnumber( cdbl(rsPayroll("OVERTIME")),0)%></td>
			<td align="right"><%=formatnumber( cdbl(rsPayroll("PREMI")),0)%></td>
			<td align="right"><%=formatnumber( cdbl(rsPayroll("BORONGAN")),0)%></td>
			<td align="right"><%=formatnumber( cdbl(rsPayroll("TOTALGAJI")),0)%></td>
			<td align="right"><%=formatnumber( cdbl(rsPayroll("P_JAMSOSTEK")),0)%></td>
			<td align="right"><%=formatnumber( cdbl(rsPayroll("P_BPJSKES")),0)%></td>
			<td align="right"><%=formatnumber( cdbl(rsPayroll("P_PPH21")),0)%></td>
			<td align="right"><%=formatnumber( cdbl(rsPayroll("PINJAMAN")),0)%></td>
			<td align="right"><%=formatnumber( cdbl(rsPayroll("DENDA")),0)%></td>
			<td align="right"><%=formatnumber( cdbl(to_nol(totalpotongan)),0)%></td>
			<td align="right"><%=formatnumber( cdbl(totalsisaupah),0)%></td>
		<%if noline  mod 2 = 1 then%>
			<td align="left" style="background-color:#ccc;" width="500" colspan="3"><font face="Verdana" size="1">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>			
			<td align="left" style="background-color:#fff;" width="500" colspan="3"><font face="Verdana" size="1"><%=NO%>.&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>			
		<%Else%>
			<td align="left" style="background-color:#fff;" width="500" colspan="3"><font face="Verdana" size="1"><%=NO%>.&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>			
			<td align="left" style="background-color:#ccc;" width="500" colspan="3"><font face="Verdana" size="1">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font></td>			
		<%End if%>
		  </tr>	
		      
    <%
			noline			= noline		+ 1
			nGet_Hk 		= nGet_Hk 		+ cdbl(rsPayroll("GET_HK"))
			nGet_Hk_kj 		= nGet_Hk_kj 	+ cdbl(rsPayroll("GET_HK_KJ"))
			nget_mkp1 		= nget_mkp1		+ cdbl(rsPayroll("get_mkp1"))
			nBase_Salary 	= nBase_Salary 	+ cdbl(rsPayroll("BASE_SALARY"))
			' nrapelgaji		= nrapelgaji 	+ cdbl(rsPayroll("rapelgaji"))
			nOvertime 		= nOvertime 	+ cdbl(rsPayroll("OVERTIME"))
			' nTunjangan 		= nTunjangan 	+ cdbl(rsPayroll("TUNJANGAN"))
			nBorongan 		= nBorongan 	+ cdbl(rsPayroll("BORONGAN"))
			nP_Jamsostek 	= nP_Jamsostek 	+ cdbl(rsPayroll("P_JAMSOSTEK"))
			nP_bpjskes = nP_bpjskes + cdbl(rsPayroll("P_BPJSKES"))
			' nP_Mk_P1 		= nP_Mk_P1 		+ cdbl(rsPayroll("P_MK_P1"))
			' nP_Panen 		= nP_Panen 		+ cdbl(rsPayroll("P_PANEN"))
			' nP_Pph21 		= nP_Pph21 		+ cdbl(rsPayroll("P_PPH21"))
			' nP_koperasi = nP_koperasi + cdbl(rsPayroll("P_KOPERASI"))
			' nP_Lain 		= nP_Lain 		+ cdbl(rsPayroll("P_LAIN"))
			nTotal 			= nTotal 		+ cdbl(rsPayroll("TOTAL"))
			nDenda 			= nDenda		+ cdbl(rsPayroll("DENDA"))
			nTotalgaji 		= nTotalgaji		+ cdbl(rsPayroll("TOTALGAJI"))
			nPremi 			= nPremi 		+ cdbl(rsPayroll("PREMI"))
			nPinjaman 		= nPinjaman 	+ cdbl(rsPayroll("PINJAMAN"))
			nTotalPotongan 		= nTotalPotongan 	+ cdbl(totalpotongan)

			rsPayroll.MoveNext
		wend
	Else
	%>
	<tr>
		<td Colspan = "22">NO DATA AVAILABLE</td>
	</tr>
	<%End If%>
	<%
	rsPayroll.Close
	set rsPayroll = Nothing
	%>
	<tr style="background-color:#F4F4A0;">	 
		<td align="center" colspan="7" ><strong>Total</strong></td>
		<td align="right"><strong><%=formatnumber(nGet_Hk,1)%></strong></td>
		<td align="right"><strong><%=formatnumber(nBase_Salary,0)%></strong></td>
		<td align="right"><strong><%=formatnumber(nOvertime,0)%></strong></td>
		<td align="right"><strong><%=formatnumber(nPremi,0)%></strong></td> 
		<td align="right"><strong><%=formatnumber(nBorongan,0)%></strong></td> 
		<td align="right"><strong><%=formatnumber(nTotalgaji,0)%></strong></td>
		<td align="right"><strong><%=formatnumber(nP_Jamsostek,0)%></strong></td>
		<td align="right"><strong><%=formatnumber(nP_bpjskes,0)%></strong></td>
		<td align="right"><strong><%=formatnumber(nP_Pph21,0)%></strong></td>
		<td align="right"><strong><%=formatnumber(nPinjaman,0)%></strong></td>
		<td align="right"><strong><%=formatnumber(nDenda,0)%></strong></td>
		<td align="right"><strong><%=formatnumber(nTotalPotongan,0)%></strong></td>
		<td align="right"><strong><%=formatnumber(nTotal,0)%></strong></td>
		<td colspan="6"></td>
	</tr>
	<tr>
		<td align="center" colspan="7" ><strong>TERBILANG</strong></td>
		<%if nTotal > 0 Then %>
		<td colspan="22">
			<%  txtSQL= "select upper(replace(terbilang_indo("& nTotal &"),'-',' ')) terbilang from dual"
			    Set rsCats = Server.CreateObject("ADODB.Recordset")
				rsCats.Open txtSQL, connstr
				if Not rsCats.eof then
					v_terbilang =  rsCats("terbilang") 
				end if
				rsCats.close
				set rsCats = Nothing
			%>
			<strong style="font-style: italic;"> <%=v_terbilang%>  </strong>
		</td>
		<%Else%>
			<td colspan="22">&nbsp;</td>
		<%End If%>
		</tr>
	</tbody>
</table>
<div id="footer_note">	
<div class="block">
</div>
	<div class="block"><h3>Prepared by</h3>,
        <br/><br/><br/><br/>
		<input class="isian" type="text" value="">
        <hr/>
        <input class="isian" type="text" value="">
    </div>
    <div class="block"><h3>Checked by</h3>,
        <br/><br/><br/><br/>
		<input class="isian" type="text" value="">
        <hr/>
        <input class="isian" type="text" value="">
    </div>
    <div class="block"><h3>Approved by</h3>,
        <br/><br/><br/><br/>
        <input class="isian" type="text" value="<%=varManagerName%>">
        <hr/>
        <input class="isian" type="text" value="MANAGER">
    </div>
 </div>
 <div class="clear"></div>
	<!--
<%
 if len(trim(txtAssistqGang)) > 0 then%>
	

   <table width="90%" class="data" style="margin:auto;">
   	   <tr> 
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	   </tr>
	   <tr>
 		<td width="30%" align="center"><font face="Verdana" size="1">Approved by,</font></td>
		<td width="30%" align="center"><font face="Verdana" size="1">Directed by,</font></td>
		<td width="30%" align="center"><font face="Verdana" size="1">Prepared by,</font></td>
	   </tr>
	   <tr> 
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	   </tr>
	   <tr> 
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	   </tr>
	   <tr> 
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	   </tr>
	   <tr> 
		<td align="center"><font face="Verdana" size="2"><strong><u>&nbsp;&nbsp;<%=Session("FCMANAGER")%>&nbsp;&nbsp;</u></strong></td>
		<td align="center"><font face="Verdana" size="2"><strong><u>&nbsp;&nbsp;<%=Session("KTU")%>&nbsp;&nbsp;</u></strong></td>
		<td align="center" ><font face="Verdana" size="2"><strong><u>&nbsp;&nbsp;<%=txtAssistqGang%>&nbsp;&nbsp;</u></strong></td>
	   </tr>
	   	  <tr> 
		<td align="center"><font face="Verdana" size="1"><strong><i>MANAGER</i></strong></td>
		<td align="center"><font face="Verdana" size="1"><strong><i>KTU</i></strong></td>
		<td align="center"><font face="Verdana" size="1"><strong><i><%=txtPostqGang%></i></strong></td>
	   </tr>
   </table>
   <%else%>
   <table width="90%" class="data" style="margin:auto;">
   	   <tr> 
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	   </tr>
	   <tr>
 		<td width="50%" align="center"><font face="Verdana" size="1">Approved by,</font></td>
		<td width="50%" align="center"><font face="Verdana" size="1">Prepared by,</font></td>
	   </tr>
	   <tr> 
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	   </tr>
	   <tr> 
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	   </tr>
	   <tr> 
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	   </tr>
	   <tr> 
		<td align="center"><font face="Verdana" size="2"><strong><u>&nbsp;&nbsp;<%=Session("FCMANAGER")%>&nbsp;&nbsp;</u></strong></td>
		<td align="center"><font face="Verdana" size="2"><strong><u>&nbsp;&nbsp;<%=Session("KTU")%>&nbsp;&nbsp;</u></strong></td>
	   </tr>
	   	  <tr> 
		<td align="center"><font face="Verdana" size="1"><strong><i>MANAGER</i></strong></td>
		<td align="center"><font face="Verdana" size="1"><strong><i>KTU</i></strong></td>
	   </tr>
   </table>

<%end if%>  -->
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