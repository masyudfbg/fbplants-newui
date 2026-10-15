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
nSpanTotal = 0
Dim v_paymenttype : v_paymenttype = Request("paymenttype")

strDefaultBa = " select" & _
               "     fccode as fcba" & _
               " from businessunit bu" & _
               " inner join user_fcba_tab uft on (bu.fccode = uft.fcba) " & _
               " where isdefault = 'TRUE' and bu.IsActive= 'TRUE' " & _
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

Function IsPostBack()
     If Request.ServerVariables("Request_Method") = "POST" Then
          IsPostBack = True
     Else
          IsPostBack = False
     End If
	End Function

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
	oSql =  " SELECT '--ALL--' as STATUS FROM DUAL UNION SELECT DISTINCT PRESENTGRADE as STATUS FROM EMPLOYEE WHERE FCBA = '"& Session("fcba") &"' ORDER BY Status ASC"
	set oRs = server.CreateObject("Adodb.recordset")
	oRs.Open oSql, connstr
	if Not oRs.Eof then
		arr_indeks2 = oRs.getrows()
		arr_indeks_eof = false
	End if
	oRs.Close
	Set oRs = Nothing
end if

'strcurrentdate = " select currentdate" & _
'                 " from periodcontrol" & _                 
'                 " where fcba = '" & v_txtFcba & "'"

'Set rscurrentdate = Server.CreateObject("ADODB.Recordset")
'rscurrentdate.Open strcurrentdate, connstr, adOpenStatic

'If Not rscurrentdate.EOF Then
'	varDefaultYear	=	Year(rscurrentdate("currentdate"))
'	varDefaultMonth	=	Month(rscurrentdate("currentdate"))
'Else
'	varDefaultYear	=	Year(Session("v_Employee"))
'	varDefaultMonth	=	Month(Session("v_Employee"))
	varDefaultYear	=	Year(date)
	varDefaultMonth	=	Month(date)
	
'End If
'rscurrentdate.Close
'set rscurrentdate = Nothing

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
		whereCons = whereCons & "AND INSTR('"& v_Emptype &"', SUBSTR(emp_type,1,12) ) <> 0"
	else
		nSpanTotal = nSpanTotal - 1  
	end if
' End Rev Emp Type By SAG 7 Agustus 2019

	Dim rsPayroll, strPayroll, noline 
	Dim nGet_Hk, nGet_Hk_kj, nget_mkp1, nBase_Salary, nrapelgaji, nOvertime, nPremi, nTunjangan, nBorongan, nP_Jamsostek, nP_bpjskes, nP_Mk_P1, nP_Panen, nP_Pph21, nP_koperasi, nP_Lain, nTotal
			strPayroll ="  SELECT year,month,fcba,fccode,emp_type,gender,subgrade,taxstatus,fcname,gangcode,presentgrade2,sectionname,paymenttype,identitycard," & _
						" position,dateofbirth,dateofjoin, " & _
						"         get_hk,base_salary,overtime,premi,borongan,tunjangan,thr_bonus,totalgaji," & _
                        "         p_jamsostek,p_bpjskes,p_pph21,denda,pinjaman,totalpotongan," & _ 
                        "         total" & _
						"  FROM VIEW_BKMPAYROLL " & _ 
						"  WHERE  year = " & v_txtYear & "" & _
						"     AND month = " & v_txtMonth & "" & _
						"     AND fcba = '" & v_txtFcba & "'" & whereCons & "" & _
						"     AND total >  0"
								
			if  len(trim(v_txtsectionname)) > 0 then
				strPayroll = strPayroll & " and sectionname = '" & v_txtsectionname & "'"
			end if
			if  len(trim(v_txtGangCode)) > 0 then
				strPayroll = strPayroll & " and gangcode = '" & v_txtGangCode & "'"
			end if
			if  len(trim(v_txtStatus)) > 0 then
				strPayroll = strPayroll & " AND INSTR('"& v_txtStatus &"', SUBSTR(presentgrade2,1,3) ) <> 0"
			end if
			if len(trim(request("PAYMENTTYPE"))) > 0 then
				strPayroll = strPayroll & " and paymenttype = '" & v_paymenttype & "'" 
			end if
			
			strPayroll = strPayroll & " order by sectionname,gangcode, presentgrade2,fccode " 
			'response.write strPayroll
			
		
			
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

}

</style>
<link href="/plugins/chosen/chosen.min.css" rel="stylesheet" />
</head>

<body>
	<div id="wrapper">	
		<div id="header-wrapper">
			<div class="printhide">
			
				<!--#include file="print_excel_blob.asp"--><input type="button" class="tableToPdf" value="Export PDF" onClick="window.print()" style="font-family: 'FontAwesome';margin-top:-10px;">
				
				<!--#include file="REPORT/menutop.asp"-->
				
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
							<select name="txtStatus" class="chosen-select" multiple style="width:270px;">
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
				<th colspan="28" style="background-color:#FFFFFF; border:none;">
					<div id="header-wrapper">
					<div id="head_judul" style="margin-bottom:-30px; position:static; text-align:left;">
						<h4>User print: <%=Session("CodeName")%></h4>
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
			<th rowspan="2" style="color:#fff;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Divisi&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</th>
			<th rowspan="2" style="color:#fff;">NO</th>
			<th rowspan="2" style="color:#fff;">NIK</th>
			<th rowspan="2" style="color:#fff;">CardId</th>
			<th rowspan="2" style="color:#fff;">Nama Karyawan</th>
			<th rowspan="2" style="color:#fff;">Gang</th>
			<th rowspan="2" style="color:#fff;">Jenis Kelamin</th>
			<th rowspan="2" style="color:#fff;">Status</th>
			<th rowspan="2" style="color:#fff;">Golongan</th>
			<th rowspan="2" style="color:#fff;">Tipe</th>
			<th rowspan="2" style="color:#fff;">Jabatan</th>
			<th rowspan="2" style="color:#fff;">Tgl Lahir</th>
			<th rowspan="2" style="color:#fff;">TMK</th>
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
	if IsPostBack then
			Set rsPayroll = Server.CreateObject("ADODB.Recordset")
			rsPayroll.Open strPayroll, connstr, adOpenStatic

			if rsPayroll.BOF and rsPayroll.EOF then
%>
		<tr>
			<td colspan="25" align="left">NO DATA AVAILABLE</td>
		</tr>
<%			end if
		NO = 0
		rows = 1
		while not rsPayroll.eof
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
		
		totalpotongan = cdbl(to_nol(rsPayroll("p_jamsostek"))) + cdbl(to_nol(rsPayroll("p_bpjskes"))) + cdbl(to_nol(rsPayroll("p_pph21"))) + cdbl(to_nol(rsPayroll("pinjaman"))) + cdbl(to_nol(rsPayroll("denda"))) 
		totalsisaupah = cdbl(to_nol(rsPayroll("totalgaji"))) - (-1 * cdbl(to_nol(totalpotongan)) )
%>
		<tr class="<%=cClass%>">
			<% if rows = 1 then
					sqlrow = "  select count(sectionname) rowsd  from ( " & strPayroll & " ) where sectionname = '" & rsPayroll("sectionname") & "'"
					' response.write sqlrow
					Set rsrow = Server.CreateObject("ADODB.Recordset")
					rsrow.Open sqlrow, connstr, adOpenStatic
					if not rsrow.EOF then
						temprows = CDbl(rsrow("rowsd"))
					else
						temprows = 0
					end if
				end if 
			no = NO+1
			%>
			<%if rows = 1 then%>
				<td align="center" Rowspan="<%=0+temprows%>"><%=rsPayroll("sectionname")%></td> 
			<%end if%>
			<td align="center"><%=No%></td>
			<td align="left"><%=rsPayroll("FCCODE")%></td>
			<td align="left"><span style='mso-number-format:"\@";'><%=rsPayroll("identitycard")%></span></td>
			<td align="left"><%=rsPayroll("FCNAME")%></td>
			<td align="left"><%=rsPayroll("gangcode")%></td>
			<td align="center"><%=rsPayroll("GENDER")%></td>
			<td align="center"><%=rsPayroll("taxstatus")%></td>
			<td align="center"><%=rsPayroll("PRESENTGRADE2")%></td>
			<td align="center"><%=rsPayroll("EMP_TYPE")%></td>
			<td align="LEFT"><%=rsPayroll("POSITION")%></td>
			<td align="center"><%=rsPayroll("DATEOFBIRTH")%></td>
			<td align="center"><%=rsPayroll("DATEOFJOIN")%></td>
			<td align="right"><%=formatnumber( cdbl(rsPayroll("GET_HK")), 1)%></td>
			<td align="right"><%=formatnumber( cdbl(rsPayroll("BASE_SALARY")),2)%></td>
			<td align="right"><%=formatnumber( cdbl(rsPayroll("OVERTIME")),2)%></td>
			<td align="right"><%=formatnumber( cdbl(rsPayroll("PREMI")),2)%></td>
			<td align="right"><%=formatnumber( cdbl(rsPayroll("BORONGAN")),2)%></td>
			<td align="right"><%=formatnumber( cdbl(rsPayroll("TOTALGAJI")),2)%></td>
			<td align="right"><%=formatnumber( cdbl(rsPayroll("P_JAMSOSTEK")),2)%></td>
			<td align="right"><%=formatnumber( cdbl(rsPayroll("P_BPJSKES")),2)%></td>
			<td align="right"><%=formatnumber( cdbl(rsPayroll("P_PPH21")),2)%></td>
			<td align="right"><%=formatnumber( cdbl(rsPayroll("PINJAMAN")),2)%></td>
			<td align="right"><%=formatnumber( cdbl(rsPayroll("DENDA")),2)%></td>
			<td align="right"><%=formatnumber( cdbl(rsPayroll("TOTALPOTONGAN")),2)%></td>
			<td align="right"><%=formatnumber( cdbl(rsPayroll("TOTAL")),2)%></td>
			<%if noline  mod 2 = 1 then%>
				<td align="left" style="background-color:#ccc;" width="500" colspan="3"><font face="Verdana" size="1">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>			
				<td align="left" style="background-color:#fff;" width="500" colspan="3"><font face="Verdana" size="1"><%=NO%>.&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>			
			<%Else%>
				<td align="left" style="background-color:#fff;" width="500" colspan="3"><font face="Verdana" size="1"><%=NO%>.&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>			
				<td align="left" style="background-color:#ccc;" width="500" colspan="3"><font face="Verdana" size="1">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font></td>			
			<%End if%> 
		</tr>
<%	if rows = temprows then
			noline			= noline		+ 1
			nGet_Hk 		= nGet_Hk 		+ cdbl(rsPayroll("GET_HK"))
			nBase_Salary 	= nBase_Salary 	+ cdbl(rsPayroll("BASE_SALARY"))
			nOvertime 		= nOvertime 	+ cdbl(rsPayroll("OVERTIME"))
			nBorongan 		= nBorongan 	+ cdbl(rsPayroll("BORONGAN"))
			nP_Jamsostek 	= nP_Jamsostek 	+ cdbl(rsPayroll("P_JAMSOSTEK"))
			nP_bpjskes 		= nP_bpjskes 	+ cdbl(rsPayroll("P_BPJSKES"))
			nTotal 			= nTotal 		+ cdbl(rsPayroll("TOTAL"))
			nDenda 			= nDenda		+ cdbl(rsPayroll("DENDA"))
			nTotalgaji 		= nTotalgaji	+ cdbl(rsPayroll("TOTALGAJI"))
			nPremi 			= nPremi 		+ cdbl(rsPayroll("PREMI"))
			nPinjaman 		= nPinjaman 	+ cdbl(rsPayroll("PINJAMAN"))
			nP_PPH21 		= nP_PPH21 		+ cdbl(rsPayroll("P_PPH21"))
			nTotalPotongan  = nTotalPotongan + cdbl(rsPayroll("TOTALPOTONGAN"))
			ntotalsisaupah = ntotalsisaupah + cdbl(rsPayroll("TOTAL"))
			rows = 1
			
	%>
		<tr style="background-color:#F4F4A0;">	 
			<td align="center" colspan="13" ><strong>Sub Total</strong></td>
			<td align="right"><strong><%=formatnumber(nGet_Hk,1)%></strong></td>
			<td align="right"><strong><%=formatnumber(nBase_Salary,2)%></strong></td>
			<td align="right"><strong><%=formatnumber(nOvertime,2)%></strong></td>
			<td align="right"><strong><%=formatnumber(nPremi,2)%></strong></td> 
			<td align="right"><strong><%=formatnumber(nBorongan,2)%></strong></td> 
			<td align="right"><strong><%=formatnumber(nTotalgaji,2)%></strong></td>
			<td align="right"><strong><%=formatnumber(nP_Jamsostek,2)%></strong></td>
			<td align="right"><strong><%=formatnumber(nP_bpjskes,2)%></strong></td>
			<td align="right"><strong><%=formatnumber(nP_Pph21,2)%></strong></td>
			<td align="right"><strong><%=formatnumber(nPinjaman,2)%></strong></td>
			<td align="right"><strong><%=formatnumber(nDenda,2)%></strong></td>
			<td align="right"><strong><%=formatnumber(nTotalPotongan,2)%></strong></td>
			<td align="right"><strong><%=formatnumber(ntotalsisaupah,2)%></strong></td>
			<td colspan="6"></td>
		</tr>
		<tr>
			<td align="center" colspan="<%= 23+ nSpanTotal %>" ></td>
		</tr>
		<%	
			noline = 1
			nGet_Hk = cdbl("0")
			nBase_Salary = cdbl("0")
			nOvertime = cdbl("0")
			nPremi = cdbl("0")
			nTunjangan = cdbl("0")
			nBorongan = cdbl("0")
			nP_Jamsostek = cdbl("0")
			nP_Mk_P1 = cdbl("0")
			nP_Panen = cdbl("0")
			nP_Pph21 = cdbl("0")
			nP_Lain = cdbl("0")
			nTotal = cdbl("0")
			nDenda = cdbl("0")
			nTotalgaji = cdbl("0")
			nPinjaman = cdbl("0")
			nTotalPotongan = cdbl("0")
			ntotalsisaupah = cdbl("0")
			nP_bpjskes = cdbl("0")
			 
		else
		' cari sub total
			noline			= noline		+ 1
			nGet_Hk 		= nGet_Hk 		+ cdbl(rsPayroll("GET_HK"))
			nBase_Salary 	= nBase_Salary 	+ cdbl(rsPayroll("BASE_SALARY"))
			nOvertime 		= nOvertime 	+ cdbl(rsPayroll("OVERTIME"))
			nBorongan 		= nBorongan 	+ cdbl(rsPayroll("BORONGAN"))
			nP_Jamsostek 	= nP_Jamsostek 	+ cdbl(rsPayroll("P_JAMSOSTEK"))
			nP_bpjskes 		= nP_bpjskes 	+ cdbl(rsPayroll("P_BPJSKES"))
			nTotal 			= nTotal 		+ cdbl(rsPayroll("TOTAL"))
			nDenda 			= nDenda		+ cdbl(rsPayroll("DENDA"))
			nTotalgaji 		= nTotalgaji	+ cdbl(rsPayroll("TOTALGAJI"))
			nPremi 			= nPremi 		+ cdbl(rsPayroll("PREMI"))
			nPinjaman 		= nPinjaman 	+ cdbl(rsPayroll("PINJAMAN"))
			nP_PPH21 		= nP_PPH21 		+ cdbl(rsPayroll("P_PPH21"))
			nTotalPotongan 		= nTotalPotongan 	+ cdbl(rsPayroll("TOTALPOTONGAN"))
			ntotalsisaupah 		= ntotalsisaupah 	+ cdbl(rsPayroll("TOTAL"))
			rows = rows + 1
		End if
							  
			' cari grand total 
				gt_GET_HK = cdbl(gt_GET_HK) + cdbl(rsPayroll("GET_HK"))
				gt_BASE_SALARY = cdbl(gt_BASE_SALARY) + cdbl(rsPayroll("BASE_SALARY"))
				gt_OVERTIME = cdbl(gt_OVERTIME) + cdbl(rsPayroll("OVERTIME")) 
				gt_BORONGAN = cdbl(gt_BORONGAN) + cdbl(rsPayroll("BORONGAN"))
				gt_P_JAMSOSTEK = cdbl(gt_P_JAMSOSTEK) + cdbl(rsPayroll("P_JAMSOSTEK"))
				gt_P_BPJSKES = cdbl(gt_P_BPJSKES) + cdbl(rsPayroll("P_BPJSKES")) 
				gt_TOTAL = cdbl(gt_TOTAL) + cdbl(rsPayroll("TOTAL")) 
				gt_DENDA = cdbl(gt_DENDA) + cdbl(rsPayroll("DENDA")) 
				gt_TOTALGAJI = cdbl(gt_TOTALGAJI) + cdbl(rsPayroll("TOTALGAJI")) 
				gt_PREMI = cdbl(gt_PREMI) + cdbl(rsPayroll("PREMI")) 
				gt_PINJAMAN = cdbl(gt_PINJAMAN) + cdbl(rsPayroll("PINJAMAN"))  
				gt_P_PPH21 = cdbl(gt_P_PPH21) + cdbl(rsPayroll("P_PPH21")) 
				gt_totalpotongan = cdbl(gt_totalpotongan) + cdbl(rsPayroll("TOTALPOTONGAN"))
				gt_totalsisaupah = cdbl(gt_totalsisaupah) + cdbl(rsPayroll("TOTAL"))				
		 
			rsPayroll.MoveNext			
			Wend
		rsPayroll.Close
		set rsPayroll = Nothing
end if  'end isposback
	%> 
	<tr style="background-color:#27AE60;">	 
		<td align="center" colspan="13" ><strong>Grand Total</strong></td>
		<td align="right"><strong><%=formatnumber(gt_GET_HK,1)%></strong></td>  
		<td align="right"><strong><%=formatnumber(gt_BASE_SALARY,2)%></strong></td> 
		<td align="right"><strong><%=formatnumber(gt_OVERTIME,2)%></strong></td> 
		<td align="right"><strong><%=formatnumber(gt_PREMI,2)%></strong></td>
		<td align="right"><strong><%=formatnumber(gt_BORONGAN,2)%></strong></td>
		<td align="right"><strong><%=formatnumber(gt_TOTALGAJI,2)%></strong></td>
		<td align="right"><strong><%=formatnumber(gt_P_JAMSOSTEK,2)%></strong></td>
		<td align="right"><strong><%=formatnumber(gt_P_BPJSKES,2)%></strong></td>
		<td align="right"><strong><%=formatnumber(gt_P_PPH21,2)%></strong></td>
		<td align="right"><strong><%=formatnumber(gt_PINJAMAN,2)%></strong></td>
		<td align="right"><strong><%=formatnumber(gt_DENDA,2)%></strong></td>
		<td align="right"><strong><%=formatnumber(gt_totalpotongan,2)%></strong></td>
		<td align="right"><strong><%=formatnumber(gt_totalsisaupah,2)%></strong></td>
		<td colspan="6"></td>
	</tr>
	<tr>
		<td align="center" colspan="13" ><strong>TERBILANG</strong></td>
		<%if gt_totalsisaupah > 0 Then %>
		<td colspan="25">
			<%  txtSQL= "select upper(replace(terbilang_indo("& gt_totalsisaupah &"),'-',' ')) terbilang from dual"
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
			<td colspan="25">&nbsp;</td>
		<%End If%>
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