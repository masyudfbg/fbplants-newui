
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
ScriptName		= Request.ServerVariables("SCRIPT_NAME")

If "" & txtYear = "" Then txtYear = Year(Session("v_Costbook"))
If "" & txtMonth = "" Then txtMonth = Month(Session("v_Costbook"))

txttdate = Request("txttdate")
txttdate = textDate(txtMonth&"/01/"&txtYear)

nSpanTotal = 91

Dim v_paymenttype : v_paymenttype = Request("paymenttype")

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
%>
<html>
<head>
	<title>ePlaS Reporting - Payroll Karyawan</title>
	<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
	<link rel="stylesheet" href="/css/font/font-awesome.css"/>
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
#footer_note .block { font-size:10px; float:left; width:24%; height:90px; background-color:#fff; margin-left:1%; text-align:center;}
input[type="text"].isian {border:0px;box-shadow:none; padding:0px;font-size: 10px;color:#000;height: 20px;width:100%; text-align:center;}
hr{border-top:1px #000000 solid; border-bottom:0px; border-left:0px; border-right:0px;}
}
@media print
{
#footer_note {margin-top:20px;}
#footer_note .block { font-size:10px; float:left; width:24%; height:90px; background-color:#fff; margin-left:1%; text-align:center;}
input[type="text"].isian {border:0px;box-shadow:none; padding:0px;font-size: 10px;color:#000;height: 20px;width:100%; text-align:center;}
hr{border-top:1px #000000 solid; border-bottom:0px; border-left:0px; border-right:0px;}
.printhide{display:none;}
@page { 
 size:Letter landscape;
}
}

</style>
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
						strsectionname	=	" select fccode, fcname from sectionname order by orderdata asc"
						
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
							<font face="verdana" size="1">GANGCODE</font>
						</td>
						<td>
							<select name="txtGangCode">
								<option value = "">--ALL--</option>
								<%
								Dim rsGangCode, strGangCode
								strGangCode	=	" SELECT FCCODE, FCNAME FROM GANG WHERE ACTIVATION = 'Y' AND FCBA = '" & v_txtFcba & "' ORDER BY FCCODE ASC"
								
								Set rsGangCode = Server.CreateObject("ADODB.Recordset")
								rsGangCode.Open strGangCode, connstr, adOpenStatic 
								
								if not rsGangCode.eof then
									
								Do While Not rsGangCode.eof
								if rsGangCode("FCCODE") = v_txtGangCode then%>
										<option value="<%=rsGangCode("FCCODE")%>" Selected><%=rsGangCode("FCCODE")%>-<%=rsGangCode("FCNAME")%></option> 
									<%Else%>
										<option value="<%=rsGangCode("FCCODE")%>"><%=rsGangCode("FCCODE")%>-<%=rsGangCode("FCNAME")%></option>  
									<%End if
									rsGangCode.movenext
								loop
								end if
								rsGangCode.Close
								set rsGangCode = Nothing
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
								<option value = "KHT" <%If v_txtStatus = "KHT" Then %> Selected <%End If%> >KHT</option>
								<option value = "KBT" <%If v_txtStatus = "KBT" Then %> Selected <%End If%> >KBT</option>
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
				<th colspan="100" style="background-color:#FFFFFF; border:none;">
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
			<%
			arr_master_eof = true
			strPayroll = " select fccode, fcname from allowded where fcba = '"& v_txtFcba &"' order by fccode"

			Set rsPayroll = Server.CreateObject("ADODB.Recordset")
			rsPayroll.Open strPayroll, connstr, adOpenStatic
   			
   			' response.write strPayroll
			if Not rsPayroll.eof then
				arr_master = rsPayroll.getrows()
				arr_master_eof = false
			End If
			x = "'zzz'"
			if not arr_master_eof then
				x = ""
				for xyz = 0 to ubound(arr_master,2)
					if xyz = ubound(arr_master,2) then
						x = x & "'" & arr_master(0,xyz) &"'"
					else
						x = x & "'" & arr_master(0,xyz) &"',"
				end if
				
				next
			end if %>
			<th rowspan="2" style="color:#fff;">NIK</th>
			<th rowspan="2" style="color:#fff;">Nama Karyawan</th>
			<th rowspan="2" style="color:#fff;">Gang</th>
			<th rowspan="2" style="color:#fff;">Golongan</th>
			<th rowspan="2" style="color:#fff;">Status</th>
			<th colspan="2" style="color:#fff;">HK</th>
			<th rowspan="2" style="color:#fff;">GAJI POKOK</th>
			<th rowspan="2" style="color:#fff;">RAPEL GAJI</th>
			<th rowspan="2" style="color:#fff;">FRINGE CUTI</th>
			<th rowspan="2" style="color:#fff;">FRINGE SAKIT</th>
			<th rowspan="2" style="color:#fff;">FRINGE SAKIT DIRAWAT</th>
			<th rowspan="2" style="color:#fff;">FRINGE CUTI HAID</th>
			<th rowspan="2" style="color:#fff;">FRINGE CUTI HAMIL</th>
			<th rowspan="2" style="color:#fff;">FRINGE IZIN SOSIAL</th>
			<th rowspan="2" style="color:#fff;">FRINGE IZIN PRIBADI</th>
			<th rowspan="2" style="color:#fff;">FRINGE IZIN PEMERINTAH</th>
			<th rowspan="2" style="color:#fff;">FRINGE LIBUR</th>
			<th rowspan="2" style="color:#fff;">BERAS PEKERJA</th>
			<th rowspan="2" style="color:#fff;">BERAS PASANGAN</th>
			<th rowspan="2" style="color:#fff;">BERAS TANGGUNGAN</th>
			<th rowspan="2" style="color:#fff;">TUNJANGAN FUNGSIONAL</th>
			<th rowspan="2" style="color:#fff;">TUNJANGAN BBM</th>
			<th rowspan="2" style="color:#fff;">TUNJANGAN JAMSOSTEK</th>
			<th rowspan="2" style="color:#fff;">LEMBUR</th>
			<th rowspan="2" style="color:#fff;">PREMI MUAT PUPUK</th>
			<th rowspan="2" style="color:#fff;">PREMI LANGSIR PUPUK</th>
			<th rowspan="2" style="color:#fff;">PREMI BURU BABI</th>
			<th rowspan="2" style="color:#fff;">PREMI KEHADIRAN</th>
			<th rowspan="2" style="color:#fff;">PREMI KERAJINAN</th>
			<th rowspan="2" style="color:#fff;">PREMI LEBIH BORONG</th>
			<th rowspan="2" style="color:#fff;">PREMI TANDAN HARI LIBUR</th>
			<th rowspan="2" style="color:#fff;">PREMI BRONDOLAN</th>
			<th rowspan="2" style="color:#fff;">PREMI SIAP BORONG</th>
			<th rowspan="2" style="color:#fff;">PREMI ANGKUT BUAH</th>
			<th rowspan="2" style="color:#fff;">PREMI ANGKUT BUAH LEBIH BASIS</th>
			<th rowspan="2" style="color:#fff;">PREMI TRANSPORT</th>
			<th rowspan="2" style="color:#fff;">PREMI TRANSPORT LEBIH BASIS</th>
			<th rowspan="2" style="color:#fff;">PREMI OPERATOR</th>
			<th rowspan="2" style="color:#fff;">PREMI MEKANIK</th>
			<th rowspan="2" style="color:#fff;">PREMI KANTOR</th>
			<th rowspan="2" style="color:#fff;">PREMI KEAMANAN</th>
			<th rowspan="2" style="color:#fff;">PREMI MANDOR PANEN (OP)</th>
			<th rowspan="2" style="color:#fff;">PREMI KERANI PANEN (OP)</th>
			<th rowspan="2" style="color:#fff;">PREMI MANDOR I (OP)</th>
			<th rowspan="2" style="color:#fff;">PREMI KERANI DIVISI (OP)</th>
			<th rowspan="2" style="color:#fff;">GAJI BORONGAN</th>
			<th rowspan="2" style="color:#fff;">PREMI SEMPROT</th>
			<th rowspan="2" style="color:#fff;">PREMI MANDOR SEMPROT</th>
			<th rowspan="2" style="color:#fff;">PREMI PUPUK</th>
			<th rowspan="2" style="color:#fff;">PREMI MANDOR PUPUK</th>
			<th rowspan="2" style="color:#fff;">PREMI UNTIL</th>
			<th rowspan="2" style="color:#fff;">PREMI TETAP</th>
			<th rowspan="2" style="color:#fff;">TUNJANGAN JABATAN</th>
			<th rowspan="2" style="color:#fff;">PREMI KONTANAN HARI LIBUR - MANDOR (SUPERVISOR 2)</th>
			<th rowspan="2" style="color:#fff;">PREMI KONTANAN HARI LIBUR - KRANI (SUPERVISOR 3)</th>
			<th rowspan="2" style="color:#fff;">PREMI KONTANAN HARI LIBUR - MANDOR1 (SUPERVISOR 1)</th>
			<th rowspan="2" style="color:#fff;">PREMI TUNAS</th>
			<th rowspan="2" style="color:#fff;">PREMI LEBIH BASIS TUNAS</th>
			<th rowspan="2" style="color:#fff;">PREMI TUNAS MANDOR</th>
			<th rowspan="2" style="color:#fff;">TUNJANGAN PPH 21</th>
			<th rowspan="2" style="color:#fff;">TUNJANGAN HARI RAYA</th>
			<th rowspan="2" style="color:#fff;">TUNJANGAN BPJS KESEHATAN</th>
			<th rowspan="2" style="color:#fff;">BONUS</th>
			<th rowspan="2" style="color:#fff;">POTONGAN MANGKIR</th>
			<th rowspan="2" style="color:#fff;">POTONGAN JAMSOSTEK</th>
			<th rowspan="2" style="color:#fff;">POTONGAN PPH 21</th>
			<th rowspan="2" style="color:#fff;">POTONGAN PPH21 - JABATAN</th>
			<th rowspan="2" style="color:#fff;">POTONGAN PPH21 - ASTEK</th>
			<th rowspan="2" style="color:#fff;">POTONGAN GAJI</th>
			<th rowspan="2" style="color:#fff;">POTONGAN ALAT KERJA</th>
			<th rowspan="2" style="color:#fff;">POTONGAN PINJAMAN UANG</th>
			<th rowspan="2" style="color:#fff;">POTONGAN PINJAMAN MATERIAL</th>
			<th rowspan="2" style="color:#fff;">POTONGAN LAIN-LAIN</th>
			<th rowspan="2" style="color:#fff;">POTONGAN BUAH MENTAH</th>
			<th rowspan="2" style="color:#fff;">POTONGAN TANGKAI PANJANG</th>
			<th rowspan="2" style="color:#fff;">POTONGAN PENGOBATAN DLT</th>
			<th rowspan="2" style="color:#fff;">POTONGAN BUAH TIDAK DIPANEN</th>
			<th rowspan="2" style="color:#fff;">POTONGAN PELEPAH TDK DIPOTONG</th>
			<th rowspan="2" style="color:#fff;">POTONGAN PELEPAH TDK DISUSUN RAPI</th>
			<th rowspan="2" style="color:#fff;">POTONGAN BRONDOLAN DIPIRINGAN</th>
			<th rowspan="2" style="color:#fff;">POTONGAN BUAH TINGGAL DIPIRINGAN</th>
			<th rowspan="2" style="color:#fff;">POTONGAN BUAH MATAHARI</th>
			<th rowspan="2" style="color:#fff;">POTONGAN IJIN PRIBADI</th>
			<th rowspan="2" style="color:#fff;">POTONGAN TIDAK DAPAT BASIS</th>
			<th rowspan="2" style="color:#fff;">POTONGAN KONTANAN HARI LIBUR - EMPLOYEE</th>
			<th rowspan="2" style="color:#fff;">POTONGAN KONTANAN HARI LIBUR - MANDOR(SUPERVISOR2)</th>
			<th rowspan="2" style="color:#fff;">POTONGAN KONTANAN HARI LIBUR - KRANI(SUPERVISIOR3)</th>
			<th rowspan="2" style="color:#fff;">POTONGAN KONTANAN HARI LIBUR - MANDOR1(SPRSV1)</th>
			<th rowspan="2" style="color:#fff;">POTONGAN BPJS KESEHATAN</th>
			<th rowspan="2" style="color:#fff;">POTONGAN IURAN KOPERASI</th>
			<th rowspan="2" style="color:#fff;">POTONGAN BON KOPERASI</th>
			<th rowspan="2" style="color:#fff;">POTONGAN BUAH MENTAH DIPERAM</th>
			<th rowspan="2" style="color:#fff;">POTONGAN OVER PRUNING</th>
			<th rowspan="2" style="color:#fff;">POTONGAN BRONDOLAN TIDAK DI ALAS</th>
			<th rowspan="2" style="color:#fff;">POTONGAN BRONDOLAN KOTOR/SAMPAH</th>
			<th rowspan="2" style="color:#fff;">POTONGAN BUAH DIBELAH</th>
			<th rowspan="2" style="color:#fff;">POTONGAN BUAH TIDAK DISUSUN RAPI DI TPH</th>
			<th rowspan="2" style="color:#fff;">Total</th>
			<th rowspan="2" colspan="2" style="color:#fff;">Tanda Tangan</th>
		</tr>
		<tr>
			<th style="color:#fff;">Sebulan</th>
			<th style="color:#fff;">Kerja</th>	
		</tr>
	</thead>
	<tbody>
	<%
	Dim rsPayroll, strPayroll, noline 
	Dim nGet_Hk, nGet_Hk_kj, nget_mkp1, nBase_Salary, nrapelgaji, nOvertime, nPremi, nTunjangan, nBorongan, nP_Jamsostek, nP_bpjskes, nP_Mk_P1, nP_Panen, nP_Pph21, nP_koperasi, nP_Lain, nTotal
			
			strPayroll = 	"   SELECT * FROM (SELECT year," & _
							"          month," & _
							"          fccode," & _
							"          fcba," & _
							"          NVL (get_hk, 0) get_hk," & _
							"          NVL (get_hk_kj, 0) get_hk_kj" & _
							"       FROM (    SELECT TO_NUMBER (SUBSTR (t.yearmonth, 1, 4)) year," & _
							"                  TO_NUMBER (SUBSTR (t.yearmonth, 5, 2)) month," & _
							"                  t.fcba," & _
							"                  t.fccode," & _
							"                  SUM (get_hk) get_hk," & _
							"                  SUM (get_hk_kj) get_hk_kj" & _
							"               FROM view_payrolldetail t" & _
							"            WHERE   TO_NUMBER (SUBSTR (t.yearmonth, 1, 4)) = " & v_txtYear & "" & _
							"                  AND TO_NUMBER (SUBSTR (t.yearmonth, 5, 2)) = " & v_txtMonth & "" & _
							"                  AND t.fcba = '" & v_txtFcba & "'" & _
							"             GROUP BY TO_NUMBER (SUBSTR (t.yearmonth, 1, 4))," & _
							"                  TO_NUMBER (SUBSTR (t.yearmonth, 5, 2))," & _
							"                  t.fcba," & _
							"                  t.fccode)) a" & _
							"      INNER JOIN (SELECT    EMD.ALLOWDEDCODE ," & _
							"              EMD.EMPCODE," & _
							"			   EMP.FCNAME," & _
							"              EMD.AMOUNT," & _
							"              EMD.FCBA," & _
							"              EMP.SUBGRADE," & _
							"              EMP.SECTIONNAME," & _
							"              EMP.GANGCODE," & _
							"              EMP.PAYMENTTYPE," & _
							"              EMP.PRESENTGRADE," & _
							"			   SUM(AMOUNT) AS TOTAL" & _
							"             FROM EMPALLDED EMD" & _
							"              INNER JOIN ALLOWDED ALD    ON (EMD.ALLOWDEDCODE = ALD.FCCODE AND EMD.FCBA = ALD.FCBA)" & _
							"              INNER JOIN EMPLOYEE EMP    ON (EMD.EMPCODE = EMP.FCCODE AND EMD.fcba = EMP.fcba)" & _
							"              WHERE EMD.FCBA = '" & v_txtFcba & "'" & _
							"              AND TO_CHAR(ADHOCDATE,'MM/DD/YYYY') = TO_CHAR('"& txttdate &"') " & _
							" 				GROUP BY  EMD.ALLOWDEDCODE, EMD.EMPCODE,EMP.FCNAME,EMD.AMOUNT, EMD.FCBA,EMP.SUBGRADE,EMP.SECTIONNAME, EMP.GANGCODE, EMP.PAYMENTTYPE, EMP.PRESENTGRADE)" & _
							"          PIVOT" & _
							"           (SUM (amount)" & _
							"           AS amount" & _
							"           FOR (allowdedcode)" & _
							"           IN ("& x &"))b" & _
							"             ON (a.fccode = b.empcode and a.fcba = b.fcba)" & _
							"      WHERE a.fcba = '" & v_txtFcba & "'"
			' " select *" & _
   '           " from ( select year" & _
   '           "         , month" & _
   '           "         , fcba" & _
   '           "         , fccode" & _
   '           "         , employee_hist_api.get_fcname(year, month, fcba, fccode) fcname" & _
   '           "         , employee_hist_api.get_gangcode(year, month, fcba, fccode) gangcode" & _
   '           "         , employee_hist_api.get_presentgrade(year, month, fcba, fccode) presentgrade" & _
   '           "         , employee_hist_api.get_sectionname(year, month, fcba, fccode) sectionname" & _
   '           "         , employee_hist_api.get_paymenttype(year, month, fcba, fccode) paymenttype" & _
   '           "         , nvl(get_hk, 0) get_hk" & _
   '           "         , nvl(get_hk_kj, 0) get_hk_kj" & _
   '           "         , nvl(get_mkp1, 0) get_mkp1" & _
   '           "         , nvl(base_salary, 0) base_salary" & _
   '           "         , nvl(rapelgaji, 0) rapelgaji" & _
   '           "         , nvl(overtime, 0) overtime" & _
   '           "         , nvl(premi, 0) premi" & _
   '           "         , nvl(tunjangan, 0) tunjangan" & _
   '           "         , nvl(borongan, 0) borongan" & _
   '           "         , nvl(p_jamsostek, 0) p_jamsostek" & _
   '           "         , nvl(p_bpjskes, 0) p_bpjskes" & _
   '           "         , nvl(p_mk_p1, 0) p_mk_p1" & _
   '           "         , nvl(p_panen, 0) p_panen" & _
   '           "         , nvl(p_pph21, 0) p_pph21" & _
   '           "         , nvl(p_koperasi, 0) p_koperasi" & _
   '           "         , nvl(p_lain, 0) p_lain" & _
   '           "         , nvl(total, 0) total" & _
   '           "  from ( select to_number(substr(t.yearmonth, 1, 4)) year" & _
   '           "          , to_number(substr(t.yearmonth, 5, 2)) month" & _
   '           "          , t.fcba" & _
   '           "          , t.fccode" & _
   '           "          , sum(get_hk) get_hk" & _
   '           "          , sum(get_hk_kj) get_hk_kj" & _
   '           "          , sum(get_mkp1) get_mkp1" & _
   '           "          , sum(t.base_salary) base_salary" & _
   '           "          , sum(t.rapelgaji) rapelgaji" & _
   '           "          , sum(t.overtime) overtime" & _
   '           "          , sum(t.premi) premi" & _
   '           "          , sum(t.tunjangan) tunjangan" & _
   '           "          , sum(t.borongan) borongan" & _
   '           "          , sum(t.p_jamsostek) p_jamsostek" & _
   '           "          , sum(t.p_bpjskes) p_bpjskes" & _
   '           "          , sum(t.p_mk_p1) p_mk_p1" & _
   '           "          , sum(t.p_panen) p_panen" & _
   '           "          , sum(t.p_pph21) p_pph21" & _
   '           "          , sum(t.p_koperasi) p_koperasi" & _
   '           "          , sum(t.p_lain) p_lain" & _
   '           "          , sum(t.base_salary)" & _
   '           "          + sum(t.rapelgaji)" & _
   '           "          + sum(t.overtime)" & _
   '           "          + sum(t.premi)" & _
   '           "          + sum(t.tunjangan)" & _
   '           "          + sum(t.borongan)" & _
   '           "          + sum(t.p_jamsostek)" & _
   '           "          + sum(t.p_bpjskes)" & _
   '           "          + sum(t.p_mk_p1)" & _
   '           "          + sum(t.p_panen)" & _
   '           "          + sum(t.p_pph21)" & _
   '           "          + sum(t.p_koperasi)" & _
   '           "          + sum(t.p_lain) total " & _
   '           "         from view_payrolldetail t" & _
   '           "         where to_number(substr(t.yearmonth,1,4)) = " & v_txtYear & "" & _
   '           "         and to_number(substr(t.yearmonth,5,2)) = " & v_txtMonth & "" & _
   '           "         and t.fcba = '" & v_txtFcba & "'" & _
   '           "         group by to_number(substr(t.yearmonth, 1, 4))" & _
   '           "      , to_number(substr(t.yearmonth, 5, 2))" & _
   '           "      , t.fcba" & _
   '           "      , t.fccode" & _
   '           "  )" & _
   '           " ) a " & _
   '      	 "      INNER JOIN (  SELECT * " & _
   '           "      FROM   (SELECT EMD.ALLOWDEDCODE as ALLOWDEDCODE, EMD.EMPCODE EMPCODE, EMD.AMOUNT, EMD.FCBA" & _
   '           "         FROM EMPALLDED EMD " & _
   '           "         INNER JOIN ALLOWDED ALD ON (EMD.ALLOWDEDCODE = ALD.FCCODE AND EMD.FCBA = ALD.FCBA)" & _
   '           "         WHERE EMD.FCBA = '" & v_txtFcba & "')" & _
   '           "       PIVOT  (SUM(amount) AS amount FOR (allowdedcode) IN ("& x &"))" & _
   '           "  )b ON (a.fccode = b.empcode and a.fcba = b.fcba)" & _
   '           "		 INNER JOIN EMPLOYEE EMD ON (EMD.FCCODE = b.empcode and emd.fcba = b.fcba)" & _
   '           " 		 where a.fcba = '" & v_txtFcba & "'"

										
										
			if  len(trim(v_txtsectionname)) > 0 then
				strPayroll = strPayroll & " and b.sectionname = '" & v_txtsectionname & "'"
			end if
			if  len(trim(v_txtGangCode)) > 0 then
				strPayroll = strPayroll & " and b.gangcode = '" & v_txtGangCode & "'"
			end if
			if  len(trim(v_txtStatus)) > 0 then
				strPayroll = strPayroll & " and b.presentgrade = '" & v_txtStatus & "'"
			end if
			if len(trim(request("PAYMENTTYPE"))) > 0 then
				strPayroll = strPayroll & " and b.paymenttype = '" & v_paymenttype & "'" 
			end if
			
			strPayroll = strPayroll & " order by a.fccode" 
			response.write strPayroll
			Set rsPayroll = Server.CreateObject("ADODB.Recordset")
			rsPayroll.Open strPayroll, connstr, adOpenStatic
			arr_data_eof  = true
			if Not rsPayroll.eof then
				arr_data = rsPayroll.getrows()
				arr_data_eof = false
			end if

			if Not arr_data_eof then
				for xyz = 0 to ubound(arr_data,2) 
					response.write "<tr>"
					%>
						<td align="left"><%=arr_data(6,xyz)%></td>
						<td align="left"><%=arr_data(7,xyz)%></td>
						<td align="left"><%=arr_data(11,xyz)%></td>
						<td align="left"><%=arr_data(9,xyz)%></td>
						<td align="left"><%=arr_data(13,xyz)%></td>
						<td align="left"><%=arr_data(4,xyz)%></td>
						<td align="left"><%=arr_data(5,xyz)%></td>
					<%	
						for col = 15 to 105 
			        	
			        %>
			            <td align="left"><%=arr_data(col,xyz)%></td>
			        <%	
				        next
						for col = 14 to 14 
					%>
			            <td align="left"><%=arr_data(14,xyz)%></td>
					<%
			          next  
			      next
			      end if
		    Set rsPayroll = Server.CreateObject("ADODB.Recordset")
			rsPayroll.Open strPayroll, connstr, adOpenStatic

	if Not arr_data_eof then
		noline = 1
		nGet_Hk = 0
		nGet_Hk_kj = 0
		' nBase_Salary = 0
		' nOvertime = 0
		' nPremi = 0
		' nTunjangan = 0
		' nBorongan = 0
		' nP_Jamsostek = 0
		' nP_Mk_P1 = 0
		' nP_Panen = 0
		' nP_Pph21 = 0
		' nP_Lain = 0
		' nTotal = 0

		while arr_data_eof  %>
		     
		<%if noline  mod 2 = 1 then%>
			<td align="left" style="background-color:#ccc;" width="50"><font face="Verdana" size="1">&nbsp;</td>			
			<td align="left" style="background-color:#fff;" width="50"><font face="Verdana" size="1">&nbsp;</td>			
		<%Else%>
			<td align="left" style="background-color:#fff;" width="50"><font face="Verdana" size="1">&nbsp;</td>			
			<td align="left" style="background-color:#ccc;" width="50" bgcolor="#CCCCCC"><font face="Verdana" size="1">&nbsp;</font></td>			
		<%End if%>
		  </tr>	
		      
    <%	
			noline			= noline		+ 1
			' nGet_Hk 		= nGet_Hk 		+ cdbl(rsPayroll("GET_HK"))
			' nGet_Hk_kj 		= nGet_Hk_kj 	+ cdbl(rsPayroll("GET_HK_KJ"))
			' nget_mkp1 		= nget_mkp1		+ cdbl(rsPayroll("get_mkp1"))
			' nBase_Salary 	= nBase_Salary 	+ cdbl(rsPayroll("BASE_SALARY"))
			' nrapelgaji		= nrapelgaji 	+ cdbl(rsPayroll("rapelgaji"))
			' nOvertime 		= nOvertime 	+ cdbl(rsPayroll("OVERTIME"))
			' nPremi 			= nPremi 		+ cdbl(rsPayroll("PREMI"))
			' nTunjangan 		= nTunjangan 	+ cdbl(rsPayroll("TUNJANGAN"))
			' nBorongan 		= nBorongan 	+ cdbl(rsPayroll("BORONGAN"))
			' nP_Jamsostek 	= nP_Jamsostek 	+ cdbl(rsPayroll("P_JAMSOSTEK"))
			' nP_bpjskes = nP_bpjskes + cdbl(rsPayroll("P_BPJSKES"))
			' nP_Mk_P1 		= nP_Mk_P1 		+ cdbl(rsPayroll("P_MK_P1"))
			' nP_Panen 		= nP_Panen 		+ cdbl(rsPayroll("P_PANEN"))
			' nP_Pph21 		= nP_Pph21 		+ cdbl(rsPayroll("P_PPH21"))
			' nP_koperasi = nP_koperasi + cdbl(rsPayroll("P_KOPERASI"))
			' nP_Lain 		= nP_Lain 		+ cdbl(rsPayroll("P_LAIN"))
			' nTotal 			= nTotal 		+ cdbl(rsPayroll("TOTAL"))

			rsPayroll.MoveNext
		wend
	Else
	%>
	<tr>
		<td Colspan = "100">NO DATA AVAILABLE</td>
	</tr>
	<%End If%>
	<%
	rsPayroll.Close
	set rsPayroll = Nothing
	%>
	<tr style="background-color:#F4F4A0;">	 
		<td align="center" colspan="5" ><strong>Total</strong></td>
		<td align="right"><strong><%=formatnumber(nGet_Hk,1)%></strong></td>
		<td align="right"><strong><%=formatnumber(nGet_Hk_kj,1)%></strong></td>
		<td align="right"><strong><%=formatnumber(nget_mkp1,0)%></strong></td>
		<td align="right"><strong><%=formatnumber(nBase_Salary,0)%></strong></td>
		<td align="right"><strong><%=formatnumber(nrapelgaji,0)%></strong></td>
		<td align="right"><strong><%=formatnumber(nOvertime,0)%></strong></td>
		<td align="right"><strong><%=formatnumber(nPremi,0)%></strong></td>
		<td align="right"><strong><%=formatnumber(nTunjangan,0)%></strong></td>
		<td align="right"><strong><%=formatnumber(nBorongan,0)%></strong></td>
		<td align="right"><strong><%=formatnumber(nP_Jamsostek,0)%></strong></td>
		<td align="right"><strong><%=formatnumber(nP_bpjskes,0)%></strong></td>
		<td align="right"><strong><%=formatnumber(nP_Mk_P1,0)%></strong></td>
		<td align="right"><strong><%=formatnumber(nP_Panen,0)%></strong></td>
		<td align="right"><strong><%=formatnumber(nP_Pph21,0)%></strong></td>
		<td align="right"><strong><%=formatnumber(nP_koperasi,0)%></strong></td>
		<td align="right"><strong><%=formatnumber(nP_Lain,0)%></strong></td>
		<td align="right"><strong><%=formatnumber(nTotal,0)%></strong></td>
		<td align="right" colspan="2"><strong>&nbsp;</strong></td>
	</tr>
	</tbody>
</table>
<div id="footer_note">	
	<div class="block">Prepared by,
  <br/><br/><br/><br/>
		<input class="isian" type="text" value="">
  <hr/>
  <input class="isian" type="text" value="">
    </div>
    <div class="block">Checked by,
  <br/><br/><br/><br/>
		<input class="isian" type="text" value="">
  <hr/>
  <input class="isian" type="text" value="">
    </div>
    <div class="block">Approved by,
  <br/><br/><br/><br/>
  <input class="isian" type="text" value="<%=varManagerName%>">
  <hr/>
  <input class="isian" type="text" value="MANAGER">
    </div>
	<div class="block">Acknowledged by,
  <br/><br/><br/><br/>
  <input class="isian" type="text" value="" >
  <hr/>
  <input class="isian" type="text" value="">
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
   </body>
   </html>