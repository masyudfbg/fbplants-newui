<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
Dim v_txtYear, v_txtMonth, v_txtGangCode, v_txtstatus, v_txtsectionname
Dim ScriptName
Dim rsCatuBeras, strCatuBeras
Dim NO, cClass, gt_TotalPekerja, gt_TotalPasangan, gt_TotalTanggungan, gt_total

v_txtYear = Request("txtYear")
v_txtMonth = Request("txtMonth")
v_txtsectionname = Request("txtsectionname")
v_txtGangCode = Request("txtGangCode")
v_txtstatus = request("txtStatus")
v_txtfcba = request("txtfcba")
ScriptName = Request.ServerVariables("SCRIPT_NAME")

if len(trim(v_txtYear)) = 0 then v_txtYear = Year(Session("v_Employee"))
if len(trim(v_txtMonth)) = 0 then v_txtMonth = Month(Session("v_Employee"))

arr_potongan_eof = true
strCatuBeras = " SELECT DISTINCT fddate, employeecode, attendance" & _
               " FROM attendance_gad" & _
               " WHERE fcba = '" & v_txtfcba & "'" & _
               "       AND TO_CHAR(fddate, 'yyyymm') = '" & v_txtYear & "'||LPAD('" & v_txtMonth & "',2,0) AND attendance IN ('MK', 'P1')"
Set rsCatuBeras = Server.CreateObject("ADODB.Recordset")
rsCatuBeras.Open strCatuBeras, connstr, adOpenStatic 
if not rsCatuBeras.eof then
	arr_potongan = rsCatuBeras.getrows()
	arr_potongan_eof = false
end if

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
	FcbaName = rsBaInfo("FCBANAME")
	CompanyName = rsBaInfo("COMPANYNAME")
	varCompanyCode = rsBaInfo("fccompanycode")
Else
	FcbaName = ""
	CompanyName = ""
	varCompanyCode = ""
End If
rsBaInfo.Close
set rsBaInfo = Nothing

strDefaultBa = " SELECT" & _
				"     FCCODE as FCBA" & _
				" FROM BUSINESSUNIT BU" & _
				" INNER JOIN USER_FCBA_TAB UFT ON (BU.FCCODE = UFT.FCBA) " & _
				" WHERE" & _
				"     ISDEFAULT = 'TRUE'" & _
				"     AND FCTYPE = 'E'" & _
				"     AND FCUSERCODE = '" & Session("CodeName") & "'"

Set rsDefaultBa = Server.CreateObject("ADODB.Recordset")
rsDefaultBa.Open strDefaultBa, connstr, adOpenStatic

If Not rsDefaultBa.EOF Then
	varDefaultBa = rsDefaultBa("FCBA")
Else
	varDefaultBa = ""
End If
rsDefaultBa.Close
set rsDefaultBa = Nothing

If (v_txtFcba = "" Or Len(Trim(v_txtFcba)) = 0) Then
	v_txtFcba = varDefaultBa
End If

function data_hk(arr_, employeecode_)
	result = 0
	if not arr_potongan_eof then
		for xy=0 to ubound(arr_,2)
			if cstr(""&employeecode_) = cstr(""&arr_(1,xy)) then
				result = result +1
			end if
		next
	end if
	data_hk = result
end function

param_pembagi = 30
strCatuBeras = " SELECT NVL(rate_1,30) pembagi" & _
               " FROM parameter" & _
               " WHERE fcba = '" & v_txtfcba & "'" & _
               "       AND parcode = 'CATUBERAS' AND fccode = 'PEMBAGI'"
Set rsCatuBeras = Server.CreateObject("ADODB.Recordset")
rsCatuBeras.Open strCatuBeras, connstr, adOpenStatic 
if not rsCatuBeras.eof then
	param_pembagi = cdbl(rsCatuBeras("pembagi"))
end if

	
strCatuBeras	=	" select fcba," & _
					"        employeecode," & _
					"        employeename," & _
					"        sectionname," & _
					"        gangcode," & _
					"        position," & _
					"        presentgrade," & _
					"        maritalstatus," & _
					"        pekerja," & _
					"        pasangan," & _
					"        tanggungan," & _
					"        total" & _
					" from (" & _
					"     select em.fcba fcba," & _
					"            em.fccode employeecode," & _
					"            em.fcname employeename," & _
					"            em.sectionname sectionname," & _
					"            em.gangcode gangcode," & _
					"            ep.description position," & _
					"            em.presentgrade presentgrade," & _
					"            em.maritalstatus maritalstatus," & _
					"            prm.rate_2 pekerja," & _
					"            prm.rate_3 pasangan," & _
					"            prm.rate_4 tanggungan," & _
					"            prm.rate_5 total" & _
					"     from employee_hist_tab em, parameter prm, empposition ep" & _
					"     where (em.fcba = prm.fcba and em.maritalstatus = prm.fccode)" & _
					"           and (em.fcba = ep.fcba and em.position = ep.position)" & _
					"           and em.dateterminate is null and em.naturetype = 'NA'" & _
					"           and em.year = '"& v_txtYear &"' and em.Month = '"& v_txtMonth &"'" & _
					"     )" & _
					" where fcba = '" & v_txtfcba & "'"
					
If  Len(Trim(v_txtsectionname)) > 0 Then
	strCatuBeras	=	strCatuBeras & " and sectionname = '" & v_txtsectionname & "'"
End If
If  Len(Trim(v_txtGangCode)) > 0 Then
	strCatuBeras	=	strCatuBeras & " and gangcode = '" & v_txtGangCode & "'"
End If
If  Len(Trim(v_txtstatus)) > 0 Then
	strCatuBeras	=	strCatuBeras & " and presentgrade  = '" & v_txtstatus & "'"
End If

	strCatuBeras	=	strCatuBeras & " order by sectionname, employeecode"

' response.write strCatuBeras
Set rsCatuBeras = Server.CreateObject("ADODB.Recordset")
rsCatuBeras.Open strCatuBeras, connstr, adOpenStatic 
%>
<html>

<head>
<title>ePlaS Reporting - Rencana Kerja Bulanan</title>
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
		<form action="<%=ScriptName%>" method="post" name="frmEntry">
		<div class="printhide">
		<table id="form-table" class="centered-margin">
			<tr>
				<td><font face="verdana" size="1">FCBA</font></td>
				<td>
					<select name="txtFcba" id="vFcba" onchange="javascript:filtering()">
						<%
						Dim rsFcba
						Set rsFcba = Server.CreateObject("ADODB.Recordset")
						rsFcba.CursorLocation = adUseClient
						strsql =    " select fccode fcba from businessunit " & _
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
				mon = Month(Session("v_Costbook"))
				If len(v_txtMonth)=0 then v_txtMonth=mon

				Yr = Year(Session("v_Costbook"))
				MinYr = 2015
				MaxYr = Yr  + 2
				If len(v_txtYear)= 0 then v_txtYear = Yr
			%>
				<td><font face="verdana" size="1">YEAR</font></td>
				<td width="4">
					<select name="txtYear">
						<% For i = MinYr to MaxYr %>
							<Option Value="<%=i%>" <%If cint(i) = cint(v_txtYear) then%> Selected <%End If%>><%=i%></option>
						<% Next %>
					</select>
				</td>
				<td><font face="verdana" size="1">MONTH</font></td>
				<td width="3">
							<select name="txtMonth">
						<%
							For i = 1 to 12
						%>
								<Option Value="<%=i%>" <%If cint(i) = cint(v_txtMonth) then%> Selected <%End If%>><%=BulanAP(i)%></option>
						<%
							Next
						%>
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
				
				<td><font face="verdana" size="1">GANGCODE</font></td>
				<td><select name="txtGangCode" ><option value=""></option>
					<%
					Dim rsGangCode, strGangCode
					strGangCode	=	" SELECT FCCODE, FCNAME FROM GANG WHERE ACTIVATION = 'Y' AND FCBA = '" & Session("FCBA") & "' ORDER BY FCCODE ASC"
					
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
				<select name="txtStatus" >
					<option value="">--All--</option>
					<%
					strAfdeling	=	"SELECT distinct presentgrade FROM EMPLOYEE ORDER BY presentgrade"
					
					Set rsAfdeling = Server.CreateObject("ADODB.Recordset")
					rsAfdeling.Open strAfdeling, connstr, adOpenStatic 
					
					if not rsAfdeling.eof then
					Do While Not rsAfdeling.eof
					if rsAfdeling("presentgrade") = v_txtStatus then%>
					<option value="<%=rsAfdeling("presentgrade")%>" Selected><%=rsAfdeling("presentgrade")%> </option> 
					<%Else%>
					<option value="<%=rsAfdeling("presentgrade")%>"><%=rsAfdeling("presentgrade")%></option>  
					<%End if
						rsAfdeling.movenext
					loop
					end if
					rsAfdeling.Close
					set rsAfdeling = Nothing
					%>
				</select>
			</td> 
			<td><input type="submit" value="View" /></td>
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
		<div id="head_judul" style="margin-bottom:-30px; position:static; text-align:left;">
		<h4>PT. <%=companyname%></h4>
		<h5><%=fcbaname%></h5>
	</div>
	<div id="subheader" class="centered">
		<h2>Rencana Kerja Bulanan</h2>
		<br/>
		<h4>PERIODE&nbsp;:&nbsp;<%=Bulan(v_txtMonth)%>&nbsp;<%=v_txtYear%></h4>
	</div>
	</div>
 </th>
  </tr>
	<tr>
		<!--<th rowspan="2">NO</th>-->
		<th rowspan="4">Jobcode</th>
		<th rowspan="4">Job Name</th>
		<th rowspan="4">UOM</th>
		<th colspan="5">RKB Progress Fisik</th>
		<th colspan="10">RKB Progress Biaya</th>
	</tr>
	<tr>
		<th rowspan="3">Budget</th>
		<th colspan="2">Realisasi</th>
		<th>RKB</th>
		<th rowspan="3">Saldo Budget <%=v_txtYear%></th>
		<th rowspan="2" colspan="2">Budget <%=v_txtYear%></th>
		<th colspan="4">Realisasi</th>
		<th rowspan="2" colspan="2">RKB Bulan Ini</th>
		<th rowspan="2" colspan="2">Saldo Budget <%=v_txtYear%></th>
	</tr>
	<tr>
		<th rowspan="2"><%=MonthName(v_txtmonth)%></th>
		<th rowspan="2">Jan - <%=MonthName(v_txtmonth)%></th>
		<th rowspan="2">This Month</th>
		<th colspan="2"><%=MonthName(v_txtmonth)%></th>
		<th colspan="2">Jan - <%=MonthName(v_txtmonth)%></th>
	</tr>
	<tr>
		<th>Upah</th>
		<th>Bahan</th> 
		<th>Upah</th>
		<th>Bahan</th> 
		<th>Upah</th>
		<th>Bahan</th> 
		<th>Upah</th>
		<th>Bahan</th> 
		<th>Upah</th>
		<th>Bahan</th>  
	</tr>
</thead>
<tbody>
<%if rsCatuBeras.BOF and rsCatuBeras.EOF then%>
	<tr>
		<td colspan="14" align="left">NO DATA AVAILABLE</td>
	</tr>
<%end if
		NO = 1
		gt_TotalPekerja = 0
		gt_TotalPasangan = 0
		gt_TotalTanggungan = 0
		gt_total = 0
		r_mkp1 = 0
		r_pot_beras_kg = 0
		r_total_kg = 0
		while not rsCatuBeras.eof
			r_mkp1 = data_hk(arr_potongan, rsCatuBeras("EMPLOYEECODE"))
			r_pot_beras_kg = 1/param_pembagi *cdbl(to_nol(rsCatuBeras("PEKERJA"))) *r_mkp1
			r_total_kg = cdbl(to_nol(rsCatuBeras("TOTAL"))) - r_pot_beras_kg
%>
			<tr class="<%=cClass%>">
				<!--<td align="center" height="30" width="3%"><%=NO%></td>-->
				<td align="left"><%=rsCatuBeras("SECTIONNAME")%></td>
				<td align="center" width="4%"><%=rsCatuBeras("EMPLOYEECODE")%></td>
				<td align="left"><%=rsCatuBeras("EMPLOYEENAME")%></td>
				<td align="center"><%=rsCatuBeras("PRESENTGRADE")%></td>
				<td align="center"><%=rsCatuBeras("MARITALSTATUS")%></td>
				<td align="right" width="4%"><%=FbpNum2(rsCatuBeras("PEKERJA"))%></td>
				<td align="right" width="4%"><%=FbpNum2(rsCatuBeras("PASANGAN"))%></td>
				<td align="right" width="4%"><%=FbpNum2(rsCatuBeras("TANGGUNGAN"))%></td>
				<td align="right" width="4%"><%=FbpNum2(rsCatuBeras("TOTAL"))%></td>
				<td align="right" width="4%"><%=FbpNum2(r_mkp1)%></td>
				<td align="right" width="4%"><%=FbpNum2(r_pot_beras_kg)%></td>
				<td align="right" width="4%"><%=FbpNum2(r_total_kg)%></td>
				
			<%if no  mod 2 = 1 then%>
				   <td align="left" width="100" bgcolor="#ccc"><font face="Verdana" size="1">&nbsp;</font></td>			
				   <td align="left" width="100" bgcolor="#fff"><font face="Verdana" size="1">&nbsp;</font></td>			
			   <%Else%>
				   <td align="left" width="100" bgcolor="#fff"><font face="Verdana" size="1">&nbsp;</font></td>			
				   <td align="left" width="100" bgcolor="#ccc"><font face="Verdana" size="1">&nbsp;</font></td>			
			   <%End if%>
			</tr>
<%
			NO = NO + 1
			
			gt_TotalPekerja = gt_TotalPekerja + Cdbl(rsCatuBeras("PEKERJA"))
			gt_TotalPasangan = gt_TotalPasangan + Cdbl(rsCatuBeras("PASANGAN"))
			gt_TotalTanggungan = gt_TotalTanggungan + Cdbl(rsCatuBeras("TANGGUNGAN"))
			
			gt_mkp1 = gt_mkp1 +r_mkp1
			gt_pot_beras_kg = gt_pot_beras_kg +r_pot_beras_kg
			gt_total_kg = gt_total_kg +r_total_kg
			gt_total = gt_total + Cdbl(rsCatuBeras("TOTAL"))
			rsCatuBeras.MoveNext
		Wend
	rsCatuBeras.Close
	set rsCatuBeras = Nothing
	
			gt_TotalPekerja = Round(gt_TotalPekerja, 1)
			gt_TotalPasangan = Round(gt_TotalPasangan, 1)
			gt_TotalTanggungan = Round(gt_TotalTanggungan, 1)
			gt_total = Round(gt_total, 1)
%>
			<tr bgcolor="#F4F4A0">
				<td align="center" colspan="5"><strong>TOTAL</strong></td>
				<td align="right"><strong><%=FbpNum2(gt_TotalPekerja)%></strong></td>
				<td align="right"><strong><%=FbpNum2(gt_TotalPasangan)%></strong></td>
				<td align="right"><strong><%=FbpNum2(gt_TotalTanggungan)%></strong></td>
				<td align="right"><strong><%=FbpNum2(gt_total)%></strong></td>
				<td align="right"><strong><%=FbpNum2(gt_mkp1)%></strong></td>
				<td align="right"><strong><%=FbpNum2(gt_pot_beras_kg)%></strong></td>
				<td align="right"><strong><%=FbpNum2(gt_total_kg)%></strong></td>
				<td colspan="2">&nbsp;</td>
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
        <input class="isian" type="text" value="<%=Session("FCMANAGER")%>">
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
   <table width="90%" class="data">
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
		<td width="50%" align="center"><font face="Verdana" size="2"><strong><u><%=Session("FCMANAGER")%></u></strong></td>
		<td width="50%" align="center"><font face="Verdana" size="2"><strong><u><%=Session("KTU")%></u></strong></td>
	   </tr>
	   	  <tr> 
		<td width="50%" align="center"><font face="Verdana" size="1"><strong><i>MANAGER</i></strong></td>
		<td width="50%" align="center"><font face="Verdana" size="1"><strong><i>KTU</i></strong></td>
	   </tr>
   </table>
-->
   </body>
   </html>
