<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/ConndB1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%
Dim MM, yyyy, mon, Yr, MinYr, MaxYr, i, rsBa, strsql,strsql1,strsql2,strsql3,strsql4,strsql5,strsql6,strsql7,strsql8
Dim ScriptName, v_txtGangCode

v_txtGangCode	=	trim(Request("txtGangCode"))
ScriptName	=	Request.ServerVariables("SCRIPT_NAME")
%>
<html>

<head>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>iPla'S Reporting - Rekap Payroll</title>
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
</head>

<body>
	<div id="wrapper">
	<div id="header-wrapper">

		<!--#include file="menutop.asp"-->

<table id="form-table" class="centered-margin">
	<tr>
	<form name="frmAR" method="post" action="<%=ScriptName%>">
		<td><font face="verdana" size="1">TAHUN</font></td>
		<td>
			<font size="2"  color="#0000FF" ></font>
<% 
MM=Request.Form("vMonth")
yyyy=Request.Form("vYear")

mon = Month(Session("v_Costbook"))
If len(MM)=0 then MM=mon

Yr = Year(Session("v_Costbook"))
MinYr = Yr - 2
MaxYr = Yr  + 2
If len(yyyy)= 0 then yyyy = Yr
%>
			<select name="vYear">
				<%
					For i = MinYr to MaxYr
				%>
						<Option Value="<%=i%>" <%If cint(i) = cint(yyyy) then%> Selected <%End If%>><%=i%></option>
				<%
					Next
				%>
			</select>
		</td>
		<td><font face="verdana" size="1">BULAN</font></td>
		<td>
			<font size="2" color="#0000FF" ></font>
				<select name="vMonth">
					<%
						For i = 1 to 12
					%>
							<Option Value="<%=i%>" <%If cint(i) = cint(MM) then%> Selected <%End If%>><%=MonthName(i)%></option>
					<%
						Next
					%>
				</select>
		</td>		

		<td><font face="verdana" size="1">GANGCODE</font></td><td><select name="txtGangCode" ><option value=""></option>
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
		</select></td>


						
		<td><input type="submit" value="View" /></td>
		</form>
	</tr>
</table>

<%
	Dim rsqGang, qGang, txtqGang
	qGang	=	" SELECT FCNAME FROM GANG WHERE FCBA = '" & Session("FCBA") & "' AND FCCODE = '"& v_txtGangCode & "'"
	
	Set rsqGang = Server.CreateObject("ADODB.Recordset")
	rsqGang.Open qGang, connstr, adOpenStatic
	
	if not rsqGang.eof then
		txtqGang	=	rsqGang("FCNAME")
	end if
		rsqGang.Close
		set rsqGang = Nothing
%>
		<div id="header">
			<h4>PT. <%=Session("COMPANY")%></h4>
			<h5><%=Session("BA")%></h5>
		</div>
		<div id="subheader" class="centered">
			<% Dim ia,rekap
		        ia = request.Form("vMonth")
			%>	
			<h2>DAFTAR UPAH KARYAWAN</h2>
			<% IF request.Form("vMonth") <> "" then %>
			<h5>PERIODE     : <%=MonthName(cint(ia))%> &nbsp; <%=request.Form("vYear")%></h5>
			<% END IF %>
			<% IF request.Form("txtGangCode") <> "" then %>
			<h4>GANG : <%=request.Form("txtGangCode")%> - <%=txtqGang%></h4>
			<%END IF%>
			 
		</div>
	</div>
	
<table border="0" width="100%" id="product-table">
     <tr>
		<th rowspan="2" style="color:#fff;">NO</th>
		<th rowspan="2" style="color:#fff;">NIK</th>		
		<th rowspan="2" style="color:#fff;">NAMA KARYAWAN</th>
		<!--<th rowspan="2" style="color:#fff;">JABATAN</th>-->
		<th rowspan="2" style="color:#fff;">STATUS</th>
		<th rowspan="2" style="color:#fff;">HK</th>
		<th rowspan="2" style="color:#fff;">GAJI</th>
		<th rowspan="2" style="color:#fff;">LEMBUR</th>
		<th rowspan="2" style="color:#fff;">PREMI</th>
		<th rowspan="2" style="color:#fff;">TUNJANGAN</th>
		<th rowspan="2" style="color:#fff;">BORONGAN</th>		
		<th colspan="5" style="color:#fff;">POTONGAN</th>		
		<th rowspan="2" style="color:#fff;">TOTAL</th>
		<th rowspan="2" colspan="2" style="color:#fff;">TANDA TANGAN</th>
	</tr>
	<tr>		
		<th style="color:#fff;">Jamsostek</th>
		<th style="color:#fff;">MK/P1</th>
		<th style="color:#fff;">Panen</th>
		<th style="color:#fff;">PPh21</th>
		<th style="color:#fff;">Lain-Lain</th>		
	</tr>
	<%
	Dim rsPayroll, strPayroll, noline 
	Dim nGet_Hk, nBase_Salary, nOvertime, nPremi, nTunjangan, nBorongan, nP_Jamsostek, nP_Mk_P1, nP_Panen, nP_Pph21, nP_Lain, nTotal
			strPayroll	=	" SELECT T.YEARMONTH, T.FCBA, T.FCCODE, T.FCNAME, EPOS.DESCRIPTION, T.PRESENTGRADE," & _
						" SUM(GET_HK) GET_HK, SUM(T.BASE_SALARY) BASE_SALARY, SUM(T.OVERTIME) OVERTIME, SUM(T.PREMI) PREMI," & _
						" SUM(T.TUNJANGAN) TUNJANGAN,SUM(T.BORONGAN) BORONGAN,SUM(T.P_JAMSOSTEK) P_JAMSOSTEK," & _
						" SUM(T.P_MK_P1) P_MK_P1, SUM(T.P_PANEN) P_PANEN, SUM(T.P_PPH21) P_PPH21, SUM(T.P_LAIN) P_LAIN," & _
						" SUM(T.BASE_SALARY + T.OVERTIME + T.PREMI + T.TUNJANGAN + T.BORONGAN + T.P_JAMSOSTEK + T.P_MK_P1 + T.P_PANEN + T.P_PPH21 + T.P_LAIN ) TOTAL " & _
						" FROM VIEW_PAYROLLDETAIL T INNER JOIN EMPLOYEE EM ON (T.FCBA = EM.FCBA AND T.FCCODE = EM.FCCODE)" & _
						" INNER JOIN EMPPOSITION EPOS ON (EM.FCBA = EPOS.FCBA AND EM.POSITION = EPOS.POSITION)" & _
						" WHERE  TO_NUMBER(SUBSTR(T.YEARMONTH,1,4)) = " & YYYY & "" & _
						" AND TO_NUMBER(SUBSTR(T.YEARMONTH,5,2)) = " & MM & " AND T.FCBA = '" & SESSION("FCBA") & "'"
		if  len(trim(v_txtGangCode)) > 0 then
			strPayroll	=	strPayroll &  " AND T.GANGCODE = '" & v_txtGangCode & "'"
		end if 

			strPayroll	=	strPayroll & 	" GROUP BY T.YEARMONTH, T.FCBA, T.FCCODE, T.FCNAME, EPOS.DESCRIPTION, T.PRESENTGRADE" & _
									" ORDER BY T.YEARMONTH, T.FCBA, T.FCCODE" 
	
	Set rsPayroll = Server.CreateObject("ADODB.Recordset")
	rsPayroll.Open strPayroll, connstr, adOpenStatic
    

	if Not rsPayroll.eof then
		noline = 1
		nGet_Hk = 0
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

		while Not rsPayroll.eof  %>
		           
		  <tr>
			<td align="center"><%=noline%></td>
			<td align="left"><%=rsPayroll("FCCODE")%></td>
			<td align="left"><%=rsPayroll("FCNAME")%></td>
			<!--<td align="left"><%=rsPayroll("DESCRIPTION")%></td>-->
			<td align="center"><%=rsPayroll("PRESENTGRADE")%></td>
			<td align="right"><%=formatnumber( cdbl(rsPayroll("GET_HK")),1)%></td>
			<td align="right"><%=formatnumber( cdbl(rsPayroll("BASE_SALARY")),0)%></td>
			<td align="right"><%=formatnumber( cdbl(rsPayroll("OVERTIME")),0)%></td>
			<td align="right"><%=formatnumber( cdbl(rsPayroll("PREMI")),0)%></td>
			<td align="right"><%=formatnumber( cdbl(rsPayroll("TUNJANGAN")),0)%></td>
			<td align="right"><%=formatnumber( cdbl(rsPayroll("BORONGAN")),0)%></td>
			<td align="right"><%=formatnumber( cdbl(rsPayroll("P_JAMSOSTEK")),0)%></td>
			<td align="right"><%=formatnumber( cdbl(rsPayroll("P_MK_P1")),0)%></td>
			<td align="right"><%=formatnumber( cdbl(rsPayroll("P_PANEN")),0)%></td>
			<td align="right"><%=formatnumber( cdbl(rsPayroll("P_PPH21")),0)%></td>
			<td align="right"><%=formatnumber( cdbl(rsPayroll("P_LAIN")),0)%></td>
			<td align="right"><%=formatnumber( cdbl(rsPayroll("TOTAL")),0)%></td>
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
			nGet_Hk 		= nGet_Hk 		+ cdbl(rsPayroll("GET_HK"))
			nBase_Salary 	= nBase_Salary 	+ cdbl(rsPayroll("BASE_SALARY"))
			nOvertime 		= nOvertime 	+ cdbl(rsPayroll("OVERTIME"))
			nPremi 			= nPremi 		+ cdbl(rsPayroll("PREMI"))
			nTunjangan 		= nTunjangan 	+ cdbl(rsPayroll("TUNJANGAN"))
			nBorongan 		= nBorongan 	+ cdbl(rsPayroll("BORONGAN"))
			nP_Jamsostek 	= nP_Jamsostek 	+ cdbl(rsPayroll("P_JAMSOSTEK"))
			nP_Mk_P1 		= nP_Mk_P1 		+ cdbl(rsPayroll("P_MK_P1"))
			nP_Panen 		= nP_Panen 		+ cdbl(rsPayroll("P_PANEN"))
			nP_Pph21 		= nP_Pph21 		+ cdbl(rsPayroll("P_PPH21"))
			nP_Lain 		= nP_Lain 		+ cdbl(rsPayroll("P_LAIN"))
			nTotal 			= nTotal 		+ cdbl(rsPayroll("TOTAL"))

			rsPayroll.MoveNext
		wend
	end if
		%> 
	<tr style="background-color:#F4F4A0;">	 
		<td align="center" colspan="4" ><strong>Total</strong></td>
		<td align="right"><strong><%=formatnumber(nGet_Hk,1)%></strong></td>
		<td align="right"><strong><%=formatnumber(nBase_Salary,0)%></strong></td>
		<td align="right"><strong><%=formatnumber(nOvertime,0)%></strong></td>
		<td align="right"><strong><%=formatnumber(nPremi,0)%></strong></td>
		<td align="right"><strong><%=formatnumber(nTunjangan,0)%></strong></td>
		<td align="right"><strong><%=formatnumber(nBorongan,0)%></strong></td>
		<td align="right"><strong><%=formatnumber(nP_Jamsostek,0)%></strong></td>
		<td align="right"><strong><%=formatnumber(nP_Mk_P1,0)%></strong></td>
		<td align="right"><strong><%=formatnumber(nP_Panen,0)%></strong></td>
		<td align="right"><strong><%=formatnumber(nP_Pph21,0)%></strong></td>
		<td align="right"><strong><%=formatnumber(nP_Lain,0)%></strong></td>
		<td align="right"><strong><%=formatnumber(nTotal,0)%></strong></td>
		<td align="right" colspan="2"><strong>&nbsp;</strong></td>
	</tr>
</table>

   <table width="90%" class="data">
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
		<td width="50%" align="center"><font face="Verdana" size="2"><strong><u><%=Session("FCMANAGER")%></u></strong></td>
		<td width="50%" align="center"><font face="Verdana" size="2"><strong><u><%=Session("KTU")%></u></strong></td>
	   </tr>
	   	  <tr> 
		<td width="50%" align="center"><font face="Verdana" size="1"><strong><i>MANAGER</i></strong></td>
		<td width="50%" align="center"><font face="Verdana" size="1"><strong><i>KTU</i></strong></td>
	   </tr>
   </table>
   </body>
   </html>