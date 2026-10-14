<%@Language=VBScript %>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
Dim v_txtYear, v_txtMonth, v_txtGangCode, v_txtGrade,v_txtComp,n,V_Periode
Dim txtYear, txtMonth, txtGangCode, txtGrade,txtComp,txtFcba, txtSelect
Dim ScriptName
Dim rsAFD,strAFD,sqlAFD,sqlsummAFD,sqlTOTAFD,sqlsummTOTAFD,numAFD,nAFD(),nfieldAFD(),ntotAFD(100),ntotAFDD(100),ntotAFDK(100)
Dim rsTBS, strTBS,rsTBSD, strTBSD,rsTBSK, strTBSK, rsAtt, strAtt, rsKJ, strKJ,strcomp,strSelect,strGroup,strSummary,strsummSelect
Dim SprevCompany,SprevRegion,SprevGroupRegion,SprevPlantation, strAFDDebit ,strAFDKredit, rsAFDDebit, rsAFDKredit	
Dim tempfcba,tempCompany,tempRegion,tempGroupRegion,tempPlantation
Dim NO, cClass, totalDay,nselect
Dim nAFD1
Dim nAFD2
Dim nAFD3
Dim nAFD4
Dim nAFD5
Dim nAFD6
Dim nAFD7
Dim nAFD8
Dim nH1
Dim nH2
Dim nCT
Dim nMK
Dim nML
Dim nWH
Dim nWS
Dim nUN
Dim nKK
Dim nTOTAL, nTOTALD, nTOTALK




txtComp = Request("txtComp")
txtSelect = Request("txtSelect")
v_txtYear	=	Request("txtYear")
v_txtMonth	=	Request("txtMonth")
if Len(Trim(v_txtYear)) = 0 then
	v_txtYear	= Year(Session("v_Costbook"))
end if
if Len(Trim(v_txtMonth)) = 0 then
	v_txtMonth	= Month(Session("v_Costbook"))
end if

V_Periode =  v_txtYear & right("00" & v_txtMonth,2)


v_txtComp 		=	txtComp


ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

strAFD = " SELECT dynamicquery_api.dynamic_CostActivity_Periode('"& txtComp &"','1114000','ALL','" & V_Periode &"') queryy from dual "
Set rsAFD = Server.CreateObject("ADODB.Recordset")
rsAFD.Open strAFD, connstr, adOpenStatic
if NOT rsAFD.eof then
   
   strTBS = rsAFD("queryy")
   
   
end if
rsAFD.Close
set rsAFD = Nothing

strAFDDebit =  "SELECT dynamicquery_api.dynamic_CostActivity_Periode('"& txtComp &"','1114000','DEBET','" & V_Periode &"') queryy from dual"
Set rsAFDDebit = Server.CreateObject("ADODB.Recordset")
rsAFDDebit.Open strAFDDebit, connstr, adOpenStatic
if NOT rsAFDDebit.eof then
   
   strTBSD = rsAFDDebit("queryy")
   
   
end if
rsAFDDebit.Close
set rsAFDDebit = Nothing

strAFDKredit=  " SELECT dynamicquery_api.dynamic_CostActivity_Periode('"& txtComp &"','1114000','KREDIT','" & V_Periode &"') queryy from dual "
Set rsAFDKredit= Server.CreateObject("ADODB.Recordset")
rsAFDKredit.Open strAFDKredit, connstr, adOpenStatic
if NOT rsAFDKredit.eof then
   
   strTBSK = rsAFDKredit("queryy")
   
   
end if
rsAFDKredit.Close
set rsAFDKredit = Nothing


dim kj
dim CountFcba, rsCountFcba, coloumnname , rscoloumn
dim kjTot
coloumnname = "select FCCODE from businessunit a where a.fccompanycode = '"& txtComp &"' order by fccode"
CountFcba = " select count (*) total from businessunit a where a.fccompanycode = '"& txtComp &"'"
Set rsCountFcba = Server.CreateObject("ADODB.Recordset")
'response.write coloumnname
rsCountFcba.Open CountFcba, connstr, adOpenStatic
if NOT rsCountFcba.eof then
   
   kjTot = rsCountFcba("total")
   
   
end if
rsCountFcba.Close
set rsCountFcba = Nothing
  
%>		  
<html>

<head>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>ePlaS Reporting - Report Intraco</title>
<script language="JavaScript" src='/js/x.js'></script>
<script language="JavaScript" src='/js/tall_xtableheaderfixed.js'></script>
<script language="JavaScript">
xAddEventListener(window, 'load',
  function() {
    new xTableHeaderFixed('xthf-site', window );
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
@media print
{
body{font-family:'Arial Narrow';font-size:12px;}
.printhide{display:none;}
}
REM #wrapper,body,html{
REM display:table;
REM min-width:100%
REM }

</style>

</head>

<body id="both">
	<div id="wrapper" >
	<!--#include file="../print_excel.asp"-->
	<div id="header-wrapper">
	<div class="printhide">
	<!--#include file="menutop.asp"-->
	</div>
	
	</div>
<div class="printhide">
<table id="form-table" class="centered-margin">
	<form action="<%=ScriptName%>" method="post" name="frmEntry">
	<tr>
	            <td><font face="verdana" size="1">COMPANY</font></td>
				<td>
					<select name="txtComp">
					<%
					Dim rsComp,strsql
					Set rsComp = Server.CreateObject("ADODB.Recordset")
					rsComp.CursorLocation = adUseClient
					strsql = " SELECT '----CHOOSE-----' as companycode from DUAL UNION " & _
						     " select FCCOMPANYCODE companycode from businessunit where USER_COMPANY_API.IsAuthorizedFcba('" & Session("CodeName") & "', fccode) = 'TRUE'"
								         


					'select distinct USER_COMPANY_API.GetCompanyCode('" & Session("CodeName") & "',t.fcba) fccompanycode from USER_FCBA_TAB t where fcusercode= '" & Session("CodeName") & "' and USER_COMPANY_API.GetCompanyCode('" & Session("CodeName") & "',t.fcba) is not null )"
					
					
					rsComp.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
					if not rsComp.eof then
					DO while not rsComp.eof
					   if v_txtComp = rsComp("companycode") then
					%>			
					    <option value="<%=rsComp("companycode")%>" Selected><%=rsComp("companycode")%></option>
					<% Else %>	
					    <option value="<%=rsComp("companycode")%>"><%=rsComp("companycode")%></option>
					<% End If
					
					rsComp.movenext
					loop
					End if
					rsComp.Close
					set rsComp = Nothing
					%>
					</select>
				</td>	
				<td><font face="verdana" size="1">YEAR</font></td><td width="4"><input type="text" name="txtYear" width="4" value="<%=v_txtYear%>"></td>
				<td><font face="verdana" size="1">MONTH</font></td><td width="2"><input type="text" name="txtMonth" width="2" value="<%=v_txtMonth%>"></td>
				<td><font face="verdana" size="1">VIEW</font></td>
				      <td width="4">
						<select name="txtSelect">
							
							<Option Value="ALL" Selected > Db-Cr With Total </option>
							<Option Value="Db-Cr"  > Db-Cr </option>
							<Option Value="Db"  > Db</option>
							<Option Value="Cr"  > Cr</option>
							<Option Value="JustTotal"  > Just Total</option>
						</select>
					</td>
		 <td><input type="submit" value="View" /></td>
	</tr>
	</form>
</table>
</div>
<table border="0" width="100%" id="product-table" class="xthf-site">
<thead>
	<tr>
	<%
	dim colspantot
	colspantot =0
	if txtSelect="ALL" then 
	colspantot = CInt(kjTot )*3 +3
	elseif txtSelect="Db-Cr" then
	colspantot = CInt(kjTot )*2 +3
	else
	colspantot = CInt(kjTot ) + 3
	end if
	%>
    <th colspan="<%= colspantot %>" style="background-color:#FFFFFF; border:none;">  
	<div id="header-wrapper">
		<div id="head_judul">
		<h4>PT. <%=Session("COMPANY")%></h4>
		<h5><%=Session("BA")%></h5>
	    </div>
	<div id="subheader" class="centered" style="width:50%; margin:auto;">
		<h2>LAPORAN INTRACO</h3>
		<br/>
		<h4>PERIODE&nbsp;:&nbsp;<%=Bulan(v_txtMonth)%>&nbsp;<%=v_txtYear%></h4>
		<h4>COMPANY&nbsp;:&nbsp;<%=v_txtComp%>&nbsp;&nbsp;&nbsp;</h4>
	</div>
	</div>
</th>
  </tr>
	<tr>
		<th style="color:#FFFFFF" rowspan="3" width="10%">JOBCODE</th>
		<th style="color:#FFFFFF" rowspan="3" width="300">JOBNAME</th>
		 <%
 if txtSelect = "ALL" then
	 %>
	 <th style="color:#FFFFFF" colspan="<%=CInt(kjTot)*3 %>" width="10%">BUSINESSUNIT</th>
<%
elseif txtSelect = "Db-Cr" then
%>
	<th style="color:#FFFFFF" colspan="<%=CInt(kjTot)*2 %>" width="10%">BUSINESSUNIT</th>
<%
else
%>
<th style="color:#FFFFFF" colspan="<%=CInt(kjTot) %>" width="10%">BUSINESSUNIT</th>
<%
end if
%>
		
		<th style="color:#FFFFFF" rowspan="3" width="10%">SALDO</th>
	</tr>
	<tr>

<% 
	Set rscoloumn = Server.CreateObject("ADODB.Recordset")
	rscoloumn.Open coloumnname, connstr, adOpenStatic
	if not rscoloumn.eof then
	DO while not rscoloumn.eof
	%>
		 <%
 if txtSelect = "ALL" then
	 %>
	 <th style="color:#FFFFFF" colspan="3" width="10%"><%=rscoloumn("FCCODE") %></th>
<%
elseif txtSelect = "Db-Cr" then
%>
<th style="color:#FFFFFF" colspan="2" width="10%"><%=rscoloumn("FCCODE") %></th>
<%
elseif txtSelect = "Db" then
%>
<th style="color:#FFFFFF" colspan="1" width="10%"><%=rscoloumn("FCCODE") %></th>
<%
elseif txtSelect = "Cr" then
%>
<th style="color:#FFFFFF" colspan="1" width="10%"><%=rscoloumn("FCCODE") %></th>
<%
else
%>
<th style="color:#FFFFFF"  width="10%"><%=rscoloumn("FCCODE") %></th>
<%
end if

		rscoloumn.MoveNext
		loop
		end if
rscoloumn.Close
set rscoloumn = Nothing
		%>
	 </tr>
	 <%
	 if txtSelect = "ALL" then
	 %>
	 <tr>
	<% For kj = 1 to CInt(kjTot) step 1 %>
		<th style="color:#FFFFFF" width="10%">DEBET</th>
		<th style="color:#FFFFFF" width="10%">KREDIT</th>
		<th style="color:#FFFFFF" width="10%">TOTAL(D+K)</th>
		<%next%>
	 </tr>
<%
elseif txtSelect = "Db-Cr" then 
%>
<tr>
	<% For kj = 1 to CInt(kjTot) step 1 %>
		<th style="color:#FFFFFF" width="10%">DEBET</th>
		<th style="color:#FFFFFF" width="10%">KREDIT</th>
		<%next%>
	 </tr>
<%
elseif txtSelect = "Db" then 
%>
<tr>
	<% For kj = 1 to CInt(kjTot) step 1 %>
		<th style="color:#FFFFFF" width="10%">DEBET</th>
		<%next%>
	 </tr>
<%
elseif txtSelect ="Cr" then
%>
	<tr>
	<% For kj = 1 to CInt(kjTot) step 1 %>
		<th style="color:#FFFFFF" width="10%">KREDIT</th>
		<%next%>
	 </tr>
<%
else

end if
%>
</thead>
<tbody>
 
<%
dim xtotal
dim xtotalK
dim xtotalD
dim nDebit, nKREDIT
Set rsTBSD = Server.CreateObject("ADODB.Recordset")
'response.write  strTBSD & "<br>"
rsTBSD.Open strTBSD, connstr, adOpenStatic
Set rsTBSK = Server.CreateObject("ADODB.Recordset")
'response.write  strTBSK & "<br>"
rsTBSK.Open strTBSK, connstr, adOpenStatic
Set rsTBS = Server.CreateObject("ADODB.Recordset")
'response.write  strTBS & "<br>"
rsTBS.Open strTBS, connstr, adOpenStatic
n=0
if txtSelect = "ALL" then
n=0
if rsTBS.bof and rsTBS.eof then
n=-1
%>
	<tr>
		<td colspan="12">NO DATA AVAILABLE</td>
	</tr>
<%end if
		NO = 1
		while not rsTBS.eof and not rsTBSD.eof and not rsTBSK.eof
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
			
%>
			<tr class="<%=cClass%>">
								   
				   <td align="center"><%=rsTBS("JOBCODE")%></td>
				 <td align="center"><a href="P_ReportIntraco.asp?txtComp=<%="txtComp"%>">
								   <%=rsTBS("JOBNAME")%>
								   </a>
							 </td>
				<% 
			xtotalD = 0
			xtotalK = 0		
			xtotal = 0
				For kj = 1 to CInt(kjTot) step 1 
				if rsTBSD("JOBCODE") = rsTBS("JOBCODE") then %>					
					<td align="right"><%=FormatNumber(rsTBSD.fields(1 + kj),0)%></td>
					<%
              				ntotAFDD(kj)	= Cdbl(ntotAFDD(kj)) + Cdbl(rsTBSD.fields(1 + kj))
							xTOTALD = xTOTALD + Cdbl(rsTBSD.fields(1 + kj))
							nTOTALD =  nTOTALD + Cdbl(rsTBSD.fields(1 + kj))
							nDebit = 0
							else
							%>
					
					<td align="right">0</td>
					<%
							nDebit = -1
							end if
						if rsTBSK("JOBCODE") = rsTBS("JOBCODE") then
					%>
					
					<td align="right"><%=FormatNumber(rsTBSK.fields(1 + kj),0)%></td>
					<%
              				ntotAFDK(kj)	= Cdbl(ntotAFDK(kj)) + Cdbl(rsTBSK.fields(1 + kj))
							xTOTALK = xTOTALK + Cdbl(rsTBSK.fields(1 + kj))
							nTOTALK =  nTOTALK+ Cdbl(rsTBSK.fields(1 + kj))
							nKREDIT = 0
							else
							%>
					
					<td align="right">0</td>
					<%
							nKREDIT = -1
							end if
					%>	
					<td align="right"><%=FormatNumber(rsTBS.fields(1 + kj),0)%></td>
					<%
              				ntotAFD(kj)	= Cdbl(ntotAFD(kj)) + Cdbl(rsTBS.fields(1 + kj))
							xTOTAL = xTOTAL + Cdbl(rsTBS.fields(1 + kj))
							nTOTAL =  nTOTAL + Cdbl(rsTBS.fields(1 + kj))
					%>
				<% Next %>
				
				<td align="right"><%=FormatNumber(Cdbl(xtotal))%></td>
				</tr>
<%
			NO = NO + 1
			
			if nDebit = 0 then
			rsTBSD.MoveNext 
			end if
			if nKREDIT = 0 then
			rsTBSK.MoveNext 
			end if
			rsTBS.MoveNext
		Wend
rsTBSD.Close
set rsTBSD = Nothing
rsTBSK.Close
set rsTBSK = Nothing
rsTBS.Close
set rsTBS = Nothing
%>
			<tr bgcolor="#F4F4A0">
				      <% if n=-1 then  %>
				    <td align="center" colspan="3"><strong>TOTAL</td>
				<% else %>
				     <td align="center" colspan="2"><strong>TOTAL</td>
				<% end if %>
				<% For kj = 1 to CInt(kjTot) step 1 %>					
				<td align="right"><%=FormatNumber(Cdbl(ntotAFDD(kj)),0)%></td>
				<td align="right"><%=FormatNumber(Cdbl(ntotAFDK(kj)),0)%></td>
					<td align="right"><%=FormatNumber(Cdbl(ntotAFD(kj)),0)%></td>
				<% Next %>
					<td align="right"><%=FormatNumber(Cdbl(nTOTAL),0)%></td>
			</tr>
<%
elseif txtSelect = "Db-Cr" then
n=0
if rsTBSD.bof and rsTBSD.eof then
%>
	<tr>
		<td colspan="12">NO DATA AVAILABLE</td>
		<% n=-1 %>
	</tr>
<%end if
		NO = 1
		while not rsTBSD.eof and not rsTBSK.eof
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
			
%>
			<tr class="<%=cClass%>">
								   
				   <td align="center"><%=rsTBSD("JOBCODE")%></td>
				 <td align="center"><a href="P_ReportIntraco.asp?txtComp=<%="txtComp"%>">
								   <%=rsTBSD("JOBNAME")%>
								   </a>
							 </td>
				<% 
			xtotalD = 0
			xtotalK = 0		
			xtotal = 0
				For kj = 1 to CInt(kjTot) step 1  %>
					<td align="right"><%=FormatNumber(rsTBSD.fields(1 + kj),0)%></td>
					<%
              				ntotAFDD(kj)	= Cdbl(ntotAFDD(kj)) + Cdbl(rsTBSD.fields(1 + kj))
							xTOTALD = xTOTALD + Cdbl(rsTBSD.fields(1 + kj))
							nTOTALD =  nTOTALD + Cdbl(rsTBSD.fields(1 + kj))
							%>
					
					
					<%
						if rsTBSK("JOBCODE") = rsTBSD("JOBCODE") then
					%>
					
					<td align="right"><%=FormatNumber(rsTBSK.fields(1 + kj),0)%></td>
					<%
              				ntotAFDK(kj)	= Cdbl(ntotAFDK(kj)) + Cdbl(rsTBSK.fields(1 + kj))
							xTOTALK = xTOTALK + Cdbl(rsTBSK.fields(1 + kj))
							nTOTALK =  nTOTALK+ Cdbl(rsTBSK.fields(1 + kj))
							nKREDIT = 0
							else
							%>
					
					<td align="right">0</td>
					<%
							nKREDIT = -1
							end if
					%>	
				<% Next %>
				
				<td align="right"><%=FormatNumber(Cdbl(xtotalD)+CDbl(xtotalK))%></td>
				</tr>
<%
			NO = NO + 1
			
			
			rsTBSD.MoveNext 
			
			if nKREDIT = 0 then
			rsTBSK.MoveNext 
			end if
			
		Wend
rsTBSD.Close
set rsTBSD = Nothing
rsTBSK.Close
set rsTBSK = Nothing
%>
			<tr bgcolor="#F4F4A0">
				      <% if n=-1 then  %>
				    <td align="center" colspan="3"><strong>TOTAL</td>
				<% else %>
				     <td align="center" colspan="2"><strong>TOTAL</td>
				<% end if %>
				<% For kj = 1 to CInt(kjTot) step 1 %>					
				<td align="right"><%=FormatNumber(Cdbl(ntotAFDD(kj)),0)%></td>
				<td align="right"><%=FormatNumber(Cdbl(ntotAFDK(kj)),0)%></td>
				<% Next %>
					<td align="right"><%=FormatNumber(Cdbl(nTOTALD)+CDbl(nTOTALK),0)%></td>
			</tr>
<%
elseif txtSelect = "Db" then
n=0
if rsTBSD.bof and rsTBSD.eof then
%>
	<tr>
		<td colspan="12">NO DATA AVAILABLE</td>
		 <% 
            n=-1
		 %>
	</tr>
<%end if
		NO = 1
		while not rsTBSD.eof 
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
			
%>
			<tr class="<%=cClass%>">
								 
				   <td align="center"><%=rsTBSD("JOBCODE")%></td>
				 <td align="center">  <a href="P_ReportIntraco.asp?txtComp=<%="txtComp"%>">
								   <%=rsTBSD("JOBNAME")%>
								   </a>
							 </td>
				<% 
			xtotalD = 0
			xtotalK = 0		
			xtotal = 0
				For kj = 1 to CInt(kjTot) step 1 %>
					<td align="right"><%=FormatNumber(rsTBSD.fields(1 + kj),0)%></td>
					<%
              				ntotAFDD(kj)	= Cdbl(ntotAFDD(kj)) + Cdbl(rsTBSD.fields(1 + kj))
							xTOTALD = xTOTALD + Cdbl(rsTBSD.fields(1 + kj))
							nTOTALD =  nTOTALD + Cdbl(rsTBSD.fields(1 + kj))
							%>
					
				<% Next %>
				
				<td align="right"><%=FormatNumber(Cdbl(xtotalD))%></td>
				</tr>
<%
			NO = NO + 1
			
			
			rsTBSD.MoveNext 
			
		Wend
rsTBSD.Close
set rsTBSD = Nothing
%>
			<tr bgcolor="#F4F4A0">
				     <% if n=-1 then  %>
				    <td align="center" colspan="3"><strong>TOTAL</td>
				<% else %>
				     <td align="center" colspan="2"><strong>TOTAL</td>
				<% end if %>
				<% For kj = 1 to CInt(kjTot) step 1 %>					
				<td align="right"><%=FormatNumber(Cdbl(ntotAFDD(kj)),0)%></td>
				<% Next %>
					<td align="right"><%=FormatNumber(Cdbl(nTOTALD),0)%></td>
			</tr>
<%
elseif txtSelect = "Cr" then
n=0
if rsTBSK.bof and rsTBSK.eof then
n= -1
%>
	<tr>
		<td colspan="12">NO DATA AVAILABLE</td>
	</tr>
<%end if
		NO = 1
		while not rsTBSK.eof 
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
			
%>
			<tr class="<%=cClass%>">
								   
				   <td align="center"><%=rsTBSK("JOBCODE")%></td>
				 <td align="center"><a href="P_ReportIntraco.asp?txtComp=<%="txtComp"%>">
								   <%=rsTBSK("JOBNAME")%>
								   </a>
							 </td>
				<% 
			xtotalD = 0
			xtotalK = 0		
			xtotal = 0
				For kj = 1 to CInt(kjTot) step 1  %>
					<td align="right"><%=FormatNumber(rsTBSK.fields(1 + kj),0)%></td>
					<%
              				ntotAFDK(kj)	= Cdbl(ntotAFDK (kj)) + Cdbl(rsTBSK.fields(1 + kj))
							xTOTALK = xTOTALK + Cdbl(rsTBSK.fields(1 + kj))
							nTOTALK =  nTOTALK + Cdbl(rsTBSK.fields(1 + kj))
							%>
					
				<% Next %>
				
				<td align="right"><%=FormatNumber(Cdbl(xtotalK))%></td>
			</tr>
<%
			NO = NO + 1
			
			
			rsTBSK.MoveNext 
			
		Wend
rsTBSK.Close
set rsTBSK = Nothing
%>
			<tr bgcolor="#F4F4A0">
				<% if n=-1 then  %>
				    <td align="center" colspan="3"><strong>TOTAL</td>
				<% else %>
				     <td align="center" colspan="2"><strong>TOTAL</td>
				<% end if %>
				<% For kj = 1 to CInt(kjTot) step 1 %>					
				<td align="right"><%=FormatNumber(Cdbl(ntotAFDK(kj)),0)%></td>
				<% Next %>
					<td align="right"><%=FormatNumber(Cdbl(nTOTALK),0)%></td>
			</tr>			
<%
else
n=0
if rsTBS.bof and rsTBS.eof then
%>
	<tr>
		<td colspan="12">NO DATA AVAILABLE</td>
		 <% 
            n=-1
		 %>
	</tr>
<%end if
		NO = 1
		while not rsTBS.eof 
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
			
%>
			<tr class="<%=cClass%>">
								  
				   <td align="center"><%=rsTBS("JOBCODE")%></td>
				 <td align="center"> <a href="P_ReportIntraco.asp?txtComp=<%="txtComp"%>">
								   <%=rsTBS("JOBNAME")%>
								   </a>
							 </td>
				<% 
			xtotalD = 0
			xtotalK = 0		
			xtotal = 0
				For kj = 1 to CInt(kjTot) step 1  %>
					<td align="right"><%=FormatNumber(rsTBS.fields(1 + kj),0)%></td>
					<%
              				ntotAFD(kj)	= Cdbl(ntotAFD (kj)) + Cdbl(rsTBS.fields(1 + kj))
							xTOTAL = xTOTAL + Cdbl(rsTBS.fields(1 + kj))
							nTOTAL =  nTOTAL + Cdbl(rsTBS.fields(1 + kj))
							%>
					
				<% Next %>
				
				<td align="right"><%=FormatNumber(Cdbl(xtotal))%></td>
				</tr>
<%
			NO = NO + 1
			
			
			rsTBS.MoveNext 
			
		Wend
rsTBS.Close
set rsTBS = Nothing
%>
			<tr bgcolor="#F4F4A0">
			    <% if n=-1 then  %>
				    <td align="center" colspan="3"><strong>TOTAL</td>
				<% else %>
				     <td align="center" colspan="2"><strong>TOTAL</td>
				<% end if %>
				
				    
				<% For kj = 1 to CInt(kjTot) step 1 %>					
				<td align="right"><%=FormatNumber(Cdbl(ntotAFD(kj)),0)%></td>
				<% Next %>
					<td align="right"><%=FormatNumber(Cdbl(nTOTAL),0)%></td>
			</tr>	
<%
END IF
%>
        </tbody>
</table>
</div>
</body>
</html>
