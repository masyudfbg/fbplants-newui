<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
Dim v_txtYear, v_txtMonth, v_txtGangCode, v_txtGrade,v_txtComp,v_txtFcba
Dim txtYear, txtMonth, txtGangCode, txtGrade,txtComp,txtFcba
Dim ScriptName
Dim rsAFD,strAFD,sqlAFD,sqlsummAFD,sqlTOTAFD,sqlsummTOTAFD,numAFD,nAFD(),nfieldAFD(),ntotAFD()
ReDim nAFD(-1),nfieldAFD(-1),ntotAFD(-1)
Dim rsTBS, strTBS, rsAtt, strAtt, rsKJ, strKJ,strcomp,strSelect,strGroup,strSummary,strsummSelect
Dim SprevCompany,SprevRegion,SprevGroupRegion,SprevPlantation
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
Dim nTOTAL



txtYear   =	Request("txtYear") 
txtMonth  =	Request("txtMonth")
txtComp = Request("txtComp")
txtFcba =	Request("txtFcba")

v_txtYear		=	txtYear
v_txtMonth		=	txtMonth
v_txtComp 		=	txtComp
v_txtFcba		=	txtFcba


ScriptName	=	Request.ServerVariables("SCRIPT_NAME")
'response.write v_txtComp
    
	if v_txtComp = "--SUMMARY--" then
	   strSelect = " RECEPTIONDATE,PRODUCTCODE "
	   strsummSelect = "PRODUCTCODE "
	   strGroup = " RECEPTIONDATE,PRODUCTCODE "
	   nselect = "2"
	   strcomp = " "
	elseif v_txtComp <> "--ALL--" then
	   strSelect = " RECEPTIONDATE,USER_COMPANY_API.GETCOMPANYCODE ('" & Session("CodeName") & "', FCBA) COMPANY,FCBA,PRODUCTCODE "
	   strsummSelect = "COMPANY,FCBA,PRODUCTCODE "
	   strGroup = " RECEPTIONDATE,USER_COMPANY_API.GETCOMPANYCODE ('" & Session("CodeName") & "', FCBA),FCBA,PRODUCTCODE "
	   nselect = "4"
	   strcomp = " USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "', FCBA) = '" & v_txtComp & "'  and "
	 else 
	   strSelect = " RECEPTIONDATE,USER_COMPANY_API.GETCOMPANYCODE ('" & Session("CodeName") & "', FCBA) COMPANY,PRODUCTCODE "
	   strsummSelect = "COMPANY,PRODUCTCODE "
	   strGroup = " RECEPTIONDATE,USER_COMPANY_API.GETCOMPANYCODE ('" & Session("CodeName") & "', FCBA) ,PRODUCTCODE "
	   strcomp = ""
	   nselect = "3"
	  
	end if
'response.write v_txtFcba
if v_txtFcba <> "--ALL--" then
  if  v_txtFcba <> "--SUMMARY--" then
   
   strcomp = strcomp & " FCBA ='" & v_txtFcba &"'  and "
  end if
end if

strAFD = " SELECT DISTINCT DIVISION FROM FIELD WHERE " & strcomp & "  user_company_api.IsAuthorizedFcba('" & Session("CodeName") & "',fcba) is not null " & _
		 " AND division in (select distinct fccode from sectionname where fccode in ('AFD-01','AFD-02','AFD-03','AFD-04','AFD-05','AFD-06','AFD-07','AFD-08','AFD-09','AFD-10','AFD-11','AFD-12','AFD-13')) ORDER BY DIVISION "
Set rsAFD = Server.CreateObject("ADODB.Recordset")
'response.write strAFD
rsAFD.Open strAFD, connstr, adOpenStatic
if NOT rsAFD.eof then
   sqlAFD = ","
   sqlsummAFD = ","
   sqlTOTAFD = "SUM(DECODE(AFDELING, "
   sqlsummTOTAFD = "SUM( "
   while not rsAFD.eof
     sqlAFD = sqlAFD & "SUM(DECODE(AFDELING, '" & rsAFD("DIVISION") & "', qty, 0))  " & REPLACE(rsAFD("DIVISION"),"-","_") & " ,"
	 sqlsummAFD = sqlsummAFD & "SUM(" & REPLACE(rsAFD("DIVISION"),"-","_") & ") " & REPLACE(rsAFD("DIVISION"),"-","_") & ", "
	 sqlTOTAFD = sqlTOTAFD & "'" &  rsAFD("DIVISION") & "', qty,"
	 sqlsummTOTAFD = sqlsummTOTAFD &   REPLACE(rsAFD("DIVISION"),"-","_") & "+"
	 ReDim Preserve nAFD(UBound(nAFD) + 1)
	 ReDim Preserve nfieldAFD(UBound(nfieldAFD) + 1)
	 ReDim Preserve ntotAFD(UBound(ntotAFD) + 1)
	 nAFD(UBound(nAFD)) = rsAFD("DIVISION")
	 nfieldAFD(UBound(nfieldAFD)) = "FormatNumber(rsTBS(""" & nAFD(kj) &"""),0)"
	 ntotAFD(UBound(ntotAFD)) = 0
	 
     rsAFD.MoveNext
   Wend
   sqlTOTAFD = sqlTOTAFD & " 0 )) TOTAL"
   sqlsummTOTAFD = left(sqlsummTOTAFD,len(sqlsummTOTAFD) - 1) & " ) TOTAL"
end if
rsAFD.Close
set rsAFD = Nothing

numAFD = UBound(nAFD) + 1
dim kj
'Response.Write sqlsummAFD
For kj = 0 to UBound(nAFD)
   ' Response.Write "(" & kj & ") Value " & nAFD(kj) & "<br>"
	'Response.Write nfieldAFD(kj)
Next
'RESPONSE.WRITE sqlsummTOTAFD
'response.write len(sqlAFD)

if request("txtYear")="" then
   v_txtYear	=	Year(Session("v_Employee"))
end if
if request("txtMonth")="" then
   v_txtMonth	=	Month(Session("v_Employee"))
end if
       	 strTBS	= " SELECT " & strSelect &  sqlAFD & sqlTOTAFD & _
				  " FROM ( SELECT FCBA_ESTATE FCBA, HARVESTDATE RECEPTIONDATE, 'TBS' PRODUCTCODE, 'OP' CROPCODE, FIELDCODE, FIELD_API.get_Afdeling (FIELDCODE, FCBA_ESTATE) AFDELING, " & _
				  " BUNCH_MILL_WEIGHT_NETTO QTY FROM HARVESTING_MILL_WEIGHT_NETTO) " & _
			      " WHERE " & strcomp & " user_company_api.IsAuthorizedFcba('" & Session("CodeName") & "',fcba) is not null  AND TO_NUMBER(TO_CHAR(RECEPTIONDATE, 'MM')) = " & v_txtMonth & " AND TO_NUMBER(TO_CHAR(RECEPTIONDATE, 'YYYY')) = " & v_txtYear & _
                  " GROUP BY " & strGroup & " order by RECEPTIONDATE "
				  
  ' response.write (strTBS)
%>		  
<html>

<head>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>iPla'S Reporting - TBS ke PKS</title>
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
</style>

</head>

<body>
	<div id="wrapper">
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
					<option value="--ALL--" <%if request("v_txtComp")="--ALL--" then%> selected <%End if%>>--ALL--</option>
					<%
					Dim rsComp,strsql
					Set rsComp = Server.CreateObject("ADODB.Recordset")
					rsComp.CursorLocation = adUseClient
					strsql = " SELECT '--SUMMARY--' as companycode from DUAL UNION " & _
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
				<td><font face="verdana" size="1">SITE</font></td>
				<td>
				    
					<select name="txtFcba">
					<option value="--ALL--" <%if request("v_txtFcba")="--ALL--" then%> selected <%End if%>>--ALL--</option>
					<%
					Dim rsFcba
					Set rsFcba = Server.CreateObject("ADODB.Recordset")
					rsFcba.CursorLocation = adUseClient
					strsql =  " SELECT * From ( SELECT '--SUMMARY--' as fcba, '1' company from DUAL UNION  ( select Fcba,USER_COMPANY_API.GetCompanyCode('" & Session("CodeName") & "',t.fcba) Company from USER_FCBA_TAB t where fcusercode= '" & Session("CodeName") & "' and USER_COMPANY_API.GetCompanyCode('" & Session("CodeName") & "',t.fcba) is not null ) ) order by company , fcba"
					rsFcba.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
					if not rsFcba.eof then
					DO while not rsFcba.eof
					   if v_txtfcba = rsFcba("Fcba") then
						   %>			
					    <option value="<%=rsFcba("Fcba")%>" Selected><%=rsFcba("Fcba")%></option>
					<% Else %>	
							<option value="<%=rsFcba("Fcba")%>"><%=rsFcba("Fcba")%></option>
					<% End If
					
					  rsFcba.movenext
					loop
					End if
					rsFcba.Close
					set rsFcba = Nothing
					%>
					</select>
				</td>
		        <td><font face="verdana" size="1">YEAR</font></td>
						<% 
							Dim mon, Yr, MinYr, MaxYr, i, MM
							mon = Month(Session("v_Harvesting"))
							If len(v_txtMonth)=0 then v_txtMonth=mon

							Yr = Year(Session("v_Harvesting"))
							MinYr = 2015
							MaxYr = Yr  + 2
							If len(v_txtYear)= 0 then v_txtYear = Yr
						%>
				      <td width="4">
						<select name="txtYear">
							<%
								For i = MinYr to MaxYr
							%>
									<Option Value="<%=i%>" <%If cint(i) = cint(v_txtYear) then%> Selected <%End If%>><%=i%></option>
							<%
								Next
							%>
						</select>
					</td>
					<td><font face="verdana" size="1">MONTH</font></td>
					<td width="2">
						<select name="txtMonth">
							<%
								For i = 1 to 12
							%>
									<Option Value="<%=i%>" <%If cint(i) = cint(v_txtMonth) then%> Selected <%End If%>><%=Bulan(i)%></option>
							<%
								Next
							%>
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
    <th colspan="50" style="background-color:#FFFFFF; border:none;">  
	<div id="header-wrapper">
		<div id="head_judul">
		<h4>PT. <%=Session("COMPANY")%></h4>
		<h5><%=Session("BA")%></h5>
	    </div>
	<div id="subheader" class="centered" style="width:50%; margin:auto;">
		<h2>LAPORAN PENGIRIMAN TBS</h3>
		<br/>
		<h4>PERIODE&nbsp;:&nbsp;<%=Bulan(v_txtMonth)%>&nbsp;<%=v_txtYear%></h4>
		<h4>COMPANY&nbsp;:&nbsp;<%=v_txtComp%>&nbsp;&nbsp;&nbsp;BUSINESS UNIT&nbsp;:&nbsp;<%=v_txtFcba%></h4>
	</div>
	</div>
</th>
  </tr>
	<tr>
		<th style="color:#FFFFFF" rowspan="2" width="10%">TANGGAL</th>
		<% if v_txtComp <> "--SUMMARY--" then %>
				<% if v_txtComp = "--ALL--" then %>
				   <th style="color:#FFFFFF" rowspan="2" width="10%">COMPANY</th>
				<% Else %>
				   <th style="color:#FFFFFF" rowspan="2" width="10%">COMPANY</th>
				   <th style="color:#FFFFFF" rowspan="2" width="10%">BUSSINES UNIT</th>
				<% end if 
		   end if		
		%>
		<th style="color:#FFFFFF" colspan=<%=numAFD%> width="10%">PRODUKSI AFDELING</th>
		<th style="color:#FFFFFF" rowspan="2" width="10%">TOTAL</th>
	</tr>
	<tr>
		<% For kj = 0 to UBound(nAFD) %>
			<th style="color:#FFFFFF" width="10%"><%=nAFD(kj)%></th>
		<% Next %>
	 </tr>
</thead>
<tbody>
 
<%
dim rcvdate_,company_
Set rsTBS = Server.CreateObject("ADODB.Recordset")
'response.write  strTBS
rsTBS.Open strTBS, connstr, adOpenStatic
if rsTBS.bof and rsTBS.eof then
%>
	<tr>
		<td colspan="12">NO DATA AVAILABLE</td>
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
			    <% If rcvdate_ <> rsTBS("RECEPTIONDATE") then %>
				   <td align="center"><%=rsTBS("RECEPTIONDATE")%></td>
				<% else %>
                   <td align="center"></td>			
				<% end if %>
				<% IF v_txtComp <> "--SUMMARY--" then %>
						<% IF v_txtComp = "--ALL--" then %>
						   <td align="center">
						   <a href="p_TbsToPom.asp?txtComp=<%=rsTBS("COMPANY")%>&txtfcba=<%=v_txtFcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>">
						   <%=rsTBS("COMPANY")%>
						   </a>
						   </td>
						<% ELSE 
							 if company_ <> rsTBS("COMPANY")  then %>
								<td align="center">
								   <a href="p_TbsToPom.asp?txtComp=<%=rsTBS("COMPANY")%>&txtfcba=ALL&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>">
								   <%=rsTBS("COMPANY")%>
								   </a>
								 </td>
							 <% else %>
								 <td align="center"></td>	
							 <% end if %> 	
							
							 <td align="center">
								   <a href="p_TbsToPom.asp?txtComp=<%=rsTBS("COMPANY")%>&txtfcba=<%=rsTBS("FCBA")%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>">
								   <%=rsTBS("FCBA")%>
								   </a>
							 </td>
						<% END IF 
					End If %>
				
				<% For kj = 0 to UBound(nAFD) %>					
					<td align="right"><%=FormatNumber(rsTBS.fields(cint(nselect) + kj),0)%></td>
					<%
              				ntotAFD(kj)	= Cdbl(ntotAFD(kj)) + Cdbl(rsTBS.fields(cint(nselect)  + kj))
					%>
				<% Next %>
				
				<td align="right"><%=FormatNumber(rsTBS("TOTAL"),0)%></td>
<%
			NO = NO + 1
			
			nTOTAL = nTOTAL + Cdbl(rsTBS("TOTAL"))
			rcvdate_=rsTBS("RECEPTIONDATE")
			if v_txtComp <> "--SUMMARY--" then
			   company_=rsTBS("COMPANY")
			end if   
			rsTBS.MoveNext
		Wend
rsTBS.Close
set rsTBS = Nothing
%>
			<tr bgcolor="#F4F4A0">
			    <% if v_txtComp = "--SUMMARY--" then %>
				   <td align="center" colspan="1"><strong>TOTAL</td>
			    <% elseif v_txtComp = "--ALL--" then %> 
				      <td align="center" colspan="2"><strong>TOTAL</td>
				<% else %>
				     <td align="center" colspan="3"><strong>TOTAL</td>
				<%end if %>
				<% For kj = 0 to UBound(nAFD) %>					
					<td align="right"><%=FormatNumber(Cdbl(ntotAFD(kj)),0)%></td>
				<% Next %>
				<% if len(sqlAFD) = 0 then %>
				<td align="right"><strong><%=FormatNumber(nTOTAL,0)%></strong></td>
				<% end if %>
				<td align="right"><strong><%=FormatNumber(nTOTAL,0)%></strong></td>
			</tr>
			<tr>
			  <td align="center" colspan="22">  </td>
			</tr>
			<tr Bgcolor="#F4F4A0">
			  <td align="center" colspan="22"><strong>S U M M A  R Y</td>
			</tr>
			<tr>
			  <td align="center" colspan="22">  </td>
			</tr>
			<%
			    dim rsSummTBS,strsummTBS
				Set rsSummTBS = Server.CreateObject("ADODB.Recordset")
			    strSummary =  " SELECT " & strsummSelect &	sqlsummAFD  & sqlsummTOTAFD & "  from ( " & StrTBS & " ) Group By " & 	strsummSelect  & " Order by " & strsummSelect
                'response.write 	strSummary
			
				rsSummTBS.Open strSummary, connstr, adOpenStatic
				if rsSummTBS.bof and rsSummTBS.eof then
			%>
			<tr>
				<td colspan="12">NO DATA AVAILABLE</td>
			</tr>
		    <%   end if
				NO = 1
				while not rsSummTBS.eof
			%>
					<tr>
						<% if v_txtComp = "--SUMMARY--" then %>
						   <td align="center" colspan="1"><strong><%=rsSummTBS.fields(0)%></td>
						<% elseif v_txtComp = "--ALL--" then %> 
							  <td align="center" colspan="2"><strong><%=rsSummTBS.fields(0)%></td>
						<% else %>
							 <td align="center" colspan="3"><strong><%=rsSummTBS.fields(1)%></td>
						<%end if %>
						<% For kj = 0 to UBound(nAFD) %>					
							<td align="right"><%=FormatNumber(rsSummTBS.fields(cint(nselect) - 1 + kj),0)%></td>
						<% Next %>
						<td align="right"><%=FormatNumber(rsSummTBS.fields(cint(UBound(nAFD)) + cint(nselect)),0)%></td>
					</tr>
			<%
					NO = NO + 1
					rsSummTBS.MoveNext
				  Wend
				  rsSummTBS.Close 
				  set rsSummTBS = Nothing
%>
                   <tr Bgcolor="#F4F4A0">
						<% if v_txtComp = "--SUMMARY--" then %>
						   <td align="center" colspan="1"><strong>TOTAL</td>
						<% elseif v_txtComp = "--ALL--" then %> 
							  <td align="center" colspan="2"><strong>TOTAL</td>
						<% else %>
							 <td align="center" colspan="3"><strong>TOTAL</td>
						<%end if %>
						<% For kj = 0 to UBound(nAFD) %>					
							<td align="right"><%=FormatNumber(Cdbl(ntotAFD(kj)),0)%></td>
						<% Next %>
						<% if len(sqlAFD) = 0 then %>
						<td align="right"><strong><%=FormatNumber(nTOTAL,0)%></strong></td>
						<% end if %>
						<td align="right"><strong><%=FormatNumber(nTOTAL,0)%></strong></td>
					</tr>
        </tbody>
</table>
</div>
</body>
</html>
