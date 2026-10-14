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
Dim rsAFD,strAFD,sqlAFD,sqlTOTAFD,numAFD,nAFD(),nfieldAFD(),ntotAFD()
ReDim nAFD(-1),nfieldAFD(-1),ntotAFD(-1)
Dim rsTBS, strTBS, rsAtt, strAtt, rsKJ, strKJ,strcomp,strSelect,strGroup
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

if v_txtComp <> "ALL" then
   strSelect = " RECEPTIONDATE,USER_COMPANY_API.GETCOMPANYCODE ('" & Session("CodeName") & "', FCBA) COMPANY,FCBA,PRODUCTCODE "
   strGroup = " RECEPTIONDATE,USER_COMPANY_API.GETCOMPANYCODE ('" & Session("CodeName") & "', FCBA),FCBA,PRODUCTCODE "
   nselect = "4"
   strcomp = " USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "', FCBA) = '" & v_txtComp & "'  and "
else 
   strSelect = " RECEPTIONDATE,USER_COMPANY_API.GETCOMPANYCODE ('" & Session("CodeName") & "', FCBA) COMPANY,PRODUCTCODE "
   strGroup = " RECEPTIONDATE,USER_COMPANY_API.GETCOMPANYCODE ('" & Session("CodeName") & "', FCBA) ,PRODUCTCODE "
   strcomp = ""
   nselect = "3"
end if

if v_txtFcba <> "ALL" then
   
   strcomp = strcomp & " FCBA ='" & v_txtFcba &"'  and "

end if

strAFD = " SELECT DISTINCT DIVISION FROM FIELD WHERE " & strcomp & "  USER_COMPANY_API.GETCOMPANYCODE ('" & Session("CodeName") & "', FCBA) IS NOT NULL ORDER BY DIVISION "
Set rsAFD = Server.CreateObject("ADODB.Recordset")
'response.write strAFD
rsAFD.Open strAFD, connstr, adOpenStatic
if NOT rsAFD.eof then
   sqlAFD = ","
   sqlTOTAFD = "SUM(DECODE(AFDELING, "
   while not rsAFD.eof
     sqlAFD = sqlAFD & "SUM(DECODE(AFDELING, '" & rsAFD("DIVISION") & "', qty, 0))  " & REPLACE(rsAFD("DIVISION"),"-","_") & " ,"
	 sqlTOTAFD = sqlTOTAFD & "'" &  rsAFD("DIVISION") & "', qty,"
	 ReDim Preserve nAFD(UBound(nAFD) + 1)
	 ReDim Preserve nfieldAFD(UBound(nfieldAFD) + 1)
	 ReDim Preserve ntotAFD(UBound(ntotAFD) + 1)
	 nAFD(UBound(nAFD)) = rsAFD("DIVISION")
	 nfieldAFD(UBound(nfieldAFD)) = "FormatNumber(rsTBS(""" & nAFD(kj) &"""),0)"
	 ntotAFD(UBound(ntotAFD)) = 0
	 
     rsAFD.MoveNext
   Wend
   sqlTOTAFD = sqlTOTAFD & " 0 )) TOTAL"
end if
rsAFD.Close
set rsAFD = Nothing

numAFD = UBound(nAFD) + 1
dim kj
'Response.Write numAFD
For kj = 0 to UBound(nAFD)
   ' Response.Write "(" & kj & ") Value " & nAFD(kj) & "<br>"
	'Response.Write nfieldAFD(kj)
Next
'RESPONSE.WRITE sqlAFD & sqlTOTAFD
'response.write len(sqlAFD)

if request("txtYear")="" then
   v_txtYear	=	Year(Session("v_Employee"))
end if
if request("txtMonth")="" then
   v_txtMonth	=	Month(Session("v_Employee"))
end if
       	 strTBS	= " SELECT " & strSelect &  sqlAFD & sqlTOTAFD & _
				  " FROM ( SELECT FCBA,RECEPTIONDATE,PRODUCTCODE,CROPCODE,FIELDCODE,FIELD_API.get_Afdeling(FIELDCODE,FCBA) AFDELING, MILL_WEIGHT QTY  " & _
			      "	       FROM HARVESTING ) " & _
			      " WHERE " & strcomp & " USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "', FCBA) IS NOT NULL  AND TO_NUMBER(TO_CHAR(RECEPTIONDATE, 'MM')) = " & v_txtMonth & " AND TO_NUMBER(TO_CHAR(RECEPTIONDATE, 'YYYY')) = " & v_txtYear & _
                  " GROUP BY " & strGroup & " order by RECEPTIONDATE "
%>		  
<html>

<head>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>iPla'S Reporting - TBS ke PKS</title>
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
					<option value="ALL" <%if request("v_txtComp")="ALL" then%> selected <%End if%>>::ALL::</option>
					<%
					Dim rsComp,strsql
					Set rsComp = Server.CreateObject("ADODB.Recordset")
					rsComp.CursorLocation = adUseClient
					strsql =  " select t.fccompanycode from ACCESSRIGHTSCOMPANY t where t.fccode = '" & Session("CodeName") & "' ORDER BY t.fccompanycode "
					rsComp.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
					if not rsComp.eof then
					DO while not rsComp.eof
					%>			
					<option value="<%=rsComp("fccompanycode")%>" <%if request("v_txtComp")= rsComp("fccompanycode") then%> selected <%End If%>><%=rsComp("fccompanycode")%></option>
					<%rsComp.movenext
					loop
					End if
					rsComp.Close
					set rsComp = Nothing
					%>
					</select>
				</td>	
				<td><font face="verdana" size="1">KEBUN</font></td>
				<td>
				    
					<select name="txtFcba">
					<option value="ALL" <%if request("v_txtFcba")="ALL" then%> selected <%End if%>>::ALL::</option>
					<%
					Dim rsFcba
					Set rsFcba = Server.CreateObject("ADODB.Recordset")
					rsFcba.CursorLocation = adUseClient
					strsql =  "select Fcba from USER_FCBA_TAB t where USER_COMPANY_API.GetCompanyCode('" & Session("CodeName") & "',t.fcba) is not null order by USER_COMPANY_API.GetCompanyCode('" & Session("CodeName") & "',t.fcba),fcba"
					rsFcba.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
					if not rsFcba.eof then
					DO while not rsFcba.eof
					%>			
					<option value="<%=rsFcba("Fcba")%>" <%if request("v_txtFcba")= rsFcba("Fcba") then%> selected <%End If%>><%=rsFcba("Fcba")%></option>
					<%rsFcba.movenext
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
							MinYr = Yr - 2
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

<div id="header">
		
	</div>
	<div id="subheader" class="centered">
		<h2>LAPORAN PENGIRIMAN TBS</h3>
		<br/>
		<h4>PERIODE&nbsp;:&nbsp;<%=Bulan(v_txtMonth)%>&nbsp;<%=v_txtYear%></h4>
		<h4>COMPANY&nbsp;:&nbsp;<%=v_txtComp%>&nbsp;&nbsp;&nbsp;BUSINESS UNIT&nbsp;:&nbsp;<%=v_txtFcba%></h4>
	</div>
	</div>

<table border="0" width="100%" id="product-table">
	<tr>
		<th style="color:#FFFFFF" rowspan="2" width="10%">TANGGAL</th>
		<% if v_txtComp = "ALL" then %>
		   <th style="color:#FFFFFF" rowspan="2" width="10%">COMPANY</th>
		<% Else %>
		   <th style="color:#FFFFFF" rowspan="2" width="10%">COMPANY</th>
		   <th style="color:#FFFFFF" rowspan="2" width="10%">BUSSINES UNIT</th>
		<% end if %>
		<th style="color:#FFFFFF" colspan=<%=numAFD%> width="10%">PRODUKSI AFDELING</th>
		<th style="color:#FFFFFF" rowspan="2" width="10%">TOTAL</th>
	</tr>
	<tr>
		<% For kj = 0 to UBound(nAFD) %>
			<th style="color:#FFFFFF" width="10%"><%=nAFD(kj)%></th>
		<% Next %>
	 </tr>
 
<%
dim rcvdate_,company_
Set rsTBS = Server.CreateObject("ADODB.Recordset")
'response.write strtbs
rsTBS.Open strTBS, connstr, adOpenStatic
if rsTBS.bof and rsTBS.eof then
%>
	<tr>
		<td colspan="11">NO DATA AVAILABLE</td>
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
				<% if v_txtComp = "ALL" then %>
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
				<% end if %>
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
			company_=rsTBS("COMPANY")
			rsTBS.MoveNext
		Wend
rsTBS.Close
set rsTBS = Nothing
%>
			<tr bgcolor="#F4F4A0">
			    <% if v_txtComp = "ALL" then %> 
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
</table>
</div>
</body>
</html>
