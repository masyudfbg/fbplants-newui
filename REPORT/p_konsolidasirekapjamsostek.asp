<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->

<%
Dim v_txtYear, v_txtMonth, no
Dim ScriptName
Dim rsjamsostek, strjamsostek, strMain, strjamsostek2
Dim cClass, nday, xNum, nSpanTotal, FirstDate

Dim var_colspan : var_colspan = 1

Dim CompanyName, q, fcbaname, MgrName, KTUName, CropName, FcbaType

nSpanTotal = 7

v_txtcompany  			= Request("txtCompany")
v_txtfcba 				= Request("txtfcba")

Dim selectFields, strFieldCons, selectCons, whereCons, GroupbyCons, orderCons

strFieldCons	=	"   USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',fcba) company" & _
					" , BUSINESSUNIT_API.get_OrderData(fcba) NoUrut, fcba, BUSINESSUNIT_API.get_fcname(fcba) fcbaname"
					
wherecons	=	" where company is not null "


If Len(Trim(v_txtCompany)) > 0 Then
	if v_txtCompany = "--SUMMARY--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
		nSpanTotal	=	nSpanTotal	- 1
	Elseif v_txtCompany = "--ALL--" then
		selectCons	=	selectCons	& " , company"
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " , COMPANY_API.get_OrderData(company)"
		var_colspan	=	var_colspan	+ 1
	Else
		selectCons	=	selectCons	& " , company"
		whereCons	=	whereCons	& " And company = '" & v_txtCompany & "'"
		orderCons	=	orderCons	& " , COMPANY_API.get_OrderData(company)"
		var_colspan	=	var_colspan	+ 1
		
		q	=	" SELECT FCCODE" & _
				" , FCCOMPANYCODE" & _
				" , FCNAME" & _
				" , FCCOMPANYNAME" & _
				" , FCMANAGER" & _
				" , FCKTU" & _
				" , FCCROP" & _
				" , FCTYPE" & _
				" FROM BUSINESSUNIT WHERE" & _
				" FCCOMPANYCODE = '" & v_txtCompany & "'"	
		set rsCat = connStr.Execute(q)		
		if not rsCat.eof then
			CompanyName = rsCat("FCCOMPANYNAME")	
		end if
		q	=	""
		rsCat.Close
		set rsCat = Nothing
	End If
End If

If Len(Trim(v_txtfcba)) > 0 Then
	if v_txtfcba = "--SUMMARY--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
		nSpanTotal	=	nSpanTotal	- 2
	Elseif v_txtfcba = "--ALL--" then
		selectCons	=	selectCons	& " , nourut, fcba, fcbaname"
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " , nourut"
		var_colspan	=	var_colspan	+ 2
	Else
		selectCons	=	selectCons	& " , nourut, fcba, fcbaname"
		whereCons	=	whereCons	& " And fcba = '" & v_txtfcba & "'"
		orderCons	=	orderCons	& " , nourut"
		var_colspan	=	var_colspan	+ 2
		
		q	=	" SELECT FCCODE" & _
				" , FCCOMPANYCODE" & _
				" , FCNAME" & _
				" , FCCOMPANYNAME" & _
				" , FCMANAGER" & _
				" , FCKTU" & _
				" , FCCROP" & _
				" , FCTYPE" & _
				" FROM BUSINESSUNIT WHERE" & _
				" FCCODE = '" & v_txtfcba & "'"	
		set rsCat = connStr.Execute(q)		
		if not rsCat.eof then
			fcbaname = rsCat("FCNAME")
			CompanyName=rsCat("FCCOMPANYNAME")
			MgrName = rsCat("FCMANAGER")
			KTUName = rsCat("FCKTU")
			CropName=rsCat("FCCROP")
			FcbaType=rsCat("FCTYPE")
		end if
		q	=	""
		rsCat.Close
		set rsCat = Nothing
	End If
End If

v_txtYear		= Request("txtYear")
v_txtMonth		= Request("txtMonth")

If Len(Trim(v_txtYear)) = 0 Then
	v_txtYear	= Year(Session("v_Costbook"))
End If
If Len(Trim(v_txtMonth)) = 0 Then
	v_txtMonth	= Month(Session("v_Costbook"))
End If

ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

strMain		=	" select 'A' A, " & strFieldCons & _
				"        , nvl(sum(pot_bpjs_tk),0) * -1 jamsostek" & _
				" from bkmpayroll" & _
				" where " & _
				"     period_year = '" & v_txtYear & "'" & _
				"     and period_month = '" & v_txtMonth & "'" & _
				" group by fcba"

				
Dim Selection_2

Selection_2 = " "

strjamsostek2	=	"        , nvl(sum(jamsostek),0) jamsostek"

strjamsostek	=	" select  A " & selectCons & Selection_2 & strjamsostek2 & "  From (" & strMain & " )" & whereCons & _
					" group by A " & selectCons & Selection_2 & " Order by A " & orderCons & " "

'response.write strjamsostek & FirstDate
%>
<html>

<head>
<title>ePlaS Reporting - Laporan Rekap Jamsostek</title>
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
				<!--#include virtual="/REPORT/FormSelectNewCompFcba.html"-->
				
				<table id="form-table" class="centered-margin">
					<tr>
						<td>
							<font face="verdana" size="1">TAHUN</font>
						</td>
						<% 
							Dim mon, Yr, MinYr, MaxYr, i, MM
							mon = Month(Session("v_Costbook"))
							If len(v_txtMonth)=0 then v_txtMonth=mon
						
							Yr = Year(Session("v_Costbook"))
							MinYr = 2019
							MaxYr = Yr  + 10
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
							<input type="submit" value="View" />
						</td>
					</tr>
				</table>
			</div>
			</form>
      <table border="0" width="100%" id="product-table" class="xthf-site">
	<thead style="border-left:1px #000000 solid;">
        <tr>
            <th colspan="<%=var_colspan+11%>" style="background:#FFFFFF; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
            <div id="header">
			<% if  CompanyName <>  "" then %>
				<h4>PT.<%=CompanyName%></h4>
			<% End if%>
			<% If  FcbaName <> ""  then %>
				   <h5><%=FcbaName%></h5>
			<% End if%>
            </div>
            <div id="subheader" class="centered">
                <h2>LAPORAN REKAP BPJS KETENAGAKERJAAN</h3>
                <br/>
                <h4>PERIODE&nbsp;:&nbsp;<%=Bulan(v_txtMonth)%>&nbsp;<%=v_txtYear%></h4>
            </div>
		</div>

</th>
</tr>
	<tr>
		
		
		<%If Len(Trim(v_txtCompany)) > 0 And v_txtCompany <>  "--SUMMARY--" Then%>
			<th style="background-color:#00A300;color:#fff;"><%=titleCompany%></th>
		<%End If%>

		<%If Len(Trim(v_txtfcba)) > 0 And v_txtfcba <>  "--SUMMARY--" Then%>
			<th style="background-color:#00A300;color:#fff;"><%=titleSite%></th>
			<th style="background-color:#00A300;color:#fff;">Nama <%=titleSite%></th>
		<%End If%>
		
		<%If var_colspan = 1 Then%>
			<th style="background-color:#00A300;color:#fff;">&nbsp;</th>
		<%End If%>
			<th style="background-color:#00A300;color:#fff;">BPJS Ketenagakerjaan</th>
	</tr>
    </thead>
<tbody>
<%
Set rsjamsostek = Server.CreateObject("ADODB.Recordset")
rsjamsostek.Open strjamsostek, connstr, adOpenStatic

Dim subtotaljamsostek : subtotaljamsostek = 0
Dim grandtotaljamsostek : grandtotaljamsostek = 0

Dim tempCompany : tempCompany = ""
Dim tempfcba : tempfcba = ""

Dim SprevCompany : SprevCompany = ""
Dim sprevFcba : sprevFcba = ""

Dim CurrRegion : CurrRegion = ""
Dim PrevRegion : PrevRegion = ""

if rsjamsostek.bof and rsjamsostek.eof then
%>
	<tr>
		<td colspan="<%=var_colspan+11%>">NO DATA AVAILABLE</td>
	</tr>
<%
else
WHILE NOT rsjamsostek.eof
%>
<%
						subtotaljamsostek = 0
%>
						<tr>
							
							<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then %>
								<% tempCompany =rsjamsostek("Company")
								   tempfcba = "--ALL--"	   
								 %>
								
								<% if tempCompany <> SprevCompany then %>
									<td align="center">
										<a href="<%=ScriptName%>?txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtYear=<%=v_txtYear%>&txtMonth=<%=v_txtMonth%>">
										<%=rsjamsostek("company")%>
										</a>
									</td>
									<% SprevCompany = tempCompany %>	
								<% else	%>
									<td align="center"></td>					
								<% end if %>
							<% end if %>
							
							<% if v_txtFcba <>  "--SUMMARY--" and v_txtFcba <>  "" then %>
								<% 
								   if  v_txtcompany <>  "--SUMMARY--" and v_txtcompany <> "" then
									   tempCompany =rsjamsostek("Company")
								   else
									   tempCompany =""
								   end if

								   tempfcba = rsjamsostek("fcba")	   
								 %>
								
								<% if tempfcba <> sprevFcba then %>
									<td align="center">
										<a href="p_jamsostek.asp?txtfcba=<%=tempfcba%>&txtYear=<%=v_txtYear%>&txtMonth=<%=v_txtMonth%>">
										<%=rsjamsostek("fcba")%>
										</a>
									</td>
									<td align="left"><%=rsjamsostek("fcbaname")%></td>
									<% sprevFcba = tempfcba %>	
								<% else	%>
									<td align="center"></td>
									<td align="center"></td>
								<% end if %>
							<% end if %>

							<%If var_colspan = 1 Then%>
								<td align="right"></td>
							<%End If%>
								<td align="right"><%=FormatNumber(rsjamsostek("jamsostek"), 2)%></td>
					</tr>
										
<%
			subtotaljamsostek = subtotaljamsostek + cdbl(rsjamsostek("jamsostek"))
%>
			
<%
			grandtotaljamsostek = grandtotaljamsostek + Cdbl(rsjamsostek("jamsostek"))
			
			rsjamsostek.MoveNext
			PrevRegion = CurrRegion
		Wend	
	rsjamsostek.Close
	set rsjamsostek = Nothing
%>
				<tr style="background-color:#F4F4A0;">
					<td align="left" colspan="<%=var_colspan-1%>" style="padding-left:10px;"><strong>Grand Total</strong></td>
					<td align="right"><strong><%=FormatNumber(grandtotaljamsostek, 2)%></strong></td>
				</tr>
<%End if%>
     </tbody>
</table>
	
   </body>
   </html>
