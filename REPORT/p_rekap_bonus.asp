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
Dim rsrekap_bpjskes, strrekap_bpjskes, strMain, strrekap_bpjskes2
Dim cClass, nday, xNum, nSpanTotal, FirstDate

Dim var_colspan : var_colspan = 1

Dim CompanyName, q, fcbaname, MgrName, KTUName, CropName, FcbaType

nSpanTotal = 7

v_txtGroupPlantation 	= Request("txtGroupPlantation")
v_txtPlantation  		= Request("txtPlantation")
v_txtGroupRegion  		= Request("txtGroupRegion")
v_txtRegion  			= Request("txtRegion")
v_txtcompany  			= Request("txtCompany")
v_txtfcba 				= Request("txtfcba")

Dim selectFields, strFieldCons, selectCons, whereCons, GroupbyCons, orderCons

strFieldCons	=	" USER_COMPANY_API.GetGroupPlantationCode('" & Session("CodeName") & "',fcba) GroupPlantation" & _
					" , USER_COMPANY_API.GetPlantationCode('" & Session("CodeName") & "',fcba) Plantation" & _
					" , USER_COMPANY_API.GetGroupRegionCode('" & Session("CodeName") & "',fcba) groupregion" & _
					" , USER_COMPANY_API.GetRegionCode('" & Session("CodeName") & "',fcba) region" & _
					" , USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',fcba) company" & _
					" , BUSINESSUNIT_API.get_OrderData(fcba) NoUrut, fcba, BUSINESSUNIT_API.get_fcname(fcba) fcbaname"
					
wherecons	=	" where company is not null "


If Len(Trim(v_txtGroupPlantation)) > 0 Then
	if v_txtGroupPlantation = "--SUMMARY--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
		nSpanTotal	=	nSpanTotal	- 1
	Elseif v_txtGroupPlantation = "--ALL--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
		var_colspan	=	var_colspan	+ 1
	Else
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " And GroupPlantation = '" & v_txtGroupPlantation & "'"
		orderCons	=	orderCons	& " "
		var_colspan	=	var_colspan	+ 1
	End If
End If

If Len(Trim(v_txtPlantation)) > 0 Then
	if v_txtPlantation = "--SUMMARY--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
		nSpanTotal	=	nSpanTotal	- 1
	Elseif v_txtPlantation = "--ALL--" then
		selectCons	=	selectCons	& " , Plantation"
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " , PLANTATION_API.get_OrderData(Plantation)"
		var_colspan	=	var_colspan	+ 1
	Else
		selectCons	=	selectCons	& " , Plantation"
		whereCons	=	whereCons	& " And Plantation = '" & v_txtPlantation & "'"
		orderCons	=	orderCons	& " , PLANTATION_API.get_OrderData(Plantation)"
		var_colspan	=	var_colspan	+ 1
	End If
End If


If Len(Trim(v_txtGroupRegion)) > 0 Then
	if v_txtGroupRegion = "--SUMMARY--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
		nSpanTotal	=	nSpanTotal	- 1
	Elseif v_txtGroupRegion = "--ALL--" then
		selectCons	=	selectCons	& " , GroupRegion"
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " , GROUPREGION_API.get_OrderData(GroupRegion)"
		var_colspan	=	var_colspan	+ 1
	Else
		selectCons	=	selectCons	& " , GroupRegion"
		whereCons	=	whereCons	& " And GroupRegion = '" & v_txtGroupRegion & "'"
		orderCons	=	orderCons	& " , GROUPREGION_API.get_OrderData(GroupRegion)"
		var_colspan	=	var_colspan	+ 1
	End If
End If

If Len(Trim(v_txtRegion)) > 0 Then
	if v_txtRegion = "--SUMMARY--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
		nSpanTotal	=	nSpanTotal	- 1
	Elseif v_txtRegion = "--ALL--" then
		selectCons	=	selectCons	& " , Region"
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " , REGION_API.get_OrderData(Region)"
		var_colspan	=	var_colspan	+ 1
	Else
		selectCons	=	selectCons	& " , Region"
		whereCons	=	whereCons	& " And Region = '" & v_txtRegion & "'"
		orderCons	=	orderCons	& " , REGION_API.get_OrderData(Region)"
		var_colspan	=	var_colspan	+ 1
	End If
End If


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

strMain		=	" SELECT" & strFieldCons & _
				"        , sum(amount) bonus" & _
				" FROM empallded" & _
				" WHERE to_number(to_char(adhocdate,'yyyymm')) = "&v_txtYear&"||LPAD("&v_txtMonth&", 2, 0)" & _
				"     AND allowdedcode = 'A61'" & _
				" GROUP BY fcba"

				
Dim Selection_2

Selection_2 = " "

strrekap_bpjskes2	=	"        , sum(bonus) bonus"

strrekap_bpjskes	=	" select GroupPlantation " & selectCons & Selection_2 & strrekap_bpjskes2 & "  From (" & strMain & " )" & whereCons & _
					" group by GroupPlantation " & selectCons & Selection_2 & " Order by GroupPlantation " & orderCons & " "

'response.write strrekap_bpjskes & FirstDate
%>
<html>

<head>
<title>ePlaS Reporting - Rekap Bonus</title>
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
				<!--#include virtual="/REPORT/FormSelectNewSum.html"-->
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
                <h2>LAPORAN REKAP BONUS</h3>
                <br/>
                <h4>PERIODE&nbsp;:&nbsp;<%=Bulan(v_txtMonth)%>&nbsp;<%=v_txtYear%></h4>
            </div>
		</div>

</th>
</tr>
	<tr>
		
		<%If Len(Trim(v_txtGroupPlantation)) > 0 And v_txtGroupPlantation <>  "--SUMMARY--" Then%>
			<th style="background-color:#00A300;color:#fff;"><%=titleGroupPlantation%></th>
		<%End If%>

		<%If Len(Trim(v_txtPlantation)) > 0 And v_txtPlantation <>  "--SUMMARY--" Then%>
			<th style="background-color:#00A300;color:#fff;"><%=titlePlantation%></th>
		<%End If%>

		<%If Len(Trim(v_txtGroupRegion)) > 0 And v_txtGroupRegion <>  "--SUMMARY--" Then%>
			<th style="background-color:#00A300;color:#fff;"><%=titleGroupRegion%></th>
		<%End If%>

		<%If Len(Trim(v_txtRegion)) > 0 And v_txtRegion <>  "--SUMMARY--" Then%>
			<th style="background-color:#00A300;color:#fff;"><%=titleRegion%></th>
		<%End If%>

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
			<th style="background-color:#00A300;color:#fff;">Bonus</th>
	</tr>
    </thead>
<tbody>
<%
Set rsrekap_bpjskes = Server.CreateObject("ADODB.Recordset")
rsrekap_bpjskes.Open strrekap_bpjskes, connstr, adOpenStatic

Dim subtotalbonus : subtotalbonus = 0
Dim grandtotalbonus : grandtotalbonus = 0

Dim tempGroupPlantation : tempGroupPlantation = ""
Dim tempPlantation : tempPlantation = ""
Dim tempGroupRegion : tempGroupRegion = ""
Dim tempRegion : tempRegion = ""
Dim tempCompany : tempCompany = ""
Dim tempfcba : tempfcba = ""

Dim SprevGroupPlantation : SprevGroupPlantation = ""
Dim SprevPlantation : SprevPlantation = ""
Dim SprevGroupRegion : SprevGroupRegion = ""
Dim SprevRegion : SprevRegion = ""
Dim SprevCompany : SprevCompany = ""
Dim sprevFcba : sprevFcba = ""

Dim CurrRegion : CurrRegion = ""
Dim PrevRegion : PrevRegion = ""

if rsrekap_bpjskes.bof and rsrekap_bpjskes.eof then
%>
	<tr>
		<td colspan="<%=var_colspan+11%>">NO DATA AVAILABLE</td>
	</tr>
<%
else
		while not rsrekap_bpjskes.eof
			if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
				CurrRegion = rsrekap_bpjskes("region")
				If CurrRegion <> PrevRegion Then
					If Len(Trim(PrevRegion)) = 0 Then
						subtotalbonus = 0
					Else
%>
						<tr style="background-color:#F4F4A0;">
							<td align="left" colspan="<%=var_colspan-1%>" style="padding-left:10px;"><strong>Sub Total <%=PCase(PrevRegion)%></strong></td>
							<td align="Right"><strong><%=FormatNumber(subtotalbonus, 2)%></strong></td>			
						</tr>
<%
						subtotalbonus = 0
%>
					<%End If%>
					<tr>
							
							<% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <>  "" then %>
								<% tempGroupPlantation = rsrekap_bpjskes("GroupPlantation")
								   tempPlantation = "--ALL--"
								   tempGroupRegion = "--SUMMARY--"
								   tempRegion = "--SUMMARY--"
								   tempCompany ="--SUMMARY--"
								   tempfcba = "--SUMMARY--"
								%>
								<% if tempGroupPlantation <> SprevGroupPlantation then %>
									<td align="center">
										<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtYear=<%=v_txtYear%>&txtMonth=<%=v_txtMonth%>">
										<%=rsrekap_bpjskes("GroupPlantation")%>
										</a>
									</td>
									<% SprevGroupPlantation = tempGroupPlantation %>	
								<% else	%>
									<td align="center"></td>					
								<% end if %>
							<% end if %>
							
							
							<% if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  "" then %>
								<%if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
									  tempGroupPlantation = rsrekap_bpjskes("GroupPlantation")
								   else
									  tempGroupPlantation = ""
								   end if
								
								   tempPlantation = rsrekap_bpjskes("Plantation")
								   tempGroupRegion = "--ALL--"
								   tempRegion = "--SUMMARY--"
								   tempCompany ="--SUMMARY--"
								   tempfcba = "--SUMMARY--"
								%>
								<% if tempPlantation <> SprevPlantation then %>
									<td align="center">
										<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtYear=<%=v_txtYear%>&txtMonth=<%=v_txtMonth%>">
										<%=rsrekap_bpjskes("Plantation")%>
										</a>
									</td>
									<% SprevPlantation = tempPlantation %>	
								<% else	%>
									<td align="center"></td>					
								<% end if %>	
							<% end if %>
							
							<% if v_txtGroupRegion <>  "--SUMMARY--"  and  v_txtGroupRegion <>  "" then %>
								<% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
									  tempGroupPlantation = rsrekap_bpjskes("GroupPlantation")
								   else
									  tempGroupPlantation = ""
								   end if					
								   
								   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
									  tempPlantation = rsrekap_bpjskes("Plantation")
								   else
									  tempPlantation = ""
								   end if
								   
								   tempGroupRegion = rsrekap_bpjskes("GroupRegion")
								   tempRegion = "--ALL--"
								   tempCompany ="--SUMMARY--"
								   tempfcba = "--SUMMARY--"
								%>
								<% if tempGroupRegion <> SprevGroupRegion then %>
									<td align="center">
										<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtYear=<%=v_txtYear%>&txtMonth=<%=v_txtMonth%>">
										<%=rsrekap_bpjskes("GroupRegion")%>
										</a>
									</td>
									<% SprevGroupRegion= tempGroupRegion %>	
								<% else	%>
									<td align="center"></td>					
								<% end if %>	
							<% end if %>
							
							<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
								<% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
									  tempGroupPlantation = rsrekap_bpjskes("GroupPlantation")
								   else
									  tempGroupPlantation = ""
								   end if					
								   
								   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
									  tempPlantation = rsrekap_bpjskes("Plantation")
								   else
									  tempPlantation = ""
								   end if
								   
								   if  v_txtGroupregion <>  "--SUMMARY--" and v_txtGroupregion <> "" then
									   tempGroupRegion = rsrekap_bpjskes("GroupRegion")
								   else
									   tempGroupRegion = ""
								   end if
								   
								   tempRegion = rsrekap_bpjskes("Region") 
								   tempCompany ="--ALL--"
								   tempfcba = "--SUMMARY--"
								%>
								
								<% if tempRegion <> SprevRegion then %>
									<td align="center">
										<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtYear=<%=v_txtYear%>&txtMonth=<%=v_txtMonth%>">
										<%=rsrekap_bpjskes("region")%>
										</a>
									</td>
									<% SprevRegion = tempRegion %>	
								<% else	%>
									<td align="center"></td>					
								<% end if %>	
							<% end if %>
							
							<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then %>
								<% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
									  tempGroupPlantation = rsrekap_bpjskes("GroupPlantation")
								   else
									  tempGroupPlantation = ""
								   end if					
								   
								   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
									  tempPlantation = rsrekap_bpjskes("Plantation")
								   else
									  tempPlantation = ""
								   end if
								   
								   if  v_txtGroupregion <>  "--SUMMARY--" and v_txtGroupregion <> "" then
									   tempGroupRegion = rsrekap_bpjskes("GroupRegion")
								   else
									   tempGroupRegion = ""
								   end if
								   
								   if  v_txtregion <>  "--SUMMARY--" and v_txtregion <> "" then
									   tempRegion = rsrekap_bpjskes("Region") 
								   else
									   tempRegion = ""
								   end if 

								   tempCompany =rsrekap_bpjskes("Company")
								   tempfcba = "--ALL--"	   
								 %>
								
								<% if tempCompany <> SprevCompany then %>
									<td align="center">
										<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtYear=<%=v_txtYear%>&txtMonth=<%=v_txtMonth%>">
										<%=rsrekap_bpjskes("company")%>
										</a>
									</td>
									<% SprevCompany = tempCompany %>	
								<% else	%>
									<td align="center"></td>					
								<% end if %>
							<% end if %>
							
							<% if v_txtFcba <>  "--SUMMARY--" and v_txtFcba <>  "" then %>
								<% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
									  tempGroupPlantation = rsrekap_bpjskes("GroupPlantation")
								   else
									  tempGroupPlantation = ""
								   end if					
								   
								   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
									  tempPlantation = rsrekap_bpjskes("Plantation")
								   else
									  tempPlantation = ""
								   end if
								   
								   if  v_txtGroupregion <>  "--SUMMARY--" and v_txtGroupregion <> "" then
									   tempGroupRegion = rsrekap_bpjskes("GroupRegion")
								   else
									   tempGroupRegion = ""
								   end if
								   
								   if  v_txtregion <>  "--SUMMARY--" and v_txtregion <> "" then
									   tempRegion = rsrekap_bpjskes("Region") 
								   else
									   tempRegion = ""
								   end if 
								   
								   if  v_txtcompany <>  "--SUMMARY--" and v_txtcompany <> "" then
									   tempCompany =rsrekap_bpjskes("Company")
								   else
									   tempCompany =""
								   end if

								   tempfcba = rsrekap_bpjskes("fcba")	   
								 %>
								
								<% if tempfcba <> sprevFcba then %>
									<td align="center">
										<a href="p_rincian_bonus.asp?txtfcba=<%=tempfcba%>&txtYear=<%=v_txtYear%>&txtMonth=<%=v_txtMonth%>">
										<%=rsrekap_bpjskes("fcba")%>
										</a>
									</td>
									<td align="left"><%=rsrekap_bpjskes("fcbaname")%></td>
									<% sprevFcba = tempfcba %>	
								<% else	%>
									<td align="center"></td>
									<td align="center"></td>
								<% end if %>
							<% end if %>

							<%If var_colspan = 1 Then%>
								<td align="right"></td>
							<%End If%>
								<td align="right"><%=FormatNumber(rsrekap_bpjskes("bonus"), 2)%></td>
					</tr>
				<%Else%>
					<tr>
					
					<% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <>  "" then %>
						<% tempGroupPlantation = rsrekap_bpjskes("GroupPlantation")
						   tempPlantation = "--ALL--"
						   tempGroupRegion = "--SUMMARY--"
						   tempRegion = "--SUMMARY--"
						   tempCompany ="--SUMMARY--"
						   tempfcba = "--SUMMARY--"
						%>
						<% if tempGroupPlantation <> SprevGroupPlantation then %>
							<td align="center">
								<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtYear=<%=v_txtYear%>&txtMonth=<%=v_txtMonth%>">
								<%=rsrekap_bpjskes("GroupPlantation")%>
								</a>
							</td>
							<% SprevGroupPlantation = tempGroupPlantation %>	
						<% else	%>
							<td align="center"></td>					
						<% end if %>
					<% end if %>
					
					
					<% if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  "" then %>
						<%if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
							  tempGroupPlantation = rsrekap_bpjskes("GroupPlantation")
						   else
							  tempGroupPlantation = ""
						   end if
						
						   tempPlantation = rsrekap_bpjskes("Plantation")
						   tempGroupRegion = "--ALL--"
						   tempRegion = "--SUMMARY--"
						   tempCompany ="--SUMMARY--"
						   tempfcba = "--SUMMARY--"
						%>
						<% if tempPlantation <> SprevPlantation then %>
							<td align="center">
								<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtYear=<%=v_txtYear%>&txtMonth=<%=v_txtMonth%>">
								<%=rsrekap_bpjskes("Plantation")%>
								</a>
							</td>
							<% SprevPlantation = tempPlantation %>	
						<% else	%>
							<td align="center"></td>					
						<% end if %>	
					<% end if %>
					
					<% if v_txtGroupRegion <>  "--SUMMARY--"  and  v_txtGroupRegion <>  "" then %>
						<% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
							  tempGroupPlantation = rsrekap_bpjskes("GroupPlantation")
						   else
							  tempGroupPlantation = ""
						   end if					
						   
						   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
							  tempPlantation = rsrekap_bpjskes("Plantation")
						   else
							  tempPlantation = ""
						   end if
						   
						   tempGroupRegion = rsrekap_bpjskes("GroupRegion")
						   tempRegion = "--ALL--"
						   tempCompany ="--SUMMARY--"
						   tempfcba = "--SUMMARY--"
						%>
						<% if tempGroupRegion <> SprevGroupRegion then %>
							<td align="center">
								<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtYear=<%=v_txtYear%>&txtMonth=<%=v_txtMonth%>">
								<%=rsrekap_bpjskes("GroupRegion")%>
								</a>
							</td>
							<% SprevGroupRegion= tempGroupRegion %>	
						<% else	%>
							<td align="center"></td>					
						<% end if %>	
					<% end if %>
					
					<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
						<% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
							  tempGroupPlantation = rsrekap_bpjskes("GroupPlantation")
						   else
							  tempGroupPlantation = ""
						   end if					
						   
						   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
							  tempPlantation = rsrekap_bpjskes("Plantation")
						   else
							  tempPlantation = ""
						   end if
						   
						   if  v_txtGroupregion <>  "--SUMMARY--" and v_txtGroupregion <> "" then
							   tempGroupRegion = rsrekap_bpjskes("GroupRegion")
						   else
							   tempGroupRegion = ""
						   end if
						   
						   tempRegion = rsrekap_bpjskes("Region") 
						   tempCompany ="--ALL--"
						   tempfcba = "--SUMMARY--"
						%>
						
						<% if tempRegion <> SprevRegion then %>
							<td align="center">
								<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtYear=<%=v_txtYear%>&txtMonth=<%=v_txtMonth%>">
								<%=rsrekap_bpjskes("region")%>
								</a>
							</td>
							<% SprevRegion = tempRegion %>	
						<% else	%>
							<td align="center"></td>					
						<% end if %>	
					<% end if %>
					
					<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then %>
						<% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
							  tempGroupPlantation = rsrekap_bpjskes("GroupPlantation")
						   else
							  tempGroupPlantation = ""
						   end if					
						   
						   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
							  tempPlantation = rsrekap_bpjskes("Plantation")
						   else
							  tempPlantation = ""
						   end if
						   
						   if  v_txtGroupregion <>  "--SUMMARY--" and v_txtGroupregion <> "" then
							   tempGroupRegion = rsrekap_bpjskes("GroupRegion")
						   else
							   tempGroupRegion = ""
						   end if
						   
						   if  v_txtregion <>  "--SUMMARY--" and v_txtregion <> "" then
							   tempRegion = rsrekap_bpjskes("Region") 
						   else
							   tempRegion = ""
						   end if 

						   tempCompany =rsrekap_bpjskes("Company")
						   tempfcba = "--ALL--"	   
						 %>
						
						<% if tempCompany <> SprevCompany then %>
							<td align="center">
								<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtYear=<%=v_txtYear%>&txtMonth=<%=v_txtMonth%>">
								<%=rsrekap_bpjskes("company")%>
								</a>
							</td>
							<% SprevCompany = tempCompany %>	
						<% else	%>
							<td align="center"></td>					
						<% end if %>
					<% end if %>
					
					<% if v_txtFcba <>  "--SUMMARY--" and v_txtFcba <>  "" then %>
						<% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
							  tempGroupPlantation = rsrekap_bpjskes("GroupPlantation")
						   else
							  tempGroupPlantation = ""
						   end if					
						   
						   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
							  tempPlantation = rsrekap_bpjskes("Plantation")
						   else
							  tempPlantation = ""
						   end if
						   
						   if  v_txtGroupregion <>  "--SUMMARY--" and v_txtGroupregion <> "" then
							   tempGroupRegion = rsrekap_bpjskes("GroupRegion")
						   else
							   tempGroupRegion = ""
						   end if
						   
						   if  v_txtregion <>  "--SUMMARY--" and v_txtregion <> "" then
							   tempRegion = rsrekap_bpjskes("Region") 
						   else
							   tempRegion = ""
						   end if 
						   
						   if  v_txtcompany <>  "--SUMMARY--" and v_txtcompany <> "" then
							   tempCompany =rsrekap_bpjskes("Company")
						   else
							   tempCompany =""
						   end if

						   tempfcba = rsrekap_bpjskes("fcba")	   
						 %>
						
						<% if tempfcba <> sprevFcba then %>
							<td align="center">
								<a href="p_rincian_bonus.asp?txtfcba=<%=tempfcba%>&txtYear=<%=v_txtYear%>&txtMonth=<%=v_txtMonth%>">
								<%=rsrekap_bpjskes("fcba")%>
								</a>
							</td>
							<td align="left"><%=rsrekap_bpjskes("fcbaname")%></td>
							<% sprevFcba = tempfcba %>	
						<% else	%>
							<td align="center"></td>
							<td align="center"></td>
						<% end if %>
					<% end if %>

					<%If var_colspan = 1 Then%>
						<td align="right"></td>
					<%End If%>
						<td align="right"><%=FormatNumber(rsrekap_bpjskes("bonus"), 2)%></td>
					</tr>
				<%End If%>
<%
				subtotalbonus = subtotalbonus + cdbl(rsrekap_bpjskes("bonus"))
%>
			<%Else%>
				<tr>
				
				<% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <>  "" then %>
					<% tempGroupPlantation = rsrekap_bpjskes("GroupPlantation")
					   tempPlantation = "--ALL--"
					   tempGroupRegion = "--SUMMARY--"
					   tempRegion = "--SUMMARY--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					%>
					<% if tempGroupPlantation <> SprevGroupPlantation then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtYear=<%=v_txtYear%>&txtMonth=<%=v_txtMonth%>">
							<%=rsrekap_bpjskes("GroupPlantation")%>
							</a>
						</td>
						<% SprevGroupPlantation = tempGroupPlantation %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>
				<% end if %>
				
				
				<% if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  "" then %>
					<%if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsrekap_bpjskes("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if
					
					   tempPlantation = rsrekap_bpjskes("Plantation")
					   tempGroupRegion = "--ALL--"
					   tempRegion = "--SUMMARY--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					%>
					<% if tempPlantation <> SprevPlantation then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtYear=<%=v_txtYear%>&txtMonth=<%=v_txtMonth%>">
							<%=rsrekap_bpjskes("Plantation")%>
							</a>
						</td>
						<% SprevPlantation = tempPlantation %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtGroupRegion <>  "--SUMMARY--"  and  v_txtGroupRegion <>  "" then %>
					<% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsrekap_bpjskes("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsrekap_bpjskes("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   tempGroupRegion = rsrekap_bpjskes("GroupRegion")
					   tempRegion = "--ALL--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					%>
					<% if tempGroupRegion <> SprevGroupRegion then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtYear=<%=v_txtYear%>&txtMonth=<%=v_txtMonth%>">
							<%=rsrekap_bpjskes("GroupRegion")%>
							</a>
						</td>
						<% SprevGroupRegion= tempGroupRegion %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
				    <% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsrekap_bpjskes("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsrekap_bpjskes("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion <>  "--SUMMARY--" and v_txtGroupregion <> "" then
					       tempGroupRegion = rsrekap_bpjskes("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   tempRegion = rsrekap_bpjskes("Region") 
					   tempCompany ="--ALL--"
					   tempfcba = "--SUMMARY--"
					%>
					
					<% if tempRegion <> SprevRegion then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtYear=<%=v_txtYear%>&txtMonth=<%=v_txtMonth%>">
							<%=rsrekap_bpjskes("region")%>
							</a>
						</td>
						<% SprevRegion = tempRegion %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then %>
				    <% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsrekap_bpjskes("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsrekap_bpjskes("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion <>  "--SUMMARY--" and v_txtGroupregion <> "" then
					       tempGroupRegion = rsrekap_bpjskes("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   if  v_txtregion <>  "--SUMMARY--" and v_txtregion <> "" then
					       tempRegion = rsrekap_bpjskes("Region") 
					   else
					       tempRegion = ""
					   end if 

					   tempCompany =rsrekap_bpjskes("Company")
					   tempfcba = "--ALL--"	   
					 %>
					
					<% if tempCompany <> SprevCompany then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtYear=<%=v_txtYear%>&txtMonth=<%=v_txtMonth%>">
							<%=rsrekap_bpjskes("company")%>
							</a>
						</td>
						<% SprevCompany = tempCompany %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>
				<% end if %>
				
				<% if v_txtFcba <>  "--SUMMARY--" and v_txtFcba <>  "" then %>
				    <% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsrekap_bpjskes("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsrekap_bpjskes("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion <>  "--SUMMARY--" and v_txtGroupregion <> "" then
					       tempGroupRegion = rsrekap_bpjskes("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   if  v_txtregion <>  "--SUMMARY--" and v_txtregion <> "" then
					       tempRegion = rsrekap_bpjskes("Region") 
					   else
					       tempRegion = ""
					   end if 
					   
					   if  v_txtcompany <>  "--SUMMARY--" and v_txtcompany <> "" then
					       tempCompany =rsrekap_bpjskes("Company")
					   else
					       tempCompany =""
					   end if

					   tempfcba = rsrekap_bpjskes("fcba")	   
					 %>
					
					<% if tempfcba <> sprevFcba then %>
						<td align="center">
							<a href="p_rincian_bonus.asp?txtfcba=<%=tempfcba%>&txtYear=<%=v_txtYear%>&txtMonth=<%=v_txtMonth%>">
						    <%=rsrekap_bpjskes("fcba")%>
							</a>
						</td>
						<td align="left"><%=rsrekap_bpjskes("fcbaname")%></td>
						<% sprevFcba = tempfcba %>	
					<% else	%>
						<td align="center"></td>
						<td align="center"></td>
					<% end if %>
				<% end if %>

				<%If var_colspan = 1 Then%>
					<td align="right"></td>
				<%End If%>
					<td align="right"><%=FormatNumber(rsrekap_bpjskes("bonus"), 2)%></td>
				</tr>
			<%End If%>
<%
			grandtotalbonus = grandtotalbonus + Cdbl(rsrekap_bpjskes("bonus"))
			
			rsrekap_bpjskes.MoveNext
			PrevRegion = CurrRegion
		Wend	
	rsrekap_bpjskes.Close
	set rsrekap_bpjskes = Nothing
%>
			<%if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then%>
				<tr style="background-color:#F4F4A0;">
					<td align="left" colspan="<%=var_colspan-1%>" style="padding-left:10px;"><strong>Sub Total <%=PCase(PrevRegion)%></strong></td>
					<td align="right"><strong><%=FormatNumber(subtotalbonus, 2)%></strong></td>
				</tr>
			<%End If%>
				<tr style="background-color:#F4F4A0;">
					<td align="left" colspan="<%=var_colspan-1%>" style="padding-left:10px;"><strong>Grand Total</strong></td>
					<td align="right"><strong><%=FormatNumber(grandtotalbonus, 2)%></strong></td>
				</tr>
<%End if%>
     </tbody>
</table>
	
   </body>
   </html>
