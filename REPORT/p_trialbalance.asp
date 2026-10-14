<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
Dim v_txtYear, v_txtMonth, v_txtjobgroup, v_txtjobcode, v_txtnonull
Dim ScriptName
Dim rstrialbalance, strtrialbalance, strMain, strtrialbalance2
Dim NO, cClass, nday, xNum, nSpanTotal, nCountofField

Dim var_colspan : var_colspan = 1

Dim CompanyName, q, fcbaname, MgrName, KTUName, CropName, FcbaType

nSpanTotal = 9

v_txtYear = Request("txtYear")
v_txtMonth = Request("txtMonth")
v_txtjobgroup = Request("txtjobgroup")
v_txtjobcode = Request("txtjobcode")
v_txtnonull = Request("no-null")

if Request("no-null") = "" Or Len(Trim(Request("no-null"))) = 0 then v_txtnonull = "full"

if request("txtYear")="" then
   v_txtYear	=	Year(Session("v_Costbook"))
end if
if request("txtMonth")="" then
   v_txtMonth	=	Month(Session("v_Costbook"))
end if

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
					" , BUSINESSUNIT_API.get_OrderData(fcba) NoUrut, fcba"
					
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
		set rstrialbalance = connStr.Execute(q)		
		if not rstrialbalance.eof then
			CompanyName = rstrialbalance("FCCOMPANYNAME")	
		end if
		q	=	""
		rstrialbalance.Close
		set rstrialbalance = Nothing
	End If
End If

If Len(Trim(v_txtfcba)) = 0 Then v_txtfcba = "--SUMMARY--"

If Len(Trim(v_txtfcba)) > 0 Then
	if v_txtfcba = "--SUMMARY--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
		nSpanTotal	=	nSpanTotal	- 1
	Elseif v_txtfcba = "--ALL--" then
		selectCons	=	selectCons	& " , nourut, fcba"
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " , nourut, fcba"
		var_colspan	=	var_colspan	+ 1
	Else
		selectCons	=	selectCons	& " , nourut, fcba"
		whereCons	=	whereCons	& " And fcba = '" & v_txtfcba & "'"
		orderCons	=	orderCons	& " , nourut, fcba"
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
				" FCCODE = '" & v_txtfcba & "'"	
		set rstrialbalance = connStr.Execute(q)		
		if not rstrialbalance.eof then
			fcbaname = rstrialbalance("FCNAME")
			CompanyName=rstrialbalance("FCCOMPANYNAME")
			MgrName = rstrialbalance("FCMANAGER")
			KTUName = rstrialbalance("FCKTU")
			CropName=rstrialbalance("FCCROP")
			FcbaType=rstrialbalance("FCTYPE")
		end if
		q	=	""
		rstrialbalance.Close
		set rstrialbalance = Nothing
	End If
End If



ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

strMain = " select " & strFieldCons & _
          "        , fccode jobcode" & _
          "        , fcname jobname" & _
          "        , nvl(sum(opening_balance), 0) opening_balance" & _
          "        , nvl(sum(debet), 0) debet" & _
          "        , nvl(sum(credit), 0) credit" & _
          "        , nvl(sum(opening_balance)+sum(debet)+sum(credit), 0) saldo" & _
          " from trialbalance" & _
          " where period_year = " & v_txtYear & "" & _
          "       and period_month = " & v_txtMonth & ""
if  len(trim(v_txtjobgroup)) > 0 then
    strMain = strMain & " and job_api.get_jobgroupCode(fccode, fcba) = " & v_txtjobgroup & ""
end if
if  len(trim(v_txtjobcode)) > 0 then
    strMain = strMain & " and fccode = '" & v_txtjobcode & "'"
end if 
	strMain = strMain & " group by fcba" & _
                        "          , fccode" & _
                        "          , fcname"

								
strtrialbalance2 = "          , jobcode" & _
                   "          , jobname" & _
				   "          , sum(opening_balance) opening_balance" & _
                   "          , sum(debet) debet" & _
                   "          , sum(credit) credit" & _
                   "          , sum(saldo) saldo"

strtrialbalance	=	" select GroupPlantation " & selectCons & strtrialbalance2 &"  From (" & strMain & " )" & _ 
                    "" & whereCons & _
                    " group by GroupPlantation" & selectCons & ", jobcode, jobname"
if v_txtnonull = "no-null" then
	strtrialbalance = strtrialbalance & " having ( nvl(sum(opening_balance), 0) <> 0 Or nvl(sum(debet), 0) <> 0 Or nvl(sum(credit), 0) <> 0 )"
End If
	strtrialbalance = strtrialbalance & " Order by GroupPlantation" & orderCons & ", jobcode, jobname"

' Response.write strtrialbalance
%>
<head>
<title>ePlaS Reporting - Neraca Perobaan</title>
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
		<div id="header-wrapper">
			<div class="printhide">
				<!--#include file="../print_excel.asp"-->
				<!--#include file="menutop.asp"-->
			</div>
			<form action="<%=ScriptName%>" method="post" name="frmEntry">
			<div class="printhide">
				<!--#include virtual="/REPORT/FormSelectNewSum.html"-->
				<table id="form-table" class="centered-margin">
					<tr>
						<td><font face="verdana" size="1">YEAR</font></td>
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
						For i = 1 to 15
				%>
						<Option Value="<%=i%>" <%If cint(i) = cint(v_txtMonth) then%> Selected <%End If%>><%=BulanAP(i)%></option>
				<%
					Next
				%>
							</select>
						</td>
						<td><font face="verdana" size="1">JOBGROUP</font></td>
						<td><input type="text" name="txtjobgroup" value="<%=v_txtjobgroup%>"></td>
						<td><font face="verdana" size="1">JOBCODE</font></td>
						<td><input type="text" name="txtjobcode" value="<%=v_txtjobcode%>"></td>
						<td><input type="submit" value="View" /></td>
						<td><input type="hidden" name="no-null" value="<%=v_txtnonull%>"></td>
					</tr>
				</table>
			</div>
			</form>
		</div>

<table width="100%" id="product-table" class="xthf-site" style="border-bottom:1px #3A5F0B solid;">
	<thead style="border-left:1px #000000 solid;">
	<tr>
		<th colspan="<%=var_colspan+5%>" style="background:#FFFFFF; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
			<div id="header-wrapper">
				<div id="head_judul" style="margin-bottom:-30px; position:static; text-align:left;">
				<% if  CompanyName <>  "" then %>
					<h4>PT.<%=CompanyName%></h4> 
				<% End if%>
				<% If  FcbaName <> ""  then %>
					   <h5><%=FcbaName%></h5>   
				<% End if%>
				</div>
			<div id="subheader" class="centered">
				<br/><br/>
				<h2>TRIAL BALANCE</h2>
				<br/>
				<h4>PERIODE&nbsp;:&nbsp;<%=BulanAP(v_txtMonth)%>&nbsp;<%=v_txtYear%></h4>
			</div>
		  </div>
			<div class="printhide">
			<%=v_txtnonull%>
			<form name="formnull" method="post" action="<%=ScriptName%>" >
			<select name="no-null" style="float:right;margin-top:-30px;width:100px;" onChange="this.form.submit();">
			   <% if v_txtnonull = "full" then %>
				<option value="full" selected>Detail</option>
				<option value="no-null">No Detail</option>
				<%elseif v_txtnonull = "no-null" then %>
				<option value="full">Detail</option>
				<option value="no-null" selected>No Detail</option>
				<%else%>
				<option value="full" selected>Detail</option>
				<option value="no-null">No Detail</option>
				<%end if%>
			</select>
			<input type="hidden" name="txtyear" value="<%=v_txtyear%>">
			<input type="hidden" name="txtmonth" value="<%=v_txtmonth%>">
			<input type="hidden" name="txtGroupPlantation" value="<%=v_txtGroupPlantation%>">
			<input type="hidden" name="txtPlantation" value="<%=v_txtPlantation%>">
			<input type="hidden" name="txtGroupRegion" value="<%=v_txtGroupRegion%>">
			<input type="hidden" name="txtRegion" value="<%=v_txtRegion%>">
			<input type="hidden" name="txtcompany" value="<%=v_txtcompany%>">
			<input type="hidden" name="txtfcba" value="<%=v_txtfcba%>">
			<input type="hidden" name="txtjobgroup" value="<%=v_txtjobgroup%>">
			<input type="hidden" name="txtjobcode" value="<%=v_txtjobcode%>">
			</form>
			</div>
		</th>

	</tr>
	<tr>
	<%
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

	Set rstrialbalance = connStr.Execute(strtrialbalance)
	if Not rstrialbalance.eof then
	   nCountofField = rstrialbalance.fields.count - 1
	end if
	NO=1%>
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

		<%If Len(Trim(v_txtfcba)) > 0 And v_txtfcba <>  "--SUMMARY--"  Then%>
			<th style="background-color:#00A300;color:#fff;"><%=titleSite%></th>
		<%End If%>
		<th style="background-color:#00A300;color:#fff;">Account Code</th>
		<th style="background-color:#00A300;color:#fff;">Account Description</th>
		<th style="background-color:#00A300;color:#fff;">Opening Balance</th>
		<th style="background-color:#00A300;color:#fff;">Debet</th>
		<th style="background-color:#00A300;color:#fff;">Credit</th>
		<th style="background-color:#00A300;color:#fff;">Saldo</th>
	</tr>
</thead>
<tbody>
	<%
	Dim v_Opening,v_Debit,v_Credit,v_Saldo
	v_Opening= 0
	v_Debit  = 0
	v_Credit = 0
	v_Saldo  = 0
	While Not rstrialbalance.EOF
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if	
%>
		<tr class="<%=cClass%>">
				<% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <>  "" then %>
					<% tempGroupPlantation = rstrialbalance("GroupPlantation")
					   tempPlantation = "--ALL--"
					   tempGroupRegion = "--SUMMARY--"
					   tempRegion = "--SUMMARY--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					%>
					<% if tempGroupPlantation <> SprevGroupPlantation then %>
						<td align="center" style=" border-bottom:none;">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtYear%>&txtMonth=<%=v_txtMonth%>&txtjobgroup=<%=v_txtjobgroup%>&txtjobcode=<%=v_txtjobcode%>&no-null=<%=v_txtnonull%>">
							<%=rstrialbalance("GroupPlantation")%>
							</a>
						</td>
						<% SprevGroupPlantation = tempGroupPlantation %>	
					<% else	%>
						<td align="center" style="border:none;border-right:1px #3A5F0B solid; border-left:1px #3A5F0B solid;"></td>					
					<% end if %>
				<% end if %>
				
				
				<% if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  "" then %>
					<% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rstrialbalance("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if
					   
					   tempPlantation = rstrialbalance("Plantation")
					   tempGroupRegion = "--ALL--"
					   tempRegion = "--SUMMARY--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					%>
					<% if tempPlantation <> SprevPlantation then %>
						<td align="center" style=" border-bottom:none;">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtYear%>&txtMonth=<%=v_txtMonth%>&txtjobgroup=<%=v_txtjobgroup%>&txtjobcode=<%=v_txtjobcode%>&no-null=<%=v_txtnonull%>">
							<%=rstrialbalance("Plantation")%>
							</a>
						</td>
						<% SprevPlantation = tempPlantation %>	
					<% else	%>
						<td align="center" style="border:none;border-right:1px #3A5F0B solid; border-left:1px #3A5F0B solid;"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtGroupRegion <>  "--SUMMARY--"  and  v_txtGroupRegion <>  "" then %>
					<% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rstrialbalance("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rstrialbalance("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   tempGroupRegion = rstrialbalance("GroupRegion")
					   tempRegion = "--ALL--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					%>
					<% if tempGroupRegion <> SprevGroupRegion then %>
						<td align="center" style=" border-bottom:none;">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtYear%>&txtMonth=<%=v_txtMonth%>&txtjobgroup=<%=v_txtjobgroup%>&txtjobcode=<%=v_txtjobcode%>&no-null=<%=v_txtnonull%>">
							<%=rstrialbalance("GroupRegion")%>
							</a>
						</td>
						<% SprevGroupRegion= tempGroupRegion %>	
					<% else	%>
						<td align="center" style="border:none;border-right:1px #3A5F0B solid; border-left:1px #3A5F0B solid;"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
				    <% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rstrialbalance("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rstrialbalance("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion <>  "--SUMMARY--"  and  v_txtGroupRegion <> "" then
					       tempGroupRegion = rstrialbalance("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   tempRegion = rstrialbalance("Region") 
					   tempCompany ="--ALL--"
					   tempfcba = "--SUMMARY--"
					%>
					
					<% if tempRegion <> SprevRegion then %>
						<td align="center" style=" border-bottom:none;">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtYear%>&txtMonth=<%=v_txtMonth%>&txtjobgroup=<%=v_txtjobgroup%>&txtjobcode=<%=v_txtjobcode%>&no-null=<%=v_txtnonull%>">
							<%=rstrialbalance("region")%>
							</a>
						</td>
						<% SprevRegion = tempRegion %>	
					<% else	%>
						<td align="center" style="border:none;border-right:1px #3A5F0B solid; border-left:1px #3A5F0B solid;"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then %>
				    <% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rstrialbalance("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rstrialbalance("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion <>  "--SUMMARY--" and v_txtGroupregion <> "" then
					       tempGroupRegion = rstrialbalance("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   if  v_txtregion <>  "--SUMMARY--" and v_txtregion <> "" then
					       tempRegion = rstrialbalance("Region") 
					   else
					       tempRegion = ""
					   end if
					   tempCompany =rstrialbalance("Company")
					   tempfcba = "--ALL--"
					 %>
					
					<% if tempCompany <> SprevCompany then %>
						<td align="center" style=" border-bottom:none;">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtYear%>&txtMonth=<%=v_txtMonth%>&txtjobgroup=<%=v_txtjobgroup%>&txtjobcode=<%=v_txtjobcode%>&no-null=<%=v_txtnonull%>">
							<%=rstrialbalance("company")%>
							</a>
						</td>
						<% SprevCompany = tempCompany %>	
					<% else	%>
						<td align="center" style="border:none;border-right:1px #3A5F0B solid; border-left:1px #3A5F0B solid;"></td>					
					<% end if %>
				<% end if %>
				
				<% if v_txtFcba <>  "--SUMMARY--" and v_txtFcba <>  "" then %>
				    <% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rstrialbalance("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rstrialbalance("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion  <>  "--SUMMARY--" and v_txtGroupregion <> "" then
					       tempGroupRegion = rstrialbalance("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   if  v_txtregion  <>  "--SUMMARY--" and v_txtregion <> "" then
					       tempRegion = rstrialbalance("Region") 
					   else
					       tempRegion = ""
					   end if
					   
					   if  v_txtcompany  <>  "--SUMMARY--" and v_txtcompany <> "" then
					       tempCompany =rstrialbalance("Company")
					   else
					       tempCompany =""
					   end if
					   
					   tempfcba = rstrialbalance("fcba")
					 %>
					
					<% if tempfcba <> sprevFcba then %>
						<td align="center" style=" border-bottom:none;">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtYear%>&txtMonth=<%=v_txtMonth%>&txtjobgroup=<%=v_txtjobgroup%>&txtjobcode=<%=v_txtjobcode%>&no-null=<%=v_txtnonull%>">
						    <%=rstrialbalance("fcba")%>
							</a>
						</td>
						<% sprevFcba = tempfcba %>	
					<% else	%>
						<td align="center" style="border:none;border-right:1px #3A5F0B solid; border-left:1px #3A5F0B solid;"></td>					
					<% end if %>
				<% end if %>
				
		<td align="Center"><font face="Verdana" size="1"><%=rstrialbalance("jobcode")%></td>
		<td align="left"><font face="Verdana" size="1"><%=rstrialbalance("jobname")%></td>
		<td align="Right"><font face="Verdana" size="1"><%=FormatNumber(rstrialbalance("OPENING_BALANCE"), 2)%></td>
		<td align="Right"><font face="Verdana" size="1"><%=FormatNumber(rstrialbalance("Debet"), 2)%></td>
		<td align="Right"><font face="Verdana" size="1"><%=FormatNumber(rstrialbalance("Credit"), 2)%></td>
		<td align="Right"><font face="Verdana" size="1"><%=FormatNumber(rstrialbalance("Saldo"), 2)%></td>
		</tr>
	  <%
	  v_Opening=v_Opening+cdbl(rstrialbalance("OPENING_BALANCE"))
	  v_Debit  =v_Debit+(cdbl(rstrialbalance("Debet")))
	  v_Credit =v_Credit+(cdbl(rstrialbalance("Credit")))
	  v_Saldo  =v_Saldo+trim((rstrialbalance("Saldo")))
	  'response.write cdbl(trim(rstrialbalance("MTHOPENBALANCE")))+1
	  rstrialbalance.MoveNext
	  'response.Flush()
	  no=no+1
	Wend
	rstrialbalance.close()
	rstrialbalance=null%>
	</form>
	<tr bgcolor="#F4F4A0">
	<td align="Center" colspan="<%=var_colspan+1%>"> <font face="Verdana" size="1">T o t a l</td>
	<td align="Right"> <font face="Verdana" size="1"><%=FormatNumber(v_Opening,2)%></td>
	<td align="Right"> <font face="Verdana" size="1"><%=FormatNumber(v_Debit,2)%></td>
	<td align="Right"> <font face="Verdana" size="1"><%=FormatNumber(v_Credit,2)%></td>
	<td align="Right"> <font face="Verdana" size="1"><%=FormatNumber(v_Saldo,2)%></td>
	</tr>
 <tbody>
</table>
<%If Len(Trim(v_txtFcba)) > 0 And (v_txtFcba <>  Cstr("--SUMMARY--") And v_txtFcba <>  Cstr("--ALL--"))then%>
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
		<td align="center"><font face="Verdana" size="2"><strong><u>&nbsp;&nbsp;<%=MgrName%>&nbsp;&nbsp;</u></strong></td>
		<td align="center"><font face="Verdana" size="2"><strong><u>&nbsp;&nbsp;<%=KTUName%>&nbsp;&nbsp;</u></strong></td>
	   </tr>
	   	  <tr> 
		<td align="center"><font face="Verdana" size="1"><strong><i>MANAGER</i></strong></td>
		<td align="center"><font face="Verdana" size="1"><strong><i>KTU</i></strong></td>
	   </tr>
<%End If%>
   </table>