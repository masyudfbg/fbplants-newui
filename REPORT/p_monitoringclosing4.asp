<%@Language=VBScript%>
<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->

<%
Dim v_txtYear, v_txtMonth, v_txtJobcode, v_txtAfdeling, v_txtTahunTanam, v_txtBlok, v_Isactive
Dim ScriptName
Dim rsmonitoringclosing, strmonitoringclosing, strMain, strmonitoringclosing2
Dim NO, cClass, nday, xNum, nSpanTotal

Dim var_colspan : var_colspan = 1

Dim CompanyName, q, fcbaname, MgrName, KTUName, CropName, FcbaType

function f_color (datefrom__, dateto__)
	result__ = "background-color:#f53333;"
	if "" & datefrom__ <> "" and "" & dateto__ <> "" then
		if DateDiff("d",Cdate(datefrom__),CDate(dateto__)) >= 0 and DateDiff("d",Cdate(datefrom__),CDate(dateto__)) <= 2 then
			result__ = "background-color:#92d050;"
		elseif DateDiff("d",Cdate(datefrom__),CDate(dateto__)) > 2 then
			result__ = "background-color:#ffc000;"
		else
			result__ = "background-color:#f53333;"
		end if
	end if
	f_color = result__
end function

nSpanTotal = 9

v_Isactive = ""
v_txtGroupPlantation 	= Request("txtGroupPlantation")
v_txtPlantation  		= Request("txtPlantation")
v_txtGroupRegion  		= Request("txtGroupRegion")
v_txtRegion  			= Request("txtRegion")
v_txtcompany  			= Request("txtCompany")
v_txtfcba 				= Request("txtfcba")
v_Isactive              = Request("txtIsactive")

If v_Isactive = "" then
   v_Isactive="TRUE"
end if




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
		set rsCat = connStr.Execute(q)		
		if not rsCat.eof then
			CompanyName = rsCat("FCCOMPANYNAME")	
		end if
		q	=	""
		rsCat.Close
		set rsCat = Nothing
	End If
End If

If Len(Trim(v_txtfcba)) = 0 Then v_txtfcba = "--SUMMARY--"

If Len(Trim(v_txtfcba)) > 0 Then
	if v_txtfcba = "--SUMMARY--" then
		selectCons	=	selectCons	& " , nourut, fcba"
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " , nourut"
		nSpanTotal	=	nSpanTotal	- 1
	Elseif v_txtfcba = "--ALL--" then
		selectCons	=	selectCons	& " , nourut, fcba"
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " , nourut"
	Else
		selectCons	=	selectCons	& " , nourut, fcba"
		whereCons	=	whereCons	& " And fcba = '" & v_txtfcba & "'"
		orderCons	=	orderCons	& " , nourut"
		
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

If v_Isactive <> "ALL" then
   whereCons	=	whereCons	& " and businessunit_api.Isactive(fcba) = '" & v_Isactive & "'"	
end if

v_txtYear = Request("txtYear")
v_txtMonth = Request("txtMonth")
txtdate = Request("txtdate")

If "" & v_txtYear = "" Then v_txtYear = Year(Session("v_Costbook"))
If "" & v_txtMonth = "" Then v_txtMonth = Month(Session("v_Costbook"))
If "" & txtdate = "" Then txtdate = textDate(date)
v_txtYear = Year(txtdate)
v_txtMonth = Month(txtdate)

datefrom = v_txtMonth&"/01/"&v_txtYear
dateuntil = txtdate
maxday = day(txtdate)


ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

strMain		=	" SELECT" & strFieldCons & _
				"        , SUM(DECODE(TO_NUMBER(TO_CHAR(execution_date, 'dd')), 1, total, 0)) day_1" & _
				"        , SUM(DECODE(TO_NUMBER(TO_CHAR(execution_date, 'dd')), 2, total, 0)) day_2" & _
				"        , SUM(DECODE(TO_NUMBER(TO_CHAR(execution_date, 'dd')), 3, total, 0)) day_3" & _
				"        , SUM(DECODE(TO_NUMBER(TO_CHAR(execution_date, 'dd')), 4, total, 0)) day_4" & _
				"        , SUM(DECODE(TO_NUMBER(TO_CHAR(execution_date, 'dd')), 5, total, 0)) day_5" & _
				"        , SUM(DECODE(TO_NUMBER(TO_CHAR(execution_date, 'dd')), 6, total, 0)) day_6" & _
				"        , SUM(DECODE(TO_NUMBER(TO_CHAR(execution_date, 'dd')), 7, total, 0)) day_7" & _
				"        , SUM(DECODE(TO_NUMBER(TO_CHAR(execution_date, 'dd')), 8, total, 0)) day_8" & _
				"        , SUM(DECODE(TO_NUMBER(TO_CHAR(execution_date, 'dd')), 9, total, 0)) day_9" & _
				"        , SUM(DECODE(TO_NUMBER(TO_CHAR(execution_date, 'dd')), 10, total, 0)) day_10" & _
				"        , SUM(DECODE(TO_NUMBER(TO_CHAR(execution_date, 'dd')), 11, total, 0)) day_11" & _
				"        , SUM(DECODE(TO_NUMBER(TO_CHAR(execution_date, 'dd')), 12, total, 0)) day_12" & _
				"        , SUM(DECODE(TO_NUMBER(TO_CHAR(execution_date, 'dd')), 13, total, 0)) day_13" & _
				"        , SUM(DECODE(TO_NUMBER(TO_CHAR(execution_date, 'dd')), 14, total, 0)) day_14" & _
				"        , SUM(DECODE(TO_NUMBER(TO_CHAR(execution_date, 'dd')), 15, total, 0)) day_15" & _
				"        , SUM(DECODE(TO_NUMBER(TO_CHAR(execution_date, 'dd')), 16, total, 0)) day_16" & _
				"        , SUM(DECODE(TO_NUMBER(TO_CHAR(execution_date, 'dd')), 17, total, 0)) day_17" & _
				"        , SUM(DECODE(TO_NUMBER(TO_CHAR(execution_date, 'dd')), 18, total, 0)) day_18" & _
				"        , SUM(DECODE(TO_NUMBER(TO_CHAR(execution_date, 'dd')), 19, total, 0)) day_19" & _
				"        , SUM(DECODE(TO_NUMBER(TO_CHAR(execution_date, 'dd')), 20, total, 0)) day_20" & _
				"        , SUM(DECODE(TO_NUMBER(TO_CHAR(execution_date, 'dd')), 21, total, 0)) day_21" & _
				"        , SUM(DECODE(TO_NUMBER(TO_CHAR(execution_date, 'dd')), 22, total, 0)) day_22" & _
				"        , SUM(DECODE(TO_NUMBER(TO_CHAR(execution_date, 'dd')), 23, total, 0)) day_23" & _
				"        , SUM(DECODE(TO_NUMBER(TO_CHAR(execution_date, 'dd')), 24, total, 0)) day_24" & _
				"        , SUM(DECODE(TO_NUMBER(TO_CHAR(execution_date, 'dd')), 25, total, 0)) day_25" & _
				"        , SUM(DECODE(TO_NUMBER(TO_CHAR(execution_date, 'dd')), 26, total, 0)) day_26" & _
				"        , SUM(DECODE(TO_NUMBER(TO_CHAR(execution_date, 'dd')), 27, total, 0)) day_27" & _
				"        , SUM(DECODE(TO_NUMBER(TO_CHAR(execution_date, 'dd')), 28, total, 0)) day_28" & _
				"        , SUM(DECODE(TO_NUMBER(TO_CHAR(execution_date, 'dd')), 29, total, 0)) day_29" & _
				"        , SUM(DECODE(TO_NUMBER(TO_CHAR(execution_date, 'dd')), 30, total, 0)) day_30" & _
				"        , SUM(DECODE(TO_NUMBER(TO_CHAR(execution_date, 'dd')), 31, total, 0)) day_31" & _
				" FROM (" & _
				"     SELECT" & _
				"     master.fcba, master.execution_date, NVL(data.total,0) total" & _
				"     FROM (" & _
				"         SELECT bu.fccode fcba, listdate.execution_date" & _
				"         FROM ( SELECT TO_DATE ('"&datefrom&"', 'MM/DD/YYYY') +ROWNUM -1 execution_date" & _
				"                FROM ALL_OBJECTS" & _
				"                WHERE ROWNUM <= TO_DATE ('"&dateuntil&"', 'MM/DD/YYYY') - TO_DATE ('"&datefrom&"', 'MM/DD/YYYY') + 1) listdate" & _
				"              , businessunit bu" & _
				"         WHERE bu.fctype = 'M'" & _
				"     ) master" & _
				"     LEFT JOIN ( SELECT DISTINCT fcba, execution_date, 1 total" & _
				"                 FROM process_month_log" & _
				"                 WHERE year_process = '"&v_txtYear&"'" & _
				"                       AND month_process = '"&v_txtMonth&"') data" & _
				"     ON (master.fcba = data.fcba AND master.execution_date = data.execution_date) )" & _
				" GROUP BY fcba"
								
strmonitoringclosing2 = " , day_1, day_2, day_3, day_4, day_5, day_6, day_7, day_8, day_9, day_10" & _
                        " , day_11, day_12, day_13, day_14, day_15, day_16, day_17, day_18, day_19, day_20" & _
                        " , day_21, day_22, day_23, day_24, day_25, day_26, day_27, day_28, day_29, day_30, day_31"

strmonitoringclosing = " SELECT GroupPlantation, businessunit_api.Isactive(fcba) fcbastatus , businessunit_api.get_FCTYPE(fcba) fcbatype  " & selectCons & strmonitoringclosing2 &"  FROM (" & strMain & " )" & whereCons & _
" ORDER BY nourut"

'Response.write strmonitoringclosing
%>
<html>

<head>
<title>ePlaS Reporting - Monitoring Closing</title>
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
							<font face="verdana" size="1">Kebun/Pabrik Active</font>
						</td>
						<td>
							<select name="txtIsactive">
							<option value="TRUE" <%if v_Isactive="TRUE" Then%> Selected <%End if%>>TRUE</option>
							<option value="FALSE" <%if v_Isactive="FALSE" Then%> Selected <%End if%>>FALSE</option>
							<option value="ALL" <%if v_Isactive="ALL" Then%> Selected <%End if%>>ALL</option>
							</select>
						</td>
						<!--
						<td>
							<font face="verdana" size="1">TAHUN</font>
						</td>
						<% 
							Dim mon, Yr, MinYr, MaxYr, i, MM
							mon = Month(Session("v_Costbook"))
							If len(v_txtMonth)=0 then v_txtMonth=mon
						
							Yr = Year(Session("v_Costbook"))
							MinYr = Yr - 2
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
						-->
						<td>
							<font face="verdana" size="1">Tanggal</font>
						</td>
						<td width="4">
							<input class="datepicker" type="text" name="txtdate" value="<%=txtdate%>" size="10">
						</td>
						<td>
							<input type="submit" value="View" />
						</td>
					</tr>
				</table>
			</div>
			</form>
		</div>

<table width="100%" id="product-table" class="xthf-site" style="border-bottom:1px #3A5F0B solid;">
	<thead style="border-left:1px #000000 solid;">
		<tr>
			<th colspan="<%=var_colspan +4 +maxday%>" style="background:#FFFFFF; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
				<div id="header-wrapper">
		<div id="head_judul">
				<% if  CompanyName <>  "" then %>
					<h4>PT.<%=CompanyName%></h4> 
				<% End if%>
				<% If  FcbaName <> ""  then %>
					   <h5><%=FcbaName%></h5>   
				<% End if%>
				</div>
				<div id="subheader" class="centered">
					<h2>MONITORING CLOSING</h2>
					<br/>
					<h4>Per Tanggal : <%=day(txtdate) & " " & bulan(month(txtdate)) & " " & year(txtdate)%></h4>
				</div>
              </div
			</th>
		</tr>
	<tr>
	    <th style="background-color:#00A300;color:#fff;" rowspan="2">No</th>
		<%If Len(Trim(v_txtGroupPlantation)) > 0 And v_txtGroupPlantation <>  "--SUMMARY--" Then%>
			<th style="background-color:#00A300;color:#fff;" rowspan="2"><%=titleGroupPlantation%></th>
		<%End If%>

		<%If Len(Trim(v_txtPlantation)) > 0 And v_txtPlantation <>  "--SUMMARY--" Then%>
			<th style="background-color:#00A300;color:#fff;" rowspan="2"><%=titlePlantation%></th>
		<%End If%>

		<%If Len(Trim(v_txtGroupRegion)) > 0 And v_txtGroupRegion <>  "--SUMMARY--" Then%>
			<th style="background-color:#00A300;color:#fff;" rowspan="2"><%=titleGroupRegion%></th>
		<%End If%>

		<%If Len(Trim(v_txtRegion)) > 0 And v_txtRegion <>  "--SUMMARY--" Then%>
			<th style="background-color:#00A300;color:#fff;" rowspan="2"><%=titleRegion%></th>
		<%End If%>

		<%If Len(Trim(v_txtCompany)) > 0 And v_txtCompany <>  "--SUMMARY--" Then%>
			<th style="background-color:#00A300;color:#fff;" rowspan="2"><%=titleCompany%></th>
		<%End If%>

		<%If Len(Trim(v_txtfcba)) > 0 Then%>
			<th style="background-color:#00A300;color:#fff;" rowspan="2"><%=titleSite%></th>
		<%End If%>
		
		<th style="background-color:#00A300;color:#fff;" colspan="<%=maxday%>">Closing Period</th>
	</tr>
	<tr>
		<%for xn=1 to maxday%>
		<th style="background-color:#00A300;color:#fff;"><%=xn%></th>
		<%next%>
	</tr>

    </thead>
	<tbody>
<%
Set rsmonitoringclosing = Server.CreateObject("ADODB.Recordset")
rsmonitoringclosing.Open strmonitoringclosing, connstr, adOpenStatic

Dim Rp_Per_Ha_RowReal : Rp_Per_Ha_RowReal = 0
Dim Rp_Per_Ha_RowBudget : Rp_Per_Ha_RowBudget = 0
Dim RowPercentage : RowPercentage = 0

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


if rsmonitoringclosing.bof and rsmonitoringclosing.eof then
%>
	<tr>
		<td colspan="<%=var_colspan+4%>">NO DATA AVAILABLE</td>
	</tr>
<%
else
		No=0
		while not rsmonitoringclosing.eof
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
		    No = No + 1 
%>
				<tr class="<%=cClass%>">
				<td align="center"><%=No%></td>
				<% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <>  "" then %>
					<% tempGroupPlantation = rsmonitoringclosing("GroupPlantation")
					   tempPlantation = "--ALL--"
					   tempGroupRegion = "--SUMMARY--"
					   tempRegion = "--SUMMARY--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					%>
					<% if tempGroupPlantation <> SprevGroupPlantation then %>
						<td align="center" style=" border-bottom:none;">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>">
							<%=rsmonitoringclosing("GroupPlantation")%>
							</a>
						</td>
						<% SprevGroupPlantation = tempGroupPlantation %>	
					<% else	%>
						<td align="center" style="border:none;border-right:1px #3A5F0B solid; border-left:1px #3A5F0B solid;"></td>					
					<% end if %>
				<% end if %>
				
				
				<% if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  "" then %>
					<% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsmonitoringclosing("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if
					   
					   tempPlantation = rsmonitoringclosing("Plantation")
					   tempGroupRegion = "--ALL--"
					   tempRegion = "--SUMMARY--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					%>
					<% if tempPlantation <> SprevPlantation then %>
						<td align="center" style=" border-bottom:none;">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>">
							<%=rsmonitoringclosing("Plantation")%>
							</a>
						</td>
						<% SprevPlantation = tempPlantation %>	
					<% else	%>
						<td align="center" style="border:none;border-right:1px #3A5F0B solid; border-left:1px #3A5F0B solid;"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtGroupRegion <>  "--SUMMARY--"  and  v_txtGroupRegion <>  "" then %>
					<% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsmonitoringclosing("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsmonitoringclosing("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   tempGroupRegion = rsmonitoringclosing("GroupRegion")
					   tempRegion = "--ALL--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					%>
					<% if tempGroupRegion <> SprevGroupRegion then %>
						<td align="center" style=" border-bottom:none;">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>">
							<%=rsmonitoringclosing("GroupRegion")%>
							</a>
						</td>
						<% SprevGroupRegion= tempGroupRegion %>	
					<% else	%>
						<td align="center" style="border:none;border-right:1px #3A5F0B solid; border-left:1px #3A5F0B solid;"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
				    <% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsmonitoringclosing("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsmonitoringclosing("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion <>  "--SUMMARY--"  and  v_txtGroupRegion <> "" then
					       tempGroupRegion = rsmonitoringclosing("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   tempRegion = rsmonitoringclosing("Region") 
					   tempCompany ="--ALL--"
					   tempfcba = "--SUMMARY--"
					%>
					
					<% if tempRegion <> SprevRegion then %>
						<td align="center" style=" border-bottom:none;">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>">
							<%=rsmonitoringclosing("region")%>
							</a>
						</td>
						<% SprevRegion = tempRegion %>	
					<% else	%>
						<td align="center" style="border:none;border-right:1px #3A5F0B solid; border-left:1px #3A5F0B solid;"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then %>
				    <% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsmonitoringclosing("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsmonitoringclosing("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion <>  "--SUMMARY--" and v_txtGroupregion <> "" then
					       tempGroupRegion = rsmonitoringclosing("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   if  v_txtregion <>  "--SUMMARY--" and v_txtregion <> "" then
					       tempRegion = rsmonitoringclosing("Region") 
					   else
					       tempRegion = ""
					   end if
					   tempCompany =rsmonitoringclosing("Company")
					   tempfcba = "--ALL--"
					 %>
					
					<% if tempCompany <> SprevCompany then %>
						<td align="center" style=" border-bottom:none;">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>">
							<%=rsmonitoringclosing("company")%>
							</a>
						</td>
						<% SprevCompany = tempCompany %>	
					<% else	%>
						<td align="center" style="border:none;border-right:1px #3A5F0B solid; border-left:1px #3A5F0B solid;"></td>					
					<% end if %>
				<% end if %>
				
				<% if v_txtFcba <>  "" then %>
				    <% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsmonitoringclosing("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsmonitoringclosing("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion  <>  "--SUMMARY--" and v_txtGroupregion <> "" then
					       tempGroupRegion = rsmonitoringclosing("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   if  v_txtregion  <>  "--SUMMARY--" and v_txtregion <> "" then
					       tempRegion = rsmonitoringclosing("Region") 
					   else
					       tempRegion = ""
					   end if
					   
					   if  v_txtcompany  <>  "--SUMMARY--" and v_txtcompany <> "" then
					       tempCompany =rsmonitoringclosing("Company")
					   else
					       tempCompany =""
					   end if
					   
					   tempfcba = rsmonitoringclosing("fcba")
					 %>
					
					<% if tempfcba <> sprevFcba then %>
						<td align="center" style=" border-bottom:none;">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>">
						    <%=rsmonitoringclosing("fcba")%>
							</a>
						</td>
						<% sprevFcba = tempfcba %>	
					<% else	%>
						<td align="center" style="border:none;border-right:1px #3A5F0B solid; border-left:1px #3A5F0B solid;"></td>					
					<% end if %>
				<% end if %>
					
					<%for xn=1 to maxday%>
						<%if cdbl(to_nol(rsmonitoringclosing("day_"&xn))) = 1 then%>
							<td align="center" style="color:#0074D9;font-weight:bold;">v</td>
						<%else%>
							<td align="center" style="color:#FF4136;font-weight:bold;">x</td>
						<%end if%>
					<%next%>
				</tr>
<%
			rsmonitoringclosing.MoveNext
			Wend	
rsmonitoringclosing.Close
set rsmonitoringclosing = Nothing
%>
<%End if%>
	</tbody>
</table>

   </body>
   </html>