<%@Language=VBScript%>
<%Option Explicit%>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
	Dim v_txtBudgetYear, v_txtBudgetVersion, v_txtModuleVersion, v_txtdatefrom, v_txtMonth, v_txtYear
	Dim no
	
	Dim ScriptName
	Dim strMain, sqlLHPP, sqlLHPP2, rsLHPP
	Dim q, CompanyName, fcbaname, MgrName, KTUName, CropName, FcbaType

    v_txtGroupPlantation = Request("txtGroupPlantation")
	v_txtPlantation = Request("txtPlantation")
	v_txtGroupRegion = Request("txtGroupRegion")
	v_txtRegion = Request("txtRegion")
	v_txtcompany = Request("txtCompany")
	v_txtfcba = Request("txtfcba")
	
	v_txtBudgetYear = Request("txtBudgetYear")
	v_txtBudgetVersion = Request("txtBudgetVersion")
	v_txtModuleVersion = Request("txtModuleVersion")
	
	v_txtdatefrom		= Request("txtdatefrom")
	
	v_txtYear		= Request("txtYear")
	v_txtMonth		= Request("txtMonth")
	
	If Len(Trim(v_txtYear)) = 0 Then v_txtYear = Year(Session("v_Costbook"))
	If Len(Trim(v_txtMonth)) = 0 Then v_txtMonth = Month(Session("v_Costbook"))
	
	If Len(Trim(v_txtBudgetYear)) = 0 Then
		v_txtBudgetYear = Year(textDate(Date))
	Else
		v_txtBudgetYear = v_txtYear
	End If
	If Len(Trim(v_txtBudgetVersion)) = 0 Then v_txtBudgetVersion = 1
	If Len(Trim(v_txtModuleVersion)) = 0 Then v_txtModuleVersion = 1
	
	If Len(Trim(v_txtdatefrom)) = 0 Then v_txtdatefrom = textDate(Date)
	
	Dim var_colspan : var_colspan = 1
	
Dim selectFields, strFieldCons, selectCons, whereCons, GroupbyCons, orderCons

strFieldCons	=	" USER_COMPANY_API.GetGroupPlantationCode('" & Session("CodeName") & "',fcba) GroupPlantation" & _
					" , USER_COMPANY_API.GetPlantationCode('" & Session("CodeName") & "',fcba) Plantation" & _
					" , USER_COMPANY_API.GetGroupRegionCode('" & Session("CodeName") & "',fcba) groupregion" & _
					" , USER_COMPANY_API.GetRegionCode('" & Session("CodeName") & "',fcba) region" & _
					" , USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',fcba) company" & _
					" , BUSINESSUNIT_API.get_OrderData(fcba) NoUrut, fcba"
					
wherecons	=	" and businessunit_api.get_fctype(fcba) = 'M' AND USER_COMPANY_API.IsAuthorizedFcba('" & Session("CodeName") & "', fcba) =  'TRUE'"

If Len(Trim(v_txtGroupPlantation)) > 0 Then
	if v_txtGroupPlantation = "--SUMMARY--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
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
	Elseif v_txtPlantation = "--ALL--" then
		selectCons	=	selectCons	& " , Plantation"
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " , PLANTATION_API.get_OrderData(Plantation)"
		var_colspan	=	var_colspan	+ 1
	Else
		selectCons	=	selectCons	& " , Plantation"
		whereCons	=	whereCons	& " And Plantationcode = '" & v_txtPlantation & "'"
		orderCons	=	orderCons	& " , PLANTATION_API.get_OrderData(Plantation)"
		var_colspan	=	var_colspan	+ 1
	End If
End If


If Len(Trim(v_txtGroupRegion)) > 0 Then
	if v_txtGroupRegion = "--SUMMARY--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
	Elseif v_txtGroupRegion = "--ALL--" then
		selectCons	=	selectCons	& " , GroupRegion"
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " , GROUPREGION_API.get_OrderData(GroupRegion)"
		var_colspan	=	var_colspan	+ 1
	Else
		selectCons	=	selectCons	& " , GroupRegion"
		whereCons	=	whereCons	& " And GroupRegioncode = '" & v_txtGroupRegion & "'"
		orderCons	=	orderCons	& " , GROUPREGION_API.get_OrderData(GroupRegion)"
		var_colspan	=	var_colspan	+ 1
	End If
End If

If Len(Trim(v_txtRegion)) > 0 Then
	if v_txtRegion = "--SUMMARY--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
	Elseif v_txtRegion = "--ALL--" then
		selectCons	=	selectCons	& " , Region"
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " , REGION_API.get_OrderData(Region)"
		var_colspan	=	var_colspan	+ 1
	Else
		selectCons	=	selectCons	& " , Region"
		whereCons	=	whereCons	& " And Regioncode = '" & v_txtRegion & "'"
		orderCons	=	orderCons	& " , REGION_API.get_OrderData(Region)"
		var_colspan	=	var_colspan	+ 1
	End If
End If


If Len(Trim(v_txtCompany)) > 0 Then
	if v_txtCompany = "--SUMMARY--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
	Elseif v_txtCompany = "--ALL--" then
		selectCons	=	selectCons	& " , company"
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " , COMPANY_API.get_OrderData(company)"
		var_colspan	=	var_colspan	+ 1
	Else
		selectCons	=	selectCons	& " , company"
		whereCons	=	whereCons	& " And companycode = '" & v_txtCompany & "'"
		orderCons	=	orderCons	& " , COMPANY_API.get_OrderData(company)"
		var_colspan	=	var_colspan	+ 1
		
		q	=	" SELECT FCCODE" & _
				"        , FCCOMPANYCODE" & _
				"        , FCNAME" & _
				"        , FCCOMPANYNAME" & _
				"        , FCMANAGER" & _
				"        , FCKTU" & _
				"        , FCCROP" & _
				"        , FCTYPE" & _
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
	Elseif v_txtfcba = "--ALL--" then
		selectCons	=	selectCons	& " , nourut, fcba"
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " , nourut"
		var_colspan	=	var_colspan	+ 1
	Else
		selectCons	=	selectCons	& " , nourut, fcba"
		whereCons	=	whereCons	& " And fcba = '" & v_txtfcba & "'"
		orderCons	=	orderCons	& " , nourut"
		var_colspan	=	var_colspan	+ 1
		
		q	=	" SELECT FCCODE" & _
				"        , FCCOMPANYCODE" & _
				"        , FCNAME" & _
				"        , FCCOMPANYNAME" & _
				"        , FCMANAGER" & _
				"        , FCKTU" & _
				"        , FCCROP" & _
				"        , FCTYPE" & _
				" FROM BUSINESSUNIT WHERE" & _
				" FCCODE = '" & v_txtfcba & "'"	
		set rsCat = connStr.Execute(q)		
		if not rsCat.eof then
			fcbaname = rsCat("FCNAME")
			CompanyName = rsCat("FCCOMPANYNAME")
			MgrName = rsCat("FCMANAGER")
			KTUName = rsCat("FCKTU")
			CropName = rsCat("FCCROP")
			FcbaType = rsCat("FCTYPE")
		end if
		q	=	""
		rsCat.Close
		set rsCat = Nothing
	End If
End If

ScriptName = Request.ServerVariables("SCRIPT_NAME")

strMain = " "
		  
sqlLHPP2 = " "

sqlLHPP = " select case when region is not null and fcba is null then 'Total '||region" & _
          "             when region is null and fcba is null then 'Grand Total'" & _
          "             else region" & _
          "        end nextRegion," & _
          "        aa.*" & _
          " from (" & _
          "     select OrderDataRegion," & _
          "            region," & _
          "            fcba," & _
          "            capacity," & _
          "            SUM(ACT_THISMONTH) ACT_THISMONTH," & _
          "            SUM(BUD_THISMONTH) BUD_THISMONTH," & _
          "            SUM(ACT_TODATE) ACT_TODATE," & _
          "            SUM(BUD_TODATE) BUD_TODATE," & _
          "            SUM(BUD_ANNUAL) BUD_ANNUAL," & _
          "            SUM(THG_THISMONTH) THG_THISMONTH," & _
          "            SUM(THG_TODATE) THG_TODATE" & _
          "     from ( SELECT case when companycode = 'PT.THIP' then 'A'" & _
          "                                           when companycode = 'PT.CRS' then 'B'" & _
          "                                           else 'C'" & _
          "                                      end OrderDataRegion," & _
          "                                      case when companycode = 'PT.THIP' then 'THIP'" & _
          "                                           when companycode = 'PT.CRS' then 'CRS'" & _
          "                                           else 'AMS Ganda'" & _
          "                                      end region," & _
          "                                      fcba," & _
          "                                      nvl(capacity, 0) capacity," & _
          "                                      SUM(ACT_THISMONTH) ACT_THISMONTH," & _
          "                                      SUM(BUD_THISMONTH) BUD_THISMONTH," & _
          "                                      SUM(ACT_TODATE) ACT_TODATE," & _
          "                                      SUM(BUD_TODATE) BUD_TODATE," & _
          "                                      SUM(BUD_ANNUAL) BUD_ANNUAL," & _
          "                                      SUM(THG_THISMONTH) THG_THISMONTH," & _
          "                                      SUM(THG_TODATE) THG_TODATE" & _
          "                               FROM mgt2_ffbprocess_det_tab" & _
          "                               WHERE year = " & v_txtYear & "" & _
          "                                     and month = " & v_txtMonth & "" & _
          "                                     " & whereCons & "" & _
          "                               GROUP BY companycode," & _
          "                                        fcba," & _
          "                                        capacity" & _
          "     )" & _
          "     group by grouping sets( (OrderDataRegion," & _
          "                              region," & _
          "                              fcba," & _
          "                              capacity)," & _
          "                             (OrderDataRegion," & _
          "                              region)," & _
          "                             () )" & _
          " ) aa" & _
		  " order by orderdataregion, BUSINESSUNIT_API.get_OrderData (fcba)"
		  
'response.write sqlLHPP
%>


<html>

	<head>
		<title>ePlaS Reporting - Laporan Bulanan TBS Proses</title>
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
            $( "#datepicker" ).datepicker();
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
				@media print{.printhide{display:none;}}
			</style>
       </head>


	<body>
		<div id="wrapper">
		<div class="printhide">
		<!--#include file="../print_excel.asp"-->
		</div>
			<form name="frmEntry" method="post" action="<%=ScriptName%>">
				<div class="printhide">
				<!--#include file="menutop.asp"-->
				<!--#include file="FormSelectNewSumM.html"-->
				
					<table id="form-table" class="centered-margin">
						<tr>
							<td><font face="verdana" size="1">BUDGET YEAR</font></td>
							<td><font face="verdana" size="1">BUDGET VERSION</font></td>
							<td><font face="verdana" size="1">MODULE VERSION</font></td>
							<td><font face="verdana" size="1">YEAR</font></td>
							<td><font face="verdana" size="1">MONTH</font></td>
							<td>&nbsp;</td>
						</tr>
						<tr>
							<td align="center">
								<input type="text" name="txtBudgetYear" value="<%=v_txtBudgetYear%>">
							</td>
							<td align="center">
								<input type="text" name="txtBudgetVersion" value="<%=v_txtBudgetVersion%>">
							</td>
							<td align="center">
								<input type="text" name="txtModuleVersion" value="<%=v_txtModuleVersion%>">
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
							<td  align="center" width="4">
								<select name="txtYear">
									<%For i = MinYr to MaxYr%>
									<Option Value="<%=i%>" <%If cint(i) = cint(v_txtYear) then%> Selected <%End If%>><%=i%></option>
									<%Next%>
								</select>
							</td>
							<td width="2">
								<select name="txtMonth">
									<%For i = 1 to 12%>
									<Option Value="<%=i%>" <%If cint(i) = cint(v_txtMonth) then%> Selected <%End If%>><%=Bulan(i)%></option>
									<%Next%>
								</select>
							</td>
							<td align="center">
								<input type="submit" value="View" />
							</td>							
						</tr>
					</table>
				</div>
			</form>

			<table border="1" width="100%" id="product-table" class='xthf-site'>
				<thead>
					<tr>
						<th colspan="11" style="background:#ffffff; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
							<div id="header-wrapper">
								<div id="subheader" class="centered">
									<h1>AMS PLANTATION</h1>
									<div style="margin-botton:10px;"/>
									<h3>LAPORAN BULANAN TBS PROSES</h3>
									<br/>
								</div>
								<div style="text-align:left;">
									<h4>Periode : <%=Bulan(v_txtMonth)%>&nbsp;<%=v_txtYear%></h4>
								</div>
							</div>
						</th>
					</tr>
					<tr>
						<th rowspan="3">NO</th>
						<th rowspan="3">MILL</th>
						<th rowspan="3">Cap.</th>
						<th colspan="6">FFB PROCESS (Kg)</th>
						<th colspan="2">THROUGHPUT (Ton/Jam)</th>
					</tr>
					<tr>
						<th colspan="3">THIS MONTH</th>
						<th colspan="3">TO DATE</th>
						<th rowspan="2">THIS MONTH</th>
						<th rowspan="2">TO DATE</th>
					</tr>
					<tr>
						<th>ACT</th>
						<th>BGT</th>
						<th>%</th>
						<th>ACT</th>
						<th>BGT</th>
						<th>%</th>
					</tr>
				</thead>
				
				<tbody>
<%
Set rsLHPP = Server.CreateObject("ADODB.Recordset")
rsLHPP.Open sqlLHPP, connstr, adOpenStatic

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

Dim CurrOrderDataRegion : CurrOrderDataRegion = ""
Dim PrevOrderDataRegion : PrevOrderDataRegion = ""

Dim currnextregion : currnextregion = ""
Dim prevnextregion : prevnextregion = ""

Dim Row_percent_thismonth : Row_percent_thismonth = 0
Dim Row_percent_todate : Row_percent_todate = 0


if rsLHPP.BOF and rsLHPP.EOF then
%>
					<tr>
						<td colspan="<%=11%>">NO DATA AVAILABLE</td>
					</tr>
<%
else
		while not rsLHPP.eof
			CurrOrderDataRegion = rsLHPP("orderdataregion")
%>
				<%if rsLHPP("NEXTREGION") = "Grand Total" then%>
					<tr bgcolor="#ffff01" >
						<td align="left" colspan="3"><strong><%=rsLHPP("nextregion")%></strong></td>
						<td align="right"><strong><%=FormatNumber(rsLHPP("ACT_THISMONTH"), 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(rsLHPP("BUD_THISMONTH"), 0)%></strong></td>
						<%
						if cdbl(rsLHPP("BUD_THISMONTH")) = 0 then
							if cdbl(rsLHPP("ACT_THISMONTH")) = 0 then
								Row_percent_thismonth = 0
							else
								Row_percent_thismonth = 100
							end if
						else
							Row_percent_thismonth = Round( ( (cdbl(rsLHPP("ACT_THISMONTH")) - cdbl(rsLHPP("BUD_THISMONTH")))/cdbl(rsLHPP("BUD_THISMONTH")) )* 100, 2)
						end if
						%>
						<td align="right"><strong><%=FormatNumber(Row_percent_thismonth, 2)%></strong></td>
						<td align="right"><strong><%=FormatNumber(rsLHPP("ACT_TODATE"), 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(rsLHPP("BUD_TODATE"), 0)%></strong></td>
						<%
						if cdbl(rsLHPP("BUD_TODATE")) = 0 then
							if cdbl(rsLHPP("ACT_TODATE")) = 0 then
								Row_percent_todate = 0
							else
								Row_percent_todate = 100
							end if
						else
							Row_percent_todate = Round( ( (cdbl(rsLHPP("ACT_TODATE")) - cdbl(rsLHPP("BUD_TODATE")))/cdbl(rsLHPP("BUD_TODATE")) )* 100, 2)
						end if
						%>
						<td align="right"><strong><%=FormatNumber(Row_percent_todate, 2)%></strong></td>
						<td align="right"><strong><%=FormatNumber(Cdbl(rsLHPP("THG_THISMONTH"))/1000, 2)%></strong></td>
						<td align="right"><strong><%=FormatNumber(Cdbl(rsLHPP("THG_TODATE"))/1000, 2)%></strong></td>
					</tr>
				<%else%>
					<%if CurrOrderDataRegion <> PrevOrderDataRegion then%>
					<tr>
						<td align="Center" style=" border-bottom:none;"><strong><%=rsLHPP("orderdataregion")%></strong></td>
						<td align="left"><strong><%=rsLHPP("region")%></strong></td>
						<td align="right">&nbsp;</td>
						<td align="right">&nbsp;</td>
						<td align="right">&nbsp;</td>
						<td align="right">&nbsp;</td>
						<td align="right">&nbsp;</td>
						<td align="right">&nbsp;</td>
						<td align="right">&nbsp;</td>
						<td align="right">&nbsp;</td>
						<td align="right">&nbsp;</td>
					</tr>
					<tr>
						<td align="center" style="border:none;border-right:1px #3A5F0B solid; border-left:1px #3A5F0B solid;"></td>
						<td align="left">&nbsp;&nbsp;&nbsp;&nbsp;<%=rsLHPP("FCBA")%></td>
						<td align="right"><%=FormatNumber(rsLHPP("capacity"), 0)%></td>
						<td align="right"><%=FormatNumber(rsLHPP("ACT_THISMONTH"), 0)%></td>
						<td align="right"><%=FormatNumber(rsLHPP("BUD_THISMONTH"), 0)%></td>
						<%
						if cdbl(rsLHPP("BUD_THISMONTH")) = 0 then
							if cdbl(rsLHPP("ACT_THISMONTH")) = 0 then
								Row_percent_thismonth = 0
							else
								Row_percent_thismonth = 100
							end if
						else
							Row_percent_thismonth = Round( ( (cdbl(rsLHPP("ACT_THISMONTH")) - cdbl(rsLHPP("BUD_THISMONTH")))/cdbl(rsLHPP("BUD_THISMONTH")) )* 100, 2)
						end if
						%>
						<td align="right"><%=FormatNumber(Row_percent_thismonth, 2)%></td>
						<td align="right"><%=FormatNumber(rsLHPP("ACT_TODATE"), 0)%></td>
						<td align="right"><%=FormatNumber(rsLHPP("BUD_TODATE"), 0)%></td>
						<%
						if cdbl(rsLHPP("BUD_TODATE")) = 0 then
							if cdbl(rsLHPP("ACT_TODATE")) = 0 then
								Row_percent_todate = 0
							else
								Row_percent_todate = 100
							end if
						else
							Row_percent_todate = Round( ( (cdbl(rsLHPP("ACT_TODATE")) - cdbl(rsLHPP("BUD_TODATE")))/cdbl(rsLHPP("BUD_TODATE")) )* 100, 2)
						end if
						%>
						<td align="right"><%=FormatNumber(Row_percent_todate, 2)%></td>
						<td align="right"><%=FormatNumber(Cdbl(rsLHPP("THG_THISMONTH"))/1000, 2)%></td>
						<td align="right"><%=FormatNumber(Cdbl(rsLHPP("THG_TODATE"))/1000, 2)%></td>
					</tr>
						<%PrevOrderDataRegion = CurrOrderDataRegion%>
					<%else%>
						<%if Left(Cstr(rsLHPP("NEXTREGION")), 5) = "Total" Then%>
							<tr bgcolor="#F4F4A0" >
								<td bgcolor="#ffffff" align="center" style="border:none;border-right:1px #3A5F0B solid; border-left:1px #3A5F0B solid;"></td>
								<td align="left"><strong><%=rsLHPP("nextregion")%></strong></td>
								<td align="right">&nbsp;</td>
								<td align="right"><strong><%=FormatNumber(rsLHPP("ACT_THISMONTH"), 0)%></strong></td>
								<td align="right"><strong><%=FormatNumber(rsLHPP("BUD_THISMONTH"), 0)%></strong></td>
								<%
								if cdbl(rsLHPP("BUD_THISMONTH")) = 0 then
									if cdbl(rsLHPP("ACT_THISMONTH")) = 0 then
										Row_percent_thismonth = 0
									else
										Row_percent_thismonth = 100
									end if
								else
									Row_percent_thismonth = Round( ( (cdbl(rsLHPP("ACT_THISMONTH")) - cdbl(rsLHPP("BUD_THISMONTH")))/cdbl(rsLHPP("BUD_THISMONTH")) )* 100, 2)
								end if
								%>
								<td align="right"><strong><%=FormatNumber(Row_percent_thismonth, 2)%></strong></td>
								<td align="right"><strong><%=FormatNumber(rsLHPP("ACT_TODATE"), 0)%></strong></td>
								<td align="right"><strong><%=FormatNumber(rsLHPP("BUD_TODATE"), 0)%></strong></td>
								<%
								if cdbl(rsLHPP("BUD_TODATE")) = 0 then
									if cdbl(rsLHPP("ACT_TODATE")) = 0 then
										Row_percent_todate = 0
									else
										Row_percent_todate = 100
									end if
								else
									Row_percent_todate = Round( ( (cdbl(rsLHPP("ACT_TODATE")) - cdbl(rsLHPP("BUD_TODATE")))/cdbl(rsLHPP("BUD_TODATE")) )* 100, 2)
								end if
								%>
								<td align="right"><strong><%=FormatNumber(Row_percent_todate, 2)%></strong></td>
								<td align="right"><strong><%=FormatNumber(Cdbl(rsLHPP("THG_THISMONTH"))/1000, 2)%></strong></td>
								<td align="right"><strong><%=FormatNumber(Cdbl(rsLHPP("THG_TODATE"))/1000, 2)%></strong></td>
							</tr>
						<%else%>
							<tr>
								<td align="center" style="border:none;border-right:1px #3A5F0B solid; border-left:1px #3A5F0B solid;"></td>
								<td align="left">&nbsp;&nbsp;&nbsp;&nbsp;<%=rsLHPP("FCBA")%></td>
								<td align="right"><%=FormatNumber(rsLHPP("capacity"), 0)%></td>
								<td align="right"><%=FormatNumber(rsLHPP("ACT_THISMONTH"), 0)%></td>
								<td align="right"><%=FormatNumber(rsLHPP("BUD_THISMONTH"), 0)%></td>
								<%
								if cdbl(rsLHPP("BUD_THISMONTH")) = 0 then
									if cdbl(rsLHPP("ACT_THISMONTH")) = 0 then
										Row_percent_thismonth = 0
									else
										Row_percent_thismonth = 100
									end if
								else
									Row_percent_thismonth = Round( ( (cdbl(rsLHPP("ACT_THISMONTH")) - cdbl(rsLHPP("BUD_THISMONTH")))/cdbl(rsLHPP("BUD_THISMONTH")) )* 100, 2)
								end if
								%>
								<td align="right"><%=FormatNumber(Row_percent_thismonth, 2)%></td>
								<td align="right"><%=FormatNumber(rsLHPP("ACT_TODATE"), 0)%></td>
								<td align="right"><%=FormatNumber(rsLHPP("BUD_TODATE"), 0)%></td>
								<%
								if cdbl(rsLHPP("BUD_TODATE")) = 0 then
									if cdbl(rsLHPP("ACT_TODATE")) = 0 then
										Row_percent_todate = 0
									else
										Row_percent_todate = 100
									end if
								else
									Row_percent_todate = Round( ( (cdbl(rsLHPP("ACT_TODATE")) - cdbl(rsLHPP("BUD_TODATE")))/cdbl(rsLHPP("BUD_TODATE")) )* 100, 2)
								end if
								%>
								<td align="right"><%=FormatNumber(Row_percent_todate, 2)%></td>
								<td align="right"><%=FormatNumber(Cdbl(rsLHPP("THG_THISMONTH"))/1000, 2)%></td>
								<td align="right"><%=FormatNumber(Cdbl(rsLHPP("THG_TODATE"))/1000, 2)%></td>
							</tr>
						<%end if%>
					<%end if%>
				<%end if%>

<%				
				rsLHPP.MoveNext
			Wend
%>

<%
End if
rsLHPP.Close
set rsLHPP = Nothing
%>

				</tbody>
			</table>

		</div>
	</body>
</html>
