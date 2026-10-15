<%@Language=VBScript%>
<%Option Explicit%>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
	Dim v_txtBudgetYear, v_txtBudgetVersion, v_txtModuleVersion, v_txtdatefrom, v_txtMonth, v_txtYear, v_txtgrade, v_txtDivisi, v_txtemployeetype, v_txtjobcode
	Dim no
	Dim rsDefaultBa, strDefaultBa, varDefaultBa
	Dim ScriptName
	Dim strMain, sqlLHPP, sqlLHPP2, rsLHPP
	Dim q, CompanyName, fcbaname, MgrName, KTUName, CropName, FcbaType
	Dim OrderSection : OrderSection = 0
	Dim TotalMandays : TotalMandays = 0

    v_txtGroupPlantation = Request("txtGroupPlantation")
	v_txtPlantation = Request("txtPlantation")
	v_txtGroupRegion = Request("txtGroupRegion")
	v_txtRegion = Request("txtRegion")
	v_txtcompany = Request("txtCompany")
	v_txtfcba = Request("txtfcba")
	
	v_txtDivisi = Request("txtDivisi")
	v_txtjobcode = Request("txtjobcode")
	v_txtemployeetype = Request("txtemployeetype")
	
	v_txtBudgetYear = Request("txtBudgetYear")
	v_txtBudgetVersion = Request("txtBudgetVersion")
	v_txtModuleVersion = Request("txtModuleVersion")
	
	v_txtdatefrom		= Request("txtdatefrom")
	
	v_txtYear		= Request("txtYear")
	v_txtMonth		= Request("txtMonth")
	v_txtgrade		= Request("txtgrade")
	
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
	
	
	strDefaultBa = " select" & _
				   "     fccode as fcba" & _
				   " from businessunit bu" & _
				   " inner join user_fcba_tab uft on (bu.fccode = uft.fcba) " & _
				   " where isdefault = 'TRUE'" & _
				   "       and fcusercode = '" & Session("CodeName") & "'"

	Set rsDefaultBa = Server.CreateObject("ADODB.Recordset")
	rsDefaultBa.Open strDefaultBa, connstr, adOpenStatic

	If Not rsDefaultBa.EOF Then
		varDefaultBa = rsDefaultBa("fcba")
	Else
		varDefaultBa = Session("fcba")
	End If
	rsDefaultBa.Close
	set rsDefaultBa = Nothing

	if not isnull(v_txtFcba) then
		If (v_txtFcba = "" Or Len(Trim(v_txtFcba)) = 0) Then
			v_txtFcba = varDefaultBa
		End If
	End If
	
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

Dim array_data, array_eof, array_i

array_eof = true

strMain = " "
		  
sqlLHPP2 = " "

sqlLHPP = " select fcba" & _
                 "        , fddate" & _
                 "        , gang_ag" & _
                 "        , sectionname" & _
                 "        , employee_type" & _
                 "        , employeecode" & _
                 "        , employee_api.get_employeename(fcba, employeecode) employeename" & _
                 "        , grade" & _
                 "        , jobcode" & _
                 "        , job_api.get_jobcode_name(jobcode, fcba) jobname" & _
                 "        , locationtype" & _
                 "        , locationcode" & _
                 "        , attendance" & _
                 "        , mandays" & _
                 " from rpt_hke_dtl_tab" & _
                 " where fcba = '"& v_txtFcba & "'" & _
                 "       and to_char(fddate, 'yyyymm') = '" & v_txtYear & lpad(v_txtMonth, "0", "2") & "'"

if len(trim(v_txtDivisi)) <> 0 then
		sqlLHPP = sqlLHPP & " and sectionname = '" & v_txtDivisi & "'"
end if

if len(trim(v_txtgrade)) <> 0 then
		sqlLHPP = sqlLHPP & " and grade = '" & v_txtgrade & "'"
end if

if len(trim(v_txtemployeetype)) <> 0 then
		sqlLHPP = sqlLHPP & " and employee_type = '" & v_txtemployeetype & "'"
end if

if len(trim(v_txtjobcode)) <> 0 then
		sqlLHPP = sqlLHPP & " and jobcode = '" & v_txtjobcode & "'"
end if
sqlLHPP = sqlLHPP & " order by fddate, employeecode"
		  
'response.write sqlLHPP
%>


<html>

	<head>
		<title>ePlaS Reporting - Laporan Detail Hari Kerja Efektif</title>
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
				<!--#include virtual="/REPORT/FormSelectFcbaOnly.html"-->
				
					<table id="form-table" class="centered-margin">
						<tr>
							<td><font face="verdana" size="1">YEAR</font></td>
							<td><font face="verdana" size="1">MONTH</font></td>
							<td><font face="verdana" size="1">SECTION</font></td>
							<td><font face="verdana" size="1">STATUS</font></td>
							<td><font face="verdana" size="1">EMPLOYEE TYPE</font></td>
							<td><font face="verdana" size="1">JOBCODE</font></td>
							<td>&nbsp;</td>
						</tr>
						<tr>
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
							<td>
								<select name="txtDivisi" >
								<Option Value="" <%If v_txtDivisi = "" then%> Selected <%End If%>>--ALL--</option>
							<%
							Dim rsDivisi, strDivisi
							strDivisi	=	" select fccode, fcname from parameter where fcba = '" & v_txtFcba & "' and parcode = 'E0002' order by sectionname_api.get_orderdata(fccode)"

							Set rsDivisi = Server.CreateObject("ADODB.Recordset")
							rsDivisi.Open strDivisi, connstr, adOpenStatic 

							if not rsDivisi.eof then
								Do While Not rsDivisi.eof
									if rsDivisi("fccode") = v_txtDivisi then%>
										<option value="<%=rsDivisi("fccode")%>" Selected><%=rsDivisi("fccode")%> - <%=rsDivisi("fcname")%></option> 
									<%Else%>
										<option value="<%=rsDivisi("fccode")%>"><%=rsDivisi("fccode")%> - <%=rsDivisi("fcname")%></option>  
									<%End if
									rsDivisi.movenext
								loop
							end if
							rsDivisi.Close
							set rsDivisi = Nothing
							%>
							</select>
							</td>
							<td>
								<select name="txtgrade">
									<Option Value="" <%If v_txtgrade = "" then%> Selected <%End If%>>--ALL--</option>
									<Option Value="KBT" <%If v_txtgrade = "KBT" then%> Selected <%End If%>>KBT : Karyawan Bulanan Tetap</option>
									<Option Value="KHT" <%If v_txtgrade = "KHT" then%> Selected <%End If%>>KHT : Karyawan Harian Tetap</option>
									<Option Value="KHL" <%If v_txtgrade = "KHL" then%> Selected <%End If%>>KHL : Karyawan Harian Lepas</option>
								</select>
							</td>
							
							<td>
								<select name="txtemployeetype" >
								<Option Value="" <%If v_txtemployeetype = "" then%> Selected <%End If%>>--ALL--</option>
							<%
							Dim rsemployeetype, stremployeetype
							stremployeetype	=	" select fccode, fcname from employee_type order by orderdata"

							Set rsemployeetype = Server.CreateObject("ADODB.Recordset")
							rsemployeetype.Open stremployeetype, connstr, adOpenStatic 

							if not rsemployeetype.eof then
							Do While Not rsemployeetype.eof
								if rsemployeetype("fccode") = v_txtemployeetype then%>
									<option value="<%=rsemployeetype("fccode")%>" Selected><%=rsemployeetype("fccode")%> - <%=rsemployeetype("fcname")%></option> 
								<%Else%>
									<option value="<%=rsemployeetype("fccode")%>"><%=rsemployeetype("fccode")%> - <%=rsemployeetype("fcname")%></option>  
								<%End if
								rsemployeetype.movenext
								loop
								end if
							rsemployeetype.Close
							set rsemployeetype = Nothing
							%>
							</select>
							</td>
							<td>
								<select name="txtjobcode" >
								<Option Value="" <%If v_txtjobcode = "" then%> Selected <%End If%>>--ALL--</option>
							<%
							Dim rsjobcode, strjobcode
							strjobcode	=	" select fccode, fcname from job where fcba = '" & v_txtFcba & "' order by fccode"

							Set rsjobcode = Server.CreateObject("ADODB.Recordset")
							rsjobcode.Open strjobcode, connstr, adOpenStatic 

							if not rsjobcode.eof then
								Do While Not rsjobcode.eof
									if rsjobcode("fccode") = v_txtjobcode then%>
										<option value="<%=rsjobcode("fccode")%>" Selected><%=rsjobcode("fccode")%> - <%=rsjobcode("fcname")%></option> 
									<%Else%>
										<option value="<%=rsjobcode("fccode")%>"><%=rsjobcode("fccode")%> - <%=rsjobcode("fcname")%></option>  
									<%End if
									rsjobcode.movenext
								loop
							end if
							rsjobcode.Close
							set rsjobcode = Nothing
							%>
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
						<th colspan="13" style="background:#ffffff; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
							<div id="header-wrapper">
								<div id="subheader" class="centered">
									<h1>LAPORAN UPR DETAIL</h1>
									<div style="margin-botton:10px;"/>
									<h3>LAPORAN DETAIL HARI KERJA EFEKTIF (HKE)</h3>
									<br/>
								</div>
								<div style="text-align:left;">
									<h4>Periode : <%=Bulan(v_txtMonth)%>&nbsp;<%=v_txtYear%></h4>
								</div>
							</div>
						</th>
					</tr>
					<tr>
						<th>Date</th>
						<th>Gang Code</th>
						<th>Section</th>
						<th>Employee Type</th>
						<th>Employee Code</th>
						<th>Employee Name</th>
						<th>Status</th>
						<th>Job Code</th>
						<th>Job Desc</th>
						<th>Target Type</th>
						<th>Target Code</th>
						<th>Attendance</th>
						<th>Mandays</th>
					</tr>
				</thead>
				
				<tbody>
<%
Set rsLHPP = Server.CreateObject("ADODB.Recordset")
rsLHPP.Open sqlLHPP, connstr, adOpenStatic

if not rsLHPP.EOF then
	array_data = rsLHPP.getrows()
	array_eof = false
end if
rsLHPP.Close
set rsLHPP = Nothing
%>
<%if array_eof = true then%>
					<tr>
						<td colspan="<%=13%>">NO DATA AVAILABLE</td>
					</tr>
<%end if%>
<%if not array_eof then%>
	<%TotalMandays = 0%>
	<%For array_i = LBound(array_data, 2) To UBound(array_data, 2)%>
					<tr>
						<td align="left"><%=array_data(1, array_i)%></td>
						<td align="left"><%=array_data(2, array_i)%></td>
						<td align="left"><%=array_data(3, array_i)%></td>
						<td align="left"><%=array_data(4, array_i)%></td>
						<td align="left"><%=array_data(5, array_i)%></td>
						<td align="left"><%=array_data(6, array_i)%></td>
						<td align="left"><%=array_data(7, array_i)%></td>
						<td align="left"><%=array_data(8, array_i)%></td>
						<td align="left"><%=array_data(9, array_i)%></td>
						<td align="left"><%=array_data(10, array_i)%></td>
						<td align="left"><%=array_data(11, array_i)%></td>
						<td align="left"><%=array_data(12, array_i)%></td>
						<td align="right"><%=FbpNum2(array_data(13, array_i))%></td>
					</tr>
					<%TotalMandays = TotalMandays + cdbl(array_data(13, array_i))%>
	<%next%>
<%end if%>

					<tr bgcolor="#f4f4a0">
						<td align="center" colspan="12"><strong>TOTAL</strong></td>
						<td align="right"><strong><%=FbpNum2(TotalMandays)%></strong></td>
					</tr>
				</tbody>
			</table>

		</div>
	</body>
</html>
