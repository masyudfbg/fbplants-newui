<%@Language=VBScript%>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%

Function rowspan_gt(data_array__)
	result_ = 0
	arr_1 = data_array__
	for indx = LBound(arr_1, 2) to UBound(arr_1, 2)
		result_ = result_ +1
	next
	rowspan_gt = result_
End Function
	Dim v_txtBudgetYear, v_txtBudgetVersion, v_txtModuleVersion, v_txtdatefrom, v_txtMonth, v_txtYear, v_txtgrade
	Dim no
	Dim rsDefaultBa, strDefaultBa, varDefaultBa
	Dim ScriptName
	Dim strMain, sqlLHPP, sqlLHPP2, rsLHPP
	Dim q, CompanyName, fcbaname, MgrName, KTUName, CropName, FcbaType
	Dim OrderSection : OrderSection = 0

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
	v_txtgrade		= Request("txtgrade")
	v_txtketerangan = Request("txtketerangan")
	
	If Len(Trim(v_txtYear)) = 0 Then v_txtYear = Year(textDate(Date))
	If Len(Trim(v_txtMonth)) = 0 Then v_txtMonth = Month(textDate(Date))
	
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

	If (v_txtFcba = "" Or Len(Trim(v_txtFcba)) = 0) Then
		v_txtFcba = varDefaultBa
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

strMain = 	"    SELECT sectionname section,presentgrade grade,cntemp, " & _
            "    bi_kj,bi_dl,bi_wh,bi_ws,bi_kk,bi_kl,bi_kh," & _
            "    bi_p2,bi_p3,bi_s1,bi_s2,bi_h1,bi_h2,bi_ct," & _
            "    bi_mk,bi_ml,bi_p1,bi_un " & _
			" 	 FROM " & _
            "    (  " & _
			"	  SELECT * " & _
            "     FROM ( SELECT em.sectionname, em.presentgrade, attendance,Count(DISTINCT employeecode) OVER (PARTITION BY em.sectionname, em.presentgrade) cntemp" & _
            "           FROM " & _
            "           ( SELECT DISTINCT ag.fcba, hdr.bkmdate, st.YEAR,st.MONTH, ag.employeecode, ag.attendance" & _
            "                  FROM bkmupahdtl ag " & _
            "                  INNER JOIN BKMhdr hdr ON (hdr.bkmno=ag.bkmno AND hdr.fcba=ag.fcba)" & _
            "                  INNER JOIN empalldedstatus st ON (hdr.fcba=st.fcba AND hdr.bkmdate BETWEEN start_date AND end_date)" & _
            "                  WHERE hdr.postingflg='1' AND ag.fcba = '" & v_txtfcba & "'" & _
            "                  AND st.year =  " & v_txtYear & "" & _
            "                  AND st.month = " & v_txtMonth & "" & _
            "                  AND InStr('KJ,DL,WH,WS,KK,KL,KH,P2,P3,S1,S2,H1,H2,CT,MK,ML,P1,UN', ag.attendance) <> 0" & _
            "           ) ag" & _
            "           INNER JOIN employee_hist_tab em ON (ag.fcba = em.fcba AND ag.employeecode = em.fccode  AND ag.year = em.year AND ag.month = em.month)" & _
            "          ) PIVOT " & _
            "          ( COUNT(attendance)" & _
            "               FOR attendance IN " & _
            "               (" & _
            "               'KJ' bi_kj,'DL' bi_dl,'WH' bi_wh,'WS' bi_ws,'KK' bi_kk,'KL' bi_kl,'KH' bi_kh," & _
            "               'P2' bi_p2,'P3' bi_p3,'S1' bi_s1,'S2' bi_s2,'H1' bi_h1,'H2' bi_h2,'CT' bi_ct," & _
            "               'MK' bi_mk,'ML' bi_ml,'P1' bi_p1,'UN' bi_un" & _
            "               )" & _
            "          )" & _
            " )"
			
            
sqlLHPP2 = " "

sqlLHPP = " select case when section is not null and grade is null then 'Subtotal '||initcap(sectionname_api.get_fcname(section))" & _
          "             when section is null and grade is null then 'Grand Total'" & _
          "             else sectionname_api.get_fcname(section)" & _
          "        end section" & _
          "        , sectionname_api.get_orderdata(section) orderdatasection" & _
          "        , grade" & _
          "        , employeegrade_api.get_orderdata(grade) orderdatagrade" & _
		  "        , NVL(SUM(cntemp), 0) cntemp" & _
          "        , NVL(SUM(bi_kj), 0) bi_kj" & _
		  "        , NVL(SUM(bi_dl), 0) bi_dl" & _
		  "        , NVL(SUM(bi_wh), 0) bi_wh" & _
		  "        , NVL(SUM(bi_ws), 0) bi_ws" & _
		  "        , NVL(SUM(bi_kk), 0) bi_kk" & _
		  "        , NVL(SUM(bi_kl), 0) bi_kl" & _
		  "        , NVL(SUM(bi_kh), 0) bi_kh" & _
		  "        , NVL(SUM(bi_p2), 0) bi_p2" & _
          "        , NVL(SUM(bi_p3), 0) bi_p3" & _
          "        , NVL(SUM(bi_s1), 0) bi_s1" & _
          "        , NVL(SUM(bi_s2), 0) bi_s2" & _
          "        , NVL(SUM(bi_h1), 0) bi_h1" & _
          "        , NVL(SUM(bi_h2), 0) bi_h2" & _
          "        , NVL(SUM(bi_ct), 0) bi_ct" & _
          "        , NVL(SUM(bi_mk), 0) bi_mk" & _
          "        , NVL(SUM(bi_ml), 0) bi_ml" & _
          "        , NVL(SUM(bi_p1), 0) bi_p1" & _
          "        , NVL(SUM(bi_un), 0) bi_un" & _
          " from (" & strMain & ") " 
		 
if len(trim(v_txtgrade)) <> 0 then
	if v_txtgrade = "TETAP" then
		sqlLHPP = sqlLHPP & "       and (grade = 'KHT' or grade = 'KBT')"
	else
		sqlLHPP = sqlLHPP & "       and grade = '" & v_txtgrade & "'"
	end if
end if

if "" & v_txtketerangan <> "" then
	sqlLHPP = sqlLHPP & "       and employee_type = '" & v_txtketerangan & "'"
end if
sqlLHPP = sqlLHPP & " group by grouping sets ( (section" & _
                    "                           , sectionname_api.get_orderdata(section)" & _
                    "                           , grade,cntemp" & _
                    "                           , employeegrade_api.get_orderdata(grade))," & _
                    "                          (section" & _
                    "                           , sectionname_api.get_orderdata(section))," & _
                    "                          ()" & _
                    " )" & _
                    " order by orderdatasection" & _
                    "          , orderdatagrade"
		  
'response.write sqlLHPP

gt_unit = " SELECT grade, orderdatagrade,NVL(SUM(cntemp), 0) cntemp" & _
          "        , NVL(SUM(bi_kj), 0) bi_kj" & _
		  "        , NVL(SUM(bi_dl), 0) bi_dl" & _
		  "        , NVL(SUM(bi_wh), 0) bi_wh" & _
		  "        , NVL(SUM(bi_ws), 0) bi_ws" & _
		  "        , NVL(SUM(bi_kk), 0) bi_kk" & _
		  "        , NVL(SUM(bi_kl), 0) bi_kl" & _
		  "        , NVL(SUM(bi_kh), 0) bi_kh" & _
		  "        , NVL(SUM(bi_p2), 0) bi_p2" & _
          "        , NVL(SUM(bi_p3), 0) bi_p3" & _
          "        , NVL(SUM(bi_s1), 0) bi_s1" & _
          "        , NVL(SUM(bi_s2), 0) bi_s2" & _
          "        , NVL(SUM(bi_h1), 0) bi_h1" & _
          "        , NVL(SUM(bi_h2), 0) bi_h2" & _
          "        , NVL(SUM(bi_ct), 0) bi_ct" & _
          "        , NVL(SUM(bi_mk), 0) bi_mk" & _
          "        , NVL(SUM(bi_ml), 0) bi_ml" & _
          "        , NVL(SUM(bi_p1), 0) bi_p1" & _
          "        , NVL(SUM(bi_un), 0) bi_un" & _
          " FROM ("&sqlLHPP&") WHERE grade IS NOT NULL" & _
          " GROUP BY grade, orderdatagrade" & _
          " ORDER BY orderdatagrade"
'response.write gt_unit

Set rsgt_unit = Server.CreateObject("ADODB.Recordset")
rsgt_unit.Open gt_unit, connstr, adOpenStatic

arr_unit_eof = true
if not rsgt_unit.eof then
	arr_unit = rsgt_unit.getrows()
	arr_unit_eof = false
end if


post_desc_total = ""
if FcbaType = "E" then
	post_desc_total = "Kebun"
elseif FcbaType = "M" or FcbaType = "P" then
	post_desc_total = "Pabrik"
else
	post_desc_total = v_txtFcba
end if
%>


<html>

	<head>
		<title>ePlaS Reporting - Laporan Hari Kerja Efektif</title>
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
							<td><font face="verdana" size="1">BUDGET YEAR</font></td>
							<td><font face="verdana" size="1">BUDGET VERSION</font></td>
							<td><font face="verdana" size="1">MODULE VERSION</font></td>
							<td><font face="verdana" size="1">YEAR</font></td>
							<td><font face="verdana" size="1">MONTH</font></td>
							<td><font face="verdana" size="1">STATUS</font></td>
							<td><font face="verdana" size="1">Langsung/Tidak Langsung</font></td>
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
							MinYr = 2019
							MaxYr = Yr  + 3
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
								<select name="txtgrade">
									<Option Value="" <%If v_txtgrade = "" then%> Selected <%End If%>>--ALL--</option>
									<Option Value="TETAP" <%If v_txtgrade = "TETAP" then%> Selected <%End If%>>KHT + KBT: Karyawan Harian/Bulanan Tetap</option>
									<Option Value="KBT" <%If v_txtgrade = "KBT" then%> Selected <%End If%>>KBT : Karyawan Bulanan Tetap</option>
									<Option Value="KHT" <%If v_txtgrade = "KHT" then%> Selected <%End If%>>KHT : Karyawan Harian Tetap</option>
									<Option Value="KHL" <%If v_txtgrade = "KHL" then%> Selected <%End If%>>KHL : Karyawan Harian Lepas</option>
								</select>
							</td>
							
							<td>
								<select name="txtketerangan">
									<Option Value="" <%If "" & v_txtketerangan = "" then%> Selected <%End If%>>--ALL--</option>
									<Option Value="KL" <%If v_txtketerangan = "KL" then%> Selected <%End If%>>Karyawan Langsung</option>
									<Option Value="KTL" <%If v_txtketerangan = "KTL" then%> Selected <%End If%>>Karyawan Tidak Langsung</option>
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
						<th colspan="37" style="background:#ffffff; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
							<div id="header-wrapper">
								<div id="subheader" class="centered">
									<h1>LAPORAN HARI KERJA EFEKTIF (HKE)</h1>
									<div style="margin-botton:10px;"/> 
									<br/>
								</div>
								<div style="text-align:left;">
									<h4>Periode : <%=Bulan(v_txtMonth)%>&nbsp;<%=v_txtYear%></h4>
									<h4 style="margin-top:5px;"><%=PCase(fcbaname)%></h4>
								</div>
							</div>
						</th>
					</tr>
					<tr>
						<th rowspan="2">NO</th>
						<th rowspan="2">Section</th>
						<th rowspan="2">Status</th>
						<th rowspan="2">Tenaga.Kerja</th>
						<th colspan="7">Hari kerja</th>
						<th colspan="7">Ijin dibayar</th>
						<th colspan="4">Ijin tdk dibayar</th>
						<th rowspan="2">Total Hari</th>
						<th rowspan="2">HKE</th>
						<th rowspan="2">%</th>
					</tr>
					<tr>
						<th>KJ</th>
						<th>DL</th>
						<th>WH</th>
						<th>WS</th>
						<th>KK</th>
						<th>KL</th>
						<th>KH</th>
						
						<th>P2</th>
						<th>P3</th>
						<th>S1</th>
						<th>S2</th>
						<th>H1</th>
						<th>H2</th>
						<th>CT</th>
						
						<th>MK</th>
						<th>ML</th>
						<th>P1</th>
						<th>UN</th>
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

Dim currsection : currsection = ""
Dim prevsection : prevsection = ""

Dim Row_percent_thismonth : Row_percent_thismonth = 0
Dim Row_percent_todate : Row_percent_todate = 0

Dim Row_cntemp : Row_cntemp = 0

Dim Row_bi_kj : Row_bi_kj= 0
Dim Row_bi_dl : Row_bi_dl= 0
Dim Row_bi_wh : Row_bi_wh= 0
Dim Row_bi_ws : Row_bi_ws= 0
Dim Row_bi_kk : Row_bi_kk= 0
Dim Row_bi_kl : Row_bi_kl= 0
Dim Row_bi_kh : Row_bi_kh= 0

Dim Row_bi_p2 : Row_bi_p2= 0
Dim Row_bi_p3 : Row_bi_p3= 0
Dim Row_bi_s1 : Row_bi_s1= 0
Dim Row_bi_s2 : Row_bi_s2= 0
Dim Row_bi_h1 : Row_bi_h1= 0
Dim Row_bi_h2 : Row_bi_h2= 0
Dim Row_bi_ct : Row_bi_ct= 0

Dim Row_bi_mk : Row_bi_mk= 0
Dim Row_bi_ml : Row_bi_ml= 0
Dim Row_bi_p1 : Row_bi_p1= 0
Dim Row_bi_un : Row_bi_un= 0

Dim Row_bi_sh : Row_bi_sh = 0
Dim Row_bi_p23 : Row_bi_p23 = 0
Dim Row_bi_paid : Row_bi_paid = 0
Dim Row_bi_mkp1 : Row_bi_mkp1 = 0
Dim Row_bi_unpaid : Row_bi_unpaid = 0
Dim Row_bi_whws : Row_bi_whws = 0

Dim Row_bi_hke : Row_bi_hke = 0
Dim Row_bi_total : Row_bi_total = 0
Dim Row_bi_paid_percent : Row_bi_paid_percent = 0
Dim Row_bi_unpaid_percent : Row_bi_unpaid_percent = 0
Dim Row_bi_hke_percent : Row_bi_hke_percent = 0




if rsLHPP.BOF and rsLHPP.EOF then
%>
					<tr>
						<td colspan="<%=37%>">NO DATA AVAILABLE</td>
					</tr>
<%
else
		while not rsLHPP.eof
		    Row_cntemp = Cdbl(rsLHPP("cntemp"))
			Row_bi_kj= Cdbl(rsLHPP("bi_kj"))
			Row_bi_dl= Cdbl(rsLHPP("bi_dl"))
			Row_bi_wh= Cdbl(rsLHPP("bi_wh"))
			Row_bi_ws= Cdbl(rsLHPP("bi_ws"))
			Row_bi_kk= Cdbl(rsLHPP("bi_kk"))
			Row_bi_kl= Cdbl(rsLHPP("bi_kl"))
			Row_bi_kh= Cdbl(rsLHPP("bi_kh"))

			Row_bi_p2= Cdbl(rsLHPP("bi_p2"))
			Row_bi_p3= Cdbl(rsLHPP("bi_p3"))
			Row_bi_s1= Cdbl(rsLHPP("bi_s1"))
			Row_bi_s2= Cdbl(rsLHPP("bi_s2"))
			Row_bi_h1= Cdbl(rsLHPP("bi_h1"))
			Row_bi_h2= Cdbl(rsLHPP("bi_h2"))
			Row_bi_ct= Cdbl(rsLHPP("bi_ct"))

			Row_bi_mk= Cdbl(rsLHPP("bi_mk"))
			Row_bi_ml= Cdbl(rsLHPP("bi_ml"))
			Row_bi_p1= Cdbl(rsLHPP("bi_p1"))
			Row_bi_un= Cdbl(rsLHPP("bi_un"))
			
			Row_bi_total = Row_bi_kj+Row_bi_dl+Row_bi_wh+Row_bi_ws+Row_bi_kk+Row_bi_kl+Row_bi_kh+Row_bi_p2+Row_bi_p3+Row_bi_s1+Row_bi_s2+Row_bi_h1+Row_bi_h2+Row_bi_ct+Row_bi_mk+Row_bi_ml+Row_bi_p1+Row_bi_un
			
			if Row_cntemp = 0 then 
				Row_bi_hke = 0
			Else
				Row_bi_hke   = (Row_bi_kj+Row_bi_dl+Row_bi_wh+Row_bi_ws+Row_bi_kk+Row_bi_kl+Row_bi_kh) / Row_cntemp
			end If
			
			if Row_bi_total = 0 then
				Row_bi_hke_percent = 0
			else
			   if trim(rsLHPP("grade"))="KHL" then
					if Row_bi_hke = 0 then
						Row_bi_hke_percent = 0
					else
						Row_bi_hke_percent = Round( (Row_bi_hke/20) * 100, 2)
					end if
			   else
				    Row_bi_hke_percent = Round( (Row_bi_hke/24) * 100, 2)
			   end if	
			end if
			
			
%>
				<%if rsLHPP("section") = "Grand Total" then%>
					<tr>
						<td colspan="38"></td>
					</tr>
					<%
						if not arr_unit_eof then
							for xxx=0 to UBound(arr_unit,2)
								Row_bi_cntemp = Cdbl(arr_unit(2,xxx))
								
								Row_bi_kj= Cdbl(arr_unit(3,xxx))
								Row_bi_dl= Cdbl(arr_unit(4,xxx))
								Row_bi_wh= Cdbl(arr_unit(5,xxx))
								Row_bi_ws= Cdbl(arr_unit(6,xxx))
								Row_bi_kk= Cdbl(arr_unit(7,xxx))
								Row_bi_kl= Cdbl(arr_unit(8,xxx))
								Row_bi_kh= Cdbl(arr_unit(9,xxx))

								Row_bi_p2= Cdbl(arr_unit(10,xxx))
								Row_bi_p3= Cdbl(arr_unit(11,xxx))
								Row_bi_s1= Cdbl(arr_unit(12,xxx))
								Row_bi_s2= Cdbl(arr_unit(13,xxx))
								Row_bi_h1= Cdbl(arr_unit(14,xxx))
								Row_bi_h2= Cdbl(arr_unit(15,xxx))
								Row_bi_ct= Cdbl(arr_unit(16,xxx))

								Row_bi_mk= Cdbl(arr_unit(17,xxx))
								Row_bi_ml= Cdbl(arr_unit(18,xxx))
								Row_bi_p1= Cdbl(arr_unit(19,xxx))
								Row_bi_un= Cdbl(arr_unit(20,xxx))
			
								Row_bi_gttotal = Row_bi_kj+Row_bi_dl+Row_bi_wh+Row_bi_ws+Row_bi_kk+Row_bi_kl+Row_bi_kh+Row_bi_p2+Row_bi_p3+Row_bi_s1+Row_bi_s2+Row_bi_h1+Row_bi_h2+Row_bi_ct+Row_bi_mk+Row_bi_ml+Row_bi_p1+Row_bi_un
			
								if Row_bi_cntemp = 0 then 
									Row_bi_gthke = 0
								Else
									Row_bi_gthke   = (Row_bi_kj+Row_bi_dl+Row_bi_wh+Row_bi_ws+Row_bi_kk+Row_bi_kl+Row_bi_kh) / Row_cntemp
								end If
								
								if Row_bi_gttotal = 0 then
									Row_bi_gthke_percent = 0
								else
								   
								   if trim(rsLHPP("grade"))="KHL" then
										if Row_bi_gthke = 0 then
											Row_bi_gthke_percent = 0
										else
											Row_bi_gthke_percent = Round( (Row_bi_gthke/20) * 100, 2)
										end if
								   else
										Row_bi_gthke_percent = Round( (Row_bi_gthke/24) * 100, 2)
								   end if	
								end if
								
								
								response.write "<tr>"
								%>
									<%if xxx=0 then%>
									<td align="center" colspan="2" rowspan="<%=Cdbl(rowspan_gt(arr_unit))%>"><%=post_desc_total%></td>
									<%end if%>
									<td align="center"><%=arr_unit(0,xxx)%></td>
									<td align="right"> <%=arr_unit(2,xxx)%></td>
									
									<td align="right"><%=FormatNumber(arr_unit(3,xxx), 0)%></td>
									<td align="right"><%=FormatNumber(arr_unit(4,xxx), 0)%></td>
									<td align="right"><%=FormatNumber(arr_unit(5,xxx), 0)%></td>
									<td align="right"><%=FormatNumber(arr_unit(6,xxx), 0)%></td>
									<td align="right"><%=FormatNumber(arr_unit(7,xxx), 0)%></td>
									<td align="right"><%=FormatNumber(arr_unit(8,xxx), 0)%></td>
									<td align="right"><%=FormatNumber(arr_unit(9,xxx), 0)%></td>
									
									<td align="right"><%=FormatNumber(arr_unit(10,xxx), 0)%></td>
									<td align="right"><%=FormatNumber(arr_unit(11,xxx), 0)%></td>
									<td align="right"><%=FormatNumber(arr_unit(12,xxx), 0)%></td>
									<td align="right"><%=FormatNumber(arr_unit(13,xxx), 0)%></td>
									<td align="right"><%=FormatNumber(arr_unit(14,xxx), 0)%></td>
									<td align="right"><%=FormatNumber(arr_unit(15,xxx), 0)%></td>
									<td align="right"><%=FormatNumber(arr_unit(16,xxx), 0)%></td>
									
									<td align="right"><%=FormatNumber(arr_unit(17,xxx), 0)%></td>
									<td align="right"><%=FormatNumber(arr_unit(18,xxx), 0)%></td>
									<td align="right"><%=FormatNumber(arr_unit(19,xxx), 0)%></td>
									<td align="right"><%=FormatNumber(arr_unit(20,xxx), 0)%></td>
									
									
									<td align="right"><%=FormatNumber(Row_bi_gttotal, 0)%></td>
									<td align="right"><%=FormatNumber(Row_bi_gthke, 0)%></td>
									<td align="right"><%=FormatNumber(Row_bi_gthke_percent, 2)%></td>
									
								<%								
								response.write "</tr>"
							next
						end if
					%>
					
					<tr>
						<td colspan="37"></td>
					</tr>
					<tr bgcolor="#ffff01" >
						<td colspan="3" align="left"><strong><%=rsLHPP("section")%></strong></td>
						<td align="right"><strong><%=FormatNumber(Row_cntemp, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(Row_bi_kj, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(Row_bi_dl, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(Row_bi_wh, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(Row_bi_ws, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(Row_bi_kk, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(Row_bi_kl, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(Row_bi_kh, 0)%></strong></td>
						
						<td align="right"><strong><%=FormatNumber(Row_bi_p2, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(Row_bi_p3, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(Row_bi_s1, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(Row_bi_s2, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(Row_bi_h1, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(Row_bi_h2, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(Row_bi_ct, 0)%></strong></td>
						
						<td align="right"><strong><%=FormatNumber(Row_bi_mk, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(Row_bi_ml, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(Row_bi_p1, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(Row_bi_un, 0)%></strong></td>
						
						<td align="right"><strong><%=FormatNumber(Row_bi_total, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(Row_bi_hke, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(Row_bi_hke_percent, 2)%></strong></td>
					</tr>
				<%elseif Left(Cstr(rsLHPP("section")), 8) = "Subtotal" then%>
					<tr bgcolor="#F4F4A0" >
						<td colspan="3" align="left"><strong><%=rsLHPP("section")%></strong></td>
						<td align="right"><strong><%=FormatNumber(Row_cntemp, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(Row_bi_kj, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(Row_bi_dl, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(Row_bi_wh, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(Row_bi_ws, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(Row_bi_kk, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(Row_bi_kl, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(Row_bi_kh, 0)%></strong></td>
						
						<td align="right"><strong><%=FormatNumber(Row_bi_p2, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(Row_bi_p3, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(Row_bi_s1, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(Row_bi_s2, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(Row_bi_h1, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(Row_bi_h2, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(Row_bi_ct, 0)%></strong></td>
						
						<td align="right"><strong><%=FormatNumber(Row_bi_mk, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(Row_bi_ml, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(Row_bi_p1, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(Row_bi_un, 0)%></strong></td>
						
						<td align="right"><strong><%=FormatNumber(Row_bi_total, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(Row_bi_hke, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(Row_bi_hke_percent, 2)%></strong></td>
					</tr>
				<%else%>
					<%currsection = rsLHPP("section")%>
					
					<tr>
						<%if currsection <> prevsection then%>
							<%OrderSection = OrderSection + 1%>
							<td align="center" style=" border-bottom:none;"><%=OrderSection%></td>
							<td align="center" style=" border-bottom:none;"><%=rsLHPP("section")%></td>
						<%else%>
							<td align="center" style="border:none;border-right:1px #3A5F0B solid; border-left:1px #3A5F0B solid;"></td>
							<td align="center" style="border:none;border-right:1px #3A5F0B solid; border-left:1px #3A5F0B solid;"></td>
						<%end if%>
						<td align="center"><%=rsLHPP("grade")%></td>
						<td align="right"><%=rsLHPP("cntemp")%></td>
						<td align="right"><%=FormatNumber(Row_bi_kj, 0)%></td>
						<td align="right"><%=FormatNumber(Row_bi_dl, 0)%></td>
						<td align="right"><%=FormatNumber(Row_bi_wh, 0)%></td>
						<td align="right"><%=FormatNumber(Row_bi_ws, 0)%></td>
						<td align="right"><%=FormatNumber(Row_bi_kk, 0)%></td>
						<td align="right"><%=FormatNumber(Row_bi_kl, 0)%></td>
						<td align="right"><%=FormatNumber(Row_bi_kh, 0)%></td>
						
						<td align="right"><%=FormatNumber(Row_bi_p2, 0)%></td>
						<td align="right"><%=FormatNumber(Row_bi_p3, 0)%></td>
						<td align="right"><%=FormatNumber(Row_bi_s1, 0)%></td>
						<td align="right"><%=FormatNumber(Row_bi_s2, 0)%></td>
						<td align="right"><%=FormatNumber(Row_bi_h1, 0)%></td>
						<td align="right"><%=FormatNumber(Row_bi_h2, 0)%></td>
						<td align="right"><%=FormatNumber(Row_bi_ct, 0)%></td>
						
						<td align="right"><%=FormatNumber(Row_bi_mk, 0)%></td>
						<td align="right"><%=FormatNumber(Row_bi_ml, 0)%></td>
						<td align="right"><%=FormatNumber(Row_bi_p1, 0)%></td>
						<td align="right"><%=FormatNumber(Row_bi_un, 0)%></td>
						
						<td align="right"><%=FormatNumber(Row_bi_total, 0)%></td>
						<td align="right"><%=FormatNumber(Row_bi_hke, 0)%></td>
						<td align="right"><%=FormatNumber(Row_bi_hke_percent, 2)%></td>
					</tr>
					<%prevsection = currsection%>
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
