<%@Language=VBScript%>
<%Option Explicit%>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
	Dim v_txtBudgetYear, v_txtBudgetVersion, v_txtModuleVersion, v_txtdatefrom, v_txtMonth, v_txtYear, v_txtgrade, v_txtDivisi, v_txtemployeetype
	Dim no
	Dim rsDefaultBa, strDefaultBa, varDefaultBa
	Dim ScriptName
	Dim strMain, sqlLHPP, sqlLHPP2, rsLHPP
	Dim q, CompanyName, fcbaname, MgrName, KTUName, CropName, FcbaType
	Dim OrderSection : OrderSection = 0
	Dim gt_hk_bi : gt_hk_bi = 0
	Dim gt_gaji_bi : gt_gaji_bi = 0
	Dim gt_premilembur_bi : gt_premilembur_bi = 0
	Dim gt_beras_bi : gt_beras_bi = 0
	Dim gt_thr_bi : gt_thr_bi = 0
	Dim gt_bonus_bi : gt_bonus_bi = 0
	Dim gt_bpjs_bi : gt_bpjs_bi = 0
	Dim gt_pengobatan_bi : gt_pengobatan_bi = 0
	Dim gt_perumahan_bi : gt_perumahan_bi = 0
	Dim gt_seragam_bi : gt_seragam_bi = 0
	Dim gt_extrafooding_bi : gt_extrafooding_bi = 0
	Dim gt_hk_sbi : gt_hk_sbi = 0
	Dim gt_gaji_sbi : gt_gaji_sbi = 0
	Dim gt_premilembur_sbi : gt_premilembur_sbi = 0
	Dim gt_beras_sbi : gt_beras_sbi = 0
	Dim gt_thr_sbi : gt_thr_sbi = 0
	Dim gt_bonus_sbi : gt_bonus_sbi = 0
	Dim gt_bpjs_sbi : gt_bpjs_sbi = 0
	Dim gt_pengobatan_sbi : gt_pengobatan_sbi = 0
	Dim gt_perumahan_sbi : gt_perumahan_sbi = 0
	Dim gt_seragam_sbi : gt_seragam_sbi = 0
	Dim gt_extrafooding_sbi : gt_extrafooding_sbi = 0
	
	Dim row_total_bi : row_total_bi = 0
	Dim row_total_sbi : row_total_sbi = 0
	
	Dim gt_total_bi : gt_total_bi = 0
	Dim gt_total_sbi : gt_total_sbi = 0
	
	' **** Sub Total (Kebun)
	Dim rsST, strST, arrayST, arrayST_eof, iii, st_indx
	
	Dim row_st_total_bi : row_st_total_bi = 0
	Dim row_st_total_sbi : row_st_total_sbi = 0
	
	arrayST_eof = true
	' ****End Sub Total (Kebun)
	
	Dim rowspan_division : rowspan_division = 0
	Dim rsRowspan, strRowspan
	Dim posting_rowspan : posting_rowspan = 0


    v_txtGroupPlantation = Request("txtGroupPlantation")
	v_txtPlantation = Request("txtPlantation")
	v_txtGroupRegion = Request("txtGroupRegion")
	v_txtRegion = Request("txtRegion")
	v_txtcompany = Request("txtCompany")
	v_txtfcba = Request("txtfcba")
	
	v_txtDivisi = Request("txtDivisi")
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

sqlLHPP = " SELECT fcba," & _
          "        section," & _
          "        grade," & _
          "        hk_bi," & _
          "        gaji_bi," & _
          "        premilembur_bi," & _
          "        beras_bi," & _
          "        thr_bi," & _
          "        bonus_bi," & _
          "        bpjs_bi," & _
          "        pengobatan_bi," & _
          "        perumahan_bi," & _
          "        seragam_bi," & _
          "        extrafooding_bi," & _
          "        hk_sbi," & _
          "        gaji_sbi," & _
          "        premilembur_sbi," & _
          "        beras_sbi," & _
          "        thr_sbi," & _
          "        bonus_sbi," & _
          "        bpjs_sbi," & _
          "        pengobatan_sbi," & _
          "        perumahan_sbi," & _
          "        seragam_sbi," & _
          "        extrafooding_sbi," & _
		  "        targetcode" & _
          "   FROM rpt_upr_dtl_tab" & _
          " WHERE fcba = '"& v_txtFcba & "'" & _
          "       AND year = " & v_txtYear & "" & _
		  "       AND month = " & v_txtMonth & ""

if len(trim(v_txtDivisi)) <> 0 then
		sqlLHPP = sqlLHPP & " AND section = '" & v_txtDivisi & "'"
end if

if len(trim(v_txtgrade)) <> 0 then
		sqlLHPP = sqlLHPP & " AND grade = '" & v_txtgrade & "'"
end if

sqlLHPP = sqlLHPP & " order by SECTIONNAME_API.get_orderdata(section), parameterdetail_api.get_rate_1(fcba, 'RPT_UPR', 'STATUS', grade)"
		  
'response.write sqlLHPP
%>


<html>

	<head>
		<title>ePlaS Reporting - Laporan Upah Rata-Rata</title>
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
							strDivisi	=	" select distinct fccode, fcname, orderdata" & _
											" from (" & _
											"     select pd.fccode, pd.fcname, pd.rate_1 orderdata" & _
											"     from parameter pr inner join parameterdetail pd" & _
											"     on (pr.fcba = pd.fcba and pr.parcode = pd.parcode and pr.fccode = pd.parheadcode)" & _
											"     inner join ( select fcba, fccode" & _
											"                         , rpt_upr_api.mapping_loccode_to_detail (fccode, 'SECTION') division" & _
											"                  from gcmaster" & _
											"                  where fcba = '" & v_txtFcba & "'" & _
											"                        and item = 'LOC'" & _
											"                        and activation = 'Y' ) loc" & _
											"     on (pd.fcba = loc.fcba and pd.fccode = loc.division)" & _
											"     where pd.fcba = '" & v_txtFcba & "'" & _
											"           and pd.parcode = 'RPT_UPR'" & _
											"           and pd.parheadcode = 'DIVISION'" & _
											" )" & _
											" order by orderdata"

							Set rsDivisi = Server.CreateObject("ADODB.Recordset")
							rsDivisi.Open strDivisi, connstr, adOpenStatic 

							if not rsDivisi.eof then
								Do While Not rsDivisi.eof
									if rsDivisi("fccode") = v_txtDivisi then%>
										<option value="<%=rsDivisi("fccode")%>" Selected><%=rsDivisi("fccode")%> : <%=rsDivisi("fcname")%></option> 
									<%Else%>
										<option value="<%=rsDivisi("fccode")%>"><%=rsDivisi("fccode")%> : <%=rsDivisi("fcname")%></option>  
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
								<select name="txtgrade" >
								<Option Value="" <%If v_txtgrade = "" then%> Selected <%End If%>>--ALL--</option>
							<%
							Dim rsStatus, strStatus
							strStatus	=	" select pd.fccode, pd.fcname, pd.rate_1 orderdata" & _
												" from parameter pr inner join parameterdetail pd" & _
												" on (pr.fcba = pd.fcba and pr.parcode = pd.parcode and pr.fccode = pd.parheadcode)" & _
												" where pd.fcba = '" & v_txtFcba & "'" & _
												"       and pd.parcode = 'RPT_UPR'" & _
												"       and pd.parheadcode = 'STATUS'" & _
												" order by pd.rate_1"

							Set rsStatus = Server.CreateObject("ADODB.Recordset")
							rsStatus.Open strStatus, connstr, adOpenStatic 

							if not rsStatus.eof then
								Do While Not rsStatus.eof
									if rsStatus("fccode") = v_txtgrade then%>
										<option value="<%=rsStatus("fccode")%>" Selected><%=rsStatus("fccode")%> : <%=rsStatus("fcname")%></option> 
									<%Else%>
										<option value="<%=rsStatus("fccode")%>"><%=rsStatus("fccode")%> : <%=rsStatus("fcname")%></option>  
									<%End if
									rsStatus.movenext
								loop
							end if
							rsStatus.Close
							set rsStatus = Nothing
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
						<th colspan="<%=28%>" style="background:#ffffff; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
							<div id="header-wrapper">
								<div id="subheader" class="centered">
									<h1>LAPORAN UPAH RATA-RATA</h1>
									<div style="margin-botton:10px;"/>
									<h3>UPAH RATA-RATA KARYAWAN LANGSUNG</h3>
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
						<th rowspan="2">Division</th>
						<th rowspan="2">Status</th>
						<th colspan="13">Bulan Ini</th>
						<th colspan="13">Sampai Bulan Ini</th>
					</tr>
					<tr>
						<th>HK</th>
						<th>Gaji</th>
						<th>Premi/Lembur</th>
						<th>Beras</th>
						<th>THR</th>
						<th>Bonus</th>
						<th>Jamsostek</th>
						<th>Pengobatan</th>
						<th>Perumahan</th>
						<th>Seragam</th>
						<th>Extrafooding</th>
						<th>Total</th>
						<th>RP/HK</th>
						
						<th>HK</th>
						<th>Gaji</th>
						<th>Premi/Lembur</th>
						<th>Beras</th>
						<th>THR</th>
						<th>Bonus</th>
						<th>Jamsostek</th>
						<th>Pengobatan</th>
						<th>Perumahan</th>
						<th>Seragam</th>
						<th>Extrafooding</th>
						<th>Total</th>
						<th>RP/HK</th>
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
						<td colspan="<%=28%>">NO DATA AVAILABLE</td>
					</tr>
<%end if%>
<%if not array_eof then%>
	<%
		strST = " SELECT fcba, grade," & _
					"        SUM(hk_bi) hk_bi," & _
					"        SUM(gaji_bi) gaji_bi," & _
					"        SUM(premilembur_bi) premilembur_bi," & _
					"        SUM(beras_bi) beras_bi," & _
					"        SUM(thr_bi) thr_bi," & _
					"        SUM(bonus_bi) bonus_bi," & _
					"        SUM(bpjs_bi) bpjs_bi," & _
					"        SUM(pengobatan_bi) pengobatan_bi," & _
					"        SUM(perumahan_bi) perumahan_bi," & _
					"        SUM(seragam_bi) seragam_bi," & _
					"        SUM(extrafooding_bi) extrafooding_bi," & _
					"        SUM(hk_sbi) hk_sbi," & _
					"        SUM(gaji_sbi) gaji_sbi," & _
					"        SUM(premilembur_sbi) premilembur_sbi," & _
					"        SUM(beras_sbi) beras_sbi," & _
					"        SUM(thr_sbi) thr_sbi," & _
					"        SUM(bonus_sbi) bonus_sbi," & _
					"        SUM(bpjs_sbi) bpjs_sbi," & _
					"        SUM(pengobatan_sbi) pengobatan_sbi," & _
					"        SUM(perumahan_sbi) perumahan_sbi," & _
					"        SUM(seragam_sbi) seragam_sbi," & _
					"        SUM(extrafooding_sbi) extrafooding_sbi," & _
					"        parameterdetail_api.get_param_1(fcba, 'RPT_UPR', 'STATUS', grade) param" & _
					" FROM (" & sqlLHPP & ")" & _
					" GROUP BY fcba, grade" & _
					" ORDER BY parameterdetail_api.get_rate_1(fcba, 'RPT_UPR', 'STATUS', grade)"
					
	Set rsST = Server.CreateObject("ADODB.Recordset")
	rsST.Open strST, connstr, adOpenStatic

	if not rsST.EOF then
		arrayST = rsST.getrows()
		arrayST_eof= false
	end if
	rsST.Close
	set rsST = Nothing
	
	gt_hk_bi = 0
	gt_gaji_bi = 0
	gt_premilembur_bi = 0
	gt_beras_bi = 0
	gt_thr_bi = 0
	gt_bonus_bi = 0
	gt_bpjs_bi = 0
	gt_pengobatan_bi = 0
	gt_perumahan_bi = 0
	gt_seragam_bi = 0
	gt_extrafooding_bi = 0
	gt_hk_sbi = 0
	gt_gaji_sbi = 0
	gt_premilembur_sbi = 0
	gt_beras_sbi = 0
	gt_thr_sbi = 0
	gt_bonus_sbi = 0
	gt_bpjs_sbi = 0
	gt_pengobatan_sbi = 0
	gt_perumahan_sbi = 0
	gt_seragam_sbi = 0
	gt_extrafooding_sbi = 0
	
	rowspan_division = 0
	%>
	<%For array_i = LBound(array_data, 2) To UBound(array_data, 2)%>
					<tr>
					
					<%
					rowspan_division = rowspan_division + 1
					row_total_bi = 0
					row_total_sbi = 0
					%>
						<%if rowspan_division = 1 then%>
						<% ' Cari Rowspan
						strRowspan = " SELECT COUNT (1) rowspan FROM (" & sqlLHPP & ")" & _
											" WHERE section = '"& array_data(1, array_i) &"'"
						Set rsRowspan = Server.CreateObject("ADODB.Recordset")
						rsRowspan.Open strRowspan, connstr, adOpenStatic
						if not rsRowspan.EOF then
							posting_rowspan = CDbl(rsRowspan("rowspan"))
						end if
						rsRowspan.Close
						set rsRowspan = Nothing
						%>
							<td align="center" rowspan="<%=posting_rowspan%>"><%=array_data(1, array_i)%></td>
						<%end if%>
						
						<td align="left"><%=array_data(2, array_i)%></td>
					
					<%if IsNull(array_data(25, array_i)) or array_data(25, array_i) = "" then%>
						
						<%
						row_total_bi = row_total_bi + Cdbl( array_data(4, array_i) )
						row_total_sbi = row_total_sbi + Cdbl( array_data(15, array_i) )
						%>
					
					
						<td align="right"><%=FormatNumber(array_data(3, array_i), 1)%></td>
						<td align="right"><%=FormatNumber(array_data(4, array_i), 0)%></td>
						<td align="right" bgcolor="grey">&nbsp;</td>
						<td align="right" bgcolor="grey">&nbsp;</td>
						<td align="right" bgcolor="grey">&nbsp;</td>
						<td align="right" bgcolor="grey">&nbsp;</td>
						<td align="right" bgcolor="grey">&nbsp;</td>
						<td align="right" bgcolor="grey">&nbsp;</td>
						<td align="right" bgcolor="grey">&nbsp;</td>
						<td align="right" bgcolor="grey">&nbsp;</td>
						<td align="right" bgcolor="grey">&nbsp;</td>
						
						
						<td align="right"><%=FormatNumber(row_total_bi, 0)%></td>
						
						<%if Cdbl(array_data(3, array_i)) = 0 then%>
								<td align="right"><%=0%></td>
						<%else%>
								<td align="right"><%=FormatNumber( Round( row_total_bi/Cdbl(array_data(3, array_i)), 0), 0)%></td>
						<%end if%>
						
						<td align="right"><%=FormatNumber(array_data(14, array_i), 1)%></td>
						<td align="right"><%=FormatNumber(array_data(15, array_i), 0)%></td>
						<td align="right" bgcolor="grey">&nbsp;</td>
						<td align="right" bgcolor="grey">&nbsp;</td>
						<td align="right" bgcolor="grey">&nbsp;</td>
						<td align="right" bgcolor="grey">&nbsp;</td>
						<td align="right" bgcolor="grey">&nbsp;</td>
						<td align="right" bgcolor="grey">&nbsp;</td>
						<td align="right" bgcolor="grey">&nbsp;</td>
						<td align="right" bgcolor="grey">&nbsp;</td>
						<td align="right" bgcolor="grey">&nbsp;</td>
						
						<td align="right"><%=FormatNumber(row_total_sbi, 0)%></td>
						
						<%if Cdbl(array_data(14, array_i)) = 0 then%>
								<td align="right"><%=0%></td>
						<%else%>
								<td align="right"><%=FormatNumber( Round( row_total_sbi/Cdbl(array_data(14, array_i)), 0), 0)%></td>
						<%end if%>
					
						
					
					<%else%>
					
						<%
						row_total_bi = row_total_bi + Cdbl( array_data(4, array_i)) +  Cdbl(array_data(5, array_i)) +  Cdbl(array_data(6, array_i)) +  Cdbl(array_data(7, array_i)) +  Cdbl(array_data(8, array_i)) +  Cdbl(array_data(9, array_i)) +  Cdbl(array_data(10, array_i)) +  Cdbl(array_data(11, array_i)) +  Cdbl(array_data(12, array_i)) +  Cdbl(array_data(13, array_i) )
						row_total_sbi = row_total_sbi + Cdbl( array_data(15, array_i)) +  Cdbl(array_data(16, array_i)) +  Cdbl(array_data(17, array_i)) +  Cdbl(array_data(18, array_i)) +  Cdbl(array_data(19, array_i)) +  Cdbl(array_data(20, array_i)) +  Cdbl(array_data(21, array_i)) +  Cdbl(array_data(22, array_i)) +  Cdbl(array_data(23, array_i)) +  Cdbl(array_data(24, array_i) )
						%>
						
						<%
						gt_hk_bi = gt_hk_bi + cdbl(array_data(3, array_i))
						gt_gaji_bi = gt_gaji_bi + cdbl(array_data(4, array_i))
						gt_premilembur_bi = gt_premilembur_bi + cdbl(array_data(5, array_i))
						gt_beras_bi = gt_beras_bi + cdbl(array_data(6, array_i))
						gt_thr_bi = gt_thr_bi + cdbl(array_data(7, array_i))
						gt_bonus_bi = gt_bonus_bi + cdbl(array_data(8, array_i))
						gt_bpjs_bi = gt_bpjs_bi + cdbl(array_data(9, array_i))
						gt_pengobatan_bi = gt_pengobatan_bi + cdbl(array_data(10, array_i))
						gt_perumahan_bi = gt_perumahan_bi + cdbl(array_data(11, array_i))
						gt_seragam_bi = gt_seragam_bi + cdbl(array_data(12, array_i))
						gt_extrafooding_bi = gt_extrafooding_bi + cdbl(array_data(13, array_i))
						gt_hk_sbi = gt_hk_sbi + cdbl(array_data(14, array_i))
						gt_gaji_sbi = gt_gaji_sbi + cdbl(array_data(15, array_i))
						gt_premilembur_sbi = gt_premilembur_sbi + cdbl(array_data(16, array_i))
						gt_beras_sbi = gt_beras_sbi + cdbl(array_data(17, array_i))
						gt_thr_sbi = gt_thr_sbi + cdbl(array_data(18, array_i))
						gt_bonus_sbi = gt_bonus_sbi + cdbl(array_data(19, array_i))
						gt_bpjs_sbi = gt_bpjs_sbi + cdbl(array_data(20, array_i))
						gt_pengobatan_sbi = gt_pengobatan_sbi + cdbl(array_data(21, array_i))
						gt_perumahan_sbi = gt_perumahan_sbi + cdbl(array_data(22, array_i))
						gt_seragam_sbi = gt_seragam_sbi + cdbl(array_data(23, array_i))
						gt_extrafooding_sbi = gt_extrafooding_sbi + cdbl(array_data(24, array_i))					
						%>
					
						<td align="right"><%=FormatNumber(array_data(3, array_i), 1)%></td>
						<td align="right"><%=FormatNumber(array_data(4, array_i), 0)%></td>
						<td align="right"><%=FormatNumber(array_data(5, array_i), 0)%></td>
						<td align="right"><%=FormatNumber(array_data(6, array_i), 0)%></td>
						<td align="right"><%=FormatNumber(array_data(7, array_i), 0)%></td>
						<td align="right"><%=FormatNumber(array_data(8, array_i), 0)%></td>
						<td align="right"><%=FormatNumber(array_data(9, array_i), 0)%></td>
						<td align="right"><%=FormatNumber(array_data(10, array_i), 0)%></td>
						<td align="right"><%=FormatNumber(array_data(11, array_i), 0)%></td>
						<td align="right"><%=FormatNumber(array_data(12, array_i), 0)%></td>
						<td align="right"><%=FormatNumber(array_data(13, array_i), 0)%></td>
						
						<td align="right"><%=FormatNumber(row_total_bi, 0)%></td>
						
						<%if Cdbl(array_data(3, array_i)) = 0 then%>
								<td align="right"><%=0%></td>
						<%else%>
								<td align="right"><%=FormatNumber( Round( row_total_bi/Cdbl(array_data(3, array_i)), 0), 0)%></td>
						<%end if%>
						
						<td align="right"><%=FormatNumber(array_data(14, array_i), 1)%></td>
						<td align="right"><%=FormatNumber(array_data(15, array_i), 0)%></td>
						<td align="right"><%=FormatNumber(array_data(16, array_i), 0)%></td>
						<td align="right"><%=FormatNumber(array_data(17, array_i), 0)%></td>
						<td align="right"><%=FormatNumber(array_data(18, array_i), 0)%></td>
						<td align="right"><%=FormatNumber(array_data(19, array_i), 0)%></td>
						<td align="right"><%=FormatNumber(array_data(20, array_i), 0)%></td>
						<td align="right"><%=FormatNumber(array_data(21, array_i), 0)%></td>
						<td align="right"><%=FormatNumber(array_data(22, array_i), 0)%></td>
						<td align="right"><%=FormatNumber(array_data(23, array_i), 0)%></td>
						<td align="right"><%=FormatNumber(array_data(24, array_i), 0)%></td>
						
						<td align="right"><%=FormatNumber(row_total_sbi, 0)%></td>
						
						<%if Cdbl(array_data(14, array_i)) = 0 then%>
								<td align="right"><%=0%></td>
						<%else%>
								<td align="right"><%=FormatNumber( Round( row_total_sbi/Cdbl(array_data(14, array_i)), 0), 0)%></td>
						<%end if%>
						
					<%end if%>
					
					<%
						if rowspan_division = posting_rowspan then
							rowspan_division = 0
						end if
					%>
					
					</tr>
					
	<%next%>
<%end if%>

	<%if not arrayST_eof then%>
	<%st_indx = 0%>
		<tr><td colspan="<%=28%>"></td></tr>
		<%For iii = LBound(arrayST, 2) To UBound(arrayST, 2)%>
			<%
			st_indx = st_indx + 1
			row_st_total_bi = 0
			row_st_total_sbi = 0
			
			row_st_total_bi = row_st_total_bi + Cdbl( arrayST(3, iii)) + Cdbl( arrayST(4, iii)) +  Cdbl(arrayST(5, iii)) +  Cdbl(arrayST(6, iii)) +  Cdbl(arrayST(7, iii)) +  Cdbl(arrayST(8, iii)) +  Cdbl(arrayST(9, iii)) +  Cdbl(arrayST(10, iii)) +  Cdbl(arrayST(11, iii)) +  Cdbl(arrayST(12, iii))
			row_st_total_sbi = row_st_total_sbi + Cdbl( arrayST(14, iii)) + Cdbl( arrayST(15, iii)) +  Cdbl(arrayST(16, iii)) +  Cdbl(arrayST(17, iii)) +  Cdbl(arrayST(18, iii)) +  Cdbl(arrayST(19, iii)) +  Cdbl(arrayST(20, iii)) +  Cdbl(arrayST(21, iii)) +  Cdbl(arrayST(22, iii)) +  Cdbl(arrayST(23, iii)) 
			%>
			<tr>
				<%if st_indx = 1 then%>
					<% ' Cari Rowspan
					strRowspan = " SELECT COUNT (1) rowspan FROM (" & strST & ")"
					Set rsRowspan = Server.CreateObject("ADODB.Recordset")
					rsRowspan.Open strRowspan, connstr, adOpenStatic
					if not rsRowspan.EOF then
						posting_rowspan = CDbl(rsRowspan("rowspan"))
					end if
					rsRowspan.Close
					set rsRowspan = Nothing
					%>
					<td align="center" rowspan="<%=posting_rowspan%>">Kebun</td>
				<%end if%>
				
				<td align="left"><%=arrayST(1, iii)%></td>
				
				<td align="right"><%=FormatNumber(arrayST(2, iii), 1)%></td>
				<td align="right"><%=FormatNumber(arrayST(3, iii), 0)%></td>
				
				<%if IsNull(arrayST(24, iii)) then%>
					<td align="right"><%=FormatNumber(arrayST(4, iii), 0)%></td>
					<td align="right"><%=FormatNumber(arrayST(5, iii), 0)%></td>
					<td align="right"><%=FormatNumber(arrayST(6, iii), 0)%></td>
					<td align="right"><%=FormatNumber(arrayST(7, iii), 0)%></td>
					<td align="right"><%=FormatNumber(arrayST(8, iii), 0)%></td>
					<td align="right"><%=FormatNumber(arrayST(9, iii), 0)%></td>
					<td align="right"><%=FormatNumber(arrayST(10, iii), 0)%></td>
					<td align="right"><%=FormatNumber(arrayST(11, iii), 0)%></td>
					<td align="right"><%=FormatNumber(arrayST(12, iii), 0)%></td>
				<%else%>
					<td align="right" bgcolor="grey">&nbsp;</td>
					<td align="right" bgcolor="grey">&nbsp;</td>
					<td align="right" bgcolor="grey">&nbsp;</td>
					<td align="right" bgcolor="grey">&nbsp;</td>
					<td align="right" bgcolor="grey">&nbsp;</td>
					<td align="right" bgcolor="grey">&nbsp;</td>
					<td align="right" bgcolor="grey">&nbsp;</td>
					<td align="right" bgcolor="grey">&nbsp;</td>
					<td align="right" bgcolor="grey">&nbsp;</td>
				<%end if%>
				
				<td align="right"><%=FormatNumber(row_st_total_bi, 0)%></td>
				
				<%if Cdbl(arrayST(2, iii)) = 0 then%>
						<td align="right"><%=0%></td>
				<%else%>
						<td align="right"><%=FormatNumber( Round(row_st_total_bi/Cdbl(arrayST(2, iii)), 0), 0)%></td>
				<%end if%>
				
				<td align="right"><%=FormatNumber(arrayST(13, iii), 1)%></td>
				<td align="right"><%=FormatNumber(arrayST(14, iii), 0)%></td>
				<%if IsNull(arrayST(24, iii)) then%>
					<td align="right"><%=FormatNumber(arrayST(15, iii), 0)%></td>
					<td align="right"><%=FormatNumber(arrayST(16, iii), 0)%></td>
					<td align="right"><%=FormatNumber(arrayST(17, iii), 0)%></td>
					<td align="right"><%=FormatNumber(arrayST(18, iii), 0)%></td>
					<td align="right"><%=FormatNumber(arrayST(19, iii), 0)%></td>
					<td align="right"><%=FormatNumber(arrayST(20, iii), 0)%></td>
					<td align="right"><%=FormatNumber(arrayST(21, iii), 0)%></td>
					<td align="right"><%=FormatNumber(arrayST(22, iii), 0)%></td>
					<td align="right"><%=FormatNumber(arrayST(23, iii), 0)%></td>
				<%else%>
					<td align="right" bgcolor="grey">&nbsp;</td>
					<td align="right" bgcolor="grey">&nbsp;</td>
					<td align="right" bgcolor="grey">&nbsp;</td>
					<td align="right" bgcolor="grey">&nbsp;</td>
					<td align="right" bgcolor="grey">&nbsp;</td>
					<td align="right" bgcolor="grey">&nbsp;</td>
					<td align="right" bgcolor="grey">&nbsp;</td>
					<td align="right" bgcolor="grey">&nbsp;</td>
					<td align="right" bgcolor="grey">&nbsp;</td>
					
				<%end if%>
				
				<td align="right"><%=FormatNumber(row_st_total_sbi, 0)%></td>
				
				<%if Cdbl(arrayST(13, iii)) = 0 then%>
						<td align="right"><%=0%></td>
				<%else%>
						<td align="right"><%=FormatNumber( Round(row_st_total_sbi/Cdbl(arrayST(13, iii)), 0), 0)%></td>
				<%end if%>
				
			</tr>
		<%next%>
	<%end if%>

					<tr><td colspan="<%=28%>"></td></tr>

					<tr bgcolor="#f4f4a0">
						<td align="center" colspan="2"><strong>Grand Total</strong></td>
						<td align="right"><strong><%=FormatNumber(gt_hk_bi, 1)%></strong></td>
						<td align="right"><strong><%=FormatNumber(gt_gaji_bi, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(gt_premilembur_bi, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(gt_beras_bi, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(gt_thr_bi, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(gt_bonus_bi, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(gt_bpjs_bi, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(gt_pengobatan_bi, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(gt_perumahan_bi, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(gt_seragam_bi, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(gt_extrafooding_bi, 0)%></strong></td>
						
						<%
						gt_total_bi = gt_total_bi + gt_gaji_bi +  gt_premilembur_bi +  gt_beras_bi +  gt_thr_bi +  gt_bonus_bi +  gt_bpjs_bi +  gt_pengobatan_bi +  gt_perumahan_bi +  gt_seragam_bi +  gt_extrafooding_bi
						gt_total_sbi = gt_total_sbi + gt_gaji_sbi +  gt_premilembur_sbi +  gt_beras_sbi +  gt_thr_sbi +  gt_bonus_sbi +  gt_bpjs_sbi +  gt_pengobatan_sbi +  gt_perumahan_sbi +  gt_seragam_sbi +  gt_extrafooding_sbi
						%>
						
						<td align="right"><strong><%=FormatNumber(gt_total_bi, 0)%></strong></td>
						
						<%if gt_hk_bi = 0 then%>
								<td align="right"><strong><%=0%></strong></td>
						<%else%>
								<td align="right"><strong><%=FormatNumber( Round(gt_total_bi/gt_hk_bi, 0), 0)%></strong></td>
						<%end if%>
						
						<td align="right"><strong><%=FormatNumber(gt_hk_sbi, 1)%></strong></td>
						<td align="right"><strong><%=FormatNumber(gt_gaji_sbi, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(gt_premilembur_sbi, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(gt_beras_sbi, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(gt_thr_sbi, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(gt_bonus_sbi, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(gt_bpjs_sbi, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(gt_pengobatan_sbi, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(gt_perumahan_sbi, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(gt_seragam_sbi, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(gt_extrafooding_sbi, 0)%></strong></td>
						
						<td align="right"><strong><%=FormatNumber(gt_total_sbi, 0)%></strong></td>
						
						<%if gt_hk_sbi = 0 then%>
								<td align="right"><strong><%=0%></strong></td>
						<%else%>
								<td align="right"><strong><%=FormatNumber( Round(gt_total_sbi/gt_hk_sbi, 0), 0)%></strong></td>
						<%end if%>
					</tr>
				</tbody>
			</table>

		</div>
	</body>
</html>
