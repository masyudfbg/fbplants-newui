<%@Language=VBScript%>
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


	v_txtfcba = Request("txtfcba")
	v_txtDivisi = Request("txtDivisi")
	
	v_txtBudgetYear = Request("txtBudgetYear")
	v_txtBudgetVersion = Request("txtBudgetVersion")
	v_txtModuleVersion = Request("txtModuleVersion")
	
	v_txtdatefrom		= Request("txtdatefrom")
	
	v_txtYear = Request("txtYear")
	v_txtMonth = Request("txtMonth")
	v_txtDate = Request("txtDate")
	v_txtJobCode = Request("txtJobCode")
	v_txtjobgroup = Request("txtjobgroup")
	
	
	If Len(Trim(v_txtBudgetYear)) = 0 Then
		v_txtBudgetYear = Year(textDate(Date))
	Else
		v_txtBudgetYear = v_txtYear
	End If
	If Len(Trim(v_txtBudgetVersion)) = 0 Then v_txtBudgetVersion = 1
	If Len(Trim(v_txtModuleVersion)) = 0 Then v_txtModuleVersion = 1
	
	If Len(Trim(v_txtdatefrom)) = 0 Then v_txtdatefrom = textDate(Date)
	
	if "" & v_txtDate = "" then v_txtDate = textDate(Date)
	v_txtYear = Year(v_txtDate)
	v_txtMonth = Month(v_txtDate)
	
	
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

Dim array_data, array_eof, xx

array_eof = true

strMain =   "SELECT  " & _
			"FCBA,  KODEAKUN, NAMAAKUN, " & _
			"DIVISI, BLOK, TAHUNTANAM, LUASBLOK, STATUSBLOK,  " & _
			"SUM(HASILUN_HI) HASILUN_HI, SUM(HASILUN_SHI) HASILUN_SHI, " & _
			"SUM(HASILHA_HI) HASILHA_HI, SUM(HASILHA_SHI) HASILHA_SHI, " & _
			"SUM(HASILBO_HI) HASILBO_HI, SUM(HASILBO_SHI) HASILBO_SHI, " & _
			"SUM(HASILOT_HI) HASILOT_HI, SUM(HASILOT_SHI) HASILOT_SHI, " & _
			"SUM(HK_KBT_HI) HK_KBT_HI, SUM(HK_KBT_SHI) as HK_KBT_SHI, " & _
			"SUM(HK_KHT_HI) HK_KHT_HI, SUM(HK_KHT_SHI) as HK_KHT_SHI, " & _
			"SUM(HK_KHL_HI) HK_KHL_HI, SUM(HK_KHL_SHI) as HK_KHL_SHI, " & _
			"SUM(HK_TOT_HI) HK_TOT_HI, SUM(HK_TOT_SHI) as HK_TOT_SHI, " & _
			"SUM(GAJI_KBT_HI) Gaji_KBT_HI, SUM(GAJI_KBT_SHI) as Gaji_KBT_SHI, " & _
			"SUM(GAJI_KHT_HI) Gaji_KHT_HI, SUM(GAJI_KHT_SHI) as Gaji_KHT_SHI, " & _
			"SUM(GAJI_KHL_HI) Gaji_KHL_HI, SUM(GAJI_KHL_SHI) as Gaji_KHL_SHI, " & _
			"SUM(GAJI_PM_HI) GAJI_PM_HI, SUM(GAJI_PM_SHI) as GAJI_PM_SHI, " & _
			"SUM(GAJI_PMPNN_HI) GAJI_PMPNN_HI, SUM(GAJI_PMPNN_SHI) as GAJI_PMPNN_SHI, " & _
			"SUM(GAJI_DND_HI) GAJI_DND_HI, SUM(GAJI_DND_SHI) as GAJI_DND_SHI, " & _
			"SUM(Gaji_OT_HI) Gaji_OT_HI, SUM(Gaji_OT_SHI) as Gaji_OT_SHI, " & _
			"SUM(GAJI_BO_HI) Gaji_BO_HI, SUM(GAJI_BO_SHI) as Gaji_BO_SHI,  " & _
			"SUM(GAJI_KBT_HI)+SUM(GAJI_KHT_HI)+SUM(GAJI_KHL_HI)+SUM(GAJI_PM_HI)+SUM(GAJI_BO_HI)+ SUM(Gaji_OT_HI)+SUM(GAJI_PMPNN_HI)-SUM(GAJI_DND_HI) Gaji_TOT_HI, " & _
			"SUM(GAJI_KBT_SHI)+SUM(GAJI_KHT_SHI)+SUM(GAJI_KHL_SHI)+SUM(GAJI_PM_SHI)+SUM(GAJI_BO_SHI)+SUM(Gaji_OT_SHI)+SUM(GAJI_PMPNN_SHI)-SUM(GAJI_DND_SHI)  Gaji_TOT_SHI, " & _
			"KODESTOK, NAMASTOK,  " & _
			"SUM(QtyStok_Hi) QtyStok_Hi, SUM(QtyStok_SHI) as QtyStok_SHI, " & _
			"SUM(AmtStok_Hi) AmtStok_Hi, SUM(AmtStok_SHI) as AmtStok_SHI, " & _
			"vehiclecode,SUM(QtyKM_Hi) QtyKM_Hi, SUM(QtyKM_SHi) QtyKM_SHi " & _
			"	FROM " & _
			"		( " & _
			"		select " & _
			"		FCBA,  KODEAKUN, NAMAAKUN, " & _
			"		   DIVISI, BLOK, TAHUNTANAM, STATUSBLOK, LUASBLOK, " & _
			"		   0 HASILUN_HI, 0 HASILHA_HI, 0 HASILBO_HI, 0 HASILOT_HI, " & _
			"		   0 HK_KBT_HI,0 HK_KHT_HI, 0 HK_KHL_HI, 0 HK_TOT_HI, " & _
			"		   0 Gaji_KBT_HI,0 Gaji_KHT_HI, 0 Gaji_KHL_HI, 0 Gaji_BO_hi, 0 GAJI_PM_HI, 0 GAJI_OT_HI, 0 GAJI_PMPNN_HI,0 GAJI_DND_HI," & _
			"		   SUM(HASILUN_HI) HASILUN_SHI,SUM(HASILHA_HI) as HASILHA_SHI, SUM(HASILBO_HI) as HASILBO_SHI,  SUM(HASILOT_HI) as HASILOT_SHI, " & _
			"		   SUM(HK_KBT_HI) HK_KBT_SHI,SUM(HK_KHT_HI) as HK_KHT_SHI, SUM(HK_KHL_HI) as HK_KHL_SHI, SUM(HK_TOT_HI) as HK_TOT_SHI, " & _
			"		   SUM(GAJI_KBT_HI) as Gaji_KBT_SHI,SUM(GAJI_KHT_HI) as Gaji_KHT_SHI, SUM(GAJI_KHL_HI) as Gaji_KHL_SHI, SUM(GAJI_BO_HI) as Gaji_BO_SHI , SUM(GAJI_PM_HI) as GAJI_PM_SHI,SUM(GAJI_OT_HI) as Gaji_OT_SHI," & _
			"          SUM(GAJI_PMPNN_HI) as GAJI_PMPNN_SHI,SUM(GAJI_DND_HI) as Gaji_DND_SHI, " & _
			"          KODESTOK, NAMASTOK, " & _
			"          0 QtyStok_Hi, 0 AmtStok_Hi, " & _
			"          SUM(QtyStok_Hi) QtyStok_SHI, SUM(AmtStok_Hi) AmtStok_SHI,  " & _
			"          vehiclecode, 0 QtyKM_Hi, SUM(QtyKM_Hi) QtyKM_SHi " & _
			"		from VIEW_LHD_ALL where  fcba='"& v_txtFcba &"' " & _
			"            AND TO_CHAR(fddate, 'yyyymm') = '" & v_txtYear & "'||LPAD('" & v_txtMonth & "',2,0)" & _
            "            AND fddate <= TO_DATE('" & v_txtDate & "', 'mm/dd/yyyy')" & _
			"		group by " & _
			"		FCBA,  KODEAKUN, NAMAAKUN," & _
			"		   DIVISI, BLOK, TAHUNTANAM, STATUSBLOK, LUASBLOK, KODESTOK, NAMASTOK, vehiclecode" & _
			"		union " & _
			"		select " & _
			"		FCBA,  KODEAKUN, NAMAAKUN," & _
			"		   DIVISI, BLOK, TAHUNTANAM, STATUSBLOK, LUASBLOK, " & _
			"		   SUM(HASILUN_HI) HASILUN_HI, SUM(HASILHA_HI) HASILHA_HI, SUM(HASILBO_HI) HASILBO_HI, SUM(HASILOT_HI) HASILOT_HI, " & _
			"		   SUM(HK_KBT_HI) HK_KBT_HI,SUM(HK_KHT_HI) HK_KHT_HI, SUM(HK_KHL_HI) HK_KHL_HI, SUM(HK_TOT_HI) HK_TOT_HI, " & _
			"		   SUM(GAJI_KBT_HI) Gaji_KBT_HI,SUM(GAJI_KHT_HI) Gaji_KHT_HI, SUM(GAJI_KHL_HI) Gaji_KHL_HI, SUM(GAJI_BO_HI) Gaji_BO_hi, SUM(GAJI_PM_HI) GAJI_PM_HI,SUM(GAJI_OT_HI) GAJI_OT_HI ," & _
			"          SUM(GAJI_PMPNN_HI) GAJI_PMPNN_HI,SUM(GAJI_DND_HI) Gaji_DND_HI, " & _
			"		   0 HASILUN_SHI,0 HASILHA_SHI, 0 HASILBO_SHI, 0 HASILOT_SHI, " & _
			"		   0 HK_KBT_SHI,0 HK_KHT_SHI, 0 HK_KHL_SHI, 0 HK_TOT_SHI, " & _
			"		   0 Gaji_KBT_SHI,0 Gaji_KHT_SHI, 0 Gaji_KHL_SHI, 0 Gaji_BO_SHI, 0 GAJI_PM_SHI,0 Gaji_OT_SHI, " & _
			"          0 GAJI_PMPNN_SHI, 0 Gaji_DND_SHI, " & _
			"          KODESTOK, NAMASTOK, " & _
			"          SUM(QtyStok_Hi) QtyStok_Hi, SUM(AmtStok_Hi) AmtStok_Hi, " & _
			"          0 QtyStok_SHI, 0 AmtStok_SHI,  " & _
			"          vehiclecode, SUM(QtyKM_Hi) QtyKM_Hi, 0 QtyKM_SHi " & _
			"		from VIEW_LHD_ALL where  fcba='"& v_txtFcba &"' " & _
			"        and fddate = TO_DATE('" & v_txtDate & "', 'mm/dd/yyyy')" & _
			"		group by  " & _
			"		FCBA,  KODEAKUN, NAMAAKUN, " & _
			"		   DIVISI, BLOK, TAHUNTANAM, STATUSBLOK, LUASBLOK, KODESTOK, NAMASTOK,vehiclecode " & _
			"		)X " & _
			"   WHERE 1=1 "
			

if "" & v_txtDivisi <> "" then
		strMain = strMain & " AND DIVISI = '" & v_txtDivisi & "'"
end if


if "" & v_txtJobCode <> "" then
		strMain = strMain & " AND KODEAKUN = '" & v_txtJobCode & "'"
end if

strMain = strMain & "	GROUP BY " & _
			"	FCBA,  KODEAKUN, NAMAAKUN,DIVISI, BLOK, TAHUNTANAM, STATUSBLOK, LUASBLOK, KODESTOK, NAMASTOK,VEHICLECODE " & _
			"	ORDER BY KODEAKUN,DIVISI,BLOK,KODESTOK "

sqlLHPP = strMain
		  
' response.write sqlLHPP
%>


<html>

	<head>
		<title>ePlaS Reporting - Laporan Harian HPP Kebun</title>
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
							<td><font face="verdana" size="1">TAHUN</font></td>
							<td><font face="verdana" size="1">BULAN</font></td>
							<td><font face="verdana" size="1">TANGGAL</font></td>
							<td><font face="verdana" size="1">DIVISI</font></td>
							<!--
							<td><font face="verdana" size="1">JOBGROUP</font></td>
							-->
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
							<td align="center">
								<input class="datepicker" type="text" maxlength="10" name="txtDate" size="10" value="<%=v_txtDate%>" title="Type Date mm/dd/yyyy" >
							</td>
							<td>
								<select name="txtDivisi" >
								<Option Value="" <%If v_txtDivisi = "" then%> Selected <%End If%>>--ALL--</option>
							<%
							Dim rsDivisi, strDivisi
							strDivisi   =    " SELECT DISTINCT emp.sectionname as fccode, sc.fcname" & _
									         "    FROM employee emp" & _
									         "         INNER JOIN sectionname sc ON (emp.sectionname = sc.fccode)" & _
									         "   WHERE fcba = '"& v_txtFcba &"' ORDER BY emp.sectionname "
						
						   

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
								<select name="txtJobCode" >
									<option value="">--ALL--</option>
								<%
								Dim rsJobCode, strJobCode
								strJobCode	=	" SELECT fccode, fcname FROM job WHERE fcba = '" & v_txtFcba & "' AND job_category IN ('FF', 'MN')ORDER BY fccode ASC"
								
								Set rsJobCode = Server.CreateObject("ADODB.Recordset")
								rsJobCode.Open strJobCode, connstr, adOpenStatic 
								
								if not rsJobCode.eof then
								Do While Not rsJobCode.eof
								if rsJobCode("FCCODE") = v_txtJobCode then%>
										<option value="<%=rsJobCode("FCCODE")%>" Selected><%=rsJobCode("FCCODE")%> - <%=rsJobCode("FCNAME")%></option> 
									 <%Else%>
										<option value="<%=rsJobCode("FCCODE")%>"><%=rsJobCode("FCCODE")%> - <%=rsJobCode("FCNAME")%></option>  
									<%End if
									rsJobCode.movenext
								loop
								end if
								rsJobCode.Close
								set rsJobCode = Nothing
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
						<th colspan="<%=45%>" style="background:#ffffff; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
							<div id="header-wrapper">
								<div id="subheader" class="centered">
									<div style="margin-botton:10px;"/>
									<h3>LAPORAN HARIAN DIVISI PANEN & PERAWATAN
									</h3>
									<br/>
								</div>
								<div style="text-align:left;">
									<h4>Periode : <%=Day(v_txtDate) & " " & Bulan(v_txtMonth) & " " & v_txtYear%></h4>
									<h4 style="margin-top:5px;"><%=PCase(fcbaname)%></h4>
								</div>
							</div>
						</th>
					</tr>
					<tr>
						<th style="background-color:#00a300;color:#fff;" rowspan="3">Kode Perkiraan</th>
						<th style="background-color:#00a300;color:#fff;" rowspan="3">Nama Perkiraan</th>
						<th style="background-color:#00a300;color:#fff;" colspan="5">Lokasi</th>
						<th style="background-color:#00a300;color:#fff;" colspan="8">Jumlah</th>
						<th style="background-color:#00a300;color:#fff;" colspan="8">Tenaga Kerja</th>
						<th style="background-color:#00a300;color:#fff;" colspan="14">Biaya (Rp)</th>
						<th style="background-color:#00a300;color:#fff;" colspan="6">Material</th>
						<th style="background-color:#00a300;color:#fff;" colspan="6">Kendaraan</th>
						
					</tr>
					<tr>
						<th style="background-color:#00a300;color:#fff;" rowspan="2" >Divisi</th>
						<th style="background-color:#00a300;color:#fff;" rowspan="2">Blok</th>
						<th style="background-color:#00a300;color:#fff;" rowspan="2">TT</th>
						<th style="background-color:#00a300;color:#fff;" rowspan="2">Luas</th>
						<th style="background-color:#00a300;color:#fff;" rowspan="2">Status</th>
						
						<th style="background-color:#00a300;color:#fff;" colspan="2">Unit</th>
						<th style="background-color:#00a300;color:#fff;" colspan="2">HA</th>
						<th style="background-color:#00a300;color:#fff;" colspan="2">Borongan</th>
						<th style="background-color:#00a300;color:#fff;" colspan="2">Lembur</th>
						
						
						<th style="background-color:#00a300;color:#fff;" colspan="2">KBT</th>
						<th style="background-color:#00a300;color:#fff;" colspan="2">KHT</th>
						<th style="background-color:#00a300;color:#fff;" colspan="2">KHL</th>
						<th style="background-color:#00a300;color:#fff;" colspan="2">TOT</th>
						
						<th style="background-color:#00a300;color:#fff;" colspan="2">KBT</th>
						<th style="background-color:#00a300;color:#fff;" colspan="2">KHT</th>
						<th style="background-color:#00a300;color:#fff;" colspan="2">KHL</th>
						<th style="background-color:#00a300;color:#fff;" colspan="2">Premi</th>
						<th style="background-color:#00a300;color:#fff;" colspan="2">Premi Panen</th>
						<th style="background-color:#00a300;color:#fff;" colspan="2">Denda</th>
						<th style="background-color:#00a300;color:#fff;" colspan="2">Lembur</th>
						<th style="background-color:#00a300;color:#fff;" colspan="2">Borongan</th>
						<th style="background-color:#00a300;color:#fff;" colspan="2">TOTAL</th>
						
						<th style="background-color:#00a300;color:#fff;" rowspan="2">Kode</th>
						<th style="background-color:#00a300;color:#fff;" rowspan="2">Nama</th>
						<th style="background-color:#00a300;color:#fff;" colspan="2">Jumlah</th>
						<th style="background-color:#00a300;color:#fff;" colspan="2">Biaya</th>
						
						<th style="background-color:#00a300;color:#fff;" rowspan="2">Kode</th>
						<th style="background-color:#00a300;color:#fff;" colspan="2">HM/KM</th>
						
						
					</tr>
					<tr>
					    <!-- JUMLAH -->
					    <th style="background-color:#00a300;color:#fff;">Hi</th>
						<th style="background-color:#00a300;color:#fff;">SHi</th>
						<th style="background-color:#00a300;color:#fff;">Hi</th>
						<th style="background-color:#00a300;color:#fff;">SHi</th>
						<th style="background-color:#00a300;color:#fff;">Hi</th>
						<th style="background-color:#00a300;color:#fff;">SHi</th>
						<th style="background-color:#00a300;color:#fff;">Hi</th>
						<th style="background-color:#00a300;color:#fff;">SHi</th>
						<!-- Tenaga Kerja -->
						<th style="background-color:#00a300;color:#fff;">Hi</th>
						<th style="background-color:#00a300;color:#fff;">SHi</th>
						<th style="background-color:#00a300;color:#fff;">Hi</th>
						<th style="background-color:#00a300;color:#fff;">SHi</th>
						<th style="background-color:#00a300;color:#fff;">Hi</th>
						<th style="background-color:#00a300;color:#fff;">SHi</th>
						<th style="background-color:#00a300;color:#fff;">Hi</th>
						<th style="background-color:#00a300;color:#fff;">SHi</th>
						<!-- Biaya -->
						<th style="background-color:#00a300;color:#fff;">Hi</th>
						<th style="background-color:#00a300;color:#fff;">SHi</th>
						<th style="background-color:#00a300;color:#fff;">Hi</th>
						<th style="background-color:#00a300;color:#fff;">SHi</th>
						<th style="background-color:#00a300;color:#fff;">Hi</th>
						<th style="background-color:#00a300;color:#fff;">SHi</th>
						<th style="background-color:#00a300;color:#fff;">Hi</th>
						<th style="background-color:#00a300;color:#fff;">SHi</th>
						<th style="background-color:#00a300;color:#fff;">Hi</th>
						<th style="background-color:#00a300;color:#fff;">SHi</th>
						<th style="background-color:#00a300;color:#fff;">Hi</th>
						<th style="background-color:#00a300;color:#fff;">SHi</th>
						<th style="background-color:#00a300;color:#fff;">Hi</th>
						<th style="background-color:#00a300;color:#fff;">SHi</th>
						<th style="background-color:#00a300;color:#fff;">Hi</th>
						<th style="background-color:#00a300;color:#fff;">SHi</th>
						<th style="background-color:#00a300;color:#fff;">Hi</th>
						<th style="background-color:#00a300;color:#fff;">SHi</th>
						
						<!-- Barang -->
						<th style="background-color:#00a300;color:#fff;">Hi</th>
						<th style="background-color:#00a300;color:#fff;">SHi</th>
						<th style="background-color:#00a300;color:#fff;">Hi</th>
						<th style="background-color:#00a300;color:#fff;">SHi</th>
						<!-- Kendaraan -->
						<th style="background-color:#00a300;color:#fff;">Hi</th>
						<th style="background-color:#00a300;color:#fff;">SHi</th>
						
					</tr>
				</thead>
				
				<tbody>
<%

Function search_span(data_array__, col__, where__)
	result_ = 0
	arr_1 = data_array__
	for indx = LBound(arr_1, 2) to UBound(arr_1, 2)
		if arr_1(col__, indx) = where__ then
			result_ = result_ + 1
		end if
	next
	search_span = result_
End Function

'response.write sqlLHPP


Set rsLHPP = Server.CreateObject("ADODB.Recordset")
rsLHPP.Open sqlLHPP, connstr, adOpenStatic


if not rsLHPP.EOF then
	array_data = rsLHPP.getrows()
	array_eof = false
end if
rsLHPP.Close
set rsLHPP = Nothing
%>
<%if not array_eof then
	gt_sku = 0
	gt_khl = 0
	gt_unit_daily = 0
	gt_unit_contract = 0
	gt_sku_1 = 0
	gt_khl_1 = 0
	gt_unit_daily_1 = 0
	gt_unit_contract_1 = 0
	gt_sku_2 = 0
	gt_khl_2 = 0
	gt_unit_daily_2 = 0
	gt_unit_contract_2 = 0
	gt_bud_hk = 0
	gt_bud_ha = 0

	For xx = LBound(array_data, 2) To UBound(array_data, 2)
	
		curr_job = array_data(1, xx)
	
			r_hk_hi = 0
			r_hk_shi = 0
			r_hk_sbi = 0
			
			r_ha_hi = 0
			r_ha_shi = 0
			r_ha_sbi = 0
			
			'r_hk_hi = Cdbl(array_data(4, xx)) + Cdbl(array_data(5, xx))
			'r_hk_shi = Cdbl(array_data(8, xx)) + Cdbl(array_data(9, xx))
			'r_hk_sbi = Cdbl(array_data(12, xx)) + Cdbl(array_data(13, xx))
			
			'r_ha_hi = Cdbl(array_data(6, xx)) + Cdbl(array_data(7, xx))
			'r_ha_shi = Cdbl(array_data(10, xx)) + Cdbl(array_data(11, xx))
			'r_ha_sbi = Cdbl(array_data(14, xx)) + Cdbl(array_data(15, xx))
	
	
	%>
					<tr>
						<td align="Left"><%=array_data(1, xx)%></td>
						<td align="Left"><%=array_data(2, xx)%></td>
						<td align="Left"><%=array_data(3, xx)%></td>
						<td align="Left"><%=array_data(4, xx)%></td>
						<td align="Left"><%=array_data(5, xx)%></td>
						<td align="right"><%=FormatNumber(array_data(6, xx), 2)%></td>
						<td align="Left"><%=array_data(7, xx)%></td>
						
						<!-- Jumlah -->
						<td align="right"><%=FormatNumber(array_data(8, xx), 2)%></td>
						<td align="right"><%=FormatNumber(array_data(9, xx), 2)%></td>
						<td align="right"><%=FormatNumber(array_data(10, xx), 2)%></td>
						<td align="right"><%=FormatNumber(array_data(11, xx), 2)%></td>
						<td align="right"><%=FormatNumber(array_data(12, xx), 2)%></td>
						<td align="right"><%=FormatNumber(array_data(13, xx), 2)%></td>
						<td align="right"><%=FormatNumber(array_data(14, xx), 2)%></td>
						<td align="right"><%=FormatNumber(array_data(15, xx), 2)%></td>
						
						<!-- Tenaga Kerja -->
						<td align="right"><%=FormatNumber(array_data(16, xx), 2)%></td>
						<td align="right"><%=FormatNumber(array_data(17, xx), 2)%></td>
						<td align="right"><%=FormatNumber(array_data(18, xx), 2)%></td>
						<td align="right"><%=FormatNumber(array_data(19, xx), 2)%></td>
						<td align="right"><%=FormatNumber(array_data(20, xx), 2)%></td>
						<td align="right"><%=FormatNumber(array_data(21, xx), 2)%></td>
						<td align="right"><%=FormatNumber(array_data(22, xx), 0)%></td>
						<td align="right"><%=FormatNumber(array_data(23, xx), 0)%></td>
						<!--  Biaya -->
						<td align="right"><%=FormatNumber(array_data(24, xx), 0)%></td>
						<td align="right"><%=FormatNumber(array_data(25, xx), 0)%></td>
						<td align="right"><%=FormatNumber(array_data(26, xx), 0)%></td>
						<td align="right"><%=FormatNumber(array_data(27, xx), 0)%></td>
						<td align="right"><%=FormatNumber(array_data(28, xx), 0)%></td>
						<td align="right"><%=FormatNumber(array_data(29, xx), 0)%></td>
						<td align="right"><%=FormatNumber(array_data(30, xx), 0)%></td>
						<td align="right"><%=FormatNumber(array_data(31, xx), 0)%></td>
						<td align="right"><%=FormatNumber(array_data(32, xx), 0)%></td>
						<td align="right"><%=FormatNumber(array_data(33, xx), 0)%></td>
						<td align="right"><%=FormatNumber(array_data(34, xx), 0)%></td>
						<td align="right"><%=FormatNumber(array_data(35, xx), 0)%></td>
						<td align="right"><%=FormatNumber(array_data(36, xx), 0)%></td>
						<td align="right"><%=FormatNumber(array_data(37, xx), 0)%></td>
						<td align="right"><%=FormatNumber(array_data(38, xx), 0)%></td>
						<td align="right"><%=FormatNumber(array_data(39, xx), 0)%></td>
						<td align="right"><%=FormatNumber(array_data(40, xx), 0)%></td>
						<td align="right"><%=FormatNumber(array_data(41, xx), 0)%></td>
						
						<td align="right"><%=array_data(42, xx)%></td>
						<td align="right"><%=array_data(43, xx)%></td>
						<td align="right"><%=FormatNumber(array_data(44, xx), 0)%></td>
						<td align="right"><%=FormatNumber(array_data(45, xx), 0)%></td>
						<td align="right"><%=FormatNumber(array_data(46, xx), 0)%></td>
						<td align="right"><%=FormatNumber(array_data(47, xx), 0)%></td>
						
						<td align="right"><%=array_data(48, xx)%></td>
						<td align="right"><%=FormatNumber(array_data(49, xx), 0)%></td>
						<td align="right"><%=FormatNumber(array_data(50, xx), 0)%></td>
						
					</tr>
					
<%
		'gt_sku = gt_sku + Cdbl(array_data(4, xx))
		'gt_khl = gt_khl + Cdbl(array_data(5, xx))
		'gt_unit_daily = gt_unit_daily + Cdbl(array_data(6, xx))
		'gt_unit_contract = gt_unit_contract + Cdbl(array_data(7, xx))
		'gt_sku_1 = gt_sku_1 + Cdbl(array_data(8, xx))
		'gt_khl_1 = gt_khl_1 + Cdbl(array_data(9, xx))
		'gt_unit_daily_1 = gt_unit_daily_1 + Cdbl(array_data(10, xx))
		'gt_unit_contract_1 = gt_unit_contract_1 + Cdbl(array_data(11, xx))
		'gt_sku_2 = gt_sku_2 + Cdbl(array_data(12, xx))
		'gt_khl_2 = gt_khl_2 + Cdbl(array_data(13, xx))
		'gt_unit_daily_2 = gt_unit_daily_2 + Cdbl(array_data(14, xx))
		'gt_unit_contract_2 = gt_unit_contract_2 + Cdbl(array_data(15, xx))
		'gt_bud_hk = gt_bud_hk + Cdbl(array_data(16, xx))
		'gt_bud_ha = gt_bud_ha + Cdbl(array_data(17, xx))
		
		prev_job = curr_job

	next
	
	gt_hk_hi = 0
	gt_hk_shi = 0
	gt_hk_sbi = 0

	gt_ha_hi = 0
	gt_ha_shi = 0
	gt_ha_sbi = 0

	'gt_hk_hi = gt_sku + gt_khl
	'gt_hk_shi = gt_sku_1 + gt_khl_1
	'gt_hk_sbi = gt_sku_2 + gt_khl_2

	'gt_ha_hi = gt_unit_daily + gt_unit_contract
	'gt_ha_shi = gt_unit_daily_1 + gt_unit_contract_1
	'gt_ha_sbi = gt_unit_daily_2 + gt_unit_contract_2
%>
					<tr><td colspan="25"></td></tr>
					<tr style="background-color:#f4f4a0;">					
						<td align="center" colspan="5"><b>Total</b></td>
						<td align="right"><b><%=FormatNumber(gt_sku, 0)%></b></td>
						<td align="right"><b><%=FormatNumber(gt_khl, 0)%></b></td>
						<td align="right"><b><%=FormatNumber(gt_hk_hi, 0)%></b></td>
						<td align="right"><b><%=FormatNumber(gt_unit_daily, 0)%></b></td>
						<td align="right"><b><%=FormatNumber(gt_unit_contract, 0)%></b></td>
						<td align="right"><b><%=FormatNumber(gt_ha_hi, 0)%></b></td>

						<%if gt_ha_hi = 0 then%>
							<td align="right"><b><%=FormatNumber(0, 2)%></b></td>
						<%else%>
							<td align="right"><b><%=FormatNumber(gt_hk_hi/gt_ha_hi, 2)%></b></td>
						<%end if%>
						
						<td align="right"><b><%=FormatNumber(gt_sku_1, 0)%></b></td>
						<td align="right"><b><%=FormatNumber(gt_khl_1, 0)%></b></td>
						<td align="right"><b><%=FormatNumber(gt_hk_shi, 0)%></b></td>
						<td align="right"><b><%=FormatNumber(gt_unit_daily_1, 0)%></b></td>
						<td align="right"><b><%=FormatNumber(gt_unit_contract_1, 0)%></b></td>
						<td align="right"><b><%=FormatNumber(gt_ha_shi, 0)%></b></td>
						
						<%if gt_ha_shi = 0 then%>
							<td align="right"><b><%=FormatNumber(0, 2)%></b></td>
						<%else%>
							<td align="right"><b><%=FormatNumber(gt_hk_shi/gt_ha_shi, 2)%></b></td>
						<%end if%>
						
						<td align="right"><b><%=FormatNumber(gt_sku_2, 0)%></b></td>
						<td align="right"><b><%=FormatNumber(gt_khl_2, 0)%></b></td>
						<td align="right"><b><%=FormatNumber(gt_hk_sbi, 0)%></b></td>
						<td align="right"><b><%=FormatNumber(gt_unit_daily_2, 0)%></b></td>
						<td align="right"><b><%=FormatNumber(gt_unit_contract_2, 0)%></b></td>
						<td align="right"><b><%=FormatNumber(gt_ha_sbi, 0)%></b></td>
						
						<%if gt_ha_sbi = 0 then%>
							<td align="right"><b><%=FormatNumber(0, 2)%></b></td>
						<%else%>
							<td align="right"><b><%=FormatNumber(gt_hk_sbi/gt_ha_sbi, 2)%></b></td>
						<%end if%>
						
						<%if gt_bud_ha = 0 then%>
							<td align="right"><b><%=FormatNumber(0, 2)%></b></td>
						<%else%>
							<td align="right"><b><%=FormatNumber(gt_bud_hk/gt_bud_ha, 2)%></b></td>
						<%end if%>
					
					</tr>
<%else%>
					<tr>
						<td colspan="<%=25%>">NO DATA AVAILABLE</td>
					</tr>
<%end if%>
				</tbody>
			</table>

		</div>
	</body>
</html>
