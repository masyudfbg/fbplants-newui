<%@Language=VBScript%>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
	Dim v_txtBudgetYear, v_txtBudgetVersion, v_txtModuleVersion, v_txtdatefrom, v_txtdateto, v_txtMonth, v_txtYear, v_txtgrade, v_txtDivisi, v_txtemployeetype
	Dim no
	Dim rsDefaultBa, strDefaultBa, varDefaultBa
	Dim ScriptName
	Dim strHeader,strHeader2,strMain, sqlLHPP, sqlLHPP2, rsLHPP
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
	Dim array_data, array_data_sum,array_eof, xx,xx_sum

	' **** Sub Total (Kebun)
	Dim rsST, strST, arrayST, arrayST_eof, iii, st_indx
	
	Dim row_st_total_bi : row_st_total_bi = 0
	Dim row_st_total_sbi : row_st_total_sbi = 0
	
	arrayST_eof = true
	' ****End Sub Total (Kebun)
	
	Dim rowspan_division : rowspan_division = 0
	Dim rsRowspan, strRowspan
	Dim posting_rowspan : posting_rowspan = 0

    Function IsPostBack()
     If Request.ServerVariables("Request_Method") = "POST" Then
          IsPostBack = True
     Else
          IsPostBack = False
     End If
	End Function
	
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

    ScriptName = Request.ServerVariables("SCRIPT_NAME")
	v_txtfcba = Request("txtfcba")
	v_txtDivisi = Request("txtDivisi")
	
	v_txtBudgetYear = Request("txtBudgetYear")
	v_txtBudgetVersion = Request("txtBudgetVersion")
	v_txtModuleVersion = Request("txtModuleVersion")
	
	v_txtdatefrom	   = Request("txtdatefrom")
	v_txtdateto		   = Request("txtdateto")
	
	
	v_txtYear = Request("txtYear")
	v_txtMonth = Request("txtMonth")
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
	If Len(Trim(v_txtdateto))   = 0 Then v_txtdateto = textDate(Date)
	
	'if "" & v_txtdateto = "" then v_txtDate = textDate(Date)
	v_txtYear = Year(v_txtdateto)
	v_txtMonth = Month(v_txtdateto)
	
	
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





strHeader = "SELECT  " & _
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
			"vehiclecode,SUM(QtyKM_Hi) QtyKM_Hi, SUM(QtyKM_SHi) QtyKM_SHi " 

strHeader2 = "SELECT  " & _
			"FCBA,  KODEAKUN, NAMAAKUN, " & _
			"DIVISI, STATUSBLOK,  " & _
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
			"vehiclecode,SUM(QtyKM_Hi) QtyKM_Hi, SUM(QtyKM_SHi) QtyKM_SHi " 	
			
strMain =   "	FROM " & _
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
			"            AND fddate between TO_DATE('" & v_txtDatefrom & "', 'mm/dd/yyyy') AND TO_DATE('" & v_txtDateto & "', 'mm/dd/yyyy')" & _
			"		and KODEAKUN not in('1114300','1114400') group by " & _
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
			"            AND fddate between TO_DATE('" & v_txtDatefrom & "', 'mm/dd/yyyy') AND TO_DATE('" & v_txtDateto & "', 'mm/dd/yyyy')" & _
			"		and KODEAKUN not in('1114300','1114400') group by  " & _
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

sqlLHPP = strHeader & strMain & "	GROUP BY " & _
			"	FCBA,  KODEAKUN, NAMAAKUN,DIVISI, BLOK, TAHUNTANAM, STATUSBLOK, LUASBLOK, KODESTOK, NAMASTOK,VEHICLECODE " & _
			"	ORDER BY KODEAKUN,BLOK "

sqlLHPP2 = strHeader2 & strMain & "	GROUP BY " & _
			"	FCBA,  KODEAKUN, NAMAAKUN,DIVISI, STATUSBLOK, KODESTOK, NAMASTOK,VEHICLECODE " & _
			"	ORDER BY KODEAKUN,DIVISI,KODESTOK,VEHICLECODE "


Response.Write sqlLHPP

%>


<html>

	<head>
		<title>ePlaS Reporting - Laporan Harian Divisi Perawatan</title>
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
							<td><font face="verdana" size="1">TANGGAL</font></td>
							<td><font face="verdana" size="1"></font></td>
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
							<td align="center">
								<input class="datepicker" type="text" maxlength="10" name="txtdatefrom" size="10" value="<%=v_txtdatefrom%>" title="Type Date mm/dd/yyyy" >
							</td>
							<td align="center">
								<input class="datepicker" type="text" maxlength="10" name="txtdateto" size="10" value="<%=v_txtdateto%>" title="Type Date mm/dd/yyyy" >
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
						<th colspan="<%=50%>" style="background:#ffffff; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
							<div id="header-wrapper">
								<div id="subheader" class="centered">
									<div style="margin-botton:10px;"/>
									<h3>LAPORAN HARIAN DIVISI
									</h3>
									<br/>
								</div>
								<div style="text-align:left;">
									<h4>Periode : <%=v_txtDatefrom & " s/d " & v_txtdateto %></h4>
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
						<th style="background-color:#00a300;color:#fff;" colspan="18">Biaya (Rp)</th>
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



array_eof = true

if IsPostBack then
Set rsLHPP = Server.CreateObject("ADODB.Recordset")
rsLHPP.Open sqlLHPP, connstr, adOpenStatic


if not rsLHPP.EOF then
	array_data = rsLHPP.getrows()
	array_eof = false
end if
rsLHPP.Close
set rsLHPP = Nothing

Set rsLHPP = Server.CreateObject("ADODB.Recordset")
rsLHPP.Open sqlLHPP2, connstr, adOpenStatic


if not rsLHPP.EOF then
	array_data_sum = rsLHPP.getrows()
end if
rsLHPP.Close
set rsLHPP = Nothing

end if

if not array_eof then

Dim luaslokasi,ttlunithi,ttlhahi,ttlhashi,ttlborhi,ttlborshi,ttllemburhi,ttllemburshi
Dim ttlkbthi,ttlkbtshi,ttlkhthi,ttlkhtshi,ttlkhlhi,ttlkhlshi,ttltothi,ttltotshi
Dim ttlkbt1hi,ttlkbt1shi,ttlkht1hi,ttlkht1shi,ttlkhl1hi,ttlkhl1shi,ttlpremihi,ttlpremishi
Dim ttlprepanenhi,ttlprepanenshi,ttldendahi,ttldendashi,ttllembur1hi,ttllembur1shi
Dim ttlbor1hi,ttlbor1shi,ttlbiayahi,ttlbiayashi,ttlmatjumhi,ttlmatjumshi
Dim ttlbiayamathi,ttlbiayamatshi,ttlhmhi,ttlhmshi

	gt_sku = 0
	gt_khl = 0
	ttlunitshi = 0
	ttlhahi = 0
	ttlhashi = 0
	ttlborhi = 0
	ttlborshi = 0
	ttllemburhi = 0
	ttllemburshi = 0
	ttlkbthi = 0
	ttlkbtshi = 0
	ttlkhthi = 0
	ttlkhtshi = 0
	ttlkhlhi = 0
	ttlkhlshi = 0
	ttltothi = 0
	ttltotshi = 0
	ttlkbthi = 0
	ttlkbtshi = 0
	ttlkht1hi = 0
	ttlkht1shi = 0
	ttlkhl1hi = 0
	ttlkhl1shi = 0
	ttlpremihi = 0
	ttlpremishi = 0
	ttlprepanenhi = 0
	ttlprepanenshi = 0
	ttldendahi = 0
	ttldendashi = 0
	ttllembur1hi = 0
	ttllembur1shi = 0
	ttlbor1hi = 0
	ttlbor1shi = 0
	ttlbiayahi = 0
	ttlbiayashi = 0
	ttlmatjumhi = 0
	ttlmatjumshi = 0
	ttlbiayamathi = 0
	ttlbiayamatshi = 0
	ttlhmhi = 0
	ttlhmshi = 0


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

		'luaslokasi = array_data(40, xx)
		'Response.Write luaslokasi
					
			r_hk_hi = 0
			r_hk_shi = 0
			r_hk_sbi = 0
			
			r_ha_hi = 0
			r_ha_shi = 0
			r_ha_sbi = 0
	
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
		gt_sku = gt_sku + Cdbl(array_data(6, xx))
		gt_khl = gt_khl + Cdbl(array_data(8, xx))
		ttlunitshi = ttlunitshi + Cdbl(array_data(9, xx))
		ttlhahi = ttlhahi + Cdbl(array_data(10, xx))
		ttlhashi = ttlhashi + Cdbl(array_data(11, xx))
		ttlborhi = ttlborhi + Cdbl(array_data(12, xx))
		ttlborshi = ttlborshi + Cdbl(array_data(13, xx))
		ttllemburhi = ttllemburhi + Cdbl(array_data(14, xx))
		ttllemburshi = ttllemburshi + Cdbl(array_data(15, xx))
		ttlkbthi = ttlkbthi + Cdbl(array_data(16, xx))
		ttlkbtshi = ttlkbtshi + Cdbl(array_data(17, xx))
		ttlkhthi = ttlkhthi + Cdbl(array_data(18, xx))
		ttlkhtshi = ttlkhtshi + Cdbl(array_data(19, xx))
		ttlkhlhi = ttlkhlhi + Cdbl(array_data(20, xx))
		ttlkhlshi = ttlkhlshi + Cdbl(array_data(21, xx))
		ttltothi = ttltothi + Cdbl(array_data(22, xx))
		ttltotshi = ttltotshi + Cdbl(array_data(23, xx))
		ttlkbt1hi = ttlkbt1hi + Cdbl(array_data(24, xx))
		ttlkbt1shi = ttlkbt1shi + Cdbl(array_data(25, xx))
		ttlkht1hi = ttlkht1hi + Cdbl(array_data(26, xx))
		ttlkht1shi = ttlkht1shi + Cdbl(array_data(27, xx))
		ttlkhl1hi = ttlkhl1hi + Cdbl(array_data(28, xx))
		ttlkhl1shi = ttlkhl1shi + Cdbl(array_data(29, xx))
		ttlpremihi = ttlpremihi + Cdbl(array_data(30, xx))
		ttlpremishi = ttlpremishi + Cdbl(array_data(31, xx))
		ttlprepanenhi = ttlprepanenhi + Cdbl(array_data(32, xx))
		ttlprepanenshi = ttlprepanenshi + Cdbl(array_data(33, xx))
		
		ttldendahi = ttldendahi + Cdbl(array_data(34, xx))
		ttldendashi = ttldendashi + Cdbl(array_data(35, xx))
		ttllembur1hi = ttllembur1hi + Cdbl(array_data(36, xx))
		ttllembur1shi = ttllembur1shi + Cdbl(array_data(37, xx))
		
		ttlbor1hi = ttlbor1hi + Cdbl(array_data(38, xx))
		ttlbor1shi = ttlbor1shi + Cdbl(array_data(39, xx))
		ttlbiayahi = ttlbiayahi + Cdbl(array_data(40, xx))
		ttlbiayashi = ttlbiayashi + Cdbl(array_data(41, xx))
		ttlmatjumhi = ttlmatjumhi + Cdbl(array_data(44, xx))
		ttlmatjumshi = ttlmatjumshi + Cdbl(array_data(45, xx))
		ttlbiayamathi = ttlbiayamathi + Cdbl(array_data(46, xx))
		ttlbiayamatshi = ttlbiayamatshi + Cdbl(array_data(47, xx))
		
		ttlhmhi = ttlhmhi + Cdbl(array_data(49, xx))
		ttlhmshi = ttlhmshi + Cdbl(array_data(50, xx))
		

		
		prev_job = curr_job

	next
	
	gt_hk_hi = 0
	gt_hk_shi = 0
	gt_hk_sbi = 0

	gt_ha_hi = 0
	gt_ha_shi = 0
	gt_ha_sbi = 0

	gt_hk_hi = gt_sku + gt_khl

%>
					<tr><td colspan="25"></td></tr>
					<tr style="background-color:#f4f4a0;">					
						<td align="center" colspan="5"><b>Total</b></td>
						<td align="right"><b><%=FormatNumber(gt_sku, 0)%></b></td>
						<td align="right"><b></b></td>
						<td align="right"><b><%=FormatNumber(gt_khl, 0)%></b></td>
						<td align="right"><b><%=FormatNumber(ttlunitshi, 0)%></b></td>
						<td align="right"><b><%=FormatNumber(ttlhahi, 0)%></b></td>
						<td align="right"><b><%=FormatNumber(ttlhashi, 0)%></b></td>
						<td align="right"><b><%=FormatNumber(ttlborhi, 0)%></b></td>

						
						<td align="right"><b><%=FormatNumber(ttlborshi, 0)%></b></td>
						
						<td align="right"><b><%=FormatNumber(ttllemburhi, 0)%></b></td>
						<td align="right"><b><%=FormatNumber(ttllemburshi, 0)%></b></td>
						<td align="right"><b><%=FormatNumber(ttlkbthi, 0)%></b></td>
						<td align="right"><b><%=FormatNumber(ttlkbtshi, 0)%></b></td>
						<td align="right"><b><%=FormatNumber(ttlkhthi, 0)%></b></td>
						<td align="right"><b><%=FormatNumber(ttlkhtshi, 0)%></b></td>
												
						<td align="right"><b><%=FormatNumber(ttlkhlhi, 0)%></b></td>
						<td align="right"><b><%=FormatNumber(ttlkhlshi, 0)%></b></td>

						<td align="right"><b><%=FormatNumber(ttltothi, 0)%></b></td>
						<td align="right"><b><%=FormatNumber(ttltotshi, 0)%></b></td>

						<td align="right"><b><%=FormatNumber(ttlkbthi, 0)%></b></td>
						<td align="right"><b><%=FormatNumber(ttlkbtshi, 0)%></b></td>
						<td align="right"><b><%=FormatNumber(ttlkht1hi, 0)%></b></td>
						<td align="right"><b><%=FormatNumber(ttlkht1shi, 0)%></b></td>
						<td align="right"><b><%=FormatNumber(ttlkhl1hi, 0)%></b></td>
						<td align="right"><b><%=FormatNumber(ttlkhl1shi, 0)%></b></td>
						<td align="right"><b><%=FormatNumber(ttlpremihi, 0)%></b></td>
						<td align="right"><b><%=FormatNumber(ttlpremishi, 0)%></b></td>
						
						<td align="right"><b><%=FormatNumber(ttlprepanenhi, 0)%></b></td>
						<td align="right"><b><%=FormatNumber(ttlprepanenshi, 0)%></b></td>
						
						<td align="right"><b><%=FormatNumber(ttldendahi, 0)%></b></td>
						<td align="right"><b><%=FormatNumber(ttldendashi, 0)%></b></td>
						
						<td align="right"><b><%=FormatNumber(ttllembur1hi, 0)%></b></td>
						<td align="right"><b><%=FormatNumber(ttllembur1shi, 0)%></b></td>
						<td align="right"><b><%=FormatNumber(ttlbor1hi, 0)%></b></td>
						<td align="right"><b><%=FormatNumber(ttlbor1shi, 0)%></b></td>
						
						<td align="right"><b><%=FormatNumber(ttlbiayahi, 0)%></b></td>
						<td align="right"><b><%=FormatNumber(ttlbiayashi, 0)%></b></td>

						<td align="right"></td>
						<td align="right"></td>

						<td align="right"><b><%=FormatNumber(ttlmatjumhi, 0)%></b></td>
						<td align="right"><b><%=FormatNumber(ttlmatjumshi, 0)%></b></td>
						<td align="right"><b><%=FormatNumber(ttlbiayamathi, 0)%></b></td>
						<td align="right"><b><%=FormatNumber(ttlbiayamatshi, 0)%></b></td>

						<td align="right"></td>

						<td align="right"><b><%=FormatNumber(ttlhmhi, 0)%></b></td>
						<td align="right"><b><%=FormatNumber(ttlhmshi, 0)%></b></td>
						
					
						
					
					</tr>
<%else%>
					<tr>
						<td colspan="<%=25%>">NO DATA AVAILABLE</td>
					</tr>
<%end if%>
				</tbody>
			</table>
 
<!-- SUMMARY LHD -->
 
			<table border="1" width="100%" id="product-table" class='xthf-site'>
				<thead>
					<tr>
						<th colspan="<%=47%>" style="background:#ffffff; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
							<div id="header-wrapper">
								<div id="subheader" class="centered">
									<div style="margin-botton:10px;"/>
									<h3>SUMMARY HARIAN DIVISI PANEN & PERAWATAN
									</h3>
									<br/>
								</div>								
							</div>
						</th>
					</tr>
					<tr>
						<th style="background-color:#00a300;color:#fff;" rowspan="3">Kode Perkiraan</th>
						<th style="background-color:#00a300;color:#fff;" rowspan="3">Nama Perkiraan</th>
						<th style="background-color:#00a300;color:#fff;" colspan="2">Lokasi</th>
						<th style="background-color:#00a300;color:#fff;" colspan="8">Jumlah</th>
						<th style="background-color:#00a300;color:#fff;" colspan="8">Tenaga Kerja</th>
						<th style="background-color:#00a300;color:#fff;" colspan="18">Biaya (Rp)</th>
						<th style="background-color:#00a300;color:#fff;" colspan="6">Material</th>
						<th style="background-color:#00a300;color:#fff;" colspan="6">Kendaraan</th>
						
					</tr>
					<tr>
						<th style="background-color:#00a300;color:#fff;" rowspan="2">Divisi</th>
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
<%if not array_eof then 
gt_sku = 0
	gt_khl = 0
	ttlunitshi = 0
	ttlhahi = 0
	ttlhashi = 0
	ttlborhi = 0
	ttlborshi = 0
	ttllemburhi = 0
	ttllemburshi = 0
	ttlkbthi = 0
	ttlkbtshi = 0
	ttlkhthi = 0
	ttlkhtshi = 0
	ttlkhlhi = 0
	ttlkhlshi = 0
	ttltothi = 0
	ttltotshi = 0
	ttlkbthi = 0
	ttlkbtshi = 0
	ttlkht1hi = 0
	ttlkht1shi = 0
	ttlkhl1hi = 0
	ttlkhl1shi = 0
	ttlpremihi = 0
	ttlpremishi = 0
	ttlprepanenhi = 0
	ttlprepanenshi = 0
	ttldendahi = 0
	ttldendashi = 0
	ttllembur1hi = 0
	ttllembur1shi = 0
	ttlbor1hi = 0
	ttlbor1shi = 0
	ttlbiayahi = 0
	ttlbiayashi = 0
	ttlmatjumhi = 0
	ttlmatjumshi = 0
	ttlbiayamathi = 0
	ttlbiayamatshi = 0
	ttlhmhi = 0
	ttlhmshi = 0


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
	
		For xx_sum = LBound(array_data_sum, 2) To UBound(array_data_sum, 2)
	
	    r_hk_hi = 0
			r_hk_shi = 0
			r_hk_sbi = 0
			
			r_ha_hi = 0
			r_ha_shi = 0
			r_ha_sbi = 0
	

%>
					<tr>
						<td align="Left"><%=array_data_sum(1, xx_sum)%></td>
						<td align="Left"><%=array_data_sum(2, xx_sum)%></td>
						<td align="Left"><%=array_data_sum(3, xx_sum)%></td>
						<td align="Left"><%=array_data_sum(4, xx_sum)%></td>
						
						<!-- Jumlah -->
						<td align="right"><%=FormatNumber(array_data_sum(5, xx_sum), 2)%></td>
						<td align="right"><%=FormatNumber(array_data_sum(6, xx_sum), 2)%></td>
						<td align="right"><%=FormatNumber(array_data_sum(7, xx_sum), 2)%></td>
						<td align="right"><%=FormatNumber(array_data_sum(8, xx_sum), 2)%></td>
						<td align="right"><%=FormatNumber(array_data_sum(9, xx_sum), 2)%></td>
						<td align="right"><%=FormatNumber(array_data_sum(10, xx_sum), 2)%></td>
						<td align="right"><%=FormatNumber(array_data_sum(11, xx_sum), 2)%></td>
						<td align="right"><%=FormatNumber(array_data_sum(12, xx_sum), 2)%></td>
						
						<!-- Tenaga Kerja -->
						<td align="right"><%=FormatNumber(array_data_sum(13, xx_sum), 2)%></td>
						<td align="right"><%=FormatNumber(array_data_sum(14, xx_sum), 2)%></td>
						<td align="right"><%=FormatNumber(array_data_sum(15, xx_sum), 2)%></td>
						<td align="right"><%=FormatNumber(array_data_sum(16, xx_sum), 2)%></td>
						<td align="right"><%=FormatNumber(array_data_sum(17, xx_sum), 2)%></td>
						<td align="right"><%=FormatNumber(array_data_sum(18, xx_sum), 2)%></td>
						<td align="right"><%=FormatNumber(array_data_sum(19, xx_sum), 0)%></td>
						<td align="right"><%=FormatNumber(array_data_sum(20, xx_sum), 0)%></td>
						<!--  Biaya -->
						<td align="right"><%=FormatNumber(array_data_sum(21, xx_sum), 0)%></td>
						<td align="right"><%=FormatNumber(array_data_sum(22, xx_sum), 0)%></td>
						<td align="right"><%=FormatNumber(array_data_sum(23, xx_sum), 0)%></td>
						<td align="right"><%=FormatNumber(array_data_sum(24, xx_sum), 0)%></td>
						<td align="right"><%=FormatNumber(array_data_sum(25, xx_sum), 0)%></td>
						<td align="right"><%=FormatNumber(array_data_sum(26, xx_sum), 0)%></td>
						<td align="right"><%=FormatNumber(array_data_sum(27, xx_sum), 0)%></td>
						<td align="right"><%=FormatNumber(array_data_sum(28, xx_sum), 0)%></td>
						<td align="right"><%=FormatNumber(array_data_sum(29, xx_sum), 0)%></td>
						<td align="right"><%=FormatNumber(array_data_sum(30, xx_sum), 0)%></td>
						<td align="right"><%=FormatNumber(array_data_sum(31, xx_sum), 0)%></td>
						<td align="right"><%=FormatNumber(array_data_sum(32, xx_sum), 0)%></td>
						<td align="right"><%=FormatNumber(array_data_sum(33, xx_sum), 0)%></td>
						<td align="right"><%=FormatNumber(array_data_sum(34, xx_sum), 0)%></td>
						<td align="right"><%=FormatNumber(array_data_sum(35, xx_sum), 0)%></td>
						<td align="right"><%=FormatNumber(array_data_sum(36, xx_sum), 0)%></td>
						<td align="right"><%=FormatNumber(array_data_sum(37, xx_sum), 0)%></td>
						<td align="right"><%=FormatNumber(array_data_sum(38, xx_sum), 0)%></td>
						
						<td align="right"><%=array_data_sum(39, xx_sum)%></td>
						<td align="right"><%=array_data_sum(40, xx_sum)%></td>
						<td align="right"><%=FormatNumber(array_data_sum(41, xx_sum), 0)%></td>
						<td align="right"><%=FormatNumber(array_data_sum(42, xx_sum), 0)%></td>
						<td align="right"><%=FormatNumber(array_data_sum(43, xx_sum), 0)%></td>
						<td align="right"><%=FormatNumber(array_data_sum(44, xx_sum), 0)%></td>
						
						<td align="right"><%=array_data_sum(45, xx_sum)%></td>
						<td align="right"><%=FormatNumber(array_data_sum(46, xx_sum), 0)%></td>
						<td align="right"><%=FormatNumber(array_data_sum(47, xx_sum), 0)%></td>
						
					</tr>
					
<%
		'gt_sku = gt_sku + Cdbl(array_data_sum(3, xx_sum))
		gt_khl = gt_khl + Cdbl(array_data_sum(5, xx_sum))
		ttlunitshi = ttlunitshi + Cdbl(array_data_sum(6, xx_sum))
		ttlhahi = ttlhahi + Cdbl(array_data_sum(7, xx_sum))
		ttlhashi = ttlhashi + Cdbl(array_data_sum(8, xx_sum))
		ttlborhi = ttlborhi + Cdbl(array_data_sum(9, xx_sum))
		ttlborshi = ttlborshi + Cdbl(array_data_sum(10, xx_sum))
		ttllemburhi = ttllemburhi + Cdbl(array_data_sum(11, xx_sum))
		ttllemburshi = ttllemburshi + Cdbl(array_data_sum(12, xx_sum))
		ttlkbthi = ttlkbthi + Cdbl(array_data_sum(13, xx_sum))
		ttlkbtshi = ttlkbtshi + Cdbl(array_data_sum(14, xx_sum))
		ttlkhthi = ttlkhthi + Cdbl(array_data_sum(15, xx_sum))
		ttlkhtshi = ttlkhtshi + Cdbl(array_data_sum(16, xx_sum))
		ttlkhlhi = ttlkhlhi + Cdbl(array_data_sum(17, xx_sum))
		ttlkhlshi = ttlkhlshi + Cdbl(array_data_sum(18, xx_sum))
		ttltothi = ttltothi + Cdbl(array_data_sum(19, xx_sum))
		ttltotshi = ttltotshi + Cdbl(array_data_sum(20, xx_sum))
		ttlkbt1hi = ttlkbt1hi + Cdbl(array_data_sum(21, xx_sum))
		ttlkbt1shi = ttlkbt1shi + Cdbl(array_data_sum(22, xx_sum))
		ttlkht1hi = ttlkht1hi + Cdbl(array_data_sum(23, xx_sum))
		ttlkht1shi = ttlkht1shi + Cdbl(array_data_sum(24, xx_sum))
		ttlkhl1hi = ttlkhl1hi + Cdbl(array_data_sum(25, xx_sum))
		ttlkhl1shi = ttlkhl1shi + Cdbl(array_data_sum(26, xx_sum))
		ttlpremihi = ttlpremihi + Cdbl(array_data_sum(27, xx_sum))
		ttlpremishi = ttlpremishi + Cdbl(array_data_sum(28, xx_sum))
		ttlprepanenhi = ttlprepanenhi + Cdbl(array_data_sum(29, xx_sum))
		ttlprepanenshi = ttlprepanenshi + Cdbl(array_data_sum(30, xx_sum))
		
		ttldendahi = ttldendahi + Cdbl(array_data_sum(31, xx_sum))
		ttldendashi = ttldendashi + Cdbl(array_data_sum(32, xx_sum))
		ttllembur1hi = ttllembur1hi + Cdbl(array_data_sum(33, xx_sum))
		ttllembur1shi = ttllembur1shi + Cdbl(array_data_sum(34, xx_sum))
		
		ttlbor1hi = ttlbor1hi + Cdbl(array_data_sum(35, xx_sum))
		ttlbor1shi = ttlbor1shi + Cdbl(array_data_sum(36, xx_sum))
		ttlbiayahi = ttlbiayahi + Cdbl(array_data_sum(37, xx_sum))
		ttlbiayashi = ttlbiayashi + Cdbl(array_data_sum(38, xx_sum))
		ttlmatjumhi = ttlmatjumhi + Cdbl(array_data_sum(41, xx_sum))
		ttlmatjumshi = ttlmatjumshi + Cdbl(array_data_sum(42, xx_sum))
		ttlbiayamathi = ttlbiayamathi + Cdbl(array_data_sum(43, xx_sum))
		ttlbiayamatshi = ttlbiayamatshi + Cdbl(array_data_sum(44, xx_sum))
		
		ttlhmhi = ttlhmhi + Cdbl(array_data_sum(46, xx_sum))
		ttlhmshi = ttlhmshi + Cdbl(array_data_sum(47, xx_sum))
		

		
		'prev_job = curr_job

	next
	
	gt_hk_hi = 0
	gt_hk_shi = 0
	gt_hk_sbi = 0

	gt_ha_hi = 0
	gt_ha_shi = 0
	gt_ha_sbi = 0

	gt_hk_hi = gt_sku + gt_khl

%>
					<tr><td colspan="25"></td></tr>
					<tr style="background-color:#f4f4a0;">					
						<td align="center" colspan="2"><b>Total</b></td>
						<td align="right"><b></b></td>
						<td align="right"><b></b></td>
						<td align="right"><b><%=FormatNumber(gt_khl, 0)%></b></td>
						<td align="right"><b><%=FormatNumber(ttlunitshi, 0)%></b></td>
						<td align="right"><b><%=FormatNumber(ttlhahi, 0)%></b></td>
						<td align="right"><b><%=FormatNumber(ttlhashi, 0)%></b></td>
						<td align="right"><b><%=FormatNumber(ttlborhi, 0)%></b></td>

						
						<td align="right"><b><%=FormatNumber(ttlborshi, 0)%></b></td>
						
						<td align="right"><b><%=FormatNumber(ttllemburhi, 0)%></b></td>
						<td align="right"><b><%=FormatNumber(ttllemburshi, 0)%></b></td>
						<td align="right"><b><%=FormatNumber(ttlkbthi, 0)%></b></td>
						<td align="right"><b><%=FormatNumber(ttlkbtshi, 0)%></b></td>
						<td align="right"><b><%=FormatNumber(ttlkhthi, 0)%></b></td>
						<td align="right"><b><%=FormatNumber(ttlkhtshi, 0)%></b></td>
												
						<td align="right"><b><%=FormatNumber(ttlkhlhi, 0)%></b></td>
						<td align="right"><b><%=FormatNumber(ttlkhlshi, 0)%></b></td>

						<td align="right"><b><%=FormatNumber(ttltothi, 0)%></b></td>
						<td align="right"><b><%=FormatNumber(ttltotshi, 0)%></b></td>

						<td align="right"><b><%=FormatNumber(ttlkbthi, 0)%></b></td>
						<td align="right"><b><%=FormatNumber(ttlkbtshi, 0)%></b></td>
						<td align="right"><b><%=FormatNumber(ttlkht1hi, 0)%></b></td>
						<td align="right"><b><%=FormatNumber(ttlkht1shi, 0)%></b></td>
						<td align="right"><b><%=FormatNumber(ttlkhl1hi, 0)%></b></td>
						<td align="right"><b><%=FormatNumber(ttlkhl1shi, 0)%></b></td>
						<td align="right"><b><%=FormatNumber(ttlpremihi, 0)%></b></td>
						<td align="right"><b><%=FormatNumber(ttlpremishi, 0)%></b></td>
						
						<td align="right"><b><%=FormatNumber(ttlprepanenhi, 0)%></b></td>
						<td align="right"><b><%=FormatNumber(ttlprepanenshi, 0)%></b></td>
						
						<td align="right"><b><%=FormatNumber(ttldendahi, 0)%></b></td>
						<td align="right"><b><%=FormatNumber(ttldendashi, 0)%></b></td>
						
						<td align="right"><b><%=FormatNumber(ttllembur1hi, 0)%></b></td>
						<td align="right"><b><%=FormatNumber(ttllembur1shi, 0)%></b></td>
						<td align="right"><b><%=FormatNumber(ttlbor1hi, 0)%></b></td>
						<td align="right"><b><%=FormatNumber(ttlbor1shi, 0)%></b></td>
						
						<td align="right"><b><%=FormatNumber(ttlbiayahi, 0)%></b></td>
						<td align="right"><b><%=FormatNumber(ttlbiayashi, 0)%></b></td>

						<td align="right"></td>
						<td align="right"></td>

						<td align="right"><b><%=FormatNumber(ttlmatjumhi, 0)%></b></td>
						<td align="right"><b><%=FormatNumber(ttlmatjumshi, 0)%></b></td>
						<td align="right"><b><%=FormatNumber(ttlbiayamathi, 0)%></b></td>
						<td align="right"><b><%=FormatNumber(ttlbiayamatshi, 0)%></b></td>

						<td align="right"></td>

						<td align="right"><b><%=FormatNumber(ttlhmhi, 0)%></b></td>
						<td align="right"><b><%=FormatNumber(ttlhmshi, 0)%></b></td>
						
					
						
					
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
