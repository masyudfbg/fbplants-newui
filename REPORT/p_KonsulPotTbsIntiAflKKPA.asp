<%@Language=VBScript%>
<% ValidateSession %>
<!--#include virtual="/conndB1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->

<script language="JavaScript">

function validasi(form){

		  if (form.txtDateFrom.value > form.txtDateTo.value){

			alert("From Date Lebih Besar Dari To Date.");
			
			form.txtDateFrom.focus();

			return (false);

		  }
}
</script>
<%
Dim v_txtFcba, v_txtYear, v_txtMonth,v_txtDay, v_txtBlok, v_txtColor
Dim v_txtDateFrom, v_txtDateTo, v_txtPabrik, v_txtSupplier, v_txtAfdeling
Dim ScriptName
Dim rsPusingan, strPusingan, strsql
Dim NO, cClass, nday,ColorID
Dim rsDefaultBa, strDefaultBa, varDefaultBa
Dim rsBaInfo, strBaInfo, varFcbaName,varCompanyCode, varCompanyName
Dim oRS, oSQL

Dim vColor
Dim tempEstate
Dim tempAfd
Dim isKebun, isAfd, isPabrik

tempEstate =""
tempFFBType=""
tempAfd =""

cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"

v_txtDateFrom = Request("txtDateFrom")
v_txtDateTo = Request("txtDateTo")

If "" & v_txtDateFrom = "" Then v_txtDateFrom = TextDate(date)
If "" & v_txtDateTo = "" Then v_txtDateTo = TextDate(date)

'' Set Variabel Default, untuk masing2x, Filter Combobox, agar tidak berubah pada saat klik View
v_txtSupplier				=	Request("txtSupplier")	'' Supplier
v_txtFcba					=	Request("txtFcba") 		'' Business Unit
v_txtAfdeling				=	Request("txtAfdeling")	'' Afdeling
v_txtPabrik					=	Request("txtPabrik")	'' Pabrik

''Set Default FCBA ( Business Unit Kebun dan Pabrik )
strDefaultBa = " SELECT" & _
				"     FCCODE as FCBA" & _
				" FROM BUSINESSUNIT BU" & _
				" INNER JOIN USER_FCBA_TAB UFT ON (BU.FCCODE = UFT.FCBA) " & _
				" WHERE" & _
				"     ISDEFAULT = 'TRUE'" & _
	    		"     AND FCUSERCODE = '" & Session("CodeName") & "'"

''response.write strDefaultBa
Set rsDefaultBa = Server.CreateObject("ADODB.Recordset")
rsDefaultBa.Open strDefaultBa, connstr, adOpenStatic

If Not rsDefaultBa.EOF Then
	varDefaultBa = rsDefaultBa("FCBA")
Else
	varDefaultBa = ""
End If
rsDefaultBa.Close
set rsDefaultBa = Nothing

If (v_txtFcba = "" Or Len(Trim(v_txtFcba)) = 0) Then
	v_txtFcba = varDefaultBa
End If

ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

<!--Get Data Company -->
strBaInfo = " SELECT" & _
			"     FCNAME as FCBANAME" & _
			"    , fccompanycode  , FCCOMPANYNAME as COMPANYNAME" & _
			" FROM BUSINESSUNIT BU" & _
			" WHERE" & _
			"     FCCODE = '" & v_txtFcba & "'"

Set rsBaInfo = Server.CreateObject("ADODB.Recordset")
rsBaInfo.Open strBaInfo, connstr, adOpenStatic

If Not rsBaInfo.EOF Then
	varFcbaName = rsBaInfo("FCBANAME")
	varCompanyName = rsBaInfo("COMPANYNAME")
	varCompanyCode = rsBaInfo("fccompanycode")
Else
	varFcbaName = ""
	varCompanyName = ""
	varCompanyCode = ""
End If
rsBaInfo.Close
set rsBaInfo = Nothing



<!--Additional -->

	'Filter Pabrik
	If v_txtPabrik <> "ALL" And Len(Trim(v_txtPabrik)) > 0 Then
		isPabrik="true"
	Else
		isPabrik="false"
	End If
	
	'Filter Kebun
	If v_txtSupplier <> "ALL" And Len(Trim(v_txtSupplier)) > 0 Then
		isKebun="true"
	Else
		isKebun="false"
	End If
	
	'Filter Afdeling
	If v_txtAfdeling <> "ALL" And Len(Trim(v_txtAfdeling)) > 0 Then
		isAfd="true"
	Else
		isAfd="false"
	End If
			
			

strPusingan	=   "SELECT " & _
				"InitialEstate, FFBtype, Afdeling, Blok, chitno, No_Polisi," & _
				"NVL(QtyNetto,0) as QtyNetto," & _
				"SUBSTR(InitialEstate,1,8) as vSubtotal," & _
				"NVL(persenmentah,0) as persenmentah, NVL(kgmentah,0) as kgmentah, " & _
				"NVL(persenlewatmasak,0) as persenlewatmasak, NVL(kglewatmasak,0) as kglewatmasak," & _
				"NVL(persenTandanKosong,0) as persenTandanKosong, NVL(KgTandanKosong,0) as KgTandanKosong," & _
				"NVL(persenTangkaipanjang,0) as persenTangkaipanjang, NVL(KgTangkaiPanjang,0) as KgTangkaiPanjang," & _
				"NVL(PersenEatenByRat,0) as PersenEatenByRat, NVL(KgEatenByRat,0) as KgEatenByRat, " & _
				"NVL(Persensampah,0) as Persensampah, NVL(KgSampah,0) as KgSampah," & _
				"NVL(Persenair,0) as Persenair, NVL(KgAir,0) as KgAir," & _
				"NVL(Persenbrondolan,0) as Persenbrondolan, NVL(Kgbrondolan,0) as Kgbrondolan," & _
				"weight, " & _			
				"NVL(PersenTotalPot,0) as PersenTotalPot , NVL(KgTotalPot,0) as KgTotalPot " & _	
				"FROM ( " & _
				"    select " & _				
				"    	a.field_estate_estatecode AS cOrder," & _				
				"    	a.field_estate_estatecode AS InitialEstate," & _	
				" 		customer_api.GetCUSTFFBTYPE(a.cust_customercode,a.fcba) FFBtype, " & _		
				"    	a.division AS Afdeling," & _				
				"     	a.field_fieldcode AS Blok," & _				
				"     	a.chitno," & _				
				"    	a.vehicleno AS No_Polisi," & _				
				"    	(a.weightin - a.weightout) AS QtyNetto," & _				
				"    	b.f00 persenmentah," & _				
				"    	(b.f00 / 100) * (a.weightin - a.weightout) kgmentah," & _				
				"     	b.f5 persenlewatmasak," & _				
				"    	(b.f5 / 100) * (a.weightin - a.weightout) kglewatmasak," & _				
				"    	b.tkosong persenTandanKosong," & _				
				"    	(b.tkosong / 100) * (a.weightin - a.weightout) KgTandanKosong," & _				
				"    	b.tpanjang persenTangkaipanjang," & _				
				"    	(b.tpanjang / 100) * (a.weightin - a.weightout) KgTangkaiPanjang," & _				
				"    	DECODE(b.eatenbyrat,0,0,DECODE( (A .weightin - A .weightout),0,0,(b.eatenbyrat*100) / (A .weightin - A .weightout) ) ) PersenEatenByRat," & _				
				"    	b.eatenbyrat KgEatenByRat," & _		
				"    	DECODE(b.sampah,0,0,DECODE( (A .weightin - A .weightout),0,0,(b.sampah*100) / (A .weightin - A .weightout) ) )  Persensampah," & _				
				"    	b.sampah KgSampah," & _	
				"    	DECODE(b.air,0,0,DECODE( (A .weightin - A .weightout),0,0,(b.air*100) / (A .weightin - A .weightout) ) )  Persenair," & _				
				"    	b.air KgAir," & _	
				"    	DECODE(b.brondolan,0,0,DECODE( (A .weightin - A .weightout),0,0,(b.brondolan*100) / (A .weightin - A .weightout) ) )  Persenbrondolan," & _				
				"    	b.brondolan Kgbrondolan," & _	
				"    	b.weight, " & _				
				
				"    	 DECODE(((DECODE(b.f00,0,0,b.f00 / 100)) * (A .weightin - A .weightout) 	" & _	
                "    	+ (DECODE(b.f5,0,0, b.f5/ 100)) * (A .weightin - A .weightout)				" & _	
                "    	+ (DECODE(b.tkosong,0,0, b.tkosong/ 100)) * (A .weightin - A .weightout)	" & _	
                "    	+ (DECODE(b.tpanjang,0,0,b.tpanjang/ 100)) * (A .weightin - A .weightout)	" & _	
                "    	+b.eatenbyrat																" & _	
                "    	+b.sampah																	" & _	
                "    	+b.air																		" & _	
                "    	+b.brondolan),0,0,DECODE((A .weightin - A .weightout),0,0, 					" & _	
                "    	                                   (DECODE(b.f00,0,0,b.f00 / 100)) * (A .weightin - A .weightout) 	" & _	
                "    	                                + (DECODE(b.f5,0,0, b.f5/ 100)) * (A .weightin - A .weightout)		" & _	
                "    	                                + (DECODE(b.tkosong,0,0, b.tkosong/ 100)) * (A .weightin - A .weightout)  " & _	
                "    	                                + (DECODE(b.tpanjang,0,0,b.tpanjang/ 100)) * (A .weightin - A .weightout) " & _	
                "     	                                +b.eatenbyrat	" & _	
                "    	                                +b.sampah		" & _	
                "    	                                +b.air			" & _	
                "    	                                +b.brondolan) / (A .weightin - A .weightout)) * 100  as PersenTotalPot, " & _	
				
				"    	(b.f00 / 100) * (A .weightin - A .weightout) " & _
				"    	 + (b.f5 / 100) * (A .weightin - A .weightout) " & _
				"    	 + (b.tkosong / 100) * (A .weightin - A .weightout) " & _
				"    	 + (b.tpanjang / 100) * (A .weightin - A .weightout) " & _
				"    	 + b.eatenbyrat+b.sampah+b.air+b.brondolan as KgTotalPot " & _
				
				"    FROM weighbridgeticket a" & _	
				"    INNER JOIN gradingrecord b on a.chitno = b.wbt_chitno and a.fcba = b.fcba" & _				
				"    INNER JOIN BUSINESSUNIT C ON C.FCCODE = A.fcba AND C.FCCODE = b.fcba " & _
				"    WHERE 1=1 " & _		
				"    /* WHERE a.FCBA = '" & varDefaultBa & "' */" & _		
				"    AND C.FCCOMPANYCODE = '" & varCompanyCode & "'" & _
				" 	 AND a.postingdate BETWEEN TO_DATE('" & v_txtDateFrom & "', 'mm/dd/yyyy') AND TO_DATE('" & v_txtDateTo & "', 'mm/dd/yyyy')" & _				
				"    AND a.product_productcode = 'FFB'" & _
				"    AND a.division is not null " & _
				"	 AND customer_api.GetCUSTFFBTYPE(a.cust_customercode,a.fcba) in ('INTI','PLASMA','AFILIASI','KKPA') " 

				IF isPabrik="true" Then 
					strPusingan	= strPusingan	& "  AND A.fcba = '" & v_txtPabrik & "'"
				Else
					strPusingan	= strPusingan	& " "
				End If
								
				IF isKebun="true" Then 
					strPusingan	= strPusingan	& "  AND A.field_estate_estatecode = '" & v_txtSupplier & "'"
				Else
					strPusingan	= strPusingan	& " "
				End If
				
				IF isAfd="true" Then 
					strPusingan	= strPusingan	& "  AND A.division = '" & v_txtAfdeling & "'"
				Else
					strPusingan	= strPusingan	& " "
				End If		
				
				strPusingan	= strPusingan	& "   UNION ALL " & _	
				"    select " & _
				"    	a.field_estate_estatecode AS cOrder, " & _				
				"    	CONCAT('SUBTOTAL ',a.field_estate_estatecode) AS InitialEstate," & _	
				"    customer_api.GetCUSTFFBTYPE(a.cust_customercode, a.fcba)  AS FFBtype," & _								
				"    '' AS Afdeling," & _				
				"    '' AS Blok," & _					
				"    '' as chitno," & _				
				"    '' AS No_Polisi," & _	
				"    SUM((a.weightin - a.weightout)) AS QtyNetto," & _				
				"    SUM(b.f00) as persenmentah," & _					
				"    SUM((b.f00 / 100) * (a.weightin - a.weightout)) kgmentah," & _					
				"    SUM(b.f5) persenlewatmasak," & _					
				"    SUM((b.f5 / 100) * (a.weightin - a.weightout)) kglewatmasak," & _					
				"    SUM(b.tkosong) persenTandanKosong," & _					
				"    SUM((b.tkosong / 100) * (a.weightin - a.weightout)) KgTandanKosong," & _					
				"    SUM(b.tpanjang) persenTangkaipanjang," & _					
				"    SUM((b.tpanjang / 100) * (a.weightin - a.weightout)) KgTangkaiPanjang," & _					
				"    SUM (DECODE(b.eatenbyrat,0,0,DECODE( (A .weightin - A .weightout),0,0,(b.eatenbyrat*100) / (A .weightin - A .weightout) ) ) ) PersenEatenByRat," & _					
				"    SUM(b.eatenbyrat) KgEatenByRat," & _					
				"    SUM (DECODE(b.sampah,0,0,DECODE( (A .weightin - A .weightout),0,0,(b.sampah*100) / (A .weightin - A .weightout) ) ) ) Persensampah," & _					
				"    SUM(b.sampah) KgSampah," & _					
				"    SUM (DECODE(b.air,0,0,DECODE( (A .weightin - A .weightout),0,0,(b.air*100) / (A .weightin - A .weightout) ) ) ) Persenair," & _	
				"    SUM(b.air) KgAir," & _					
				"    SUM (DECODE(b.brondolan,0,0,DECODE( (A .weightin - A .weightout),0,0,(b.brondolan*100) / (A .weightin - A .weightout) ) ) ) Persenbrondolan," & _					
				"    SUM(b.brondolan) Kgbrondolan," & _					
				"    SUM(b.weight) weight," & _	
				
				"    SUM(DECODE(((DECODE(b.f00,0,0,b.f00 / 100)) * (A .weightin - A .weightout) 	" & _	
                "    	+ (DECODE(b.f5,0,0, b.f5/ 100)) * (A .weightin - A .weightout)				" & _	
                "    	+ (DECODE(b.tkosong,0,0, b.tkosong/ 100)) * (A .weightin - A .weightout)	" & _	
                "    	+ (DECODE(b.tpanjang,0,0,b.tpanjang/ 100)) * (A .weightin - A .weightout)	" & _	
                "    	+b.eatenbyrat																" & _	
                "    	+b.sampah																	" & _	
                "    	+b.air																		" & _	
                "    	+b.brondolan),0,0,DECODE((A .weightin - A .weightout),0,0, 					" & _	
                "    	                                   (DECODE(b.f00,0,0,b.f00 / 100)) * (A .weightin - A .weightout) 	" & _	
                "    	                                + (DECODE(b.f5,0,0, b.f5/ 100)) * (A .weightin - A .weightout)		" & _	
                "    	                                + (DECODE(b.tkosong,0,0, b.tkosong/ 100)) * (A .weightin - A .weightout)  " & _	
                "    	                                + (DECODE(b.tpanjang,0,0,b.tpanjang/ 100)) * (A .weightin - A .weightout) " & _	
                "     	                                +b.eatenbyrat	" & _	
                "    	                                +b.sampah		" & _	
                "    	                                +b.air			" & _	
                "    	                                +b.brondolan) / (A .weightin - A .weightout)) * 100)  as PersenTotalPot, " & _	
				
				"    SUM((b.f00 / 100) * (A .weightin - A .weightout) " & _
				"     + (b.f5 / 100) * (A .weightin - A .weightout)" & _
				"     + (b.tkosong / 100) * (A .weightin - A .weightout)" & _
				"     + (b.tpanjang / 100) * (A .weightin - A .weightout)" & _
				"     + b.eatenbyrat+b.sampah+b.air+b.brondolan) as KgTotalPot" & _
				
				"    FROM weighbridgeticket a" & _					
				"    INNER JOIN gradingrecord b on a.chitno = b.wbt_chitno and a.fcba = b.fcba" & _					
				"    INNER JOIN BUSINESSUNIT C ON C.FCCODE = A.fcba AND C.FCCODE = b.fcba " & _
				"    WHERE 1=1 " & _		
				"    /* WHERE a.FCBA = '" & varDefaultBa & "' */" & _		
				"    AND C.FCCOMPANYCODE = '" & varCompanyCode & "'" & _
				" 	 AND a.postingdate BETWEEN TO_DATE('" & v_txtDateFrom & "', 'mm/dd/yyyy') AND TO_DATE('" & v_txtDateTo & "', 'mm/dd/yyyy')" & _				
				"    AND a.product_productcode = 'FFB'" & _
				"    AND a.division is not null" & _					
				"	 AND customer_api.GetCUSTFFBTYPE(a.cust_customercode,a.fcba) in ('INTI','PLASMA','AFILIASI','KKPA') "
				
				IF isPabrik="true" Then 
					strPusingan	= strPusingan	& "  AND A.fcba = '" & v_txtPabrik & "'"
				Else
					strPusingan	= strPusingan	& " "
				End If
				
				IF isKebun="true" Then 
					strPusingan	= strPusingan	& "  AND A.field_estate_estatecode = '" & v_txtSupplier & "'"
				Else
					strPusingan	= strPusingan	& " "
				End If
				
				IF isAfd="true" Then 
					strPusingan	= strPusingan	& "  AND A.division = '" & v_txtAfdeling & "'"
				Else
					strPusingan	= strPusingan	& " "
				End If	
				
				strPusingan	= strPusingan	& "  " & _	
				"    GROUP BY  a.fcba,a.cust_customercode,a.field_estate_estatecode" & _					
				"    UNION ALL" & _
				"    select " & _
				"    	'zzz' AS cOrder, " & _				
				"    	'TOTAL'," & _
				"    '' AS FFBtype," & _					
				"    '' AS Afdeling," & _				
				"    '' AS Blok," & _					
				"    '' as chitno," & _				
				"    '' AS No_Polisi," & _	
				"    SUM((a.weightin - a.weightout)) AS QtyNetto," & _				
				"    SUM(b.f00) as persenmentah," & _					
				"    SUM((b.f00 / 100) * (a.weightin - a.weightout)) kgmentah," & _					
				"    SUM(b.f5) persenlewatmasak," & _					
				"    SUM((b.f5 / 100) * (a.weightin - a.weightout)) kglewatmasak," & _					
				"    SUM(b.tkosong) persenTandanKosong," & _					
				"    SUM((b.tkosong / 100) * (a.weightin - a.weightout)) KgTandanKosong," & _					
				"    SUM(b.tpanjang) persenTangkaipanjang," & _					
				"    SUM((b.tpanjang / 100) * (a.weightin - a.weightout)) KgTangkaiPanjang," & _					
				"    SUM (DECODE(b.eatenbyrat,0,0,DECODE( (A .weightin - A .weightout),0,0,(b.eatenbyrat*100) / (A .weightin - A .weightout) ) ) ) PersenEatenByRat," & _					
				"    SUM(b.eatenbyrat) KgEatenByRat," & _					
				"    SUM (DECODE(b.sampah,0,0,DECODE( (A .weightin - A .weightout),0,0,(b.sampah*100) / (A .weightin - A .weightout) ) ) ) Persensampah," & _					
				"    SUM(b.sampah) KgSampah," & _					
				"    SUM (DECODE(b.air,0,0,DECODE( (A .weightin - A .weightout),0,0,(b.air*100)/ (A .weightin - A .weightout) ) ) ) Persenair," & _	
				"    SUM(b.air) KgAir," & _					
				"    SUM (DECODE(b.brondolan,0,0,DECODE( (A .weightin - A .weightout),0,0,(b.brondolan*100) / (A .weightin - A .weightout) ) ) ) Persenbrondolan," & _					
				"    SUM(b.brondolan) Kgbrondolan," & _					
				"    SUM(b.weight) weight," & _	
				
				"    SUM(DECODE(((DECODE(b.f00,0,0,b.f00 / 100)) * (A .weightin - A .weightout) 	" & _	
                "    	+ (DECODE(b.f5,0,0, b.f5/ 100)) * (A .weightin - A .weightout)				" & _	
                "    	+ (DECODE(b.tkosong,0,0, b.tkosong/ 100)) * (A .weightin - A .weightout)	" & _	
                "    	+ (DECODE(b.tpanjang,0,0,b.tpanjang/ 100)) * (A .weightin - A .weightout)	" & _	
                "    	+b.eatenbyrat																" & _	
                "    	+b.sampah																	" & _	
                "    	+b.air																		" & _	
                "    	+b.brondolan),0,0,DECODE((A .weightin - A .weightout),0,0, 					" & _	
                "    	                                   (DECODE(b.f00,0,0,b.f00 / 100)) * (A .weightin - A .weightout) 	" & _	
                "    	                                + (DECODE(b.f5,0,0, b.f5/ 100)) * (A .weightin - A .weightout)		" & _	
                "    	                                + (DECODE(b.tkosong,0,0, b.tkosong/ 100)) * (A .weightin - A .weightout)  " & _	
                "    	                                + (DECODE(b.tpanjang,0,0,b.tpanjang/ 100)) * (A .weightin - A .weightout) " & _	
                "     	                                +b.eatenbyrat	" & _	
                "    	                                +b.sampah		" & _	
                "    	                                +b.air			" & _	
                "    	                                +b.brondolan) / (A .weightin - A .weightout)) * 100)  as PersenTotalPot, " & _	
				
				"    SUM((b.f00 / 100) * (A .weightin - A .weightout) " & _
				"     + (b.f5 / 100) * (A .weightin - A .weightout)" & _
				"     + (b.tkosong / 100) * (A .weightin - A .weightout)" & _
				"     + (b.tpanjang / 100) * (A .weightin - A .weightout)" & _
				"     + b.eatenbyrat+b.sampah+b.air+b.brondolan) as KgTotalPot" & _				
				
				"    FROM weighbridgeticket a" & _					
				"    INNER JOIN gradingrecord b on a.chitno = b.wbt_chitno and a.fcba = b.fcba" & _					
				"    INNER JOIN BUSINESSUNIT C ON C.FCCODE = A.fcba AND C.FCCODE = b.fcba " & _
				"    WHERE 1=1 " & _		
				"    /* WHERE a.FCBA = '" & varDefaultBa & "' */" & _		
				"    AND C.FCCOMPANYCODE = '" & varCompanyCode & "'" & _
				" 	 AND a.postingdate BETWEEN TO_DATE('" & v_txtDateFrom & "', 'mm/dd/yyyy') AND TO_DATE('" & v_txtDateTo & "', 'mm/dd/yyyy')" & _				
				"    AND a.product_productcode = 'FFB'" & _
				"    AND a.division is not null" & _
				" 	 AND customer_api.GetCUSTFFBTYPE(a.cust_customercode,a.fcba) in ('INTI','PLASMA','AFILIASI','KKPA')"
				
				IF isPabrik="true" Then 
					strPusingan	= strPusingan	& "  AND A.fcba = '" & v_txtPabrik & "'"
				Else
					strPusingan	= strPusingan	& " "
				End If
				
				IF isKebun="true" Then 
					strPusingan	= strPusingan	& "  AND A.field_estate_estatecode = '" & v_txtSupplier & "'"
				Else
					strPusingan	= strPusingan	& " "
				End If
				
				IF isAfd="true" Then 
					strPusingan	= strPusingan	& "  AND A.division = '" & v_txtAfdeling & "'"
				Else
					strPusingan	= strPusingan	& " "
				End If
				
				strPusingan	= strPusingan	& " " & _	
				"    ) X3 " & _
				"  WHERE 1=1 " 
						
				strPusingan	= strPusingan	& " ORDER BY FFBtype,cOrder, InitialEstate, Afdeling  "	

				'response.write strpusingan			
%>

<%
			Dim strTitle, rsTitle
			Dim titleGroupPlantation, titlePlantation, titleGroupRegion, titleRegion, titleCompany, titleSite
			
	        strTitle = " select PARAMETERSYSDETAIL_API.GETPARAM1 ('REPORTGROUP','GROUPPLANTATION') titleGP, " & _
							    " PARAMETERSYSDETAIL_API.GETPARAM1 ('REPORTGROUP','PLANTATION') titleP , " & _
								" PARAMETERSYSDETAIL_API.GETPARAM1 ('REPORTGROUP','GROUPREGION') titleGR, " & _
								" PARAMETERSYSDETAIL_API.GETPARAM1 ('REPORTGROUP','REGION') titleR, " & _
								" PARAMETERSYSDETAIL_API.GETPARAM1 ('REPORTGROUP','COMPANY') titleC, " & _
								" PARAMETERSYSDETAIL_API.GETPARAM1 ('REPORTGROUP','SITE') titleS from dual"
								
			Set rsTitle = Server.CreateObject("ADODB.Recordset")
			''response.write rsTitle
			rsTitle.Open strTitle, connstr, 3,3	

			if not rsTitle.EOF then
				no = 0
				titleGroupPlantation = rsTitle ("titleGP")			
				titlePlantation = rsTitle ("titleP")
				titleGroupRegion = rsTitle ("titleGR")
				titleRegion = rsTitle ("titleR")
				titleCompany = rsTitle ("titleC")
				titleSite = rsTitle ("titleS")
			else
				titleGroupPlantation = "Group Plantation"			
				titlePlantation = "Plantation"
				titleGroupRegion = "Group Region"
				titleRegion = "Region"
				titleCompany = "Company"
				titleSite = "Site"
			end if
			rsTitle.Close
			set rsTitle = Nothing
			
%>
<html>

<head>
<title>ePlaS Reporting - Potongan TBS Inti Afiliasi & KKPA</title>
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
	$( ".datepicker" ).datepicker({
		changeMonth: true,
		changeYear: true
	});
});
</script>
<!--
<script language="JavaScript" src="/js/date-picker.js"></script> 
<script language="JavaScript">

var newwin;
function launchwin(winurl,winname,winfeatures)
{
	//This launches a new window and then
	//focuses it if window.focus() is supported.
	newwin = window.open(winurl,winname,winfeatures);

}
</script>
-->

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
			<form action="<%=ScriptName%>" method="post" name="frmEntry" onSubmit="javascript:return validasi(this)" >
				<div class="printhide">
					<table id="form-table" class="centered-margin">	
					
					<tr>
						<!--<td><font face="verdana" size="1">TAHUN</font></td>
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
						<%
							For i = MinYr to MaxYr
						%>
							<Option Value="<%=i%>" <%If cint(i) = cint(v_txtYear) then%> Selected <%End If%>><%=i%></option>
						<%
							Next
						%>
								</select>
							</td>
							<td><font face="verdana" size="1">BULAN</font></td>
							<td width="2">
								<select name="txtMonth">
						<%
							For i = 1 to 12
						%>
							<Option Value="<%=i%>" <%If cint(i) = cint(v_txtMonth) then%> Selected <%End If%>><%=Bulan(i)%></option>
						<%
							Next
						%>
								</select>
							</td>-->
							
							<!-- From Date -->
							<td>
								<font face="verdana" size="1">FROM DATE</font>
							</td>
							<td width="2">
								<input class="datepicker" type="text" name="txtDateFrom" value="<%=v_txtDateFrom%>" />
							</td>
							
							<!-- To Date -->
							<td>
								<font face="verdana" size="1">TO DATE</font>
							</td>
							<td width="2">
								<input class="datepicker" type="text" name="txtDateTo" value="<%=v_txtDateTo%>" />
							</td>
							
							<!-- Pabrik / Mill -->
							<td>
								<font face="verdana" size="1">Pabrik</font>
							</td>
							<td>
								<select name="txtPabrik" >
									<option value="ALL"<%If v_txtPabrik = "ALL" then%> Selected <%End If%>>ALL</option>
									<%
									Dim rsPabrik, strPabrik																						
									strPabrik	=	" SELECT FCCODE " & _
													" FROM BUSINESSUNIT " & _
													" WHERE 1=1" & _
													" AND FCCOMPANYCODE= '" & varCompanyCode & "'" & _	
													" AND FCTYPE ='M' " & _
													" ORDER BY FCCODE " 
																	
									Set rsPabrik = Server.CreateObject("ADODB.Recordset")
									rsPabrik.Open strPabrik, connstr, adOpenStatic 
									
										if not rsPabrik.eof then
											Do While Not rsPabrik.eof
											if rsPabrik("FCCODE") = v_txtPabrik then%>
											<option value="<%=rsPabrik("FCCODE")%>" Selected><%=rsPabrik("FCCODE")%></option> 
											<%Else%>
											<option value="<%=rsPabrik("FCCODE")%>"><%=rsPabrik("FCCODE")%></option>  
											<%End if
												rsPabrik.movenext
											loop
										end if
									rsPabrik.Close
									set rsPabrik = Nothing
									%>
								</select>
							</td>
							
							<!-- Kebun / Estate -->
							<td>
								<font face="verdana" size="1">Kebun</font>
							</td>
							<td>
								<select name="txtSupplier" >
									<option value="ALL"<%If v_txtSupplier = "ALL" then%> Selected <%End If%>>ALL</option>
									<%
									Dim rsSupplier, strSupplier																						
									strSupplier	=	" SELECT FCCODE " & _
													" FROM BUSINESSUNIT " & _
													" WHERE 1=1" & _
													" AND FCTYPE ='E' " & _
													" ORDER BY FCCODE " 
																										
									Set rsSupplier = Server.CreateObject("ADODB.Recordset")
									rsSupplier.Open strSupplier, connstr, adOpenStatic 
									
										if not rsSupplier.eof then
											Do While Not rsSupplier.eof
											if rsSupplier("FCCODE") = v_txtSupplier then%>
											<option value="<%=rsSupplier("FCCODE")%>" Selected><%=rsSupplier("FCCODE")%></option> 
											<%Else%>
											<option value="<%=rsSupplier("FCCODE")%>"><%=rsSupplier("FCCODE")%></option>  
											<%End if
												rsSupplier.movenext
											loop
										end if
									rsSupplier.Close
									set rsSupplier = Nothing
									%>
								</select>
							</td>
							
							<!-- Afdeling -->
							<td>
								<font face="verdana" size="1">Afdeling</font>
							</td>
							<td>
								<select name="txtAfdeling" >
									<option value="ALL"<%If v_txtAfdeling = "ALL" then%> Selected <%End If%>>ALL</option>
									<%
									Dim rsAfdeling, strAfdeling																						
									strAfdeling	=	" SELECT DISTINCT DIVISION " & _
													" FROM FIELDDETAIL " & _
													" WHERE 1=1" & _
													" ORDER BY length(DIVISION), DIVISION " 
																										
									Set rsAfdeling = Server.CreateObject("ADODB.Recordset")
									rsAfdeling.Open strAfdeling, connstr, adOpenStatic 
									
										if not rsAfdeling.eof then
											Do While Not rsAfdeling.eof
											if rsAfdeling("DIVISION") = v_txtAfdeling then%>
											<option value="<%=rsAfdeling("DIVISION")%>" Selected><%=rsAfdeling("DIVISION")%></option> 
											<%Else%>
											<option value="<%=rsAfdeling("DIVISION")%>"><%=rsAfdeling("DIVISION")%></option>  
											<%End if
												rsAfdeling.movenext
											loop
										end if
									rsAfdeling.Close
									set rsAfdeling = Nothing
									%>
								</select>
							</td>
							
						
							<!-- View -->
							<td>
								<input type="submit" value="View" />
							</td>
						</tr>				
					</table>
				</div>
			</form>
		</div>
		
				
<table border="0" width="100%" id="product-table" class="xthf-site">
	<thead style="border-left:1px #000000 solid;">
	<tr>
         <th colspan="<%=25%>" style="background:#FFFFFF; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
		 <div id="header-wrapper">
			<div id="head_judul">
				<h4><%If Len(Trim(varCompanyName)) > 0 Then%>PT. <%End If%><%=varCompanyName%></h4>
				<h5><%=v_txtPabrik%></h5>
			</div>
			<div id="subheader" class="centered">
				<h2>Potongan TBS Inti Afiliasi dan KKPA</h3>
				<br/>
				
				<h4>Dari Tanggal&nbsp;:&nbsp;<%=v_txtDateFrom%>&nbsp; s/d Tanggal&nbsp;:&nbsp;<%=v_txtDateTo%></h4>
				<!-- 
				<h3>PERIODE&nbsp;:&nbsp;<%=Bulan(v_txtMonth)%>&nbsp;<%=v_txtYear%></h3>
				<h4>PERIODE&nbsp;:&nbsp;<%=Bulan(v_txtMonth)%>&nbsp;<%=v_txtYear%></h4>
				<h4>Dari Tanggal&nbsp;:&nbsp;<%=v_txtDay%>&nbsp;<%=Bulan(v_txtMonth)%>&nbsp;&nbsp;<%=v_txtYear%></h4>
				-->
				
				<br/>
			</div>
          </div>  
		 </th>
	</tr>
<% 
' dim selec_day
	' if len(trim(DAY(v_txtDate))) = 0 then
		' selec_day = GetLastDay(v_txtMonth, v_txtYear)
	' else
		' selec_day = DAY(v_txtDate)
	' end if
%>
	<tr>
		<th rowspan="2" style="background-color:#00A300;color:#fff;">Initial Estate</th>
		<th rowspan="2" style="background-color:#00A300;color:#fff;">FFB Type</th>
		<th rowspan="2" style="background-color:#00A300;color:#fff;">Afdeling</th>
		<th rowspan="2" style="background-color:#00A300;color:#fff;">Blok</th>
		<th rowspan="2" style="background-color:#00A300;color:#fff;">chitno</th>
		<th rowspan="2" style="background-color:#00A300;color:#fff;">No_Polisi</th>
		
		<th rowspan="2" style="background-color:#00A300;color:#fff;">Qty Netto</th>
		<th colspan="2" style="background-color:#00A300;color:#fff;">Buah Mentah</th>
		<th colspan="2" style="background-color:#00A300;color:#fff;">Buah Lewat Masak</th>
		<th colspan="2" style="background-color:#00A300;color:#fff;">Tandan Kosong</th>
		<th colspan="2" style="background-color:#00A300;color:#fff;">Tangkai Panjang</th>
		
		<th colspan="2" style="background-color:#00A300;color:#fff;">Dimakan Tikus</th>
		<th colspan="2" style="background-color:#00A300;color:#fff;">Sampah</th>
		<th colspan="2" style="background-color:#00A300;color:#fff;">Air</th>
		<th colspan="2" style="background-color:#00A300;color:#fff;">Brondolan</th>
		
		<th rowspan="2" style="background-color:#00A300;color:#fff;">% Total Pot</th>
		<th rowspan="2" style="background-color:#00A300;color:#fff;">Kg Total Pot</th>
	</tr>
	<tr>
		<th style="background-color:#00A300;color:#fff;">% Pot</th>
		<th style="background-color:#00A300;color:#fff;">Kg Pot</th>
		
		<th style="background-color:#00A300;color:#fff;">% Pot</th>
		<th style="background-color:#00A300;color:#fff;">Kg Pot</th>

		<th style="background-color:#00A300;color:#fff;">% Pot</th>
		<th style="background-color:#00A300;color:#fff;">Kg Pot</th>

		<th style="background-color:#00A300;color:#fff;">% Pot</th>
		<th style="background-color:#00A300;color:#fff;">Kg Pot</th>

		<th style="background-color:#00A300;color:#fff;">% Pot</th>
		<th style="background-color:#00A300;color:#fff;">Kg Pot</th>

		<th style="background-color:#00A300;color:#fff;">% Pot</th>
		<th style="background-color:#00A300;color:#fff;">Kg Pot</th>

		<th style="background-color:#00A300;color:#fff;">% Pot</th>
		<th style="background-color:#00A300;color:#fff;">Kg Pot</th>
		
		<th style="background-color:#00A300;color:#fff;">% Pot</th>
		<th style="background-color:#00A300;color:#fff;">Kg Pot</th>
		
	</tr>
	
    </thead>
	<tbody>
<%

''response.write strPusingan
''response.write strPabrik


Set rsPusingan = Server.CreateObject("ADODB.Recordset")
rsPusingan.Open strPusingan, connstr, adOpenStatic

Dim xNum

Dim Total_Ha : Total_Ha = 0
Dim Total_LastInterval : Total_LastInterval = 0
Dim DataIsPersen : DataIsPersen = 0
Dim CountData : CountData = 0


if rsPusingan.bof and rsPusingan.eof then

%>
	<tr>
		<td colspan="25">NO DATA AVAILABLE</td>
	</tr>
<%end if

		Dim n_netto
		Dim n_bm_percentPot, n_bm_kg	
		Dim n_blm_percentPot, n_blm_kg
		Dim n_tk_percentPot, n_tk_kg
		Dim n_tp_percentPot, n_tp_kg
		Dim	n_dt_percentPot, n_dt_kg 
		Dim n_bron_percentPot, n_bron_kg
		Dim n_sam_percentPot, n_sam_kg
		Dim n_Tot_percentPot, n_Tot_kg
			
		n_netto=0
		n_bm_percentPot=0
		n_bm_kg=0
		n_blm_percentPot=0
		n_blm_kg=0
		n_tk_percentPot=0
		n_tk_kg=0
		n_tp_percentPot=0
		n_tp_kg=0
		n_dt_percentPot=0
		n_dt_kg=0
		n_bron_percentPot=0
		n_bron_kg=0
		n_sam_percentPot=0
		n_sam_kg=0
		n_Tot_percentPot=0
		n_Tot_kg=0
				
		while not rsPusingan.eof
			CountData = CountData + 1
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
%>

			<tr class="<%=cClass%>">
				<%if rsPusingan("vSubtotal") = "SUBTOTAL" then%>  <!-- Merge Subtotal Here by AST-->
						<td align="Right" bgcolor="#F4F4A0" colspan="6" style="<%=cmdstyle%>"><font face="Verdana" size="1"><strong><%=rsPusingan("InitialEstate")%></strong></td>
						<%vColor="#F4F4A0"%> <!-- Give Color For Subtotal Here by AST-->
						<%tempEstate="" %>	
						<%tempFFBType="" %>						
						<%tempAfd="" %>
						
				<%Else %>
					<%if rsPusingan("vSubtotal") = "TOTAL" then%> <!-- Merge total Here by AST-->
						<td align="Right" bgcolor="#ffff01" colspan="6" style="<%=cmdstyle%>"><font face="Verdana" size="1"><strong><%=rsPusingan("InitialEstate")%></strong></td>
						<%vColor="#ffff01"%> <!-- Give Color For Total Here by AST-->
						<%tempEstate="" %>	
						<%tempFFBType="" %>
						<%tempAfd="" %>
						
					<%Else %>
						<%if rsPusingan("InitialEstate") = rsPusingan("vSubtotal")  then%> <!-- if Not Subtotal Or Not Total Here by AST-->
							
							<!-- InitialEstate hanya ditampilkan pertama kali, untuk nilai yang sama. by AST-->
							<%if tempEstate <> rsPusingan("InitialEstate") then %> 
								<td align="center"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsPusingan("InitialEstate")%></td>
								<%tempEstate=rsPusingan("InitialEstate")%> 
								
							<%Else %>
								<td align="center"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=""%></td>
								
							<%End if %>	
							
							<!-- FFBtype hanya ditampilkan pertama kali, untuk nilai yang sama. by AST-->
							<%if tempFFBType <> rsPusingan("FFBtype") then %> 
								<td align="center"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsPusingan("FFBtype")%></td>
								<%tempFFBType=rsPusingan("FFBtype")%> 
							<%Else %>
								<td align="center"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=""%></td>
							<%End if %>	
							
							<!-- Afdeling hanya ditampilkan pertama kali, untuk nilai yang sama. by AST-->
							<%if tempAfd <> rsPusingan("Afdeling") then %> 
								<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsPusingan("Afdeling")%></td>
								<%tempAfd=rsPusingan("Afdeling")%> 
							<%Else %>
								<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=""%></td>
							<%End if %>
					
														
							<td align="center"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsPusingan("Blok")%></td>
							<td align="center"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsPusingan("chitno")%></td>
							<td align="center"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsPusingan("No_Polisi")%></td>
							<%vColor=""%>
							
						<%End if %>	
					<%End if %>	
				<%End if %>		
				
				<td align="center" bgcolor="<%=vColor%>" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsPusingan("QtyNetto"),2)%></td>
				
				<td align="center" bgcolor="<%=vColor%>" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsPusingan("persenmentah"),2)%></td>
				<td align="center" bgcolor="<%=vColor%>" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsPusingan("kgmentah"),2)%></td>
				
				<td align="center" bgcolor="<%=vColor%>" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsPusingan("persenlewatmasak"),2)%></td>
				<td align="center" bgcolor="<%=vColor%>" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsPusingan("kglewatmasak"),2)%></td>
				
				<td align="Right" bgcolor="<%=vColor%>" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsPusingan("persenTandanKosong"),2)%></td>
				<td align="Right" bgcolor="<%=vColor%>" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsPusingan("KgTandanKosong"),2)%></td>
				
				<td align="Right" bgcolor="<%=vColor%>" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsPusingan("persenTangkaipanjang"),2)%></td>
				<td align="Right" bgcolor="<%=vColor%>" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsPusingan("KgTangkaiPanjang"),2)%></td>				
				
				<td align="Right" bgcolor="<%=vColor%>" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsPusingan("PersenEatenByRat"),2)%></td>
				<td align="Right" bgcolor="<%=vColor%>" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsPusingan("KgEatenByRat"),2)%></td>
				
				<td align="Right" bgcolor="<%=vColor%>" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsPusingan("Persensampah"),2)%></td>
				<td align="Right" bgcolor="<%=vColor%>" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsPusingan("KgSampah"),2)%></td>				

				<td align="Right" bgcolor="<%=vColor%>" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsPusingan("Persenair"),2)%></td>
				<td align="Right" bgcolor="<%=vColor%>" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsPusingan("KgAir"),2)%></td>				

				<td align="Right" bgcolor="<%=vColor%>" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsPusingan("Persenbrondolan"),2)%></td>
				<td align="Right" bgcolor="<%=vColor%>" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsPusingan("Kgbrondolan"),2)%></td>				
				
				<td align="Right" bgcolor="<%=vColor%>" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsPusingan("PersenTotalPot"),2)%></td>
				<td align="Right" bgcolor="<%=vColor%>" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsPusingan("KgTotalPot"),2)%></td>				
						
			</tr>

<%
			' n_Bruto=n_Bruto+trim(cdbl(rsPusingan("Bruto")))
			' n_Tarra=n_Tarra+trim(cdbl(rsPusingan("Tarra")))
			' n_Netto=n_Netto+trim(cdbl(rsPusingan("QtyNetto")))
			' n_TotalTariff=n_TotalTariff+trim(cdbl(rsPusingan("total_tarif")))
				
			rsPusingan.MoveNext
		Wend
			
''rsPusingan.Close
set rsPusingan = Nothing

''response.write  rsPusingan

''response.write  Total_Ha 
%>
		<!--
		<tr bgcolor="#F4F4A0">
			<td align="Right" colspan="5"> <strong><i><font face="Verdana" size="1">Grand Total</i></strong></td>
			<td align="Right"> <strong><i> <font face="Verdana" size="1"><%=FormatNumber(n_netto, 2)%></i></strong></td>
			
			<td align="Right"> <strong><i> <font face="Verdana" size="1"><%=FormatNumber(n_bm_percentPot, 2)%></i></strong></td>
			<td align="Right"> <strong><i> <font face="Verdana" size="1"><%=FormatNumber(n_bm_kg, 2)%></i></strong></td>
			
			<td align="Right"> <strong><i> <font face="Verdana" size="1"><%=FormatNumber(n_blm_percentPot, 2)%></i></strong></td>
			<td align="Right"> <strong><i> <font face="Verdana" size="1"><%=FormatNumber(n_blm_kg, 2)%></i></strong></td>
			
			<td align="Right"> <strong><i> <font face="Verdana" size="1"><%=FormatNumber(n_tk_percentPot, 2)%></i></strong></td>
			<td align="Right"> <strong><i> <font face="Verdana" size="1"><%=FormatNumber(n_tk_kg, 2)%></i></strong></td>
			
			<td align="Right"> <strong><i> <font face="Verdana" size="1"><%=FormatNumber(n_tp_percentPot, 2)%></i></strong></td>
			<td align="Right"> <strong><i> <font face="Verdana" size="1"><%=FormatNumber(n_tp_kg, 2)%></i></strong></td>
			
			<td align="Right"> <strong><i> <font face="Verdana" size="1"><%=FormatNumber(n_dt_percentPot, 2)%></i></strong></td>
			<td align="Right"> <strong><i> <font face="Verdana" size="1"><%=FormatNumber(n_dt_kg, 2)%></i></strong></td>
			
			<td align="Right"> <strong><i> <font face="Verdana" size="1"><%=FormatNumber(n_bron_percentPot, 2)%></i></strong></td>
			<td align="Right"> <strong><i> <font face="Verdana" size="1"><%=FormatNumber(n_bron_kg, 2)%></i></strong></td>
			
			<td align="Right"> <strong><i> <font face="Verdana" size="1"><%=FormatNumber(n_sam_percentPot, 2)%></i></strong></td>
			<td align="Right"> <strong><i> <font face="Verdana" size="1"><%=FormatNumber(n_sam_kg, 2)%></i></strong></td>
			
			<td align="Right"> <strong><i> <font face="Verdana" size="1"><%=FormatNumber(n_Tot_percentPot, 2)%></i></strong></td>
			<td align="Right"> <strong><i> <font face="Verdana" size="1"><%=FormatNumber(n_Tot_kg, 2)%></i></strong></td>
			
			<td align="Right"> <strong><i> <font face="Verdana" size="1"><%=FormatNumber(n_Tot_percentPot, 2)%></i></strong></td>
			<td align="Right"> <strong><i> <font face="Verdana" size="1"><%=FormatNumber(n_Tot_kg, 2)%></i></strong></td>
		</tr>
		-->
		
</tbody>
</table>

   <table width="90%" class="data">
		<br/>
		<!-- 
		   <tr>
				<td width="30%" align="center"><font face="Verdana" size="1">Prepared by,</font></td>
				<td width="30%" align="center"><font face="Verdana" size="1">Checked by,</font></td>
				<td width="30%" align="center"><font face="Verdana" size="1">AcKnowledge by,</font></td>	
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
		--> 
		
	   <%
		DIM rsttd,ttd
		
		ttd = "SELECT FCMANAGER,FCKTU FROM BusinessUnit WHERE FCCODE = '" & v_txtFcba & "'"
		
	    Set rsttd = Server.CreateObject("ADODB.Recordset")
		rsttd.Open ttd, connstr, adOpenStatic

		If Not rsttd.EOF Then
	   %>
	   <tr> 
	   <!--
		<td width="30%" align="center"><font face="Verdana" size="2"><strong><u><%=rsttd("FCKTU")%></u></strong></td>
		<td width="30%" align="center"><font face="Verdana" size="2"><strong><u><%=rsttd("FCMANAGER")%></u></strong></td>
		<td width="30%" align="center"><font face="Verdana" size="2"><strong><u><%=rsttd("FCMANAGER")%></u></strong></td>
		-->
	   </tr>
	<%  End If
%> 
	   	<!--
		<tr> 
			<td width="30%" align="center"><font face="Verdana" size="1"><strong><i>Krani Timbangan</i></strong></td>
			<td width="30%" align="center"><font face="Verdana" size="1"><strong><i>Manager</i></strong></td>
			<td width="30%" align="center"><font face="Verdana" size="1"><strong><i>General Manager</i></strong></td>
	    </tr>
		-->
   </table>
   </body>
   </html>
