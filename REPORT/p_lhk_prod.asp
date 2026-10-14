<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->

<script language="JavaScript">

function validasi(form){
		  if (form.txtDateFrom.value > form.txtDateTo.value){
			alert("From Date Tidak Boleh Lebih Besar Dari To Date.");
			form.txtDateFrom.focus();
			return (false);
		  }
}
</script>

<%
Dim v_txtDateFrom, v_txtDateTo,v_txtFCBA
Dim ScriptName, NO, cClass
Dim varCompanyName, v_txtPabrik


	if Trim(Request("txtDateFrom")) <> "" then
		v_txtDateFrom = Trim(Request("txtDateFrom"))
	 else
	   v_txtDateFrom = textDate(date)
	end if
	 
	if Trim(Request("txtDateTo")) <> "" then
		v_txtDateTo = Trim(Request("txtDateTo"))
	 else
	   v_txtDateTo = textDate(date)
	end if
	
  v_txtFCBA		= Trim(Request("txtFCBA"))
  ScriptName	= Request.ServerVariables("SCRIPT_NAME")

Dim rsSqlstr, Sqlstr
Sqlstr	=	" SELECT A.TRANS_DATE AS TANGGAL, " & _
			" 		A.STACKING_HA as STACKING_BI_AKTUAL_HI, " & _
			"		COALESCE(SUM(A.STACKING_HA) OVER(PARTITION BY A.FCBA ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING), A.STACKING_HA ) AS STACKING_BI_AKTUAL_SHI, " & _
			" 		(B.STACKING_HA / C.WORKINGDAY) as STACKING_BI_BUDGET_HI, " & _
			"  		COALESCE(SUM(B.STACKING_HA / C.WORKINGDAY) OVER (PARTITION BY B.FCBA	ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING ), (B.STACKING_HA / C.WORKINGDAY) ) AS STACKING_BI_BUDGET_SHI," & _
			"  		( " & _
			"  			COALESCE(SUM(A.STACKING_HA) OVER (PARTITION BY A.FCBA ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING ), A.STACKING_HA ) " & _
			"  			- " & _
			"  			COALESCE(SUM(B.STACKING_HA / C.WORKINGDAY) OVER (PARTITION BY B.FCBA ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING ), (B.STACKING_HA / C.WORKINGDAY))" & _
			"  		) as STACKING_BI_VarSHI_HA," & _
			"  		( " & _
			"  			( COALESCE(SUM(A.STACKING_HA) OVER (PARTITION BY A.FCBA ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING ), A.STACKING_HA ) " & _
			"  				-" & _
			"  			COALESCE(SUM(B.STACKING_HA / C.WORKINGDAY) OVER (PARTITION BY B.FCBA ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING ), (B.STACKING_HA / C.WORKINGDAY))" & _
			"  			)" & _
			"  			/ " & _
			"  			( COALESCE(SUM(B.STACKING_HA / C.WORKINGDAY) OVER (PARTITION BY B.FCBA ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING ), (B.STACKING_HA / C.WORKINGDAY)) * 100 " & _
			"  		 	) " & _
			"  		) as STACKING_BI_VarSHI_PERCENT" & _
			"  		,D.STACKING_SBI_AKTUAL ,E.STACKING_SBI_BUDGET" & _
			"  		,(D.STACKING_SBI_AKTUAL - E.STACKING_SBI_BUDGET) AS STACKING_SBI_VarSBI_HA" & _
			"  		,((D.STACKING_SBI_AKTUAL - E.STACKING_SBI_BUDGET) / (E.STACKING_SBI_BUDGET*100)) AS STACKING_SBI_VarSBI_PERCENT," & _

			"  		A.TANAMBARU_HA as TANAMBARU_BI_AKTUAL_HI," & _
			"  		COALESCE(SUM(A.TANAMBARU_HA) OVER(PARTITION BY A.FCBA ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING), A.TANAMBARU_HA ) AS TANAMBARU_BI_AKTUAL_SHI," & _
			"  		(B.TANAMBARU_HA / C.WORKINGDAY) as TANAMBARU_BI_BUDGET_HI," & _
			"  		COALESCE(SUM(B.TANAMBARU_HA / C.WORKINGDAY) OVER (PARTITION BY B.FCBA	ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING ), (B.TANAMBARU_HA / C.WORKINGDAY) ) AS TANAMBARU_BI_BUDGET_SHI," & _
			"  		( COALESCE(SUM(A.TANAMBARU_HA) OVER (PARTITION BY A.FCBA ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING ), A.TANAMBARU_HA ) " & _
			"  			- " & _
			"  		  COALESCE(SUM(B.TANAMBARU_HA / C.WORKINGDAY) OVER (PARTITION BY B.FCBA ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING ), (B.TANAMBARU_HA / C.WORKINGDAY))" & _
			"  		) as TANAMBARU_BI_VarSHI_HA," & _
			"  		(" & _
			"  			( COALESCE(SUM(A.TANAMBARU_HA) OVER (PARTITION BY A.FCBA ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING ), A.TANAMBARU_HA ) " & _
			"  				-" & _
			"  			  COALESCE(SUM(B.TANAMBARU_HA / C.WORKINGDAY) OVER (PARTITION BY B.FCBA ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING ), (B.TANAMBARU_HA / C.WORKINGDAY))" & _
			"  			)" & _
			"  			/" & _
			"  			( COALESCE(SUM(B.TANAMBARU_HA / C.WORKINGDAY) OVER (PARTITION BY B.FCBA ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING ), (B.TANAMBARU_HA / C.WORKINGDAY)) * 100 " & _
			"  			)" & _
			"  		) as TANAMBARU_BI_VarSHI_PERCENT" & _
			"  		,D.TANAMBARU_SBI_AKTUAL ,E.TANAMBARU_SBI_BUDGET" & _
			"  		,(D.TANAMBARU_SBI_AKTUAL - E.TANAMBARU_SBI_BUDGET) AS TANAMBARU_SBI_VarSBI_HA" & _
			"  		,((D.TANAMBARU_SBI_AKTUAL - E.TANAMBARU_SBI_BUDGET) / (E.TANAMBARU_SBI_BUDGET*100)) AS TANAMBARU_SBI_VarSBI_PERCENT," & _
			
			"  		A.KOMPLETING_PKK as KOMPLETING_BI_AKTUAL_HI, " & _
			"  		COALESCE(SUM(A.KOMPLETING_PKK) OVER(PARTITION BY A.FCBA ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING), A.KOMPLETING_PKK ) AS KOMPLETING_BI_AKTUAL_SHI," & _
			"  		(B.KOMPLETING_PKK / C.WORKINGDAY) as KOMPLETING_BI_BUDGET_HI," & _
			"  		 COALESCE(SUM(B.KOMPLETING_PKK / C.WORKINGDAY) OVER (PARTITION BY B.FCBA	ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING ), (B.KOMPLETING_PKK / C.WORKINGDAY) ) AS KOMPLETING_BI_BUDGET_SHI," & _
			"  		( COALESCE(SUM(A.KOMPLETING_PKK) OVER (PARTITION BY A.FCBA ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING ), A.KOMPLETING_PKK ) " & _
			"  			- " & _
			"  		  COALESCE(SUM(B.KOMPLETING_PKK / C.WORKINGDAY) OVER (PARTITION BY B.FCBA ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING ), (B.KOMPLETING_PKK / C.WORKINGDAY))" & _
			"  		) as KOMPLETING_BI_VarSHI_PKK," & _
			"  		(" & _
			"  			( COALESCE(SUM(A.KOMPLETING_PKK) OVER (PARTITION BY A.FCBA ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING ), A.KOMPLETING_PKK) " & _
			"  				-" & _
			"  			  COALESCE(SUM(B.KOMPLETING_PKK / C.WORKINGDAY) OVER (PARTITION BY B.FCBA ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING ), (B.KOMPLETING_PKK / C.WORKINGDAY))" & _
			"  			)" & _
			"  			/ " & _
			"  			( COALESCE(SUM(B.KOMPLETING_PKK / C.WORKINGDAY) OVER (PARTITION BY B.FCBA ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING ), (B.KOMPLETING_PKK / C.WORKINGDAY)) * 100 " & _
			"  			)" & _
			"  		 ) as KOMPLETING_BI_VarSHI_PERCENT" & _
			"  		,D.KOMPLETING_SBI_AKTUAL ,E.KOMPLETING_SBI_BUDGET" & _
			"  		,(D.KOMPLETING_SBI_AKTUAL - E.KOMPLETING_SBI_BUDGET) AS KOMPLETING_SBI_VarSBI_HA" & _
			"  		,((D.KOMPLETING_SBI_AKTUAL - E.KOMPLETING_SBI_BUDGET) / (E.KOMPLETING_SBI_BUDGET*100)) AS KOMPLETING_SBI_VarSBI_PERCENT," & _	
			
			"  		A.TANAMSISIP_TBM_PKK as SISIP_TBM_BI_AKTUAL_HI, " & _
			"  		COALESCE(SUM(A.TANAMSISIP_TBM_PKK) OVER(PARTITION BY A.FCBA ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING), A.TANAMSISIP_TBM_PKK ) AS SISIP_TBM_BI_AKTUAL_SHI," & _
			"  		(B.TANAMSISIP_TBM_PKK / C.WORKINGDAY) as SISIP_TBM_BI_BUDGET_HI," & _
			"  		COALESCE(SUM(B.TANAMSISIP_TBM_PKK / C.WORKINGDAY) OVER (PARTITION BY B.FCBA	ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING ), (B.TANAMSISIP_TBM_PKK / C.WORKINGDAY) ) AS SISIP_TBM_BI_BUDGET_SHI," & _
			"  		( COALESCE(SUM(A.TANAMSISIP_TBM_PKK) OVER (PARTITION BY A.FCBA ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING ), A.TANAMSISIP_TBM_PKK ) " & _
			"  			- " & _
			"  		  COALESCE(SUM(B.TANAMSISIP_TBM_PKK / C.WORKINGDAY) OVER (PARTITION BY B.FCBA ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING ), (B.TANAMSISIP_TBM_PKK / C.WORKINGDAY))" & _
			"  		) as SISIP_TBM_BI_VarSHI_PKK," & _
			"  		(" & _
			"  			( COALESCE(SUM(A.TANAMSISIP_TBM_PKK) OVER (PARTITION BY A.FCBA ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING ), A.TANAMSISIP_TBM_PKK) " & _
			"  			- " & _
			"  			COALESCE(SUM(B.TANAMSISIP_TBM_PKK / C.WORKINGDAY) OVER (PARTITION BY B.FCBA ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING ), (B.TANAMSISIP_TBM_PKK / C.WORKINGDAY))" & _
			"  			)" & _
			"  			/" & _
			"  			( COALESCE(SUM(B.TANAMSISIP_TBM_PKK / C.WORKINGDAY) OVER (PARTITION BY B.FCBA ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING ), (B.TANAMSISIP_TBM_PKK / C.WORKINGDAY)) * 100 " & _
			"  			)" & _
			"  		 ) as SISIP_TBM_BI_VarSHI_PERCENT" & _
			"  		,D.SISIP_TBM_SBI_AKTUAL	 ,E.SISIP_TBM_SBI_BUDGET" & _
			"  		,(D.SISIP_TBM_SBI_AKTUAL - E.SISIP_TBM_SBI_BUDGET) AS SISIP_TBM_SBI_VarSBI_PKK" & _
			"  		,((D.SISIP_TBM_SBI_AKTUAL - E.SISIP_TBM_SBI_BUDGET) / (E.SISIP_TBM_SBI_BUDGET*100)) AS SISIP_TBM_SBI_VarSBI_PERCENT," & _	

			"  		A.TANAMSISIP_TM_PKK as SISIP_TM_BI_AKTUAL_HI, " & _
			"  		COALESCE(SUM(A.TANAMSISIP_TM_PKK) OVER(PARTITION BY A.FCBA ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING), A.TANAMSISIP_TM_PKK ) AS SISIP_TM_BI_AKTUAL_SHI," & _
			"  		(B.TANAMSISIP_TM_PKK / C.WORKINGDAY) as SISIP_TM_BI_BUDGET_HI," & _
			"  		COALESCE(SUM(B.TANAMSISIP_TM_PKK / C.WORKINGDAY) OVER (PARTITION BY B.FCBA	ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING ), (B.TANAMSISIP_TM_PKK / C.WORKINGDAY) ) AS SISIP_TM_BI_BUDGET_SHI," & _
			"  		( COALESCE(SUM(A.TANAMSISIP_TM_PKK) OVER (PARTITION BY A.FCBA ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING ), A.TANAMSISIP_TM_PKK ) " & _
			"  		   - " & _
			"  		  COALESCE(SUM(B.TANAMSISIP_TM_PKK / C.WORKINGDAY) OVER (PARTITION BY B.FCBA ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING ), (B.TANAMSISIP_TM_PKK / C.WORKINGDAY))" & _
			"  		) as SISIP_TM_BI_VarSHI_PKK," & _
			"  		(" & _
			"  			( COALESCE(SUM(A.TANAMSISIP_TM_PKK) OVER (PARTITION BY A.FCBA ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING ), A.TANAMSISIP_TM_PKK) " & _
			"  			  - " & _
			"  			  COALESCE(SUM(B.TANAMSISIP_TM_PKK / C.WORKINGDAY) OVER (PARTITION BY B.FCBA ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING ), (B.TANAMSISIP_TM_PKK / C.WORKINGDAY))" & _
			"  			)" & _
			"  			/" & _
			"  			( COALESCE(SUM(B.TANAMSISIP_TM_PKK / C.WORKINGDAY) OVER (PARTITION BY B.FCBA ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING ), (B.TANAMSISIP_TM_PKK / C.WORKINGDAY)) * 100 " & _
			"  			)" & _
			"  		 ) as SISIP_TM_BI_VarSHI_PERCENT" & _
			"  		,D.SISIP_TM_SBI_AKTUAL ,E.SISIP_TM_SBI_BUDGET" & _
			"  		,(D.SISIP_TM_SBI_AKTUAL - E.SISIP_TM_SBI_BUDGET) AS SISIP_TM_SBI_VarSBI_PKK" & _
			"  		,((D.SISIP_TM_SBI_AKTUAL - E.SISIP_TM_SBI_BUDGET) / (E.SISIP_TM_SBI_BUDGET*100)) AS SISIP_TM_SBI_VarSBI_PERCENT," & _		

			"  		A.KASTRASI_HA as KASTRASI_BI_AKTUAL_HI, " & _
			"  		COALESCE(SUM(A.KASTRASI_HA) OVER(PARTITION BY A.FCBA ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING), A.KASTRASI_HA ) AS KASTRASI_BI_AKTUAL_SHI," & _
			"  		(B.KASTRASI_HA / C.WORKINGDAY) as KASTRASI_BI_BUDGET_HI," & _	
			"  		COALESCE(SUM(B.KASTRASI_HA / C.WORKINGDAY) OVER (PARTITION BY B.FCBA	ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING ), (B.KASTRASI_HA / C.WORKINGDAY) ) AS KASTRASI_BI_BUDGET_SHI," & _
			"  		( COALESCE(SUM(A.KASTRASI_HA) OVER (PARTITION BY A.FCBA ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING ), A.KASTRASI_HA ) " & _
			"  		  -" & _	
			"  		  COALESCE(SUM(B.KASTRASI_HA / C.WORKINGDAY) OVER (PARTITION BY B.FCBA ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING ), (B.KASTRASI_HA / C.WORKINGDAY))" & _
			"  		) as KASTRASI_BI_VarSHI_PKK," & _
			"  		(" & _	
			"  			( COALESCE(SUM(A.KASTRASI_HA) OVER (PARTITION BY A.FCBA ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING ), A.KASTRASI_HA) " & _
			"  			  -" & _
			"  			  COALESCE(SUM(B.KASTRASI_HA / C.WORKINGDAY) OVER (PARTITION BY B.FCBA ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING ), (B.KASTRASI_HA / C.WORKINGDAY))" & _	
			"  			)" & _
			"  			/" & _
			"  			( COALESCE(SUM(B.KASTRASI_HA / C.WORKINGDAY) OVER (PARTITION BY B.FCBA ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING ), (B.KASTRASI_HA / C.WORKINGDAY)) * 100 " & _	
			"  			)" & _
			"  		 ) as KASTRASI_BI_VarSHI_PERCENT" & _
			"  		,D.KASTRASI_SBI_AKTUAL ,E.KASTRASI_SBI_BUDGET" & _	
			"  		,(D.SISIP_TM_SBI_AKTUAL - E.SISIP_TM_SBI_BUDGET) AS KASTRASI_SBI_VarSBI_PKK" & _
			"  		,((D.SISIP_TM_SBI_AKTUAL - E.SISIP_TM_SBI_BUDGET) / (E.SISIP_TM_SBI_BUDGET*100)) AS KASTRASI_SBI_VarSBI_PERCENT," & _
			
			"  		A.SANITASI_HA as SANITASI_BI_AKTUAL_HI, " & _				
			"  		COALESCE(SUM(A.SANITASI_HA) OVER(PARTITION BY A.FCBA ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING), A.SANITASI_HA ) AS SANITASI_BI_AKTUAL_SHI," & _	
			"  		(B.SANITASI_HA / C.WORKINGDAY) as SANITASI_BI_BUDGET_HI," & _
			"  		COALESCE(SUM(B.SANITASI_HA / C.WORKINGDAY) OVER (PARTITION BY B.FCBA	ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING ), (B.SANITASI_HA / C.WORKINGDAY) ) AS SANITASI_BI_BUDGET_SHI," & _
			"  		( COALESCE(SUM(A.SANITASI_HA) OVER (PARTITION BY A.FCBA ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING ), A.SANITASI_HA ) " & _
			"  		   - " & _	
			"  		  COALESCE(SUM(B.SANITASI_HA / C.WORKINGDAY) OVER (PARTITION BY B.FCBA ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING ), (B.SANITASI_HA / C.WORKINGDAY))" & _
			"  		) as SANITASI_BI_VarSHI_PKK," & _
			"  		(" & _
			"  			( COALESCE(SUM(A.SANITASI_HA) OVER (PARTITION BY A.FCBA ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING ), A.SANITASI_HA) " & _	
			"  			  -" & _
			"  		      COALESCE(SUM(B.SANITASI_HA / C.WORKINGDAY) OVER (PARTITION BY B.FCBA ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING ), (B.SANITASI_HA / C.WORKINGDAY))" & _
			"  			)" & _
			"  			/" & _	
			"  			( COALESCE(SUM(B.SANITASI_HA / C.WORKINGDAY) OVER (PARTITION BY B.FCBA ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING ), (B.SANITASI_HA / C.WORKINGDAY)) * 100 " & _
			"  			)" & _
			"  		 ) as SANITASI_BI_VarSHI_PERCENT" & _
			"  		,D.SANITASI_SBI_AKTUAL	,E.SANITASI_SBI_BUDGET" & _	
			"  		,(D.SANITASI_SBI_AKTUAL - E.SANITASI_SBI_BUDGET) AS SANITASI_SBI_VarSBI_PKK" & _
			"  		,((D.SANITASI_SBI_AKTUAL - E.SANITASI_SBI_BUDGET) / (E.SANITASI_SBI_BUDGET*100)) AS SANITASI_SBI_VarSBI_PERCENT," & _

			"  		A.TK_PEMELIHARAAN_TERSEDIA AS PEMEL_AKTUAL_TERSEDIA_HI, " & _
			"  		COALESCE(SUM(A.TK_PEMELIHARAAN_TERSEDIA) OVER (PARTITION BY A.FCBA ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING ), A.TK_PEMELIHARAAN_TERSEDIA ) AS PEMEL_AKTUAL_TERSEDIA_SHI," & _	
			"  		A.TK_PEMELIHARAAN_TANAMBARU AS PEMEL_AKTUAL_TANAMBARU_HI," & _
			"  		COALESCE(SUM(A.TK_PEMELIHARAAN_TANAMBARU) OVER (PARTITION BY A.FCBA ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING ), A.TK_PEMELIHARAAN_TANAMBARU ) AS PEMEL_AKTUAL_TANAMBARU_SHI," & _
			"  		A.TK_PEMELIHARAAN_SISIP AS PEMEL_AKTUAL_SISIP_HI," & _
			"  		COALESCE(SUM(A.TK_PEMELIHARAAN_SISIP) OVER (PARTITION BY A.FCBA ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING ), A.TK_PEMELIHARAAN_SISIP ) AS PEMEL_AKTUAL_SISIP_SHI," & _	
			"  		A.TK_PEMELIHARAAN_KOMPLETING AS PEMEL_AKTUAL_KOMPLETING_HI," & _
			"  		COALESCE(SUM(A.TK_PEMELIHARAAN_KOMPLETING) OVER (PARTITION BY A.FCBA ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING), A.TK_PEMELIHARAAN_KOMPLETING ) AS PEMEL_AKTUAL_KOMPLETING_SHI," & _
			"  		A.TK_PEMELIHARAAN_PEMUPUKAN AS PEMEL_AKTUAL_PEMUPUKAN_HI," & _
			"  		COALESCE(SUM(A.TK_PEMELIHARAAN_PEMUPUKAN) OVER (PARTITION BY A.FCBA ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING ), A.TK_PEMELIHARAAN_PEMUPUKAN ) AS PEMEL_AKTUAL_PEMUPUKAN_SHI," & _	
			"  		A.TK_PEMELIHARAAN_LAINNYA AS PEMEL_AKTUAL_LAINNYA_HI," & _
			"  		COALESCE(SUM(A.TK_PEMELIHARAAN_LAINNYA) OVER (PARTITION BY A.FCBA ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING ), A.TK_PEMELIHARAAN_LAINNYA ) AS PEMEL_AKTUAL_LAINNYA_SHI," & _
			"  		(" & _
			"  		 (A.TK_PEMELIHARAAN_TANAMBARU+A.TK_PEMELIHARAAN_SISIP+A.TK_PEMELIHARAAN_KOMPLETING+A.TK_PEMELIHARAAN_PEMUPUKAN+A.TK_PEMELIHARAAN_LAINNYA)" & _
			"  		/ ( A.TK_PEMELIHARAAN_TERSEDIA * 100 ) " & _
			"  		)AS PEMEL_AKTUAL_PERCENT_HADIR_HI, " & _
			"  		COALESCE(SUM((A.TK_PEMELIHARAAN_TANAMBARU+A.TK_PEMELIHARAAN_SISIP+A.TK_PEMELIHARAAN_KOMPLETING+A.TK_PEMELIHARAAN_PEMUPUKAN+A.TK_PEMELIHARAAN_LAINNYA)" & _
			"  				/( A.TK_PEMELIHARAAN_TERSEDIA * 100 ) " & _
			"  				) OVER (PARTITION BY A.FCBA ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING ), " & _
			"  				((A.TK_PEMELIHARAAN_TANAMBARU+A.TK_PEMELIHARAAN_SISIP+A.TK_PEMELIHARAAN_KOMPLETING+A.TK_PEMELIHARAAN_PEMUPUKAN+A.TK_PEMELIHARAAN_LAINNYA)" & _
			"  				/(A.TK_PEMELIHARAAN_TERSEDIA * 100))" & _
			"  			) AS PEMEL_AKTUAL_PERCENT_HADIR_SHI," & _
			
			"  		A.PEMEL_JALAN_TIMBUN_MR_MTR as TIM_JL_MR_BI_AKTUAL_HI," & _
			"  		COALESCE(SUM(A.PEMEL_JALAN_TIMBUN_MR_MTR) OVER(PARTITION BY A.FCBA ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING), A.PEMEL_JALAN_TIMBUN_MR_MTR ) AS TIM_JL_MR_BI_AKTUAL_SHI," & _
			"  		(B.PEMEL_JALAN_TIMBUN_MR_MTR / C.WORKINGDAY) as TIM_JL_MR_BI_BUDGET_HI," & _
			"  		COALESCE(SUM(B.PEMEL_JALAN_TIMBUN_MR_MTR / C.WORKINGDAY) OVER (PARTITION BY B.FCBA	ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING ), (B.PEMEL_JALAN_TIMBUN_MR_MTR / C.WORKINGDAY) ) AS TIM_JL_MR_BI_BUDGET_SHI," & _	
			"  		( COALESCE(SUM(A.PEMEL_JALAN_TIMBUN_MR_MTR) OVER (PARTITION BY A.FCBA ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING ), A.PEMEL_JALAN_TIMBUN_MR_MTR ) " & _
			"  		  -" & _
			"  		  COALESCE(SUM(B.PEMEL_JALAN_TIMBUN_MR_MTR / C.WORKINGDAY) OVER (PARTITION BY B.FCBA ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING ), (B.PEMEL_JALAN_TIMBUN_MR_MTR / C.WORKINGDAY))" & _
			"  		) as TIM_JL_MR_BI_VarSHI_MTR," & _	
			"  		(" & _
			"  			( COALESCE(SUM(A.PEMEL_JALAN_TIMBUN_MR_MTR) OVER (PARTITION BY A.FCBA ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING ), A.PEMEL_JALAN_TIMBUN_MR_MTR) " & _
			"  			  -" & _
			"  		      COALESCE(SUM(B.PEMEL_JALAN_TIMBUN_MR_MTR / C.WORKINGDAY) OVER (PARTITION BY B.FCBA ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING ), (B.PEMEL_JALAN_TIMBUN_MR_MTR / C.WORKINGDAY))" & _	
			"  			)" & _
			"  			/" & _
			"  			( COALESCE(SUM(B.PEMEL_JALAN_TIMBUN_MR_MTR / C.WORKINGDAY) OVER (PARTITION BY B.FCBA ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING ), (B.PEMEL_JALAN_TIMBUN_MR_MTR / C.WORKINGDAY)) * 100 " & _		
			"  			)" & _
			"  		) as TIM_JL_MR_BI_VarSHI_PERCENT" & _
			"  		,D.TIM_JL_MR_SBI_AKTUAL	,E.TIM_JL_MR_SBI_BUDGET" & _
			"  		,(D.TIM_JL_MR_SBI_AKTUAL - E.TIM_JL_MR_SBI_BUDGET) AS TIM_JL_MR_SBI_VarSBI_MTR" & _	
			"  		,((D.TIM_JL_MR_SBI_AKTUAL - E.TIM_JL_MR_SBI_BUDGET) / (E.TIM_JL_MR_SBI_BUDGET*100)) AS TIM_JL_MR_SBI_VarSBI_PERCENT," & _

			"  		A.PEMEL_JALAN_TIMBUN_CR_MTR as TIM_JL_CR_BI_AKTUAL_HI," & _
			"  		COALESCE(SUM(A.PEMEL_JALAN_TIMBUN_CR_MTR) OVER(PARTITION BY A.FCBA ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING), A.PEMEL_JALAN_TIMBUN_CR_MTR ) AS TIM_JL_CR_BI_AKTUAL_SHI," & _
			"  		(B.PEMEL_JALAN_TIMBUN_CR_MTR / C.WORKINGDAY) as TIM_JL_CR_BI_BUDGET_HI," & _	
			"  		 COALESCE(SUM(B.PEMEL_JALAN_TIMBUN_CR_MTR / C.WORKINGDAY) OVER (PARTITION BY B.FCBA	ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING ), (B.PEMEL_JALAN_TIMBUN_CR_MTR / C.WORKINGDAY) ) AS TIM_JL_CR_BI_BUDGET_SHI," & _
			"  		( COALESCE(SUM(A.PEMEL_JALAN_TIMBUN_CR_MTR) OVER (PARTITION BY A.FCBA ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING ), A.PEMEL_JALAN_TIMBUN_CR_MTR ) " & _
			"  		  - " & _
			"  		  COALESCE(SUM(B.PEMEL_JALAN_TIMBUN_CR_MTR / C.WORKINGDAY) OVER (PARTITION BY B.FCBA ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING ), (B.PEMEL_JALAN_TIMBUN_CR_MTR / C.WORKINGDAY))" & _	
			"  		) as TIM_JL_CR_BI_VarSHI_MTR," & _
			"  		(" & _
			"  			( COALESCE(SUM(A.PEMEL_JALAN_TIMBUN_CR_MTR) OVER (PARTITION BY A.FCBA ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING ), A.PEMEL_JALAN_TIMBUN_CR_MTR) " & _				
			"  			  - " & _		
			"  			  COALESCE(SUM(B.PEMEL_JALAN_TIMBUN_CR_MTR / C.WORKINGDAY) OVER (PARTITION BY B.FCBA ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING ), (B.PEMEL_JALAN_TIMBUN_CR_MTR / C.WORKINGDAY))" & _
			"  			) " & _
			"  			/ " & _
			"  			( COALESCE(SUM(B.PEMEL_JALAN_TIMBUN_CR_MTR / C.WORKINGDAY) OVER (PARTITION BY B.FCBA ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING ), (B.PEMEL_JALAN_TIMBUN_CR_MTR / C.WORKINGDAY)) * 100 " & _	
			"  			)" & _
			"  		 ) as TIM_JL_CR_BI_VarSHI_PERCENT" & _
			"  		,D.TIM_JL_CR_SBI_AKTUAL	,E.TIM_JL_CR_SBI_BUDGET" & _
			"  		,(D.TIM_JL_CR_SBI_AKTUAL - E.TIM_JL_CR_SBI_BUDGET) AS TIM_JL_CR_SBI_VarSBI_MTR" & _	
			"  		,((D.TIM_JL_CR_SBI_AKTUAL - E.TIM_JL_CR_SBI_BUDGET) / (E.TIM_JL_CR_SBI_BUDGET*100)) AS TIM_JL_CR_SBI_VarSBI_PERCENT," & _
			
			"  		A.PEMEL_JALAN_PENGERASAN_MR_MTR as KERAS_JL_MR_BI_AKTUAL_HI," & _
			"  		COALESCE(SUM(A.PEMEL_JALAN_PENGERASAN_MR_MTR) OVER(PARTITION BY A.FCBA ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING), A.PEMEL_JALAN_PENGERASAN_MR_MTR ) AS KERAS_JL_MR_BI_AKTUAL_SHI," & _
			"  		(B.PEMEL_JALAN_PENGERASAN_MR_MTR / C.WORKINGDAY) as KERAS_JL_MR_BI_BUDGET_HI," & _	
			"  		COALESCE(SUM(B.PEMEL_JALAN_PENGERASAN_MR_MTR / C.WORKINGDAY) OVER (PARTITION BY B.FCBA	ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING ), (B.PEMEL_JALAN_PENGERASAN_MR_MTR / C.WORKINGDAY) ) AS KERAS_JL_MR_BI_BUDGET_SHI," & _
			"  		( COALESCE(SUM(A.PEMEL_JALAN_PENGERASAN_MR_MTR) OVER (PARTITION BY A.FCBA ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING ), A.PEMEL_JALAN_PENGERASAN_MR_MTR ) " & _
			"  		  -" & _		
			"  		  COALESCE(SUM(B.PEMEL_JALAN_PENGERASAN_MR_MTR / C.WORKINGDAY) OVER (PARTITION BY B.FCBA ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING ), (B.PEMEL_JALAN_PENGERASAN_MR_MTR / C.WORKINGDAY))" & _
			"  		) as KERAS_JL_MR_BI_VarSHI_MTR," & _	
			"  		(" & _
			"  			( COALESCE(SUM(A.PEMEL_JALAN_PENGERASAN_MR_MTR) OVER (PARTITION BY A.FCBA ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING ), A.PEMEL_JALAN_PENGERASAN_MR_MTR) " & _	
			"  			  - " & _
			"  			  COALESCE(SUM(B.PEMEL_JALAN_PENGERASAN_MR_MTR / C.WORKINGDAY) OVER (PARTITION BY B.FCBA ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING ), (B.PEMEL_JALAN_PENGERASAN_MR_MTR / C.WORKINGDAY))" & _	
			"  			)" & _
			"  			/" & _	
			"  			( COALESCE(SUM(B.PEMEL_JALAN_PENGERASAN_MR_MTR / C.WORKINGDAY) OVER (PARTITION BY B.FCBA ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING ), (B.PEMEL_JALAN_PENGERASAN_MR_MTR / C.WORKINGDAY)) * 100 " & _
			"  			)" & _	
			"  		) as KERAS_JL_MR_BI_VarSHI_PERCENT" & _
			"  		,D.KERAS_JL_MR_SBI_AKTUAL ,E.KERAS_JL_MR_SBI_BUDGET" & _	
			"  		,(D.KERAS_JL_MR_SBI_AKTUAL - E.KERAS_JL_MR_SBI_BUDGET) AS KERAS_JL_MR_SBI_VarSBI_MTR" & _
			"  		,((D.KERAS_JL_MR_SBI_AKTUAL - E.KERAS_JL_MR_SBI_BUDGET) / (E.KERAS_JL_MR_SBI_BUDGET*100)) AS KERAS_JL_MR_SBI_VarSBI_PERCENT," & _	

			"  		A.PEMEL_JALAN_PENGERASAN_CR_MTR as KERAS_JL_CR_BI_AKTUAL_HI, " & _
			"  		COALESCE(SUM(A.PEMEL_JALAN_PENGERASAN_CR_MTR) OVER(PARTITION BY A.FCBA ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING), A.PEMEL_JALAN_PENGERASAN_CR_MTR ) AS KERAS_JL_CR_BI_AKTUAL_SHI," & _	
			"  		(B.PEMEL_JALAN_PENGERASAN_CR_MTR / C.WORKINGDAY) as KERAS_JL_CR_BI_BUDGET_HI," & _
			"  		COALESCE(SUM(B.PEMEL_JALAN_PENGERASAN_CR_MTR / C.WORKINGDAY) OVER (PARTITION BY B.FCBA	ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING ), (B.PEMEL_JALAN_PENGERASAN_CR_MTR / C.WORKINGDAY)) AS KERAS_JL_CR_BI_BUDGET_SHI," & _	
			"  		( COALESCE(SUM(A.PEMEL_JALAN_PENGERASAN_CR_MTR) OVER (PARTITION BY A.FCBA ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING ), A.PEMEL_JALAN_PENGERASAN_CR_MTR ) " & _
			"  		  -" & _	
			"  		  COALESCE(SUM(B.PEMEL_JALAN_PENGERASAN_CR_MTR / C.WORKINGDAY) OVER (PARTITION BY B.FCBA ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING ), (B.PEMEL_JALAN_PENGERASAN_CR_MTR / C.WORKINGDAY))" & _
			"  		) as KERAS_JL_CR_BI_VarSHI_MTR," & _		
			"  		(" & _
			"  			( COALESCE(SUM(A.PEMEL_JALAN_PENGERASAN_CR_MTR) OVER (PARTITION BY A.FCBA ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING ), A.PEMEL_JALAN_PENGERASAN_CR_MTR) " & _	
			"  			  - " & _
			"  			  COALESCE(SUM(B.PEMEL_JALAN_PENGERASAN_CR_MTR / C.WORKINGDAY) OVER (PARTITION BY B.FCBA ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING ), (B.PEMEL_JALAN_PENGERASAN_CR_MTR / C.WORKINGDAY))" & _	
			"  			)" & _
			"  			/" & _	
			"  			(" & _
			"  			  COALESCE(SUM(B.PEMEL_JALAN_PENGERASAN_CR_MTR / C.WORKINGDAY) OVER (PARTITION BY B.FCBA ORDER BY A.TRANS_DATE ROWS BETWEEN UNBOUNDED PRECEDING AND 0 PRECEDING ), (B.PEMEL_JALAN_PENGERASAN_CR_MTR / C.WORKINGDAY)) * 100 " & _	
			"  			)" & _
			"  		 ) as KERAS_JL_CR_BI_VarSHI_PERCENT" & _	
			"  		,D.KERAS_JL_CR_SBI_AKTUAL ,E.KERAS_JL_CR_SBI_BUDGET" & _
			"  		,(D.KERAS_JL_CR_SBI_AKTUAL - E.KERAS_JL_CR_SBI_BUDGET) AS KERAS_JL_CR_SBI_VarSBI_MTR" & _	
			"  		,((D.KERAS_JL_CR_SBI_AKTUAL - E.KERAS_JL_CR_SBI_BUDGET) / (E.KERAS_JL_CR_SBI_BUDGET*100)) AS KERAS_JL_CR_SBI_VarSBI_PERCENT" & _			
			" FROM LHK_NON_PROD A " & _
			" INNER JOIN LHK_NON_PROD_BUDGET B ON B.FCBA = A.FCBA AND TO_CHAR(B.TRANS_DATE,'yyyymm') = TO_CHAR(A.TRANS_DATE,'yyyymm') " & _
			" INNER JOIN ( " & _
			"  		SELECT COUNT(*) as WORKINGDAY, FCBA, BUDGET_YEAR " & _
			"  		FROM BUDGET_TRANSACTION_WORKINGDAY " & _
			"  		WHERE FCBA ='" & Session("fcBA") & "'" & _	
			"  		AND BUDGET_YEAR =TO_CHAR(to_date('" & v_txtDateTo & "', 'mm/dd/yyyy'),'yyyy')" & _		
			"  		AND TO_CHAR(WORKINGDATE,'yyyymm') = TO_CHAR(to_date('" & v_txtDateTo & "', 'mm/dd/yyyy'),'yyyymm')" & _	
			"  		GROUP BY FCBA, BUDGET_YEAR " & _
			"  		) C ON C.FCBA = B.FCBA " & _
			" INNER JOIN ( " & _
			"  		SELECT " & _
			"  			SUM(STACKING_HA) AS STACKING_SBI_AKTUAL," & _
			"  			SUM(TANAMBARU_PKK) AS TANAMBARU_SBI_AKTUAL, " & _
			"  			SUM(KOMPLETING_PKK) AS KOMPLETING_SBI_AKTUAL, " & _
			"  			SUM(TANAMSISIP_TBM_PKK) AS SISIP_TBM_SBI_AKTUAL," & _
			"  			SUM(TANAMSISIP_TM_PKK) AS SISIP_TM_SBI_AKTUAL," & _
			"  			SUM(KASTRASI_HA) AS KASTRASI_SBI_AKTUAL," & _			
			"  			SUM(SANITASI_HA) AS SANITASI_SBI_AKTUAL," & _
			"  			SUM(PEMEL_JALAN_TIMBUN_MR_MTR) AS TIM_JL_MR_SBI_AKTUAL," & _
			"  			SUM(PEMEL_JALAN_TIMBUN_CR_MTR) AS TIM_JL_CR_SBI_AKTUAL," & _
			"  			SUM(PEMEL_JALAN_PENGERASAN_MR_MTR) AS KERAS_JL_MR_SBI_AKTUAL," & _
			"  			SUM(PEMEL_JALAN_PENGERASAN_CR_MTR) AS KERAS_JL_CR_SBI_AKTUAL," & _
			"  			FCBA " & _
			"  		FROM LHK_NON_PROD" & _
			"  		WHERE FCBA ='" & Session("fcBA") & "'" & _
			"  		AND TO_CHAR(TRANS_DATE,'yyyymm') <= TO_CHAR(to_date('" & v_txtDateTo & "', 'mm/dd/yyyy'),'yyyymm')" & _	
			"  		GROUP BY FCBA" & _
			"  		) D ON D.FCBA = B.FCBA" & _
			" INNER JOIN ( 	" & _
			"  		SELECT " & _
			"  			SUM(STACKING_HA) AS STACKING_SBI_BUDGET, " & _
			"  			SUM(TANAMBARU_PKK) AS TANAMBARU_SBI_BUDGET, " & _
			"  			SUM(KOMPLETING_PKK) AS KOMPLETING_SBI_BUDGET," & _
			"  			SUM(TANAMSISIP_TBM_PKK) AS SISIP_TBM_SBI_BUDGET," & _
			"  			SUM(TANAMSISIP_TM_PKK) AS SISIP_TM_SBI_BUDGET," & _
			"  			SUM(KASTRASI_HA) AS KASTRASI_SBI_BUDGET," & _
			"  			SUM(SANITASI_HA) AS SANITASI_SBI_BUDGET," & _
			"  			SUM(PEMEL_JALAN_TIMBUN_MR_MTR) AS TIM_JL_MR_SBI_BUDGET," & _
			"  			SUM(PEMEL_JALAN_TIMBUN_CR_MTR) AS TIM_JL_CR_SBI_BUDGET," & _
			"  			SUM(PEMEL_JALAN_PENGERASAN_MR_MTR) AS KERAS_JL_MR_SBI_BUDGET," & _
			"  			SUM(PEMEL_JALAN_PENGERASAN_CR_MTR) AS KERAS_JL_CR_SBI_BUDGET," & _
			"  			FCBA" & _		
			"  		FROM LHK_NON_PROD_BUDGET" & _
			"  		WHERE FCBA ='" & Session("fcBA") & "'" & _		
			"  		AND TO_CHAR(TRANS_DATE,'yyyymm') <= TO_CHAR(to_date('" & v_txtDateTo & "', 'mm/dd/yyyy'),'yyyymm')" & _				
			"  		GROUP BY FCBA" & _
			"  		) E ON E.FCBA = B.FCBA" & _
			" WHERE A.FCBA ='" & Session("fcBA") & "'" & _
			" AND TO_CHAR(A.TRANS_DATE,'yyyymm') = TO_CHAR(to_date('" & v_txtDateTo & "', 'mm/dd/yyyy'),'yyyymm')" & _
			" AND TO_CHAR(A.TRANS_DATE,'yyyymmdd') <= TO_CHAR(to_date('" & v_txtDateTo & "', 'mm/dd/yyyy'),'yyyymmdd')"			

			
' if  len(trim(v_txtDateFrom)) > 0 then
	' Sqlstr = Sqlstr & " AND PRDATE >=	to_Date('" & v_txtDateFrom & "','MM/DD/YYYY')"
' end if
' if  len(trim(v_txtDateTo)) > 0 then
	' Sqlstr = Sqlstr &  " AND PRDATE <= to_Date('" & v_txtDateTo & "','MM/DD/YYYY')"
' end if
' if  len(trim(v_txtFCBA)) > 0 then
	' Sqlstr = Sqlstr &  " AND Erp_Report_Api.GetBusinessUnit_FCBA(SUBSTR(p.PLANT, 1, 2), p.STORLOC) = '" & v_txtFCBA & "'"
' end if

''response.write Sqlstr

Set rsSqlstr = Server.CreateObject("ADODB.Recordset")
rsSqlstr.Open Sqlstr, connstr, adOpenStatic

Dim Total_1, Total_2, Total_3, Total_4, Total_5, Total_6, Total_7, Total_8, Total_9, Total_10
Dim Total_11, Total_12, Total_13, Total_14, Total_15, Total_16, Total_17, Total_18, Total_19, Total_20
Dim Total_21, Total_22, Total_23, Total_24, Total_25, Total_26, Total_27, Total_28, Total_29, Total_30
Dim Total_31, Total_32, Total_33, Total_34, Total_35, Total_36, Total_37, Total_38, Total_39, Total_40
Dim Total_41, Total_42, Total_43, Total_44, Total_45, Total_46, Total_47, Total_48, Total_49, Total_50
Dim Total_51, Total_52, Total_53, Total_54, Total_55, Total_56, Total_57, Total_58, Total_59, Total_60
Dim Total_61, Total_62, Total_63, Total_64, Total_65, Total_66, Total_67, Total_68, Total_69, Total_70
Dim Total_71, Total_72, Total_73, Total_74, Total_75, Total_76, Total_77, Total_78, Total_79, Total_80
Dim Total_81, Total_82, Total_83, Total_84, Total_85, Total_86, Total_87, Total_88, Total_89, Total_90
Dim Total_91, Total_92, Total_93, Total_94, Total_95, Total_96, Total_97, Total_98, Total_99, Total_100
Dim Total_101, Total_102, Total_103, Total_104, Total_105, Total_106, Total_107, Total_108, Total_109, Total_110
Dim Total_111, Total_112, Total_113, Total_114, Total_115, Total_116, Total_117, Total_118, Total_119, Total_120
Dim Total_121, Total_122, Total_123, Total_124

%>
<html>

<head>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>ePlaS Reporting - LHK Produksi</title>
<link rel="stylesheet" href="/css/jquery-ui.css" />
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
    <script src="/js/jquery-1.9.1.js"></script>
	<script src="/js/jquery-ui.js"></script>
    <script>
		$(function() {
		$( ".datepicker" ).datepicker();
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

th.rotate-270 {
  height: 200px; /* 1 */
  white-space: nowrap;
  min-width:40px;
  max-width:40px;
  overflow:hidden;
}
th.rotate-270 > div {
	transform: translate(8px /* 2 */ , 80px /* 3 */ ) rotate(-90deg);
	-webkit-transform: translate(8px /* 2 */ , 80px /* 3 */ ) rotate(-90deg);
	-ms-transform: translate(8px /* 2 */ , 80px /* 3 */ ) rotate(-90deg);
    width:15px; /* 4 */
	font-size:10px;
	font-weight:normal;
}

/*
1. Mengatur tinggi table head.
2. mengatur posisi X div
3. mengatur posisi Y div.
4. mengatur lebarnya DIV yang akan mempengaruhi Lebarnya Table Head.
*/
</style>


</head>

<body>
	<div id="wrapper">
	<!--#include file="../print_excel.asp"-->
	<div id="header-wrapper">
	<!--
	<div class="printhide">
		#include file="menutop.asp"
	</div>
    -->
	<table id="form-table" class="centered-margin">
		<form action="<%=ScriptName%>" method="post" name="frmEntry" onSubmit="javascript:return validasi(this)" >
			<tr>

				
				<td><font face="verdana" size="1">Sampai Hari ini</font></td>
				<td align="center">
					<% if v_txtDateTo <> "" and v_txtDateTo = v_txtDateTo then %>
						<input class="datepicker" type="text" maxlength="10" name="txtDateTo" size="10" value="<%=v_txtDateTo%>" title="Type Date mm/dd/yyyy" >
					<% else %>
						<input class="datepicker" type="text" maxlength="10" name="txtDateTo" size="10" value="<%=textDate(Date)%>" title="Type Date mm/dd/yyyy" >
					<% end if%>
				</td>
				
				<td><font face="verdana" size="1">FCBA</font></td>
				<td>
					<select name="txtFCBA" ><option value=""></option>
						<%
							Dim rsFCBA, strFCBA
							strFCBA	=	" SELECT FCBA FROM PARAMETER WHERE  parcode = 'PR' AND fccode = 'PRINFO' "
							strFCBA	=	strFCBA & " ORDER BY FCBA ASC"
							
							
							Set rsFCBA = Server.CreateObject("ADODB.Recordset")
							rsFCBA.Open strFCBA, connstr, adOpenStatic 
							
							if not rsFCBA.eof then
							Do While Not rsFCBA.eof
							if rsFCBA("FCBA") = v_txtFCBA then%>
									<option value="<%=rsFCBA("FCBA")%>" Selected><%=rsFCBA("FCBA")%></option> 
								 <%Else%>
									<option value="<%=rsFCBA("FCBA")%>"><%=rsFCBA("FCBA")%></option>  
								<%End if
								rsFCBA.movenext
							loop
							end if
							rsFCBA.Close
							set rsFCBA = Nothing
						%>
					</select>
				</td>
				  
				<td><input type="submit" value="View" /></td>
			</tr>
			
			<!--
			<div class="printhide">
				<table id="form-table" class="centered-margin">
				</table>
			</div>
			 -->
			
		</form>
	</table>

<table border="0" width="100%" id="product-table" class="xthf-site" >  <!-- -->
	<thead style="border-left:1px #000000 solid;">
		<tr>
			<th colspan="125" style="background:#FFFFFF; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
				<div id="header-wrapper">
					<div id="head_judul">
						<h4><%If Len(Trim(varCompanyName)) > 0 Then%>PT. <%End If%><%=varCompanyName%></h4>
						<h5><%=v_txtPabrik%></h5>
					</div>
				
					<div id="subheader" class="centered"  style="width:10%; margin:auto;">
						<h1>Laporan LHK Produksi</h1>
						<br/>
						<h4>Sampai Hari ini &nbsp;:&nbsp;<%=v_txtDateTo%></h4>
					</div>
				</div>
			</th>
	   </tr>
	         
			<!-- Pengaturan Baris 1-->
			<tr>
				<th rowspan="4" style="background-color:#00A300;color:#fff;">TANGGAL</th>
				<th colspan="10" style="background-color:#00A300;color:#fff;">STACKING (Ha)</th>
				<th colspan="10" style="background-color:#00A300;color:#fff;">TANAM BARU (Pkk)</th>
				<th colspan="10" style="background-color:#00A300;color:#fff;">KOMPLETING (Pkk)</th>
				
				<th colspan="10" style="background-color:#00A300;color:#fff;">TANAM SISIP TBM (Pkk)</th>
				<th colspan="10" style="background-color:#00A300;color:#fff;">TANAM SISIP TM (Pkk)</th>
				<th colspan="10" style="background-color:#00A300;color:#fff;">KASTRASI (Ha)</th>
				<th colspan="10" style="background-color:#00A300;color:#fff;">SANITASI (Ha)</th>

				<th colspan="14" style="background-color:#00A300;color:#fff;">TK PEMELIHARAAN</th>

				<th colspan="10" style="background-color:#00A300;color:#fff;">TIMBUN JALAN MR (MTR)</th>
				<th colspan="10" style="background-color:#00A300;color:#fff;">TIMBUN JALAN CR (MTR)</th>
				<th colspan="10" style="background-color:#00A300;color:#fff;">PENGERASAN JALAN MR (MTR)</th>
				<th colspan="10" style="background-color:#00A300;color:#fff;">PENGERASAN JALAN CR (MTR)</th>
				
			</tr>	
			
			<!-- Pengaturan Baris 2-->
			<tr>
				<!-- Stacking-->
				<th colspan="6" style="background-color:#00A300;color:#fff;">BULAN INI  (BI)</th>
				<th colspan="4" style="background-color:#00A300;color:#fff;">SAMPAI BULAN INI (SBI)</th>
				
				<!--TANAM BARU-->
				<th colspan="6" style="background-color:#00A300;color:#fff;">BULAN INI  (BI)</th>
				<th colspan="4" style="background-color:#00A300;color:#fff;">SAMPAI BULAN INI (SBI)</th>
				
				<!--KOMPLETING-->
				<th colspan="6" style="background-color:#00A300;color:#fff;">BULAN INI  (BI)</th>
				<th colspan="4" style="background-color:#00A300;color:#fff;">SAMPAI BULAN INI (SBI)</th>
				
				<!--TANAM SISIP TBM (Pkk)-->
				<th colspan="6" style="background-color:#00A300;color:#fff;">BULAN INI  (BI)</th>
				<th colspan="4" style="background-color:#00A300;color:#fff;">SAMPAI BULAN INI (SBI)</th>
				
				<!--TANAM SISIP TM (Pkk)-->
				<th colspan="6" style="background-color:#00A300;color:#fff;">BULAN INI  (BI)</th>
				<th colspan="4" style="background-color:#00A300;color:#fff;">SAMPAI BULAN INI (SBI)</th>
				
				<!--KASTRASI (Ha)-->
				<th colspan="6" style="background-color:#00A300;color:#fff;">BULAN INI  (BI)</th>
				<th colspan="4" style="background-color:#00A300;color:#fff;">SAMPAI BULAN INI (SBI)</th>
				
				<!--SANITASI (Ha)-->
				<th colspan="6" style="background-color:#00A300;color:#fff;">BULAN INI  (BI)</th>
				<th colspan="4" style="background-color:#00A300;color:#fff;">SAMPAI BULAN INI (SBI)</th>
				
				<!--TK PEMELIHARAAN-->
				<th colspan="14" style="background-color:#00A300;color:#fff;">AKTUAL</th>
				
				<!--TIMBUN JALAN MR (MTR)-->
				<th colspan="6" style="background-color:#00A300;color:#fff;">BULAN INI  (BI)</th>
				<th colspan="4" style="background-color:#00A300;color:#fff;">SAMPAI BULAN INI (SBI)</th>
				
				<!--TIMBUN JALAN CR (MTR)-->
				<th colspan="6" style="background-color:#00A300;color:#fff;">BULAN INI  (BI)</th>
				<th colspan="4" style="background-color:#00A300;color:#fff;">SAMPAI BULAN INI (SBI)</th>
				
				<!--PENGERASAN JALAN MR (MTR)-->
				<th colspan="6" style="background-color:#00A300;color:#fff;">BULAN INI  (BI)</th>
				<th colspan="4" style="background-color:#00A300;color:#fff;">SAMPAI BULAN INI (SBI)</th>
				
				<!--PENGERASAN JALAN CR (MTR)-->
				<th colspan="6" style="background-color:#00A300;color:#fff;">BULAN INI  (BI)</th>
				<th colspan="4" style="background-color:#00A300;color:#fff;">SAMPAI BULAN INI (SBI)</th>
			</tr>
			
			<!-- Pengaturan Baris 3-->
			<tr>
				<!-- Stacking-->
				<th colspan="2" style="background-color:#00A300;color:#fff;">AKTUAL</th>
				<th colspan="2" style="background-color:#00A300;color:#fff;">BUDGET</th>
				<th colspan="2" style="background-color:#00A300;color:#fff;">Var SHI</th>
				<th rowspan="2" style="background-color:#00A300;color:#fff;">AKTUAL</th>
				<th rowspan="2" style="background-color:#00A300;color:#fff;">BUDGET</th>
				<th colspan="2" style="background-color:#00A300;color:#fff;">Var SBI</th>
				
				<!--TANAM BARU-->
				<th colspan="2" style="background-color:#00A300;color:#fff;">AKTUAL</th>
				<th colspan="2" style="background-color:#00A300;color:#fff;">BUDGET</th>
				<th colspan="2" style="background-color:#00A300;color:#fff;">Var SHI</th>
				<th rowspan="2" style="background-color:#00A300;color:#fff;">AKTUAL</th>
				<th rowspan="2" style="background-color:#00A300;color:#fff;">BUDGET</th>
				<th colspan="2" style="background-color:#00A300;color:#fff;">Var SBI</th>
				
				<!--KOMPLETING-->
				<th colspan="2" style="background-color:#00A300;color:#fff;">AKTUAL</th>
				<th colspan="2" style="background-color:#00A300;color:#fff;">BUDGET</th>
				<th colspan="2" style="background-color:#00A300;color:#fff;">Var SHI</th>
				<th rowspan="2" style="background-color:#00A300;color:#fff;">AKTUAL</th>
				<th rowspan="2" style="background-color:#00A300;color:#fff;">BUDGET</th>
				<th colspan="2" style="background-color:#00A300;color:#fff;">Var SBI</th>
				
				<!--TANAM SISIP TBM (Pkk)-->
				<th colspan="2" style="background-color:#00A300;color:#fff;">AKTUAL</th>
				<th colspan="2" style="background-color:#00A300;color:#fff;">BUDGET</th>
				<th colspan="2" style="background-color:#00A300;color:#fff;">Var SHI</th>
				<th rowspan="2" style="background-color:#00A300;color:#fff;">AKTUAL</th>
				<th rowspan="2" style="background-color:#00A300;color:#fff;">BUDGET</th>
				<th colspan="2" style="background-color:#00A300;color:#fff;">Var SBI</th>
				
				<!--TANAM SISIP TM (Pkk)-->
				<th colspan="2" style="background-color:#00A300;color:#fff;">AKTUAL</th>
				<th colspan="2" style="background-color:#00A300;color:#fff;">BUDGET</th>
				<th colspan="2" style="background-color:#00A300;color:#fff;">Var SHI</th>
				<th rowspan="2" style="background-color:#00A300;color:#fff;">AKTUAL</th>
				<th rowspan="2" style="background-color:#00A300;color:#fff;">BUDGET</th>
				<th colspan="2" style="background-color:#00A300;color:#fff;">Var SBI</th>
				
				<!--KASTRASI (Ha)-->
				<th colspan="2" style="background-color:#00A300;color:#fff;">AKTUAL</th>
				<th colspan="2" style="background-color:#00A300;color:#fff;">BUDGET</th>
				<th colspan="2" style="background-color:#00A300;color:#fff;">Var SHI</th>
				<th rowspan="2" style="background-color:#00A300;color:#fff;">AKTUAL</th>
				<th rowspan="2" style="background-color:#00A300;color:#fff;">BUDGET</th>
				<th colspan="2" style="background-color:#00A300;color:#fff;">Var SBI</th>
				
				<!--SANITASI (Ha)-->
				<th colspan="2" style="background-color:#00A300;color:#fff;">AKTUAL</th>
				<th colspan="2" style="background-color:#00A300;color:#fff;">BUDGET</th>
				<th colspan="2" style="background-color:#00A300;color:#fff;">Var SHI</th>
				<th rowspan="2" style="background-color:#00A300;color:#fff;">AKTUAL</th>
				<th rowspan="2" style="background-color:#00A300;color:#fff;">BUDGET</th>
				<th colspan="2" style="background-color:#00A300;color:#fff;">Var SBI</th>
				
				<!--TK PEMELIHARAAN-->
				<th colspan="2" style="background-color:#00A300;color:#fff;">TK TERSEDIA</th>
				<th colspan="2" style="background-color:#00A300;color:#fff;">TK TANAM BARU</th>
				<th colspan="2" style="background-color:#00A300;color:#fff;">TK SISIP</th>
				<th colspan="2" style="background-color:#00A300;color:#fff;">TK KOMPLETING</th>
				<th colspan="2" style="background-color:#00A300;color:#fff;">TK PEMUPUKAN</th>
				<th colspan="2" style="background-color:#00A300;color:#fff;">TK PEMEL LAINNYA</th>
				<th colspan="2" style="background-color:#00A300;color:#fff;">% KEHADIRAN TK PEMEL</th>
				
				<!--TIMBUN JALAN MR (MTR)-->
				<th colspan="2" style="background-color:#00A300;color:#fff;">AKTUAL</th>
				<th colspan="2" style="background-color:#00A300;color:#fff;">BUDGET</th>
				<th colspan="2" style="background-color:#00A300;color:#fff;">Var SHI</th>
				<th rowspan="2" style="background-color:#00A300;color:#fff;">AKTUAL</th>
				<th rowspan="2" style="background-color:#00A300;color:#fff;">BUDGET</th>
				<th colspan="2" style="background-color:#00A300;color:#fff;">Var SBI</th>
				
				<!--TIMBUN JALAN CR (MTR)-->
				<th colspan="2" style="background-color:#00A300;color:#fff;">AKTUAL</th>
				<th colspan="2" style="background-color:#00A300;color:#fff;">BUDGET</th>
				<th colspan="2" style="background-color:#00A300;color:#fff;">Var SHI</th>
				<th rowspan="2" style="background-color:#00A300;color:#fff;">AKTUAL</th>
				<th rowspan="2" style="background-color:#00A300;color:#fff;">BUDGET</th>
				<th colspan="2" style="background-color:#00A300;color:#fff;">Var SBI</th>
				
				<!--PENGERASAN JALAN MR (MTR)-->
				<th colspan="2" style="background-color:#00A300;color:#fff;">AKTUAL</th>
				<th colspan="2" style="background-color:#00A300;color:#fff;">BUDGET</th>
				<th colspan="2" style="background-color:#00A300;color:#fff;">Var SHI</th>
				<th rowspan="2" style="background-color:#00A300;color:#fff;">AKTUAL</th>
				<th rowspan="2" style="background-color:#00A300;color:#fff;">BUDGET</th>
				<th colspan="2" style="background-color:#00A300;color:#fff;">Var SBI</th>
				
				<!--PENGERASAN JALAN CR (MTR)-->
				<th colspan="2" style="background-color:#00A300;color:#fff;">AKTUAL</th>
				<th colspan="2" style="background-color:#00A300;color:#fff;">BUDGET</th>
				<th colspan="2" style="background-color:#00A300;color:#fff;">Var SHI</th>
				<th rowspan="2" style="background-color:#00A300;color:#fff;">AKTUAL</th>
				<th rowspan="2" style="background-color:#00A300;color:#fff;">BUDGET</th>
				<th colspan="2" style="background-color:#00A300;color:#fff;">Var SBI</th>
			</tr>
	
			<!-- Pengaturan Baris 4-->
			<tr>	
				<!-- Stacking-->
				<th  style="background-color:#00A300;color:#fff;">HI</th>
				<th  style="background-color:#00A300;color:#fff;">SHI</th>
				<th  style="background-color:#00A300;color:#fff;">HI</th>
				<th  style="background-color:#00A300;color:#fff;">SHI</th>
				<th  style="background-color:#00A300;color:#fff;">HA</th>
				<th  style="background-color:#00A300;color:#fff;">%</th>
				<th  style="background-color:#00A300;color:#fff;">HA</th>
				<th  style="background-color:#00A300;color:#fff;">%</th>
				
				<!--TANAM BARU-->
				<th  style="background-color:#00A300;color:#fff;">HI</th>
				<th  style="background-color:#00A300;color:#fff;">SHI</th>
				<th  style="background-color:#00A300;color:#fff;">HI</th>
				<th  style="background-color:#00A300;color:#fff;">SHI</th>
				<th  style="background-color:#00A300;color:#fff;">HA</th>
				<th  style="background-color:#00A300;color:#fff;">%</th>
				<th  style="background-color:#00A300;color:#fff;">HA</th>
				<th  style="background-color:#00A300;color:#fff;">%</th>
				
				<!--KOMPLETING-->
				<th  style="background-color:#00A300;color:#fff;">HI</th>
				<th  style="background-color:#00A300;color:#fff;">SHI</th>
				<th  style="background-color:#00A300;color:#fff;">HI</th>
				<th  style="background-color:#00A300;color:#fff;">SHI</th>
				<th  style="background-color:#00A300;color:#fff;">HA</th>
				<th  style="background-color:#00A300;color:#fff;">%</th>
				<th  style="background-color:#00A300;color:#fff;">HA</th>
				<th  style="background-color:#00A300;color:#fff;">%</th>
				
				<!--TANAM SISIP TBM (Pkk)-->
				<th  style="background-color:#00A300;color:#fff;">HI</th>
				<th  style="background-color:#00A300;color:#fff;">SHI</th>
				<th  style="background-color:#00A300;color:#fff;">HI</th>
				<th  style="background-color:#00A300;color:#fff;">SHI</th>
				<th  style="background-color:#00A300;color:#fff;">HA</th>
				<th  style="background-color:#00A300;color:#fff;">%</th>
				<th  style="background-color:#00A300;color:#fff;">HA</th>
				<th  style="background-color:#00A300;color:#fff;">%</th>
				
				<!--TANAM SISIP TM (Pkk)-->
				<th  style="background-color:#00A300;color:#fff;">HI</th>
				<th  style="background-color:#00A300;color:#fff;">SHI</th>
				<th  style="background-color:#00A300;color:#fff;">HI</th>
				<th  style="background-color:#00A300;color:#fff;">SHI</th>
				<th  style="background-color:#00A300;color:#fff;">HA</th>
				<th  style="background-color:#00A300;color:#fff;">%</th>
				<th  style="background-color:#00A300;color:#fff;">HA</th>
				<th  style="background-color:#00A300;color:#fff;">%</th>
				
				<!--KASTRASI (Ha)-->
				<th  style="background-color:#00A300;color:#fff;">HI</th>
				<th  style="background-color:#00A300;color:#fff;">SHI</th>
				<th  style="background-color:#00A300;color:#fff;">HI</th>
				<th  style="background-color:#00A300;color:#fff;">SHI</th>
				<th  style="background-color:#00A300;color:#fff;">HA</th>
				<th  style="background-color:#00A300;color:#fff;">%</th>
				<th  style="background-color:#00A300;color:#fff;">HA</th>
				<th  style="background-color:#00A300;color:#fff;">%</th>
				
				<!--SANITASI (Ha)-->
				<th  style="background-color:#00A300;color:#fff;">HI</th>
				<th  style="background-color:#00A300;color:#fff;">SHI</th>
				<th  style="background-color:#00A300;color:#fff;">HI</th>
				<th  style="background-color:#00A300;color:#fff;">SHI</th>
				<th  style="background-color:#00A300;color:#fff;">HA</th>
				<th  style="background-color:#00A300;color:#fff;">%</th>
				<th  style="background-color:#00A300;color:#fff;">HA</th>
				<th  style="background-color:#00A300;color:#fff;">%</th>
				
				<!--TK PEMELIHARAAN-->
				<th  style="background-color:#00A300;color:#fff;">HI</th>
				<th  style="background-color:#00A300;color:#fff;">SHI</th>
				<th  style="background-color:#00A300;color:#fff;">HI</th>
				<th  style="background-color:#00A300;color:#fff;">SHI</th>
				<th  style="background-color:#00A300;color:#fff;">HI</th>
				<th  style="background-color:#00A300;color:#fff;">SHI</th>
				<th  style="background-color:#00A300;color:#fff;">HI</th>
				<th  style="background-color:#00A300;color:#fff;">SHI</th>
				<th  style="background-color:#00A300;color:#fff;">HI</th>
				<th  style="background-color:#00A300;color:#fff;">SHI</th>
				<th  style="background-color:#00A300;color:#fff;">HI</th>
				<th  style="background-color:#00A300;color:#fff;">SHI</th>
				<th  style="background-color:#00A300;color:#fff;">HI</th>
				<th  style="background-color:#00A300;color:#fff;">SHI</th>
				
				<!--TIMBUN JALAN MR (MTR)-->
				<th  style="background-color:#00A300;color:#fff;">HI</th>
				<th  style="background-color:#00A300;color:#fff;">SHI</th>
				<th  style="background-color:#00A300;color:#fff;">HI</th>
				<th  style="background-color:#00A300;color:#fff;">SHI</th>
				<th  style="background-color:#00A300;color:#fff;">MTR</th>
				<th  style="background-color:#00A300;color:#fff;">%</th>
				<th  style="background-color:#00A300;color:#fff;">MTR</th>
				<th  style="background-color:#00A300;color:#fff;">%</th>
				
				<!--TIMBUN JALAN CR (MTR)-->
				<th  style="background-color:#00A300;color:#fff;">HI</th>
				<th  style="background-color:#00A300;color:#fff;">SHI</th>
				<th  style="background-color:#00A300;color:#fff;">HI</th>
				<th  style="background-color:#00A300;color:#fff;">SHI</th>
				<th  style="background-color:#00A300;color:#fff;">MTR</th>
				<th  style="background-color:#00A300;color:#fff;">%</th>
				<th  style="background-color:#00A300;color:#fff;">MTR</th>
				<th  style="background-color:#00A300;color:#fff;">%</th>
				
				<!--PENGERASAN JALAN MR (MTR)-->
				<th  style="background-color:#00A300;color:#fff;">HI</th>
				<th  style="background-color:#00A300;color:#fff;">SHI</th>
				<th  style="background-color:#00A300;color:#fff;">HI</th>
				<th  style="background-color:#00A300;color:#fff;">SHI</th>
				<th  style="background-color:#00A300;color:#fff;">MTR</th>
				<th  style="background-color:#00A300;color:#fff;">%</th>
				<th  style="background-color:#00A300;color:#fff;">MTR</th>
				<th  style="background-color:#00A300;color:#fff;">%</th>
				
				<!--PENGERASAN JALAN CR (MTR)-->
				<th  style="background-color:#00A300;color:#fff;">HI</th>
				<th  style="background-color:#00A300;color:#fff;">SHI</th>
				<th  style="background-color:#00A300;color:#fff;">HI</th>
				<th  style="background-color:#00A300;color:#fff;">SHI</th>
				<th  style="background-color:#00A300;color:#fff;">MTR</th>
				<th  style="background-color:#00A300;color:#fff;">%</th>
				<th  style="background-color:#00A300;color:#fff;">MTR</th>
				<th  style="background-color:#00A300;color:#fff;">%</th>
			</tr>
			
		</thead>
	<tbody>
<%
if rsSqlstr.bof and rsSqlstr.eof then
%>
	<tr><td colspan="125" align="center" >NO DATA AVAILABLE</td></tr>
<%else
		''For Calculate Subtotal Or Total, Set variable Before Looping
		'' Set Variabel Total HI
		Total_1  = 0 : Total_3 = 0 : Total_5 = 0 : Total_9 = 0 :  		'' Get Total Stacking
		Total_11 = 0 : Total_13 = 0 : Total_15 = 0 : Total_19 = 0 :		'' Get Total Tanam Baru
		Total_21 = 0 : Total_23 = 0 : Total_25 = 0 : Total_29 = 0 :		'' Get Total Kompleting
		Total_31 = 0 : Total_33 = 0 : Total_35 = 0 : Total_39 = 0 :		'' Get Total Tanam Sisip Tbm
		Total_41 = 0 : Total_43 = 0 : Total_45 = 0 : Total_49 = 0 :		'' Get Total Tanam Sisip Tm
		Total_51 = 0 : Total_53 = 0 : Total_55 = 0 : Total_59 = 0 :		'' Get Total Kastrasi
		Total_61 = 0 : Total_63 = 0 : Total_65 = 0 : Total_69 = 0 :		'' Get Total Sanitasi
		Total_71 = 0 : Total_73 = 0 : Total_75 = 0 : Total_77 = 0 : Total_79 = 0 : Total_81 = 0 :	'' Get Total TK Pemeliharaan
		Total_85 = 0 : Total_87 = 0 : Total_89 = 0 : Total_93 = 0 :		'' Get Total Timbun Jalan MR
		Total_95 = 0 : Total_97 = 0 : Total_99 = 0 : Total_103 = 0 :	'' Get Total Timbun Jalan CR
		Total_105 = 0 : Total_107 = 0 : Total_109 = 0 : Total_113 = 0 :	'' Get Total Pengerasan Jalan MR
		Total_115 = 0 : Total_117 = 0 : Total_119 = 0 : Total_123 = 0 :	'' Get Total Pengerasan Jalan CR
		
		NO = 1
		Do while not rsSqlstr.eof
			if NO  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
			
		'' Set Variabel Total SHI
		Total_2 = 0 : Total_4 = 0 : Total_6 = 0 : Total_7 = 0 : Total_8 = 0 : Total_10 = 0 : '' Get Total Stacking
		Total_12 = 0 : Total_14 = 0 : Total_16 = 0 : Total_17 = 0 : Total_18 = 0 : Total_20 = 0 :'' Get Total Tanam Baru
		Total_22 = 0 : Total_24 = 0 : Total_26 = 0 : Total_27 = 0 : Total_28 = 0 : Total_30 = 0 :
		Total_32 = 0 : Total_34 = 0 : Total_36 = 0 : Total_37 = 0 : Total_38 = 0 : Total_40 = 0 :
		Total_42 = 0 : Total_44 = 0 : Total_46 = 0 : Total_47 = 0 : Total_48 = 0 : Total_50 = 0 :
		Total_52 = 0 : Total_54 = 0 : Total_56 = 0 : Total_57 = 0 : Total_58 = 0 : Total_60 = 0 :
		Total_62 = 0 : Total_64 = 0 : Total_66 = 0 : Total_67 = 0 : Total_68 = 0 : Total_70 = 0 :
		Total_72 = 0 : Total_74 = 0 : Total_76 = 0 : Total_78 = 0 : Total_80 = 0 : Total_82 = 0 :
		Total_83 = 0 : Total_84 = 0 : Total_86 = 0 : Total_88 = 0 : Total_90 = 0 : Total_91 = 0 :
		Total_92 = 0 : Total_94 = 0 : Total_96 = 0 : Total_98 = 0 : Total_100 = 0 : Total_101 = 0 :
		Total_102 = 0 : Total_104 = 0 : Total_106 = 0 : Total_108 = 0 : Total_110 = 0 : Total_111 = 0 :
		Total_112 = 0 : Total_114 = 0 : Total_116 = 0 : Total_118 = 0 : Total_120 = 0 : Total_121 = 0
		Total_122 = 0 : Total_124 = 0
			
%>
	
	<tr class="<%=cClass%>">
		<td align="right"><%=rsSqlstr("TANGGAL")%></td>
		<!--STACKING -->		
		<td align="right"><%=FormatNumber(rsSqlstr("STACKING_BI_AKTUAL_HI"),0)%></td>
		<td align="right"><%=FormatNumber(rsSqlstr("STACKING_BI_AKTUAL_SHI"),0)%></td>
		<td align="right"><%=FormatNumber(rsSqlstr("STACKING_BI_BUDGET_HI"),0)%></td>
		<td align="right"><%=FormatNumber(rsSqlstr("STACKING_BI_BUDGET_SHI"),0)%></td>		
		<td align="right"><%=FormatNumber(rsSqlstr("STACKING_BI_VarSHI_HA"),0)%></td>
		<td align="right"><%=FormatNumber(rsSqlstr("STACKING_BI_VarSHI_PERCENT"),2)%></td>
		
		<td align="right"><%=FormatNumber(rsSqlstr("STACKING_SBI_AKTUAL"),0)%></td>		
		<td align="right"><%=FormatNumber(rsSqlstr("STACKING_SBI_BUDGET"),0)%></td>		
		<td align="right"><%=FormatNumber(rsSqlstr("STACKING_SBI_VarSBI_HA"),0)%></td>
		<td align="right"><%=FormatNumber(rsSqlstr("STACKING_SBI_VarSBI_PERCENT"),2)%></td>	
		<!--TANAM BARU -->	
		<td align="right"><%=FormatNumber(rsSqlstr("TANAMBARU_BI_AKTUAL_HI"),0)%></td>
		<td align="right"><%=FormatNumber(rsSqlstr("TANAMBARU_BI_AKTUAL_SHI"),0)%></td>
		<td align="right"><%=FormatNumber(rsSqlstr("TANAMBARU_BI_BUDGET_HI"),0)%></td>
		<td align="right"><%=FormatNumber(rsSqlstr("TANAMBARU_BI_BUDGET_SHI"),0)%></td>
		<td align="right"><%=FormatNumber(rsSqlstr("TANAMBARU_BI_VarSHI_HA"),0)%></td>
		<td align="right"><%=FormatNumber(rsSqlstr("TANAMBARU_BI_VarSHI_PERCENT"),2)%></td>
		
		<td align="right"><%=FormatNumber(rsSqlstr("TANAMBARU_SBI_AKTUAL"),0)%></td>
		<td align="right"><%=FormatNumber(rsSqlstr("TANAMBARU_SBI_BUDGET"),0)%></td>
		<td align="right"><%=FormatNumber(rsSqlstr("TANAMBARU_SBI_VarSBI_HA"),0)%></td>
		<td align="right"><%=FormatNumber(rsSqlstr("TANAMBARU_SBI_VarSBI_PERCENT"),2)%></td>		
		<!--KOMPLETING -->	
		<td align="right"><%=FormatNumber(rsSqlstr("KOMPLETING_BI_AKTUAL_HI"),0)%></td>
		<td align="right"><%=FormatNumber(rsSqlstr("KOMPLETING_BI_AKTUAL_SHI"),0)%></td>
		<td align="right"><%=FormatNumber(rsSqlstr("KOMPLETING_BI_BUDGET_HI"),0)%></td>
		<td align="right"><%=FormatNumber(rsSqlstr("KOMPLETING_BI_BUDGET_SHI"),0)%></td>
		<td align="right"><%=FormatNumber(rsSqlstr("KOMPLETING_BI_VarSHI_PKK"),0)%></td>
		<td align="right"><%=FormatNumber(rsSqlstr("KOMPLETING_BI_VarSHI_PERCENT"),2)%></td>
		
		<td align="right"><%=FormatNumber(rsSqlstr("KOMPLETING_SBI_AKTUAL"),0)%></td>
		<td align="right"><%=FormatNumber(rsSqlstr("KOMPLETING_SBI_BUDGET"),0)%></td>	
		<td align="right"><%=FormatNumber(rsSqlstr("KOMPLETING_SBI_VarSBI_HA"),0)%></td>
		<td align="right"><%=FormatNumber(rsSqlstr("KOMPLETING_SBI_VarSBI_PERCENT"),2)%></td>	
		<!--TANAM SISIP TBM -->		
		<td align="right"><%=FormatNumber(rsSqlstr("SISIP_TBM_BI_AKTUAL_HI"),0)%></td>	
		<td align="right"><%=FormatNumber(rsSqlstr("SISIP_TBM_BI_AKTUAL_SHI"),0)%></td>			
		<td align="right"><%=FormatNumber(rsSqlstr("SISIP_TBM_BI_BUDGET_HI"),0)%></td>	
		<td align="right"><%=FormatNumber(rsSqlstr("SISIP_TBM_BI_BUDGET_SHI"),0)%></td>		
		<td align="right"><%=FormatNumber(rsSqlstr("SISIP_TBM_BI_VarSHI_PKK"),0)%></td>	
		<td align="right"><%=FormatNumber(rsSqlstr("SISIP_TBM_BI_VarSHI_PERCENT"),2)%></td>	
		
		<td align="right"><%=FormatNumber(rsSqlstr("SISIP_TBM_SBI_AKTUAL"),0)%></td>	
		<td align="right"><%=FormatNumber(rsSqlstr("SISIP_TBM_SBI_BUDGET"),0)%></td>		
		<td align="right"><%=FormatNumber(rsSqlstr("SISIP_TBM_SBI_VarSBI_PKK"),0)%></td>	
		<td align="right"><%=FormatNumber(rsSqlstr("SISIP_TBM_SBI_VarSBI_PERCENT"),2)%></td>		
		<!--TANAM SISIP TM -->
		<td align="right"><%=FormatNumber(rsSqlstr("SISIP_TM_BI_AKTUAL_HI"),0)%></td>	
		<td align="right"><%=FormatNumber(rsSqlstr("SISIP_TM_BI_AKTUAL_SHI"),0)%></td>			
		<td align="right"><%=FormatNumber(rsSqlstr("SISIP_TM_BI_BUDGET_HI"),0)%></td>	
		<td align="right"><%=FormatNumber(rsSqlstr("SISIP_TM_BI_BUDGET_SHI"),0)%></td>		
		<td align="right"><%=FormatNumber(rsSqlstr("SISIP_TM_BI_VarSHI_PKK"),0)%></td>	
		<td align="right"><%=FormatNumber(rsSqlstr("SISIP_TM_BI_VarSHI_PERCENT"),2)%></td>		
		
		<td align="right"><%=FormatNumber(rsSqlstr("SISIP_TM_SBI_AKTUAL"),0)%></td>	
		<td align="right"><%=FormatNumber(rsSqlstr("SISIP_TM_SBI_BUDGET"),0)%></td>		
		<td align="right"><%=FormatNumber(rsSqlstr("SISIP_TM_SBI_VarSBI_PKK"),0)%></td>	
		<td align="right"><%=FormatNumber(rsSqlstr("SISIP_TM_SBI_VarSBI_PERCENT"),2)%></td>	
		<!--KASTRASI -->
		<td align="right"><%=FormatNumber(rsSqlstr("KASTRASI_BI_AKTUAL_HI"),0)%></td>	
		<td align="right"><%=FormatNumber(rsSqlstr("KASTRASI_BI_AKTUAL_SHI"),0)%></td>			
		<td align="right"><%=FormatNumber(rsSqlstr("KASTRASI_BI_BUDGET_HI"),0)%></td>	
		<td align="right"><%=FormatNumber(rsSqlstr("KASTRASI_BI_BUDGET_SHI"),0)%></td>		
		<td align="right"><%=FormatNumber(rsSqlstr("KASTRASI_BI_VarSHI_PKK"),0)%></td>	
		<td align="right"><%=FormatNumber(rsSqlstr("KASTRASI_BI_VarSHI_PERCENT"),2)%></td>
		
		<td align="right"><%=FormatNumber(rsSqlstr("KASTRASI_SBI_AKTUAL"),0)%></td>	
		<td align="right"><%=FormatNumber(rsSqlstr("KASTRASI_SBI_BUDGET"),0)%></td>		
		<td align="right"><%=FormatNumber(rsSqlstr("KASTRASI_SBI_VarSBI_PKK"),0)%></td>	
		<td align="right"><%=FormatNumber(rsSqlstr("KASTRASI_SBI_VarSBI_PERCENT"),2)%></td>
		<!--SANITASI -->
		<td align="right"><%=FormatNumber(rsSqlstr("SANITASI_BI_AKTUAL_HI"),0)%></td>	
		<td align="right"><%=FormatNumber(rsSqlstr("SANITASI_BI_AKTUAL_SHI"),0)%></td>			
		<td align="right"><%=FormatNumber(rsSqlstr("SANITASI_BI_BUDGET_HI"),0)%></td>	
		<td align="right"><%=FormatNumber(rsSqlstr("SANITASI_BI_BUDGET_SHI"),0)%></td>		
		<td align="right"><%=FormatNumber(rsSqlstr("SANITASI_BI_VarSHI_PKK"),0)%></td>	
		<td align="right"><%=FormatNumber(rsSqlstr("SANITASI_BI_VarSHI_PERCENT"),2)%></td>		
		
		<td align="right"><%=FormatNumber(rsSqlstr("SANITASI_SBI_AKTUAL"),0)%></td>	
		<td align="right"><%=FormatNumber(rsSqlstr("SANITASI_SBI_BUDGET"),0)%></td>		
		<td align="right"><%=FormatNumber(rsSqlstr("SANITASI_SBI_VarSBI_PKK"),0)%></td>	
		<td align="right"><%=FormatNumber(rsSqlstr("SANITASI_SBI_VarSBI_PERCENT"),2)%></td>
		<!--PEMELIHARAAN  -->
		<td align="right"><%=FormatNumber(rsSqlstr("PEMEL_AKTUAL_TERSEDIA_HI"),0)%></td>	
		<td align="right"><%=FormatNumber(rsSqlstr("PEMEL_AKTUAL_TERSEDIA_SHI"),0)%></td>			
		<td align="right"><%=FormatNumber(rsSqlstr("PEMEL_AKTUAL_TANAMBARU_HI"),0)%></td>	
		<td align="right"><%=FormatNumber(rsSqlstr("PEMEL_AKTUAL_TANAMBARU_SHI"),0)%></td>		
		<td align="right"><%=FormatNumber(rsSqlstr("PEMEL_AKTUAL_SISIP_HI"),0)%></td>	
		<td align="right"><%=FormatNumber(rsSqlstr("PEMEL_AKTUAL_SISIP_SHI"),0)%></td>		
		<td align="right"><%=FormatNumber(rsSqlstr("PEMEL_AKTUAL_KOMPLETING_HI"),0)%></td>	
		<td align="right"><%=FormatNumber(rsSqlstr("PEMEL_AKTUAL_KOMPLETING_SHI"),0)%></td>		
		<td align="right"><%=FormatNumber(rsSqlstr("PEMEL_AKTUAL_PEMUPUKAN_HI"),0)%></td>	
		<td align="right"><%=FormatNumber(rsSqlstr("PEMEL_AKTUAL_PEMUPUKAN_SHI"),0)%></td>
		<td align="right"><%=FormatNumber(rsSqlstr("PEMEL_AKTUAL_LAINNYA_HI"),0)%></td>
		<td align="right"><%=FormatNumber(rsSqlstr("PEMEL_AKTUAL_LAINNYA_SHI"),0)%></td>
		<td align="right"><%=FormatNumber(rsSqlstr("PEMEL_AKTUAL_PERCENT_HADIR_HI"),0)%></td>
		<td align="right"><%=FormatNumber(rsSqlstr("PEMEL_AKTUAL_PERCENT_HADIR_SHI"),0)%></td>		
		<!--TIMBUN JALAN MR -->
		<td align="right"><%=FormatNumber(rsSqlstr("TIM_JL_MR_BI_AKTUAL_HI"),0)%></td>	
		<td align="right"><%=FormatNumber(rsSqlstr("TIM_JL_MR_BI_AKTUAL_SHI"),0)%></td>			
		<td align="right"><%=FormatNumber(rsSqlstr("TIM_JL_MR_BI_BUDGET_HI"),0)%></td>	
		<td align="right"><%=FormatNumber(rsSqlstr("TIM_JL_MR_BI_BUDGET_SHI"),0)%></td>		
		<td align="right"><%=FormatNumber(rsSqlstr("TIM_JL_MR_BI_VarSHI_MTR"),0)%></td>	
		<td align="right"><%=FormatNumber(rsSqlstr("TIM_JL_MR_BI_VarSHI_PERCENT"),2)%></td>		
		
		<td align="right"><%=FormatNumber(rsSqlstr("TIM_JL_MR_SBI_AKTUAL"),0)%></td>	
		<td align="right"><%=FormatNumber(rsSqlstr("TIM_JL_MR_SBI_BUDGET"),0)%></td>		
		<td align="right"><%=FormatNumber(rsSqlstr("TIM_JL_MR_SBI_VarSBI_MTR"),0)%></td>	
		<td align="right"><%=FormatNumber(rsSqlstr("TIM_JL_MR_SBI_VarSBI_PERCENT"),2)%></td>
		<!--TIMBUN JALAN CR -->
		<td align="right"><%=FormatNumber(rsSqlstr("TIM_JL_CR_BI_AKTUAL_HI"),0)%></td>	
		<td align="right"><%=FormatNumber(rsSqlstr("TIM_JL_CR_BI_AKTUAL_SHI"),0)%></td>			
		<td align="right"><%=FormatNumber(rsSqlstr("TIM_JL_CR_BI_BUDGET_HI"),0)%></td>	
		<td align="right"><%=FormatNumber(rsSqlstr("TIM_JL_CR_BI_BUDGET_SHI"),0)%></td>		
		<td align="right"><%=FormatNumber(rsSqlstr("TIM_JL_CR_BI_VarSHI_MTR"),0)%></td>	
		<td align="right"><%=FormatNumber(rsSqlstr("TIM_JL_CR_BI_VarSHI_PERCENT"),2)%></td>		
		
		<td align="right"><%=FormatNumber(rsSqlstr("TIM_JL_CR_SBI_AKTUAL"),0)%></td>	
		<td align="right"><%=FormatNumber(rsSqlstr("TIM_JL_CR_SBI_BUDGET"),0)%></td>		
		<td align="right"><%=FormatNumber(rsSqlstr("TIM_JL_CR_SBI_VarSBI_MTR"),0)%></td>	
		<td align="right"><%=FormatNumber(rsSqlstr("TIM_JL_CR_SBI_VarSBI_PERCENT"),2)%></td>
		<!--PENGERASAN JALAN MR -->
		<td align="right"><%=FormatNumber(rsSqlstr("KERAS_JL_MR_BI_AKTUAL_HI"),0)%></td>	
		<td align="right"><%=FormatNumber(rsSqlstr("KERAS_JL_MR_BI_AKTUAL_SHI"),0)%></td>			
		<td align="right"><%=FormatNumber(rsSqlstr("KERAS_JL_MR_BI_BUDGET_HI"),0)%></td>	
		<td align="right"><%=FormatNumber(rsSqlstr("KERAS_JL_MR_BI_BUDGET_SHI"),0)%></td>		
		<td align="right"><%=FormatNumber(rsSqlstr("KERAS_JL_MR_BI_VarSHI_MTR"),0)%></td>	
		<td align="right"><%=FormatNumber(rsSqlstr("KERAS_JL_MR_BI_VarSHI_PERCENT"),2)%></td>	
		
		<td align="right"><%=FormatNumber(rsSqlstr("KERAS_JL_MR_SBI_AKTUAL"),0)%></td>	
		<td align="right"><%=FormatNumber(rsSqlstr("KERAS_JL_MR_SBI_BUDGET"),0)%></td>		
		<td align="right"><%=FormatNumber(rsSqlstr("KERAS_JL_MR_SBI_VarSBI_MTR"),0)%></td>	
		<td align="right"><%=FormatNumber(rsSqlstr("KERAS_JL_MR_SBI_VarSBI_PERCENT"),2)%></td>
		<!--PENGERASAN JALAN CR -->
		<td align="right"><%=FormatNumber(rsSqlstr("KERAS_JL_CR_BI_AKTUAL_HI"),0)%></td>	
		<td align="right"><%=FormatNumber(rsSqlstr("KERAS_JL_CR_BI_AKTUAL_SHI"),0)%></td>			
		<td align="right"><%=FormatNumber(rsSqlstr("KERAS_JL_CR_BI_BUDGET_HI"),0)%></td>	
		<td align="right"><%=FormatNumber(rsSqlstr("KERAS_JL_CR_BI_BUDGET_SHI"),0)%></td>		
		<td align="right"><%=FormatNumber(rsSqlstr("KERAS_JL_CR_BI_VarSHI_MTR"),0)%></td>	
		<td align="right"><%=FormatNumber(rsSqlstr("KERAS_JL_CR_BI_VarSHI_PERCENT"),2)%></td>	
		
		<td align="right"><%=FormatNumber(rsSqlstr("KERAS_JL_CR_SBI_AKTUAL"),0)%></td>	
		<td align="right"><%=FormatNumber(rsSqlstr("KERAS_JL_CR_SBI_BUDGET"),0)%></td>		
		<td align="right"><%=FormatNumber(rsSqlstr("KERAS_JL_CR_SBI_VarSBI_MTR"),0)%></td>	
		<td align="right"><%=FormatNumber(rsSqlstr("KERAS_JL_CR_SBI_VarSBI_PERCENT"),2)%></td>
		<!-- -->
		
	</tr>
	
<%
			 NO = NO + 1
			Total_1 = Total_1 +  Cdbl(rsSqlstr("STACKING_BI_AKTUAL_HI"))
			Total_2 = Total_2 +  Cdbl(rsSqlstr("STACKING_BI_AKTUAL_SHI"))
			Total_3 = Total_3 +  Cdbl(rsSqlstr("STACKING_BI_BUDGET_HI"))  
			Total_4 = Total_4 +  Cdbl(rsSqlstr("STACKING_BI_BUDGET_SHI"))
			Total_5 = Total_5 +  Cdbl(rsSqlstr("STACKING_BI_VarSHI_HA"))
			
			''Total_6 = Total_6 +  Cdbl(rsSqlstr("STACKING_BI_VarSHI_PERCENT"))
			Total_6 = (Total_5 / ( Total_4 * 100 ))
			''Total_7 = Total_7 +  Cdbl(rsSqlstr("STACKING_SBI_AKTUAL"))
			Total_7 = Total_4 + Total_6
			''Total_8 = Total_8 +  Cdbl(rsSqlstr("STACKING_SBI_BUDGET"))
			
			if Total_6 = 0 then 
			Total_8 = (Total_7 /100)
			else
			Total_8 = (Total_7 /(Total_6 *100))
			end if
			
			Total_9 = Total_9 +  Cdbl(rsSqlstr("STACKING_SBI_VarSBI_HA"))
			''Total_10 = Total_10 +  Cdbl(rsSqlstr("STACKING_SBI_VarSBI_PERCENT"))
			Total_10 = (Total_9 / ( Total_8 * 100 ))
			
			''Hitung Tanam Baru
			Total_11 = Total_11 +  Cdbl(rsSqlstr("TANAMBARU_BI_AKTUAL_HI"))
			Total_12 = Total_12 +  Cdbl(rsSqlstr("TANAMBARU_BI_AKTUAL_SHI"))
			Total_13 = Total_13 +  Cdbl(rsSqlstr("TANAMBARU_BI_BUDGET_HI"))
			Total_14 = Total_14 +  Cdbl(rsSqlstr("TANAMBARU_BI_BUDGET_SHI"))
			Total_15 = Total_15 +  Cdbl(rsSqlstr("TANAMBARU_BI_VarSHI_HA"))
			
			''Total_16 = Total_16 +  Cdbl(rsSqlstr("TANAMBARU_BI_VarSHI_PERCENT"))
			Total_16 =(Total_15 / ( Total_14 * 100 ))
			''Total_17 = Total_17 +  Cdbl(rsSqlstr("TANAMBARU_SBI_AKTUAL"))
			Total_17 = Total_14 + Total_16
			
			if Total_16 = 0 then 
			Total_18 = (Total_17 /100)
			else
			''Total_18 = Total_18 +  Cdbl(rsSqlstr("TANAMBARU_SBI_BUDGET"))
			Total_18 = (Total_17 /(Total_16 *100))
			end if
			
			Total_19 = Total_19 +  Cdbl(rsSqlstr("TANAMBARU_SBI_VarSBI_HA"))
			''Total_20 = Total_20 +  Cdbl(rsSqlstr("TANAMBARU_SBI_VarSBI_PERCENT"))
			Total_20 = (Total_19 / ( Total_18 * 100 ))
			
			''Hitung Kompleting
			Total_21 = Total_21 +  Cdbl(rsSqlstr("KOMPLETING_BI_AKTUAL_HI"))
			Total_22 = Total_22 +  Cdbl(rsSqlstr("KOMPLETING_BI_AKTUAL_SHI"))
			Total_23 = Total_23 +  Cdbl(rsSqlstr("KOMPLETING_BI_BUDGET_HI"))
			Total_24 = Total_24 +  Cdbl(rsSqlstr("KOMPLETING_BI_BUDGET_SHI"))
			Total_25 = Total_25 +  Cdbl(rsSqlstr("KOMPLETING_BI_VarSHI_PKK"))
			
			''Total_26 = Total_26 +  Cdbl(rsSqlstr("KOMPLETING_BI_VarSHI_PERCENT"))
			Total_26 =(Total_25 / ( Total_24 * 100 ))
			
			''Total_27 = Total_27 +  Cdbl(rsSqlstr("KOMPLETING_SBI_AKTUAL"))
			Total_27 = Total_24 + Total_26
			
			if Total_26 = 0 then 
			Total_28 = (Total_27 /100)
			else
			''Total_28 = Total_28 +  Cdbl(rsSqlstr("KOMPLETING_SBI_BUDGET"))
			Total_28 = (Total_27 /(Total_26 *100))
			end if
							
			Total_29 = Total_29 +  Cdbl(rsSqlstr("KOMPLETING_SBI_VarSBI_HA"))
			''Total_30 = Total_30 +  Cdbl(rsSqlstr("KOMPLETING_SBI_VarSBI_PERCENT"))
			Total_30 = (Total_29 / ( Total_28 * 100 ))
			
			''Hitung SISIP TBM
			Total_31 = Total_31 +  Cdbl(rsSqlstr("SISIP_TBM_BI_AKTUAL_HI"))
			Total_32 = Total_32 +  Cdbl(rsSqlstr("SISIP_TBM_BI_AKTUAL_SHI"))
			Total_33 = Total_33 +  Cdbl(rsSqlstr("SISIP_TBM_BI_BUDGET_HI"))
			Total_34 = Total_34 +  Cdbl(rsSqlstr("SISIP_TBM_BI_BUDGET_SHI"))
			Total_35 = Total_35 +  Cdbl(rsSqlstr("SISIP_TBM_BI_VarSHI_PKK"))
			
			''Total_36 = Total_36 +  Cdbl(rsSqlstr("SISIP_TBM_BI_VarSHI_PERCENT"))
			Total_36 =(Total_35 / ( Total_34 * 100 ))
			
			''Total_37 = Total_37 +  Cdbl(rsSqlstr("SISIP_TBM_SBI_AKTUAL"))
			Total_37 = Total_34 + Total_36
			
			if Total_36 = 0 then 
			Total_38 = (Total_37 /100)
			else
			''Total_38 = Total_38 +  Cdbl(rsSqlstr("SISIP_TBM_SBI_BUDGET"))
			Total_38 = (Total_37 /(Total_36 *100))
			end if
								
			Total_39 = Total_39 +  Cdbl(rsSqlstr("SISIP_TBM_SBI_VarSBI_PKK"))
			''Total_40 = Total_40 +  Cdbl(rsSqlstr("SISIP_TBM_SBI_VarSBI_PERCENT"))
			Total_40 = (Total_39 / ( Total_38 * 100 ))
			
			''Hitung SISIP TM
			Total_41 = Total_41 +  Cdbl(rsSqlstr("SISIP_TM_BI_AKTUAL_HI"))
			Total_42 = Total_42 +  Cdbl(rsSqlstr("SISIP_TM_BI_AKTUAL_SHI"))
			Total_43 = Total_43 +  Cdbl(rsSqlstr("SISIP_TM_BI_BUDGET_HI"))
			Total_44 = Total_44 +  Cdbl(rsSqlstr("SISIP_TM_BI_BUDGET_SHI"))
			Total_45 = Total_45 +  Cdbl(rsSqlstr("SISIP_TM_BI_VarSHI_PKK"))
			
			''Total_46 = Total_46 +  Cdbl(rsSqlstr("SISIP_TM_BI_VarSHI_PERCENT"))
			Total_46 =(Total_45 / ( Total_44 * 100 ))
			
			''Total_47 = Total_47 +  Cdbl(rsSqlstr("SISIP_TM_SBI_AKTUAL"))
			Total_47 = Total_44 + Total_46
			
			if Total_46 = 0 then 
			Total_48 = (Total_47 /100)
			else
			''Total_48 = Total_48 +  Cdbl(rsSqlstr("SISIP_TM_SBI_BUDGET"))
			Total_48 = (Total_47 /(Total_46 *100))
			end if
			
			Total_49 = Total_49 +  Cdbl(rsSqlstr("SISIP_TM_SBI_VarSBI_PKK"))
			''Total_50 = Total_50 +  Cdbl(rsSqlstr("SISIP_TM_SBI_VarSBI_PERCENT"))
			Total_50 = (Total_49 / ( Total_48 * 100 ))
			
			''Hitung Kastrasi
			Total_51 = Total_51 +  Cdbl(rsSqlstr("KASTRASI_BI_AKTUAL_HI"))
			Total_52 = Total_52 +  Cdbl(rsSqlstr("KASTRASI_BI_AKTUAL_SHI"))
			Total_53 = Total_53 +  Cdbl(rsSqlstr("KASTRASI_BI_BUDGET_HI"))
			Total_54 = Total_54 +  Cdbl(rsSqlstr("KASTRASI_BI_BUDGET_SHI"))
			Total_55 = Total_55 +  Cdbl(rsSqlstr("KASTRASI_BI_VarSHI_PKK"))
			
			''Total_56 = Total_56 +  Cdbl(rsSqlstr("KASTRASI_BI_VarSHI_PERCENT"))
			Total_56 =(Total_55 / ( Total_54 * 100 ))
			
			''Total_57 = Total_57 +  Cdbl(rsSqlstr("KASTRASI_SBI_AKTUAL"))
			Total_57 = Total_54 + Total_56
			
			if Total_56 = 0 then 
			Total_58 = (Total_57 /100)
			else
			''Total_58 = Total_58 +  Cdbl(rsSqlstr("KASTRASI_SBI_BUDGET"))
			Total_58 = (Total_57 /(Total_56 *100))
			end if
			
			Total_59 = Total_59 +  Cdbl(rsSqlstr("KASTRASI_SBI_VarSBI_PKK"))
			''Total_60 = Total_60 +  Cdbl(rsSqlstr("KASTRASI_SBI_VarSBI_PERCENT"))
			Total_60 = (Total_59 / ( Total_58 * 100 ))
			
			''Hitung Sanitasi
			Total_61 = Total_61 +  Cdbl(rsSqlstr("SANITASI_BI_AKTUAL_HI"))
			Total_62 = Total_62 +  Cdbl(rsSqlstr("SANITASI_BI_AKTUAL_SHI"))
			Total_63 = Total_63 +  Cdbl(rsSqlstr("SANITASI_BI_BUDGET_HI"))
			Total_64 = Total_64 +  Cdbl(rsSqlstr("SANITASI_BI_BUDGET_SHI"))
			Total_65 = Total_65 +  Cdbl(rsSqlstr("SANITASI_BI_VarSHI_PKK"))
			
			''Total_66 = Total_66 +  Cdbl(rsSqlstr("SANITASI_BI_VarSHI_PERCENT"))
			Total_66 =(Total_65 / ( Total_64 * 100 ))
			
			''Total_67 = Total_67 +  Cdbl(rsSqlstr("SANITASI_SBI_AKTUAL"))
			Total_67 = Total_64 + Total_66
			
			if Total_66 = 0 then 
			Total_68 = (Total_67 /100)
			else
			''Total_68 = Total_68 +  Cdbl(rsSqlstr("SANITASI_SBI_BUDGET"))
			Total_68 = (Total_67 /(Total_66 *100))
			end if
			
			Total_69 = Total_69 +  Cdbl(rsSqlstr("SANITASI_SBI_VarSBI_PKK"))
			''Total_70 = Total_70 +  Cdbl(rsSqlstr("SANITASI_SBI_VarSBI_PERCENT"))
			Total_70 = (Total_69 / ( Total_68 * 100 ))
			
			''Hitung TK Pemeliharaan
			Total_71 = Total_71 +  Cdbl(rsSqlstr("PEMEL_AKTUAL_TERSEDIA_HI"))
			Total_72 = Total_72 +  Cdbl(rsSqlstr("PEMEL_AKTUAL_TERSEDIA_SHI"))
			Total_73 = Total_73 +  Cdbl(rsSqlstr("PEMEL_AKTUAL_TANAMBARU_HI"))
			Total_74 = Total_74 +  Cdbl(rsSqlstr("PEMEL_AKTUAL_TANAMBARU_SHI"))
			Total_75 = Total_75 +  Cdbl(rsSqlstr("PEMEL_AKTUAL_SISIP_HI"))
			Total_76 = Total_76 +  Cdbl(rsSqlstr("PEMEL_AKTUAL_SISIP_SHI"))
			Total_77 = Total_77 +  Cdbl(rsSqlstr("PEMEL_AKTUAL_KOMPLETING_HI"))
			Total_78 = Total_78 +  Cdbl(rsSqlstr("PEMEL_AKTUAL_KOMPLETING_SHI"))
			Total_79 = Total_79 +  Cdbl(rsSqlstr("PEMEL_AKTUAL_PEMUPUKAN_HI"))
			Total_80 = Total_80 +  Cdbl(rsSqlstr("PEMEL_AKTUAL_PEMUPUKAN_SHI"))
			Total_81 = Total_81 +  Cdbl(rsSqlstr("PEMEL_AKTUAL_LAINNYA_HI"))
			Total_82 = Total_82 +  Cdbl(rsSqlstr("PEMEL_AKTUAL_LAINNYA_SHI"))
			
			''Total_83 = Total_83 +  Cdbl(rsSqlstr("PEMEL_AKTUAL_PERCENT_HADIR_HI"))
			Total_83 = ((Total_73 + Total_75 + Total_77 + Total_79 + Total_81) / (Total_71*100))
			''Total_84 = Total_84 +  Cdbl(rsSqlstr("PEMEL_AKTUAL_PERCENT_HADIR_SHI"))
			Total_84 = ((Total_74 + Total_76 + Total_78 + Total_80 + Total_82) / (Total_72*100))
			
			''Hitung Timbun Jalan MR
			Total_85 = Total_85 +  Cdbl(rsSqlstr("TIM_JL_MR_BI_AKTUAL_HI"))
			Total_86 = Total_86 +  Cdbl(rsSqlstr("TIM_JL_MR_BI_AKTUAL_SHI"))
			Total_87 = Total_87 +  Cdbl(rsSqlstr("TIM_JL_MR_BI_BUDGET_HI"))
			Total_88 = Total_88 +  Cdbl(rsSqlstr("TIM_JL_MR_BI_BUDGET_SHI"))
			Total_89 = Total_89 +  Cdbl(rsSqlstr("TIM_JL_MR_BI_VarSHI_MTR"))
			
			''Total_90 = Total_90 +  Cdbl(rsSqlstr("TIM_JL_MR_BI_VarSHI_PERCENT"))
			Total_90 =(Total_89 / ( Total_88 * 100 ))
			
			''Total_91 = Total_91 +  Cdbl(rsSqlstr("TIM_JL_MR_SBI_AKTUAL"))
			Total_91 = Total_88 + Total_90
			
			if Total_90 = 0 then 
			Total_92 = (Total_91 /100)
			else
			''Total_92 = Total_92 +  Cdbl(rsSqlstr("TIM_JL_MR_SBI_BUDGET"))
			Total_92 = (Total_91 /(Total_90 *100))
			end if
						
			Total_93 = Total_93 +  Cdbl(rsSqlstr("TIM_JL_MR_SBI_VarSBI_MTR"))
			''Total_94 = Total_94 +  Cdbl(rsSqlstr("TIM_JL_MR_SBI_VarSBI_PERCENT"))
			Total_70 = (Total_93 / ( Total_92 * 100 ))
			
			''Hitung Timbun Jalan CR
			Total_95 = Total_95 +  Cdbl(rsSqlstr("TIM_JL_CR_BI_AKTUAL_HI"))
			Total_96 = Total_96 +  Cdbl(rsSqlstr("TIM_JL_CR_BI_AKTUAL_SHI"))
			Total_97 = Total_97 +  Cdbl(rsSqlstr("TIM_JL_CR_BI_BUDGET_HI"))
			Total_98 = Total_98 +  Cdbl(rsSqlstr("TIM_JL_CR_BI_BUDGET_SHI"))
			Total_99 = Total_99 +  Cdbl(rsSqlstr("TIM_JL_CR_BI_VarSHI_MTR"))
			
			''Total_100 = Total_100 +  Cdbl(rsSqlstr("TIM_JL_CR_BI_VarSHI_PERCENT"))
			Total_100 =(Total_99 / ( Total_98 * 100 ))
			
			''Total_101 = Total_101 +  Cdbl(rsSqlstr("TIM_JL_CR_SBI_AKTUAL"))
			Total_101 = Total_98 + Total_100
			
			if Total_100 = 0 then 
			Total_102 = (Total_101 /100)
			else
			''Total_102 = Total_102 +  Cdbl(rsSqlstr("TIM_JL_CR_SBI_BUDGET"))
			Total_102 = (Total_101 /(Total_100 *100))
			end if
						
			Total_103 = Total_103 +  Cdbl(rsSqlstr("TIM_JL_CR_SBI_VarSBI_MTR"))
			''Total_104 = Total_104 +  Cdbl(rsSqlstr("TIM_JL_CR_SBI_VarSBI_PERCENT"))
			Total_104 = (Total_103 / ( Total_102 * 100 ))
			
			''Hitung Pengerasan Jalan MR
			Total_105 = Total_105 +  Cdbl(rsSqlstr("KERAS_JL_MR_BI_AKTUAL_HI"))
			Total_106 = Total_106 +  Cdbl(rsSqlstr("KERAS_JL_MR_BI_AKTUAL_SHI"))
			Total_107 = Total_107 +  Cdbl(rsSqlstr("KERAS_JL_MR_BI_BUDGET_HI"))
			Total_108 = Total_108 +  Cdbl(rsSqlstr("KERAS_JL_MR_BI_BUDGET_SHI"))
			Total_109 = Total_109 +  Cdbl(rsSqlstr("KERAS_JL_MR_BI_VarSHI_MTR"))
			
			''Total_110 = Total_110 +  Cdbl(rsSqlstr("KERAS_JL_MR_BI_VarSHI_PERCENT"))
			Total_110 =(Total_109 / ( Total_108 * 100 ))
			
			''Total_111 = Total_111 +  Cdbl(rsSqlstr("KERAS_JL_MR_SBI_AKTUAL"))
			Total_111 = Total_108 + Total_110
						
			if Total_110 = 0 then 
			Total_112 = (Total_111 /100)
			else
			''Total_112 = Total_112 +  Cdbl(rsSqlstr("KERAS_JL_MR_SBI_BUDGET"))
			Total_112 = (Total_111 /(Total_110 *100))
			end if
						
			Total_113 = Total_113 +  Cdbl(rsSqlstr("KERAS_JL_MR_SBI_VarSBI_MTR"))
			''Total_114 = Total_114 +  Cdbl(rsSqlstr("KERAS_JL_MR_SBI_VarSBI_PERCENT"))
			Total_114 = (Total_113 / ( Total_112 * 100 ))
			
			''Hitung Pengerasan Jalan CR
			Total_115 = Total_115 +  Cdbl(rsSqlstr("KERAS_JL_CR_BI_AKTUAL_HI"))
			Total_116 = Total_116 +  Cdbl(rsSqlstr("KERAS_JL_CR_BI_AKTUAL_SHI"))
			Total_117 = Total_117 +  Cdbl(rsSqlstr("KERAS_JL_CR_BI_BUDGET_HI"))
			Total_118 = Total_118 +  Cdbl(rsSqlstr("KERAS_JL_CR_BI_BUDGET_SHI"))
			Total_119 = Total_119 +  Cdbl(rsSqlstr("KERAS_JL_CR_BI_VarSHI_MTR"))
			
			''Total_120 = Total_120 +  Cdbl(rsSqlstr("KERAS_JL_CR_BI_VarSHI_PERCENT"))
			Total_120 =(Total_119 / ( Total_118 * 100 ))
			
			''Total_121 = Total_121 +  Cdbl(rsSqlstr("KERAS_JL_CR_SBI_AKTUAL"))
			Total_121 = Total_118 + Total_120
			
			if Total_120 = 0 then 
			Total_122 = (Total_121 /100)
			else
			''Total_122 = Total_122 +  Cdbl(rsSqlstr("KERAS_JL_CR_SBI_BUDGET"))
			Total_122 = (Total_121 /(Total_120 *100))
			end if
						
			Total_123 = Total_123 +  Cdbl(rsSqlstr("KERAS_JL_CR_SBI_VarSBI_MTR"))
			''Total_124 = Total_124 +  Cdbl(rsSqlstr("KERAS_JL_CR_SBI_VarSBI_PERCENT"))
			Total_124 = (Total_123 / ( Total_122 * 100 ))
						 
			rsSqlstr.MoveNext
		loop
end if
rsSqlstr.Close
set rsSqlstr = Nothing

%>

	<tr>
		<td colspan="1" align="center"><strong>TOTAL</strong></td>
	
		<td align="right"><%=Total_1 %></td>
		<td align="right"><%=Total_2%></td>
		<td align="right"><%=Total_3%></td>
		<td align="right"><%=Total_4%></td>
		<td align="right"><%=Total_5%></td>
		<td align="right"><%=Total_6%></td>
		<td align="right"><%=Total_7%></td>
		<td align="right"><%=Total_8%></td>
		<td align="right"><%=Total_9%></td>
		<td align="right"><%=Total_10%></td>
		<td align="right"><%=Total_11%></td>
		<td align="right"><%=Total_12%></td>
		<td align="right"><%=Total_13%></td>
		<td align="right"><%=Total_14%></td>
		<td align="right"><%=Total_15%></td>
		<td align="right"><%=Total_16%></td>
		<td align="right"><%=Total_17%></td>
		<td align="right"><%=Total_18%></td>
		<td align="right"><%=Total_19%></td>
		<td align="right"><%=Total_20%></td>
		<td align="right"><%=Total_21%></td>
		<td align="right"><%=Total_22%></td>
		<td align="right"><%=Total_23%></td>
		<td align="right"><%=Total_24%></td>
		<td align="right"><%=Total_25%></td>
		<td align="right"><%=Total_26%></td>
		<td align="right"><%=Total_27%></td>
		<td align="right"><%=Total_28%></td>
		<td align="right"><%=Total_29%></td>
		<td align="right"><%=Total_30%></td>
		<td align="right"><%=Total_31%></td>
		<td align="right"><%=Total_32%></td>
		<td align="right"><%=Total_33%></td>
		<td align="right"><%=Total_34%></td>
		<td align="right"><%=Total_35%></td>
		<td align="right"><%=Total_36%></td>
		<td align="right"><%=Total_37%></td>
		<td align="right"><%=Total_38%></td>
		<td align="right"><%=Total_39%></td>
		<td align="right"><%=Total_40%></td>
		<td align="right"><%=Total_41%></td>
		<td align="right"><%=Total_42%></td>
		<td align="right"><%=Total_43%></td>
		<td align="right"><%=Total_44%></td>
		<td align="right"><%=Total_45%></td>
		<td align="right"><%=Total_46%></td>
		<td align="right"><%=Total_47%></td>
		<td align="right"><%=Total_48%></td>
		<td align="right"><%=Total_49%></td>
		<td align="right"><%=Total_50%></td>
		<td align="right"><%=Total_51%></td>
		<td align="right"><%=Total_52%></td>
		<td align="right"><%=Total_53%></td>
		<td align="right"><%=Total_54%></td>
		<td align="right"><%=Total_55%></td>
		<td align="right"><%=Total_56%></td>
		<td align="right"><%=Total_57%></td>
		<td align="right"><%=Total_58%></td>
		<td align="right"><%=Total_59%></td>
		<td align="right"><%=Total_60%></td>
		<td align="right"><%=Total_61%></td>
		<td align="right"><%=Total_62%></td>
		<td align="right"><%=Total_63%></td>
		<td align="right"><%=Total_64%></td>
		<td align="right"><%=Total_65%></td>
		<td align="right"><%=Total_66%></td>
		<td align="right"><%=Total_67%></td>
		<td align="right"><%=Total_68%></td>
		<td align="right"><%=Total_69%></td>
		<td align="right"><%=Total_70%></td>
		<td align="right"><%=Total_71%></td>
		<td align="right"><%=Total_72%></td>
		<td align="right"><%=Total_73%></td>
		<td align="right"><%=Total_74%></td>
		<td align="right"><%=Total_75%></td>
		<td align="right"><%=Total_76%></td>
		<td align="right"><%=Total_77%></td>
		<td align="right"><%=Total_78%></td>
		<td align="right"><%=Total_79%></td>
		<td align="right"><%=Total_80%></td>
		<td align="right"><%=Total_81%></td>
		<td align="right"><%=Total_82%></td>
		<td align="right"><%=Total_83%></td>
		<td align="right"><%=Total_84%></td>
		<td align="right"><%=Total_85%></td>
		<td align="right"><%=Total_86%></td>
		<td align="right"><%=Total_87%></td>
		<td align="right"><%=Total_88%></td>
		<td align="right"><%=Total_89%></td>
		<td align="right"><%=Total_90%></td>
		<td align="right"><%=Total_91%></td>
		<td align="right"><%=Total_92%></td>
		<td align="right"><%=Total_93%></td>
		<td align="right"><%=Total_94%></td>
		<td align="right"><%=Total_95%></td>
		<td align="right"><%=Total_96%></td>
		<td align="right"><%=Total_97%></td>
		<td align="right"><%=Total_98%></td>
		<td align="right"><%=Total_99%></td>
		<td align="right"><%=Total_100%></td>
		<td align="right"><%=Total_101%></td>
		<td align="right"><%=Total_102%></td>
		<td align="right"><%=Total_103%></td>
		<td align="right"><%=Total_104%></td>
		<td align="right"><%=Total_105%></td>
		<td align="right"><%=Total_106%></td>
		<td align="right"><%=Total_107%></td>
		<td align="right"><%=Total_108%></td>
		<td align="right"><%=Total_109%></td>
		<td align="right"><%=Total_110%></td>
		<td align="right"><%=Total_111%></td>
		<td align="right"><%=Total_112%></td>
		<td align="right"><%=Total_113%></td>
		<td align="right"><%=Total_114%></td>
		<td align="right"><%=Total_115%></td>
		<td align="right"><%=Total_116%></td>
		<td align="right"><%=Total_117%></td>
		<td align="right"><%=Total_118%></td>
		<td align="right"><%=Total_119%></td>
		<td align="right"><%=Total_120%></td>
		<td align="right"><%=Total_121%></td>
		<td align="right"><%=Total_122%></td>
		<td align="right"><%=Total_123%></td>
		<td align="right"><%=Total_124%></td>	
	</tr>
   </tbody>
</table>
   
</body>
</html>
