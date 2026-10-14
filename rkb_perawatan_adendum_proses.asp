<%validatesession%>
<!--#include file="connDb1.Inc"-->


<%
	dim rkbpost,norkb
	dim sqlhdr,rshdr 
	dim RKBNO, PERIOD_YEAR, PERIOD_MONTH, SECTIONCODE,ROWSTATE, DOCUMENT_CLASSIFICATION, TOTUPAH,CREATEDATE

	dim classification,simpan1,fcba,countrkb,sqlsimpandtl
	DIM statuspost,periodethn,periodebln,section,tglinput,ttlamountupah
	dim sqlnomor,rsnomor,tglprev,currentnomor,currentnomorfix,chargeunit
	DIM actcode,blok,ttlhk,rate,normamandays,luasdikerjakan,rotasi,totalamounthk
	DIM nQtdFields,nCount,linenokey

	rkbpost = request("rkbno")
	simpan1 = request("simpan1")

	

	if simpan1 = "SAVE" then
			
		
			RKBNO 			= request("norkbview")
			fcba 			= session("fcba")

		'	Response.Write RKBNO

			
			'looping di tabel detail
			nQtdFields 	= request.form("jmlrawat")
			nCount 		= 0
			linenokey  	= 1

			

			do while cint(nCount) <= cint(nQtdFields) 
				actcode 		= request.form("activityrwt" & nCount)
				'Response.Write actcode
			'	Response.End
		
				
				blok 			= request.form("tcrwt" & nCount)
				ttlhk 			= request.form("ttlhk" & nCount)
				rate 			= request.form("rate" & nCount)
				normamandays	= request.form("mandayform" & nCount)
				luasdikerjakan	= request.form("luasdikerjakan" & nCount)
				rotasi			= request.form("rotasi" & nCount)
				totalamounthk	= request.form("amountbck" & nCount)
				haactual 		= request.form("haactual" & nCount)

				sqlsimpandtl = "MERGE INTO RKBUPAHDTL  R " & _ 
				" USING (SELECT RKBNO, ACTCODE, LOCATIONCODE, MANDAYS, RATE, " & _
				"	NORMAMANDAYS, NORMAPREMI, VOLUME, RATEBORONGAN, ABW_BUDGET, " & _
				"	KG_BUDGET, ROTASI, BASIS_JJG, RATE_SB, RATE_LB, " & _
				"	TOTALAMOUNT_HK, TOTALAMOUNT_PREMI, TOTALAMOUNT_BORONGAN, LINENOKEY, REMARKS, " & _
				"	CREATEDATE, LASTUPDATE, FCENTRY, FCEDIT, FCIP, " & _
				"	FCBA, HAACTUAL, ROWSTATE, ADDENDUM " & _
				"	FROM RKBUPAHDTL  WHERE RKBNO='" & RKBNO & "') RK " & _
				" ON (R.RKBNO = RK.RKBNO AND R.ACTCODE=RK.ACTCODE AND R.LOCATIONCODE=RK.LOCATIONCODE AND R.FCBA=RK.FCBA) " & _
				" WHEN MATCHED THEN " & _
				"	UPDATE " & _
				"	SET R.NORMAMANDAYS=RK.NORMAMANDAYS,R.VOLUME = RK.VOLUME,R.MANDAYS=RK.MANDAYS,R.ROTASI=RK.ROTASI,R.TOTALAMOUNT_HK=RK.TOTALAMOUNT_HK " & _
				"	WHERE R.RKBNO = RK.RKBNO AND R.FCBA=RK.FCBA" & _
				" WHEN NOT MATCHED THEN " & _
				"	INSERT (RKBNO, ACTCODE, LOCATIONCODE, MANDAYS, RATE,  " & _
				" NORMAMANDAYS, VOLUME,ROTASI,   " & _
				" TOTALAMOUNT_HK, LINENOKEY,  " & _
				" CREATEDATE, LASTUPDATE, FCENTRY, FCEDIT, FCIP,  " & _ 
				" FCBA, HAACTUAL, ROWSTATE, ADDENDUM) " & _
				" VALUES ('" & RKBNO & "', '" & actcode & "','" & blok & "','" & ttlhk & "','" & rate & "','" & normamandays & "','" & luasdikerjakan & "', " & _ 
				" '" & rotasi & "','" & totalamounthk & "','" & linenokey & "',SYSDATE,SYSDATE,'" & Session("CodeName") & "', " & _ 
				" '" & Session("CodeName") & "','" & Request.ServerVariables("REMOTE_ADDR") & "','" & fcba & "','" & haactual & "','Planned','1')" 

					
			'	sqlsimpandtl = "insert into RKBUPAHDTL(RKBNO, ACTCODE, LOCATIONCODE, MANDAYS, RATE, " & _
			'	" NORMAMANDAYS, VOLUME,  " & _
			'	" ROTASI,TOTALAMOUNT_HK,LINENOKEY, " & _
			'	" CREATEDATE, LASTUPDATE, FCENTRY, FCEDIT, FCIP,FCBA,HAACTUAL,ROWSTATE,ADDENDUM) " & _
			'	" values('" & RKBNO & "','" & actcode & "','" & blok & "','" & ttlhk & "', " & _ 
			'	" '" & rate & "','" & normamandays & "','" & luasdikerjakan & "','" & rotasi & "','" & totalamounthk & "','" & linenokey & "', " & _ 
			'	" TO_DATE('" & tglinput & "', 'dd/mm/yyyy'),SYSDATE,'" & Session("CodeName") & "','" & Session("CodeName") & "','" & Request.ServerVariables("REMOTE_ADDR") & "', " & _ 
			'	" '" & fcba & "','" & haactual & "','Planned','1')"
				'connstr.BeginTrans
				'Response.Write sqlsimpandtl
				connStr.execute(sqlsimpandtl)

			
				nCount 	= nCount + 1
				linenokey 	= linenokey + 1
			loop
	'	response.redirect("rkb_perawatan_adendum.asp?rkbno="&RKBNO)
	End IF

	if simpan="APPROVED" then
		dim sqlupdate
		norkb 			= Request("norkbview")

		
		'sqlupdate = "update rkbhdr set rowstate='Approved' where rkbno='" & norkb & "'"
		'connStr.execute(sqlupdate)

		sqlupdate = "UPDATE RKBUPAHDTL set ROWSTATE='Approved' where rkbno='" & norkb & "' AND ROWSTATE='Planned' AND ADDENDUM=1"
		connStr.execute(sqlupdate)
	'	Response.Write sqlupdate
	END IF
		'response.redirect("rkb_perawatan_adendum.asp?rkbno="&norkb)

%>
