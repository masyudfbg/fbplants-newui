<%validatesession%>
<!--#include file="connDb1.Inc"-->
<html>
<body>
<%

	dim sqlhdr,sqlupah,sqlnomor,rsnomor,bkmno,sqldelkosong
	dim fcba
	dim nQtdFields,nCount,linenokey,gangcode,ScriptName
	ScriptName	=	Request.ServerVariables("SCRIPT_NAME")
	
	tglbkm 			= Request("tglbkm")
	sectionname 	= Request("sectionname")
	gangcode 		= Request("gangcode")
	clasification 	= "501"
	supervisi1 		= Request("supervisi1")
	supervisi2 		= Request("supervisi2")
	supervisi3 		= Request("supervisi3")
	supervisi4 		= Request("supervisi4")
	fcba 			= session("fcba")
	proses 			= Request("simpan")
	rkbno           = Request("rkbno")

	if proses = "SIMPAN UPAH" then
	
	

	sqlnomor = "select 'BKM/'||'" & fcba & "'||'/'||To_Char(to_date('" & tglbkm & "','mm/dd/yyyy') ,'yyyymm')||'/'||SubStr(standard_hash(('" & gangcode & "'||To_Char(to_date('" & tglbkm & "','mm/dd/yyyy') ,'yyyymmdd')||'" & fcba & "'),'MD5'),0,10) bkmno " & _
	" from dual "
	'Response.Write sqlnomor
	set rsnomor = connStr.Execute(sqlnomor)
	
	if not rsnomor.eof then
		bkmno = rsnomor("bkmno")			
	end if

	

	nQtdFields 	= request.form("jumldata")
	nCount 		= 1
	linenokey 	= 1	

	sqlhdr = "insert into bkmhdr(BKMNO, BKMDATE, SECTIONCODE, GANGCODE, SUPERVISI1, " & _
	" SUPERVISI2, SUPERVISI3, SUPERVISI4,CREATEDATE,LASTUPDATE,FCENTRY,FCEDIT,FCIP,fcba,POSTINGFLG,ROWSTATE,DOCUMENT_CLASSIFICATION,RKBNO) " & _
	" values('" & bkmno & "',TO_DATE('" & tglbkm & "', 'mm/dd/yyyy'),'" & sectionname & "','" & gangcode & "','" & supervisi1 & "','" & supervisi2 & "', " & _ 
	" '" & supervisi3 & "','" & supervisi4 & "',SYSDATE,SYSDATE,'" & Session("CodeName") & "', " & _ 
	" '" & Session("CodeName") & "','" & Request.ServerVariables("REMOTE_ADDR") & "','" & fcba & "','0','Planned','" & clasification & "','" & rkbno & "')"
	'connstr.BeginTrans
	connStr.execute(sqlhdr)

	Response.Write sqlhdr

	' insert data dtl upah   ovtimehours typeovrtime
	
	do while cint(nCount) <= cint(nQtdFields)  

			dim absenupah, lokasiupah,sqlupdateupah,hk
			lokasiupah = request.form("lokasi" & nCount)
			absenupah = request.form("absen" & nCount)
			hk 			= request.form("hk" & nCount)
			if hk = "" then 
				hk = 0
			end if
			
			if absenupah <>"" then 
		
				SQL = "insert into BKMUPAHDTL(BKMNO, EMPLOYEECODE, ATTENDANCE, ACTCODE, " & _
				" LOCATIONCODE, MANDAYS, RATE, UNIT, OUTPUT, ABW, BUNCH_WEIGHT, KG_BRONDOLAN, " & _ 
				" OVERTIME_HOURS, TYPE_OVERTIME, " & _ 
				" CHARGEUNIT, CHARGEJOB, CHARGECODE, ATTENDANCE_FIN, " & _ 
				" IN_FIN, OUT_FIN, MANDAYS_FIN, ALLOWANCE_CODE, ALLOWANCE_AMT, " & _ 
				" LINENOKEY, REMARKS, CREATEDATE, LASTUPDATE, FCENTRY, " & _ 
				" FCEDIT, FCIP, FCBA,OTHCONVERTION,RKBNO) " & _ 
				" values ('" & bkmno & "','" & request.form("fccode" & nCount) & "','" & request.form("absen" & nCount) & "', " & _
				" '" & request.form("activity" & nCount) & "','" & request.form("tcupah" & nCount) & "','" & hk & "','" & request.form("ratekb" & nCount) & "', " & _ 
				" '" & request.form("unit" & nCount) & "','" & request.form("output" & nCount) & "','0','0', " & _ 
				" '0','" & request.form("ovtimehours" & nCount) & "','" & request.form("typeovrtime" & nCount) & "','" & lokasiupah & "','','', " & _ 
				" '" & request.form("absen" & nCount) & "', " & _ 
				" DECODE('" & request.form("timeinx" & nCount) & "','null',null,TO_DATE('" & request.form("timeinx" & nCount) & "','dd/mm/yyyy hh24:mi:ss')),  " & _ 
				" DECODE('" & request.form("timeoutx" & nCount) & "','null',null,TO_DATE('" & request.form("timeoutx" & nCount) & "','dd/mm/yyyy hh24:mi:ss')),  " & _ 
				" '" & hk & "','" & request.form("allowedcode" & nCount) & "','" & request.form("ratepremi" & nCount) & "', " & _ 
				" '" & linenokey & "','" & request.form("remark" & nCount) & "',TO_DATE('" & tglbkm & "', 'mm/dd/yyyy'),SYSDATE,'" & Session("CodeName") & "','" & Session("CodeName") & "', " & _ 
				" '" & Request.ServerVariables("REMOTE_ADDR") & "','" & fcba & "','" & request.form("othconvertion" & nCount) & "','" & request.form("rkbnofcba" & nCount) & "')"
			''	Response.Write SQL
				connStr.execute(SQL)
				

				'IF lokasiupah <> fcba then 

					'sqlupdateupah = "update BKMUPAHDTL set ACTCODE='1114001',LOCATIONCODE='BS',chargeunit='" & lokasi & "',chargejob='" & request.form("activity" & nCount) & "', " & _ 
					'"CHARGECODE='" & request.form("tcupah" & nCount) & "' " & _ 
					'"WHERE BKMNO='" & bkmno & "' AND LINENOKEY='" & linenokey & "' "
					'connStr.execute(sqlupdateupah)
					'Response.Write sqlupdateupah
				'end if

			end if

		nCount 	= nCount + 1
		linenokey  = linenokey + 1
	loop

		sqldelkosong = "delete from BKMUPAHDTL where employeecode is null and BKMNO='" & bkmno & "'"
		connStr.execute(sqldelkosong)


		SQL = "begin bkmproses_upah_api.insert_bkm('" & bkmno & "',TO_DATE('" & tglbkm & "', 'mm/dd/yyyy'),'" & fcba & "','" & Session("CodeName") & "'); end;"
		connStr.execute(SQL)
		'ExecuteAndLog ScriptName, SQL


	response.redirect("bkm_operasional_edit.asp?bkmno="&bkmno)
	

	

	'-------------------Batas proses upah-------------------------------------'

	ELSE

	sqlnomor = "select 'BKM/'||'" & fcba & "'||'/'||To_Char(to_date('" & tglbkm & "','mm/dd/yyyy') ,'yyyymm')||'/'||SubStr(standard_hash(('" & gangcode & "'||To_Char(to_date('" & tglbkm & "','mm/dd/yyyy') ,'yyyymmdd')||'" & fcba & "'),'MD5'),0,10) bkmno " & _
	" from dual "
	'Response.Write sqlnomor
	set rsnomor = connStr.Execute(sqlnomor)
	
	if not rsnomor.eof then
		bkmno = rsnomor("bkmno")			
	end if

	dim sqlbahan,nQtdFieldsbahan,nCountbahan,linenokeybahan

	nQtdFieldsbahan 	= request.form("jumldataitem")
	nCountbahan 		= 1
	linenokeybahan 		= 1	

	do while cint(nCountbahan) <= cint(nQtdFieldsbahan)

		dim itembarang 
		itembarang = request.form("ITEMCODE" & nCountbahan)
		if itembarang<>"" then 
			if request.form("qtyitem" & nCountbahan) = 0 then
				response.write "Tidak Boleh isi 0! Mohon di cek kembali!"
			else
				sqlbahan = " insert into BKMBAHANDTL (BKMNO,ACTCODE,LOCATIONTYPE,LOCATIONCODE,ITEMCODE, " & _ 
				" UOM,QTY,CREATEDATE,LASTUPDATE,FCENTRY, " & _
				" FCEDIT,FCIP,FCBA,LINENOKEY) " & _
				" values ('" & bkmno & "','" & request.form("activityitem" & nCountbahan) & "','FF','" & request.form("tcitem" & nCountbahan) & "', " & _ 
				" '" & request.form("ITEMCODE" & nCountbahan) & "','" & request.form("uom" & nCountbahan) & "', " & _ 
				" '" & request.form("qtyitem" & nCountbahan) & "',TO_DATE('" & tglbkm & "', 'mm/dd/yyyy'),SYSDATE,'" & Session("CodeName") & "','" & Session("CodeName") & "', " & _ 
				" '" & Request.ServerVariables("REMOTE_ADDR") & "','" & fcba & "','" & linenokeybahan & "')"
				'Response.Write sqlbahan
				connStr.execute(sqlbahan)
			end if
		end if

	nCountbahan 	= nCountbahan + 1
	linenokeybahan 	= linenokeybahan + 1

	loop

	sqldelkosong = "delete from BKMBAHANDTL where actcode is null and BKMNO='" & bkmno & "'"
	connStr.execute(sqldelkosong)

	'-------------------Batas proses bahan-------------------------------------'

	SQL = "begin bkmproses_bahan_api.insert_bkm('" & bkmno & "',TO_DATE('" & tglbkm & "', 'mm/dd/yyyy'),'" & fcba & "','" & Session("CodeName") & "'); end;"
	connStr.execute(SQL)


	response.redirect("bkm_operasional_edit.asp?bkmno="&bkmno)



	end if
			
			
	
%>

</body>
</html>