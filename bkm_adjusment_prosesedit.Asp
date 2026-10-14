<%validatesession%>
<!--#include file="connDb1.Inc"-->
<html>
<body>
<%

			dim sqlhdr,sqlupah,sqlnomor,rsnomor,bkmno,sqldelkosong,rkbno,rowstatedata
			dim fcba
			dim nQtdFields,nCount,linenokey,gangcode,ScriptName
			ScriptName	=	Request.ServerVariables("SCRIPT_NAME")
			
			tglbkm 			= Request("tglbkm")
			sectionname 	= Request("sectionname")
			gangcode 		= Request("gangcode")
			clasification 	= "505"
			supervisi1 		= Request("supervisi1")
			supervisi2 		= Request("supervisi2")
			supervisi3 		= Request("supervisi3")
			supervisi4 		= Request("supervisi4")
			fcba 			= Request("fcbahdr")
			proses 			= Request("btnsave")
			rkbno           = Request("rkbno")
			rowstatedata    = Request("rowstatedata")

			if proses = "SAVE UPAH" THEN

			
			'PROES EDIT'
			dim delupah,bkmupdate
			bkmupdate = request("bkmno")

			'delete upah
			delupah = "delete from BKMUPAHDTL where bkmno='" & bkmupdate & "'"
			connstr.BeginTrans
			connStr.execute(delupah)

			'delete bkmproses
			'delupah = "delete from bkmproses where bkmno='" & bkmupdate & "'"
			'connStr.execute(delupah)

			'delete bahan
			'delupah = "delete from BKMBAHANDTL where bkmno='" & bkmupdate & "'"
			'connStr.execute(delupah)

			

			nQtdFields 	= request.form("jumldata")
			nCount 		= 0
			linenokey 	= 1	

			sqlhdr = "update bkmhdr set SECTIONCODE='" & sectionname & "',GANGCODE='" & gangcode & "',SUPERVISI1='" & supervisi1 & "', " & _
			" SUPERVISI2='" & supervisi2 & "', SUPERVISI3='" & supervisi3 & "',SUPERVISI4='" & supervisi4 & "', " & _ 
			" LASTUPDATE=SYSDATE,FCEDIT='" & Session("CodeName") & "',DOCUMENT_CLASSIFICATION='" & clasification & "' " & _
			" where bkmno='" & bkmupdate & "'"
			connStr.execute(sqlhdr)
			WriteLog session("username") & " Execution at: " & Now() &  " " & ScriptName & " Query executed: " & sqlhdr
			 
		' 	Response.Write sqlhdr

		 	' insert data dtl upah
			
			 do while cint(nCount) <= cint(nQtdFields)  

					dim absenupah, lokasiupah,sqlupdateupah 
					lokasiupah = request.form("lokasi" & nCount)
					absenupah = request.form("absen" & nCount)
					hk = request.form("hk" & nCount)
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
						" values ('" & bkmupdate & "','" & request.form("fccode" & nCount) & "','" & request.form("absen" & nCount) & "', " & _
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
						WriteLog session("username") & " Execution at: " & Now() &  " " & ScriptName & " Query executed: " & SQL

					'	IF lokasiupah <> fcba then 

					'		sqlupdateupah = "update BKMUPAHDTL set ACTCODE='1114001',LOCATIONCODE='BS',chargeunit='" & lokasi & "',chargejob='" & request.form("activity" & nCount) & "', " & _ 
					'		"CHARGECODE='" & request.form("tcupah" & nCount) & "' " & _ 
					'		"WHERE BKMNO='" & bkmupdate & "' AND LINENOKEY='" & linenokey & "' "
					'		connStr.execute(sqlupdateupah)
							'Response.Write sqlupdateupah
					'	end if
					end if

			 nCount 	= nCount + 1
			 linenokey  = linenokey + 1
			 loop

			

			sqldelkosong = "delete from BKMUPAHDTL where employeecode is null and BKMNO='" & bkmno & "'"
			connStr.execute(sqldelkosong)
			WriteLog session("username") & " Execution at: " & Now() &  " " & ScriptName & " Query executed: " & sqldelkosong

			If err.number = 0 Then
			connStr.CommitTrans
			End If
			If err.number <> 0 Then
			connStr.RollbackTrans
			End If

		
			SQL = "begin bkmproses_upah_api.insert_bkm('" & bkmupdate & "',TO_DATE('" & tglbkm & "', 'mm/dd/yyyy'),'" & fcba & "','" & Session("CodeName") & "'); end;"
		  	connStr.execute(SQL)
			WriteLog session("username") & " Execution at: " & Now() &  " " & ScriptName & " Query executed: " & SQL
			'ExecuteAndLog ScriptName, SQL
			

			

			response.redirect("bkm_adjusment_edit.asp?bkmno="&bkmupdate)
			end if

			'-------------------Batas proses upah-------------------------------------'

			proses 			= Request("btnsave")

			if proses = "DELETE" THEN
				bkmupdate 		= request("bkmno")
				tglbkm 			= Request("tglbkm")
				fcba 			= session("fcba")
				SQL = "begin bkmproses_upah_api.delete_bkm('" & bkmupdate & "',TO_DATE('" & tglbkm & "', 'mm/dd/yyyy'),'" & fcba & "','" & Session("CodeName") & "'); end;"
		  		connStr.execute(SQL)
				WriteLog session("username") & " Execution at: " & Now() &  " " & ScriptName & " Query executed: " & SQL
				'ExecuteAndLog ScriptName, SQL
			END IF

			'-------------------Batas DELETE upah-------------------------------------'

			proses 			= Request("btnsave")

			if proses = "SAVE BAHAN" THEN

			dim sqlbahan,nQtdFieldsbahan,nCountbahan,linenokeybahan
			bkmupdate = request("bkmno")

			'delete bkmproses
			'delupah = "delete from bkmproses where bkmno='" & bkmupdate & "'"
			'connstr.BeginTrans
			'connStr.execute(delupah)

			'delete bahan
			delupah = "delete from BKMBAHANDTL where bkmno='" & bkmupdate & "'"
			'connstr.BeginTrans
			connStr.execute(delupah)
			WriteLog session("username") & " Execution at: " & Now() &  " " & ScriptName & " Query executed: " & delupah

			nQtdFieldsbahan 	= request.form("jumldataitem")
			nCountbahan 		= 0
			linenokeybahan 		= 1	

			do while cint(nCountbahan) <= cint(nQtdFieldsbahan)

				dim itembarang 
				itembarang = request.form("ITEMCODE" & nCountbahan)
				if itembarang<>"" then 

					sqlbahan = " insert into BKMBAHANDTL (BKMNO,ACTCODE,LOCATIONTYPE,LOCATIONCODE,ITEMCODE, " & _ 
					" UOM,QTY,CREATEDATE,LASTUPDATE,FCENTRY, " & _
					" FCEDIT,FCIP,FCBA,LINENOKEY) " & _
					" values ('" & bkmupdate & "','" & request.form("activityitem" & nCountbahan) & "','-','" & request.form("tcitem" & nCountbahan) & "', " & _ 
					" '" & request.form("ITEMCODE" & nCountbahan) & "','" & request.form("uom" & nCountbahan) & "', " & _ 
					" '" & request.form("qtyitem" & nCountbahan) & "',TO_DATE('" & tglbkm & "', 'mm/dd/yyyy'),SYSDATE,'" & Session("CodeName") & "','" & Session("CodeName") & "', " & _ 
					" '" & Request.ServerVariables("REMOTE_ADDR") & "','" & fcba & "','" & linenokeybahan & "')"
					'Response.Write sqlbahan
					connStr.execute(sqlbahan)
					WriteLog session("username") & " Execution at: " & Now() &  " " & ScriptName & " Query executed: " & sqlbahan

				end if

 			nCountbahan 	= nCountbahan + 1
			linenokeybahan 	= linenokeybahan + 1

			loop

			sqldelkosong = "delete from BKMBAHANDTL where actcode is null and BKMNO='" & bkmupdate & "'"
			connStr.execute(sqldelkosong)
			WriteLog session("username") & " Execution at: " & Now() &  " " & ScriptName & " Query executed: " & sqldelkosong



			'-------------------Batas proses bahan-------------------------------------'

        	SQL = "begin bkmproses_bahan_api.insert_bkm('" & bkmupdate & "',TO_DATE('" & tglbkm & "', 'mm/dd/yyyy'),'" & fcba & "','" & Session("CodeName") & "'); end;"
		  	connStr.execute(SQL)
			WriteLog session("username") & " Execution at: " & Now() &  " " & ScriptName & " Query executed: " & SQL
			'ExecuteAndLog ScriptName, SQL

			'If err.number = 0 Then
			'connStr.CommitTrans
			'End If
			'If err.number <> 0 Then
			'connStr.RollbackTrans
			'End If

			END IF

			response.redirect("bkm_adjusment_edit.asp?bkmno="&bkmupdate)
	
%>

</body>
</html>