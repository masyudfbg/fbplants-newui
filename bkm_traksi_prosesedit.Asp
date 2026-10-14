<%validatesession%>
<!--#include file="connDb1.Inc"-->
<html>
<body>
<%

			dim sqlhdr,sqlupah,sqlnomor,rsnomor,bkmno,sqldelkosong
			dim fcba
			dim nQtdFields,nCount,linenokey,gangcode
			dim oprvh, lokasivh,sqlupdatevh,ScriptName
			ScriptName	=	Request.ServerVariables("SCRIPT_NAME")
			
			tglbkm 			= Request("tglbkm")
			sectionname 	= Request("sectionname")
			gangcode 		= Request("gangcode")
			clasification 	= "502"
			supervisi		= Request("supervisi")
			kendaraan 		= Request("kendaraan")
			fcba 			= Request("fcbahdr")

			proses 			= Request("btnsave")
			if proses = "SAVE" THEN
		

			'PROES EDIT'
			dim delupah,bkmupdate
			bkmno = request("bkmno")

			'delete upah
			delupah = "delete from BKMUPAHDTL where bkmno='" & bkmno & "'"
			connstr.BeginTrans
			connStr.execute(delupah)

			'delete bkmproses
			delupah = "delete from bkmproses where bkmno='" & bkmno & "'"
			connStr.execute(delupah)

			'delete bahan
			delupah = "delete from BKMVEHICLEDTL where bkmno='" & bkmno & "'"
			connStr.execute(delupah)

			

			nQtdFields 	= request.form("jumldata")
			nCount 		= 0
			linenokey 	= 1	

			sqlhdr = "update bkmhdr set BKMDATE=TO_DATE('" & tglbkm & "', 'mm/dd/yyyy'), " & _
			" SUPERVISI1='" & supervisi & "',CREATEDATE=TO_DATE('" & tglbkm & "', 'mm/dd/yyyy'), " & _ 
			" LASTUPDATE=SYSDATE,DOCUMENT_CLASSIFICATION='" & clasification & "',VEHICLECODE='" & kendaraan & "' " & _
			" where bkmno='" & bkmno & "'"
			connStr.execute(sqlhdr)
			WriteLog session("username") & " Execution at: " & Now() &  " " & ScriptName & " Query executed: " & sqlhdr

		' 	Response.Write sqlhdr
	
			do while cint(nCount) <= cint(nQtdFields)  

					dim absenupah, lokasiupah,sqlupdateupah 
					lokasiupah = request.form("lokasi" & nCount)
					absenupah = request.form("absen" & nCount)
					
					if absenupah <>"" then 
			 
						SQL = "insert into BKMUPAHDTL(BKMNO, EMPLOYEECODE, ATTENDANCE, ACTCODE, " & _
						" LOCATIONCODE, MANDAYS, RATE, UNIT, OUTPUT, ABW, BUNCH_WEIGHT, KG_BRONDOLAN, " & _ 
						" OVERTIME_HOURS, TYPE_OVERTIME, " & _ 
						" CHARGEUNIT, CHARGEJOB, CHARGECODE, ATTENDANCE_FIN, " & _ 
						" IN_FIN, OUT_FIN, MANDAYS_FIN, ALLOWANCE_CODE, ALLOWANCE_AMT, " & _ 
						" LINENOKEY, REMARKS, CREATEDATE, LASTUPDATE, FCENTRY, " & _ 
						" FCEDIT, FCIP, FCBA,OTHCONVERTION) " & _ 
						" values ('" & bkmno & "','" & request.form("fccode" & nCount) & "','" & request.form("absen" & nCount) & "', " & _
						" '" & request.form("activity" & nCount) & "','" & request.form("tcupah" & nCount) & "','" & request.form("hk" & nCount) & "','" & request.form("ratekb" & nCount) & "', " & _ 
						" '" & request.form("unit" & nCount) & "','" & request.form("output" & nCount) & "','0','0', " & _ 
						" '0','" & request.form("ovtimehours" & nCount) & "','" & request.form("typeovrtime" & nCount) & "','" & lokasiupah & "','','', " & _ 
						" '" & request.form("absen" & nCount) & "', " & _ 
						" DECODE('" & request.form("timeinx" & nCount) & "','null',null,TO_DATE('" & request.form("timeinx" & nCount) & "','dd/mm/yyyy hh24:mi:ss')),  " & _ 
						" DECODE('" & request.form("timeoutx" & nCount) & "','null',null,TO_DATE('" & request.form("timeoutx" & nCount) & "','dd/mm/yyyy hh24:mi:ss')),  " & _ 
						" '" & request.form("hk" & nCount) & "','" & request.form("allowedcode" & nCount) & "','" & request.form("ratepremi" & nCount) & "', " & _ 
						" '" & linenokey & "','',TO_DATE('" & tglbkm & "', 'mm/dd/yyyy'),SYSDATE,'" & Session("CodeName") & "','" & Session("CodeName") & "', " & _ 
						" '" & Request.ServerVariables("REMOTE_ADDR") & "','" & fcba & "','" & request.form("othconvertion" & nCount) & "')"
					''	Response.Write SQL
						connStr.execute(SQL)
						WriteLog session("username") & " Execution at: " & Now() &  " " & ScriptName & " Query executed: " & SQL

						'IF lokasiupah <> fcba then 

						'	sqlupdateupah = "update BKMUPAHDTL set ACTCODE='1114001',LOCATIONCODE='BS',chargeunit='" & lokasi & "',chargejob='" & request.form("activity" & nCount) & "', " & _ 
						'	"CHARGECODE='" & request.form("tcupah" & nCount) & "' " & _ 
						'	"WHERE BKMNO='" & bkmno & "' AND LINENOKEY='" & linenokey & "' "
						'	connStr.execute(sqlupdateupah)
							'Response.Write sqlupdateupah
						'end if
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

			'-------------------Batas proses upah-------------------------------------'

			dim sqlvh,nQtdFieldsvh,nCountvh,linenokeyvh

			nQtdFieldsvh 	= request.form("jumldatavh")
			nCountvh 		= 0
			linenokeyvh 	= 1	

			do while cint(nCountvh) <= cint(nQtdFieldsvh)

				
					oprvh = request.form("employeecode" & nCountvh)
					lokasivh = request.form("lokasivh" & nCountvh)
					kattransport = request.form("kattransport" & nCountvh)
					if kattransport = "" then 
						kattransport = 0
					end if

				if oprvh<>"" then

					sqlvh = "insert into BKMVEHICLEDTL(BKMNO, EMPLOYEECODE, VEHICLECODE, ACTCODE, " & _
					" LOCATIONCODE, CATEGORYTRANSPORT, TIME, HOURS,UNIT, " & _
					" OUTPUT, TRIP, CHARGEUNIT, CHARGEJOB,  " & _
					" CHARGECODE, REMARKS, CREATEDATE, LASTUPDATE, FCENTRY, " & _
					" FCEDIT, FCIP, FCBA, LINENOKEY,HELPERCODE) " & _
					" values ('" & bkmno & "','" & request.form("employeecode" & nCountvh) & "','" & kendaraan & "', " & _ 
					" '" & request.form("activityivh" & nCountvh) & "','" & request.form("tcvh" & nCountvh) & "','" & kattransport & "'," & _
					" '" & request.form("timestart" & nCountvh) & "','" & request.form("jamkerja" & nCountvh) & "','" & request.form("kmhm" & nCountvh) & "', " & _ 
					" '" & request.form("outputvh" & nCountvh) & "','" & request.form("trip" & nCountvh) & "','" & lokasivh & "','-','-','" & request.form("remarkvh" & nCountvh) & "', " & _ 
					" TO_DATE('" & tglbkm & "', 'mm/dd/yyyy'),SYSDATE,'" & Session("CodeName") & "','" & Session("CodeName") & "', " & _ 
					" '" & Request.ServerVariables("REMOTE_ADDR") & "','" & fcba & "','" & linenokeyvh & "','" & request.form("helpercode" & nCountvh) & "' )"
				'	Response.Write sqlvh
					connStr.execute(sqlvh)
					WriteLog session("username") & " Execution at: " & Now() &  " " & ScriptName & " Query executed: " & sqlvh
				'	IF lokasivh <> fcba then 

				'			sqlupdatevh = "update BKMVEHICLEDTL set ACTCODE='1114001',LOCATIONCODE='BS',chargeunit='" & lokasivh & "',chargejob='" & request.form("activityivh" & nCountvh) & "', " & _ 
				'			"CHARGECODE='" & request.form("tcvh" & nCountvh) & "' " & _ 
				'			"WHERE BKMNO='" & bkmno & "' AND LINENOKEY='" & linenokeyvh & "' "
				'			connStr.execute(sqlupdatevh)
							'Response.Write sqlupdateupah
				'	end if
				end if

			nCountvh 		= nCountvh + 1
			linenokeyvh 	= linenokeyvh + 1	

			loop

			sqldelkosong = "delete from BKMVEHICLEDTL where EMPLOYEECODE is null and BKMNO='" & bkmno & "'"
			connStr.execute(sqldelkosong)
			WriteLog session("username") & " Execution at: " & Now() &  " " & ScriptName & " Query executed: " & sqldelkosong

			'-------------------Batas proses vehicle-------------------------------------'
			' perubahan mei 2024 upah akan keluar jika sdh aproved


           	SQL = "begin bkmproses_upah_api.insert_bkm('" & bkmno & "',TO_DATE('" & tglbkm & "', 'mm/dd/yyyy'),'" & fcba & "','" & Session("CodeName") & "'); end;"
		  	connStr.execute(SQL)
			'ExecuteAndLog ScriptName, SQL
			WriteLog session("username") & " Execution at: " & Now() &  " " & ScriptName & " Query executed: " & SQL

			else

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
			end if

			'If err.number = 0 Then
			'connStr.CommitTrans
			'End If
			'If err.number <> 0 Then
			'connStr.RollbackTrans
			'End If

			response.redirect("bkm_traksi_edit.asp?bkmno="&bkmno)
			
	
%>

</body>
</html>