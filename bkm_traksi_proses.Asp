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
			clasification 	= "502"
			supervisi		= Request("supervisi")
			kendaraan 		= Request("kendaraan")
			fcba 			= session("fcba")
		

			sqlnomor = "select 'BKM/'||'" & fcba & "'||'/'||To_Char(to_date('" & tglbkm & "','mm/dd/yyyy') ,'yyyymm')||'/'||SubStr(standard_hash(('" & kendaraan & fcba & "'||To_Char(to_date('" & tglbkm & "','mm/dd/yyyy') ,'yyyymmdd')||'" & fcba & "'),'MD5'),0,9) bkmno " & _
			" from dual "
			'Response.Write sqlnomor
			set rsnomor = connStr.Execute(sqlnomor)
			' Response.Write rsnomor("bkmno")
			if not rsnomor.eof then
				bkmno = rsnomor("bkmno")			
			end if			

			nQtdFields 	= request.form("jumldata")
			nCount 		= 1
			linenokey 	= 1	

			sqlhdr = "insert into bkmhdr(BKMNO, BKMDATE, SECTIONCODE, GANGCODE, SUPERVISI1, " & _
			" SUPERVISI2, SUPERVISI3, SUPERVISI4,CREATEDATE,LASTUPDATE,FCENTRY,FCEDIT,FCIP,FCBA, " & _ 
			" POSTINGFLG,ROWSTATE,DOCUMENT_CLASSIFICATION,VEHICLECODE) " & _
			" values('" & bkmno & "',TO_DATE('" & tglbkm & "', 'mm/dd/yyyy'),'" & sectionname & "', " & _ 
			" '" & gangcode & "','" & supervisi & "','', " & _ 
			" '','',TO_DATE('" & tglbkm & "', 'mm/dd/yyyy'),SYSDATE,'" & Session("CodeName") & "', " & _ 
			" '" & Session("CodeName") & "','" & Request.ServerVariables("REMOTE_ADDR") & "', " & _ 
			" '" & fcba & "','0','Planned','" & clasification & "','" & kendaraan & "')"
			'connstr.BeginTrans
			' response.write sqlhd
			connStr.execute(sqlhdr)
			WriteLog session("username") & " Execution at: " & Now() &  " " & ScriptName & " Query executed: " & sqlhdr

			' Response.Write sqlhdr
	
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
			

			'-------------------Batas proses upah-------------------------------------'

			dim sqlvh,nQtdFieldsvh,nCountvh,linenokeyvh

			nQtdFieldsvh 	= request.form("jumldatavh")
			nCountvh 		= 1
			linenokeyvh 	= 1	

			do while cint(nCountvh) <= cint(nQtdFieldsvh)

				dim oprvh, lokasivh,sqlupdatevh 
					oprvh = request.form("employeecode" & nCountvh)
					lokasivh = request.form("lokasivh" & nCountvh)

				if oprvh<>"" then

					sqlvh = "insert into BKMVEHICLEDTL(BKMNO, EMPLOYEECODE, VEHICLECODE, ACTCODE, " & _
					" LOCATIONCODE, CATEGORYTRANSPORT, TIME, HOURS,UNIT, " & _
					" OUTPUT, TRIP, CHARGEUNIT, CHARGEJOB,  " & _
					" CHARGECODE, REMARKS, CREATEDATE, LASTUPDATE, FCENTRY, " & _
					" FCEDIT, FCIP, FCBA, LINENOKEY) " & _
					" values ('" & bkmno & "','" & request.form("employeecode" & nCountvh) & "','" & kendaraan & "', " & _ 
					" '" & request.form("activityivh" & nCountvh) & "','" & request.form("tcvh" & nCountvh) & "','" & request.form("kattransport" & nCountvh) & "'," & _
					" '" & request.form("timestart" & nCountvh) & "','" & request.form("jamkerja" & nCountvh) & "','" & request.form("kmhm" & nCountvh) & "', " & _ 
					" '" & request.form("outputvh" & nCountvh) & "','" & request.form("trip" & nCountvh) & "','" & lokasivh & "','-','-','" & request.form("remarkvh" & nCountvh) & "', " & _ 
					" TO_DATE('" & tglbkm & "', 'mm/dd/yyyy'),SYSDATE,'" & Session("CodeName") & "','" & Session("CodeName") & "', " & _ 
					" '" & Request.ServerVariables("REMOTE_ADDR") & "','" & fcba & "','" & linenokeyvh & "' )"
				''	Response.Write sqlvh
					connStr.execute(sqlvh)
					WriteLog session("username") & " Execution at: " & Now() &  " " & ScriptName & " Query executed: " & sqlvh

					'IF lokasivh <> fcba then 

							'sqlupdatevh = "update BKMVEHICLEDTL set ACTCODE='1114001',LOCATIONCODE='BS',chargeunit='" & lokasivh & "',chargejob='" & request.form("activityivh" & nCountvh) & "', " & _ 
							'"CHARGECODE='" & request.form("tcvh" & nCountvh) & "' " & _ 
							'"WHERE BKMNO='" & bkmno & "' AND LINENOKEY='" & linenokeyvh & "' "
							'connStr.execute(sqlupdatevh)
							'Response.Write sqlupdateupah
					'end if
				end if

			nCountvh 		= nCountvh + 1
			linenokeyvh 	= linenokeyvh + 1	

			loop

			sqldelkosong = "delete from BKMVEHICLEDTL where EMPLOYEECODE is null and BKMNO='" & bkmno & "'"
			connStr.execute(sqldelkosong)
			WriteLog session("username") & " Execution at: " & Now() &  " " & ScriptName & " Query executed: " & sqldelkosong
			'-------------------Batas proses vehicle-------------------------------------'


           	SQL = "begin bkmproses_upah_api.insert_bkm('" & bkmno & "',TO_DATE('" & tglbkm & "', 'mm/dd/yyyy'),'" & fcba & "','" & Session("CodeName") & "'); end;"
		  	connStr.execute(SQL)
			WriteLog session("username") & " Execution at: " & Now() &  " " & ScriptName & " Query executed: " & SQL
			'ExecuteAndLog ScriptName, SQL

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