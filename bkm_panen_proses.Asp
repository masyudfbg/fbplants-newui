<%validatesession%>
<!--#include file="connDb1.Inc"-->
<html>
<body>
<%

			dim sqlhdr,sqlupah,sqlnomor,rsnomor,bkmno,sqldelkosong,rkbno
			dim fcba
			dim nQtdFields,nCount,linenokey,gangcode,ScriptName
			ScriptName	=	Request.ServerVariables("SCRIPT_NAME")
			
			tglbkm 			= Request("tglbkm")
			sectionname 	= Request("sectionname")
			gangcode 		= Request("gangcode")
			clasification 	= "503"
			supervisi1 		= Request("supervisi1")
			supervisi2 		= Request("supervisi2")
			supervisi3 		= Request("supervisi3")
			supervisi4 		= Request("supervisi4")
			hujan 			= Request("hujan")
			fcba 			= session("fcba")
			proses 			= Request("proses")
			rkbno			= Request("rkbno")

			if hujan="on" then 
				hujan = 1
			else 
				hujan = 0
			end if

		

			sqlnomor = "select 'BKM/'||'" & fcba & "'||'/'||To_Char(to_date('" & tglbkm & "','mm/dd/yyyy') ,'yyyymm')||'/'||SubStr(standard_hash(('" & gangcode & "'||To_Char(to_date('" & tglbkm & "','mm/dd/yyyy') ,'yyyymmdd')||'" & fcba & "'),'MD5'),0,11) bkmno " & _
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
			" SUPERVISI2, SUPERVISI3, SUPERVISI4,CREATEDATE,LASTUPDATE,FCENTRY,FCEDIT,FCIP,fcba, " & _
			" POSTINGFLG,ROWSTATE,DOCUMENT_CLASSIFICATION,HUJAN, RKBNO) " & _
			" values('" & bkmno & "',TO_DATE('" & tglbkm & "', 'mm/dd/yyyy'),'" & sectionname & "','" & gangcode & "','" & supervisi1 & "','" & supervisi2 & "', " & _ 
			" '" & supervisi3 & "','" & supervisi4 & "',SYSDATE,SYSDATE,'" & Session("CodeName") & "', " & _ 
			" '" & Session("CodeName") & "','" & Request.ServerVariables("REMOTE_ADDR") & "','" & fcba & "',  " & _
			" '0','Planned','" & clasification & "','" & hujan & "','" & rkbno & "')"
			'connstr.BeginTrans
			connStr.execute(sqlhdr)

			'Response.Write sqlhdr
			
			do while cint(nCount) <= cint(nQtdFields)  

					dim absenupah, lokasiupah,hk,sqlupdateupah 
					lokasiupah 	= request.form("lokasi" & nCount)
					absenupah 	= request.form("absen" & nCount)
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
						" FCEDIT, FCIP, FCBA,OTHCONVERTION,BUNCHUN,BUNCHBA,RKBNO) " & _ 
						" values ('" & bkmno & "','" & request.form("fccode" & nCount) & "','" & request.form("absen" & nCount) & "', " & _
						" '" & request.form("activity" & nCount) & "','" & request.form("tcupah" & nCount) & "','" & hk & "','" & request.form("ratekb" & nCount) & "', " & _ 
						" '" & request.form("bunch" & nCount) & "','" & request.form("output" & nCount) & "','" & request.form("abw" & nCount) & "','" & request.form("kgtbs" & nCount) & "', " & _ 
						" '" & request.form("kgbrondolan" & nCount) & "','" & request.form("ovtimehours" & nCount) & "','" & request.form("typeovrtime" & nCount) & "','" & lokasiupah & "','','', " & _ 
						" '" & request.form("absen" & nCount) & "', " & _ 
						" DECODE('" & request.form("timeinx" & nCount) & "','null',null,TO_DATE('" & request.form("timeinx" & nCount) & "','dd/mm/yyyy hh24:mi:ss')),  " & _ 
						" DECODE('" & request.form("timeoutx" & nCount) & "','null',null,TO_DATE('" & request.form("timeoutx" & nCount) & "','dd/mm/yyyy hh24:mi:ss')),  " & _
						" '" & hk & "','" & request.form("allowedcode" & nCount) & "','" & request.form("ratepremi" & nCount) & "', " & _ 
						" '" & linenokey & "','" & request.form("remark" & nCount) & "',TO_DATE('" & tglbkm & "', 'mm/dd/yyyy'),SYSDATE,'" & Session("CodeName") & "','" & Session("CodeName") & "', " & _ 
						" '" & Request.ServerVariables("REMOTE_ADDR") & "','" & fcba & "', " & _ 
						" '" & request.form("othconvertion" & nCount) & "', " & _ 
						" '" & request.form("jjgu" & nCount) & "','" & request.form("jjgab" & nCount) & "','" & request.form("rkbnofcba" & nCount) & "')"
						'Response.Write SQL
						connStr.execute(SQL)
						

					end if

			 nCount 	= nCount + 1
			 linenokey  = linenokey + 1
			 loop

			 sqldelkosong = "delete from BKMUPAHDTL where employeecode is null and BKMNO='" & bkmno & "'"
			 connStr.execute(sqldelkosong)

			

			'-------------------Batas proses upah-------------------------------------'

			dim sqlmutu,nQtdFieldsmutu,nCountmutu,linenokeymutu

			nQtdFieldsmutu 	= request.form("jumldataitem")
			nCountmutu 		= 1
			linenokeymutu 		= 1	

			do while cint(nCountmutu) <= cint(nQtdFieldsmutu)

				dim itemmutu 
				itemmutu = request.form("tcmutu" & nCountmutu)
				if itemmutu<>"" then 

					sqlmutu = " insert into BKMMUTU (BKMNO,BLOK,BMATENG,BLMATENG,BKMATENG, " & _ 
					" BMENTAH,BBUSUK,CREATEDATE,LASTUPDATE,FCENTRY,FCEDIT,FCIP,FCBA) " & _
					" values ('" & bkmno & "','" & request.form("tcmutu" & nCountmutu) & "','" & request.form("bmateng" & nCountmutu) & "', " & _ 
					" '" & request.form("blmateng" & nCountmutu) & "','" & request.form("bkmateng" & nCountmutu) & "', " & _ 
					" '" & request.form("bmentah" & nCountmutu) & "','" & request.form("bbusuk" & nCountmutu) & "',TO_DATE('" & tglbkm & "', 'mm/dd/yyyy'),SYSDATE,'" & Session("CodeName") & "','" & Session("CodeName") & "', " & _ 
					" '" & Request.ServerVariables("REMOTE_ADDR") & "','" & fcba & "')"
					'Response.Write sqlmutu
					connStr.execute(sqlmutu)

				end if

 			nCountmutu 	= nCountmutu + 1
			linenokeymutu 	= linenokeymutu + 1

			loop

			sqldelkosong = "delete from BKMMUTU where BLOK is null and BKMNO='" & bkmno & "'"
			connStr.execute(sqldelkosong)

			'-------------------Batas proses mutu-------------------------------------'




          	SQL = "begin bkmproses_upah_api.insert_bkm('" & bkmno & "',TO_DATE('" & tglbkm & "', 'mm/dd/yyyy'),'" & fcba & "','" & Session("CodeName") & "'); end;"
		  	connStr.execute(SQL)
			'ExecuteAndLog ScriptName, SQL

			'If err.number = 0 Then
			'connStr.CommitTrans
			'End If
			'If err.number <> 0 Then
			'connStr.RollbackTrans
			'End If

			response.redirect("bkm_panen_edit.asp?bkmno="&bkmno)
		
			
	
%>

</body>
</html>