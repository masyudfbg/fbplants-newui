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
			clasification 	= "503"
			supervisi1 		= Request("supervisi1")
			supervisi2 		= Request("supervisi2")
			supervisi3 		= Request("supervisi3")
			supervisi4 		= Request("supervisi4")
			fcba 			= Request("fcbahdr")
			
			hujan 			= Request("hujan")

			proses 			= Request("btnsave")

			if proses = "SAVE" THEN

			if hujan="on" then 
				hujan = 1
			else 
				hujan = 0
			end if


			'PROES EDIT'
			dim delupah,bkmupdate
			bkmupdate = request("bkmno")

			'delete upah
			delupah = "delete from BKMUPAHDTL where bkmno='" & bkmupdate & "'"
			connstr.BeginTrans
			connStr.execute(delupah)

			'delete bkmproses
			delupah = "delete from bkmproses where bkmno='" & bkmupdate & "'"
			connStr.execute(delupah)

			'delete bahan
			delupah = "delete from BKMMUTU where bkmno='" & bkmupdate & "'"
			connStr.execute(delupah)

			'delete denda
			delupah = "delete from BKMDENDA where bkmno='" & bkmupdate & "'"
			connStr.execute(delupah)

			

			nQtdFields 	= request.form("jumldata")
			nCount 		= 0
			linenokey 	= 1	

			'sqlhdr = "update bkmhdr set BKMDATE=TO_DATE('" & tglbkm & "', 'mm/dd/yyyy'),SECTIONCODE='" & sectionname & "',GANGCODE='" & gangcode & "',SUPERVISI1='" & supervisi1 & "', " & _
			'" SUPERVISI2='" & supervisi2 & "', SUPERVISI3='" & supervisi3 & "',SUPERVISI4='" & supervisi4 & "',CREATEDATE=TO_DATE('" & tglbkm & "', 'mm/dd/yyyy'), " & _ 
			'" LASTUPDATE=SYSDATE,DOCUMENT_CLASSIFICATION='" & clasification & "',HUJAN='" & hujan & "' " & _
			'" where bkmno='" & bkmupdate & "'"
			'connStr.execute(sqlhdr)
			
			sqlhdr = "update bkmhdr set SECTIONCODE='" & sectionname & "',GANGCODE='" & gangcode & "',SUPERVISI1='" & supervisi1 & "', " & _
			" SUPERVISI2='" & supervisi2 & "', SUPERVISI3='" & supervisi3 & "',SUPERVISI4='" & supervisi4 & "', " & _ 
			" LASTUPDATE=SYSDATE,FCEDIT='" & Session("CodeName") & "',DOCUMENT_CLASSIFICATION='" & clasification & "',HUJAN='" & hujan & "' " & _
			" where bkmno='" & bkmupdate & "'"
			connStr.execute(sqlhdr)

		' 	Response.Write sqlhdr
			
			do while cint(nCount) <= cint(nQtdFields)  

					dim absenupah, lokasiupah,sqlupdateupah 
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
						" FCEDIT, FCIP, FCBA,OTHCONVERTION,BUNCHUN,BUNCHBA,RKBNO) " & _ 
						" values ('" & bkmupdate & "','" & request.form("fccode" & nCount) & "','" & request.form("absen" & nCount) & "', " & _
						" '" & request.form("activity" & nCount) & "','" & request.form("tcupah" & nCount) & "','" & hk & "','" & request.form("ratekb" & nCount) & "', " & _ 
						" '" & request.form("bunch" & nCount) & "','" & request.form("output" & nCount) & "','" & request.form("abw" & nCount) & "','" & request.form("kgtbs" & nCount) & "', " & _ 
						" '" & request.form("kgbrondolan" & nCount) & "','" & request.form("ovtimehours" & nCount) & "','" & request.form("typeovrtime" & nCount) & "','" & lokasiupah & "','','', " & _ 
						" '" & request.form("absen" & nCount) & "', " & _ 
						" DECODE('" & request.form("timeinx" & nCount) & "','null',null,TO_DATE('" & request.form("timeinx" & nCount) & "','dd/mm/yyyy hh24:mi:ss')),  " & _ 
						" DECODE('" & request.form("timeoutx" & nCount) & "','null',null,TO_DATE('" & request.form("timeoutx" & nCount) & "','dd/mm/yyyy hh24:mi:ss')),  " & _
						" '" & hk & "','" & request.form("allowedcode" & nCount) & "','" & request.form("ratepremi" & nCount) & "', " & _ 
						" '" & linenokey & "','" & request.form("remark" & nCount) & "',TO_DATE('" & tglbkm & "', 'mm/dd/yyyy'),SYSDATE,'" & Session("CodeName") & "','" & Session("CodeName") & "', " & _ 
						" '" & Request.ServerVariables("REMOTE_ADDR") & "','" & fcba & "', " & _ 
						" '" & request.form("othconvertion" & nCount) & "','" & request.form("jjgu" & nCount) & "' " & _ 
						",'" & request.form("jjgab" & nCount) & "','" & request.form("rkbnofcba" & nCount) & "')"
						'Response.Write SQL
						connStr.execute(SQL)
						

					'	IF lokasiupah <> fcba then 

						'	sqlupdateupah = "update BKMUPAHDTL set ACTCODE='1114001',LOCATIONCODE='BS',chargeunit='" & lokasi & "',chargejob='" & request.form("activity" & nCount) & "', " & _ 
						'	"CHARGECODE='" & request.form("tcupah" & nCount) & "' " & _ 
						'	"WHERE BKMNO='" & bkmupdate & "' AND LINENOKEY='" & linenokey & "' "
						'	connStr.execute(sqlupdateupah)
							'Response.Write sqlupdateupah
						'end if

					end if

			 nCount 	= nCount + 1
			 linenokey  = linenokey + 1
			 loop

			' If err.number = 0 Then
			'connStr.CommitTrans
			'End If
			'If err.number <> 0 Then
			'connStr.RollbackTrans
			'End If

			'-------------------Batas proses upah-------------------------------------'

			dim sqlmutu,nQtdFieldsmutu,nCountmutu,linenokeymutu

			nQtdFieldsmutu 	= request.form("jumldataitem")
			nCountmutu 		= 0
			linenokeymutu 	= 1	

			do while cint(nCountmutu) <= cint(nQtdFieldsmutu)

				dim itemmutu 
				itemmutu = request.form("tcmutu" & nCountmutu)
				if itemmutu<>"" then 

					sqlmutu = " insert into BKMMUTU (BKMNO,BLOK,BMATENG,BLMATENG,BKMATENG, " & _ 
					" BMENTAH,BBUSUK,CREATEDATE,LASTUPDATE,FCENTRY,FCEDIT,FCIP,FCBA) " & _
					" values ('" & bkmupdate & "','" & request.form("tcmutu" & nCountmutu) & "','" & request.form("bmateng" & nCountmutu) & "', " & _ 
					" '" & request.form("blmateng" & nCountmutu) & "','" & request.form("bkmateng" & nCountmutu) & "', " & _ 
					" '" & request.form("bmentah" & nCountmutu) & "','" & request.form("bbusuk" & nCountmutu) & "',TO_DATE('" & tglbkm & "', 'mm/dd/yyyy'),SYSDATE,'" & Session("CodeName") & "','" & Session("CodeName") & "', " & _ 
					" '" & Request.ServerVariables("REMOTE_ADDR") & "','" & fcba & "')"
					'Response.Write sqlmutu
					connStr.execute(sqlmutu)

				end if

 			nCountmutu 	= nCountmutu + 1
			linenokeymutu 	= linenokeymutu + 1

			loop

			sqldelkosong = "delete from BKMMUTU where BLOK is null and BKMNO='" & bkmupdate & "'"
			connStr.execute(sqldelkosong)

			'-------------------Batas proses mutu-------------------------------------'


			dim sqldenda,nQtdFieldsdenda,nCountdenda,linenokeydenda,empcodedenda,linedenda

			nQtdFieldsdenda 	= request.form("jumldatadenda")
			nCountdenda 		= 0
			linenokeydenda 		= 1	

			do while cint(nCountdenda) <= cint(nQtdFieldsdenda)

				  
				empcodedenda = request.form("empcode" & nCountdenda)
				linedenda = request.form("linedenda" & nCountdenda)
				
				if empcodedenda<>"" then 

					sqldenda = " insert into BKMDENDA ( BKMNO, EMPLOYEECODE, LINENOKEY, REMARKS, CREATEDATE, " & _
					" LASTUPDATE, FCENTRY, FCEDIT, FCIP, FCBA, " & _
					" TBSA, TBSBA, TBSTP, TBSBK, TBSO, " & _
					" TBSAB, TBSE, ANCAKBT, ANCAKPS, ANCAKSP, " & _
					" ANCAKBRD, ANCAKOP, ANCAKBM, ANCAKBDT) " & _
					" values ('" & bkmupdate & "','" & empcodedenda & "','" & linedenda & "','Denda Panen', " & _ 
					" TO_DATE('" & tglbkm & "', 'mm/dd/yyyy'),SYSDATE,'" & Session("CodeName") & "','" & Session("CodeName") & "', " & _ 
					" '" & Request.ServerVariables("REMOTE_ADDR") & "','" & fcba & "', " & _
					" '" & request.form("ktbsa" & nCountdenda) & "','" & request.form("ktbsba" & nCountdenda) & "', " & _ 
					" '" & request.form("ktbstp" & nCountdenda) & "','" & request.form("ktbsbk" & nCountdenda) & "', " & _ 
					" '" & request.form("ktbso" & nCountdenda) & "','" & request.form("ktbsab" & nCountdenda) & "', " & _ 
					" '" & request.form("ktbse" & nCountdenda) & "','" & request.form("kancakbt" & nCountdenda) & "', " & _ 
					" '" & request.form("kancakps" & nCountdenda) & "','" & request.form("kancaksp" & nCountdenda) & "', "  & _ 
					" '" & request.form("kancakbrd" & nCountdenda) & "','" & request.form("kancakop" & nCountdenda) & "', "  & _ 
					" '" & request.form("kancakbm" & nCountdenda) & "','" & request.form("kancakbdt" & nCountdenda) & "')"
					'Response.Write sqldenda
					connStr.execute(sqldenda)

				end if

 			nCountdenda 	= nCountdenda + 1
			linenokeydenda 	= linenokeydenda + 1

			loop

			sqldelkosong = "delete from BKMDENDA where EMPLOYEECODE is null and BKMNO='" & bkmupdate & "'"
			connStr.execute(sqldelkosong)

			If err.number = 0 Then
			connStr.CommitTrans
			End If
			If err.number <> 0 Then
			connStr.RollbackTrans
			End If

			'-------------------Batas proses denda-------------------------------------'
			' per mei 2024 upah tidak keluar jika tidak approved

           	SQL = "begin bkmproses_upah_api.insert_bkm('" & bkmupdate & "',TO_DATE('" & tglbkm & "', 'mm/dd/yyyy'),'" & fcba & "','" & Session("CodeName") & "'); end;"
		  	connStr.execute(SQL)
			'ExecuteAndLog ScriptName, SQL

			

			else 

			proses 			= Request("btnsave")

			if proses = "DELETE" THEN
				bkmupdate 		= request("bkmno")
				tglbkm 			= Request("tglbkm")
				fcba 			= session("fcba")
				SQL = "begin bkmproses_upah_api.delete_bkm('" & bkmupdate & "',TO_DATE('" & tglbkm & "', 'mm/dd/yyyy'),'" & fcba & "','" & Session("CodeName") & "'); end;"
		  		connStr.execute(SQL)
				'ExecuteAndLog ScriptName, SQL
			END IF


			end if

			response.redirect("bkm_panen_edit.asp?bkmno="&bkmupdate)
			
	
%>

</body>
</html>