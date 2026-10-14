<%validatesession%>
<!--#include file="connDb1.Inc"-->
<html>
<body>
<%

			dim sqlhdr,sqlupah,sqlnomor,rsnomor,bkmno,sqldelkosong
			dim fcba,tglbkm,submit
			dim nQtdFields,nCountwip,linenokey,gangcode,ScriptName
			ScriptName	=	Request.ServerVariables("SCRIPT_NAME")
			
			submit 			= Request("simpan")
			tglbkm 			= Request("tglbkm")
			contractorcode 	= Request("contractorcode")
			bkmno 			= Request("bkmno")
			fcba 			= session("fcba")
			clasification 	= "504"
						


			nQtdFields 	= request.form("jumldatawip")
			nCountwip 	= 0
			linenokey 	= 1	


			if submit = "Simpan" then 
			
		

			'delete wip
			delupah = "delete from BKMWIPDTL where bkmno='" & bkmno & "'"
		''	connstr.BeginTrans
			'connStr.execute(delupah)
			ExecuteAndLog "bkm_wip_prosesedit.asp", delupah



			sqlhdr = "update bkmhdr set BKMDATE=TO_DATE('" & tglbkm & "', 'mm/dd/yyyy'),CONTRACTORCODE='" & contractorcode & "', " & _
			" FCEDIT='" & Session("CodeName") & "', " & _ 
			" LASTUPDATE=SYSDATE " & _
			" where bkmno='" & bkmno & "'"
			
			'connStr.execute(sqlhdr)
			ExecuteAndLog "bkm_wip_prosesedit.asp", sqlhdr


		' 	Response.Write sqlhdr
			
			do while cint(nCountwip) <= cint(nQtdFields)  

					dim  spkwip 
					spkwip = request.form("spkno" & nCountwip)
					
					
					if spkwip <>"" then 
			 
						sqlwip = "insert into BKMWIPDTL(BKMNO,CONTRACTORCODE,AGREEMENTCODE,CHARGEUNIT,ACTCODE, " & _
						" LOCATIONTYPE,LOCATIONCODE,QTY,OUTPUT,UNIT, " & _
						" REMARKS,CREATEDATE,LASTUPDATE,FCENTRY,FCEDIT, " & _
						" FCIP,FCBA,LINENOKEY,BKMAGREEMENTLINENO,RATE)" & _
						" values('" & bkmno & "','" & contractorcode & "','" & request.form("spkno" & nCountwip) & "', " & _ 
						" '" & request.form("lokasi" & nCountwip) & "','" & request.form("activitywip" & nCountwip) & "','-','" & request.form("tcwip" & nCountwip) & "', " & _ 
						" '" & request.form("qtywip" & nCountwip) & "','" & request.form("outputwip" & nCountwip) & "', " & _ 
						" '" & request.form("unitwip" & nCountwip) & "','',TO_DATE('" & tglbkm & "', 'mm/dd/yyyy'),SYSDATE,'" & Session("CodeName") & "','" & Session("CodeName") & "', " & _ 
						" '" & Request.ServerVariables("REMOTE_ADDR") & "', " & _ 
						" '" & fcba & "','" & linenokey & "','" & request.form("linenowip" & nCountwip) & "','" & request.form("ratewip" & nCountwip) & "')"
						'Response.Write sqlwip
						'connStr.execute(sqlwip)
						ExecuteAndLog "bkm_wip_prosesedit.asp", sqlwip

					end if

			 nCountwip 	= nCountwip + 1
			 linenokey  = linenokey + 1
			 loop

			 sqldelkosong = "delete from BKMWIPDTL where AGREEMENTCODE is null and BKMNO='" & bkmno & "'"
			 'connStr.execute(sqldelkosong)
			 ExecuteAndLog "bkm_wip_prosesedit.asp", sqldelkosong
			 
			

			'-------------------Batas proses upah-------------------------------------'

			

           	SQL = "begin bkmproses_wip_api.insert_bkm('" & bkmno & "',TO_DATE('" & tglbkm & "', 'mm/dd/yyyy'),'" & fcba & "','" & Session("CodeName") & "'); end;"
		  	'connStr.execute(SQL)
			ExecuteAndLog "bkm_wip_prosesedit.asp", SQL
			'ExecuteAndLog ScriptName, SQL

			response.redirect("bkm_wip_edit.asp?bkmno="&bkmno)

			end if 
			if submit = "Delete" THEN
				bkmupdate 		= request("bkmno")
				tglbkm 			= Request("tglbkm")
				fcba 			= session("fcba")
				SQL = "begin bkmproses_upah_api.delete_bkm('" & bkmupdate & "',TO_DATE('" & tglbkm & "', 'mm/dd/yyyy'),'" & fcba & "','" & Session("CodeName") & "'); end;"
		  		'connStr.execute(SQL)
				ExecuteAndLog "bkm_wip_prosesedit.asp", SQL
				'ExecuteAndLog ScriptName, SQL

				'delete bkm wip
				SQL = "begin bkmproses_wip_api.Deleteworkinprogress('" & bkmupdate & "','" & fcba & "'); end;"
		  		'connStr.execute(SQL)
				ExecuteAndLog "bkm_wip_prosesedit.asp", SQL
				'ExecuteAndLog ScriptName, SQL

				'delete jurnal bkm wip
				SQL = "begin bkmproses_wip_api.Deljurnalwip('" & bkmupdate & "','" & fcba & "'); end;"
		  		'connStr.execute(SQL)
				ExecuteAndLog "bkm_wip_prosesedit.asp", SQL
				'ExecuteAndLog ScriptName, SQL

				response.redirect("bkm_wip_edit.asp?bkmno="&bkmno)
			END IF 

			
			
			
	
%>

</body>
</html>