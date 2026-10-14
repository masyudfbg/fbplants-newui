<%validatesession%>
<!--#include file="connDb1.Inc"-->
<html>
<body>
<%

			dim sqlhdr,sqlupah,sqlnomor,rsnomor,bkmno,sqldelkosong
			dim fcba
			dim nQtdFields,nCountwip,linenokey,gangcode,ScriptName
			ScriptName	=	Request.ServerVariables("SCRIPT_NAME")
			
			tglbkm 			= Request("tglbkm")
			contractorcode 	= Request("contractorcode")
			fcba 			= session("fcba")
			clasification 	= "504"
						

			sqlnomor = "select 'BKM/'||'" & fcba & "'||'/'||To_Char(to_date('" & tglbkm & "','mm/dd/yyyy') ,'yyyymm')||'/'||SubStr(standard_hash(('" & contractorcode & "'||To_Char(to_date('" & tglbkm & "','mm/dd/yyyy') ,'yyyymmdd')||'" & fcba & "'),'MD5'),0,10) bkmno " & _
			" from dual "
			'Response.Write sqlnomor
			set rsnomor = connStr.Execute(sqlnomor)
			
			if not rsnomor.eof then
				bkmno = rsnomor("bkmno")			
			end if


			nQtdFields 	= request.form("jumldatawip")
			nCountwip 		= 1
			linenokey 	= 1	

			sqlhdr = "insert into bkmhdr(BKMNO,BKMDATE,CONTRACTORCODE,CREATEDATE,LASTUPDATE,FCENTRY,FCEDIT,FCIP,fcba,POSTINGFLG,ROWSTATE,DOCUMENT_CLASSIFICATION) " & _
			" values('" & bkmno & "',TO_DATE('" & tglbkm & "', 'mm/dd/yyyy'),'" & contractorcode & "', " & _ 
			" TO_DATE('" & tglbkm & "', 'mm/dd/yyyy'),SYSDATE,'" & Session("CodeName") & "', " & _ 
			" '" & Session("CodeName") & "','" & Request.ServerVariables("REMOTE_ADDR") & "','" & fcba & "','0','Planned','" & clasification & "')"
			'connstr.BeginTrans
			'connStr.execute(sqlhdr)
			ExecuteAndLog "bkm_wip_proses.asp", sqlhdr

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
						" '" & Request.ServerVariables("REMOTE_ADDR") & "','" & fcba & "', " & _ 
						" '" & linenokey & "','" & request.form("linenowip" & nCountwip) & "','" & request.form("ratewip" & nCountwip) & "')"
						'connStr.execute(sqlwip)
						ExecuteAndLog "bkm_wip_proses.asp", sqlwip

					end if

			 nCountwip 	= nCountwip + 1
			 linenokey  = linenokey + 1
			 loop

			 sqldelkosong = "delete from BKMWIPDTL where AGREEMENTCODE is null and BKMNO='" & bkmno & "'"
			 'connStr.execute(sqldelkosong)
			 ExecuteAndLog "bkm_wip_proses.asp", sqldelkosong

			

			'-------------------Batas proses upah-------------------------------------'



           	SQL = "begin bkmproses_wip_api.insert_bkm('" & bkmno & "',TO_DATE('" & tglbkm & "', 'mm/dd/yyyy'),'" & fcba & "','" & Session("CodeName") & "'); end;"
		  	'connStr.execute(SQL)
			ExecuteAndLog "bkm_wip_proses.asp", SQL

			'ExecuteAndLog ScriptName, SQL

			'If err.number = 0 Then
			'connStr.CommitTrans
			'End If
			'If err.number <> 0 Then
			'connStr.RollbackTrans
			'End If

			response.redirect("bkm_wip_edit.asp?bkmno="&bkmno)
			
			
	
%>

</body>
</html>