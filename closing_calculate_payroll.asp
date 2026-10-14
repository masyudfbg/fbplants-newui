
<%
Response.CharSet="euc-kr"
Session.codepage="949"
Response.codepage="949"
Response.ContentType="text/html;charset=euc-kr"
%>
<%
Response.LCID = 1046 
%>

<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->

<!--#includes file="include/JSON.asp" --> 
<!--#includes file="include/json.objectClass.asp"-->
<!--#includes file="include/aspJSON.asp" -->

<%


dim txtcari,fcba,tahun,bulan,msg,msgerror,company,filtercompany1,ScriptName
Dim qABC, qLineMax, rsqLineMax, tempMaxLine
Dim TimeAwal,TimeAkhir,Hitung,Lama, sqlpt
Dim sqldate,rsdate,edndate
Dim q,rsq,skordata
Dim qart,rsart,statuswsvra,lockdata
Dim qpv,rspv,statuspv, sqlexec,qProcessPayroll


ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

fcba  = Session("fcBA")
tahun = Request("tahun")
bulan = Request("bulan")
msgerror = ""
msg = ""

lockdata  	= Request("lock")

if lockdata = 1 then


		TimeAwal=Time()
		Endingtime = FormatDateTime(CDate(Now))
		'response.write Endingtime

		qCheckStatus = "select payroll_closing closingStatus from EMPALLDEDSTATUS WHERE YEAR='" & tahun & "' AND MONTH='" & bulan & "' AND FCBA='" & fcba & "' "
		Set rsCheckStatus = Server.CreateObject("ADODB.Recordset")
		rsCheckStatus.CursorType = 3
		rsCheckStatus.CursorLocation = 3
		rsCheckStatus.Open qCheckStatus, connstr
		if not rsCheckStatus.eof then
			txtClosingStatus = rsCheckStatus("closingStatus")
		else
			msgerror =  "EmpalldedStatus are not defined "
			msg = msg &  "<font face=Verdana size=1 color=red>" & "EmpalldedStatus are not defined "
			txtClosingStatus="FALSE"
		end if
		rsCheckStatus.close
		set rsCheckStatus = Nothing
		msg = msg &   "<br>" 
		msg = msg &   "<font face=Verdana size=1 color=green>" & " Begin Proses Closing " & fcba & " Accounting Period: " & bulan & "/" & tahun & " Time: " & TimeAwal & "</font><br>" 



		' CEK END_DATE
		sqldate = "SELECT to_char(END_DATE,'mm/dd/yyyy')END_DATE FROM EmpalldedStatus WHERE fcba = '" & session("fcBA") & "' AND YEAR='" & tahun & "' AND MONTH='" & bulan & "'"
		set rsdate = Server.CreateObject("ADODB.Recordset")
		rsdate.Open sqldate, connstr
		if not rsdate.eof then
			edndate = (rsdate("END_DATE"))
		else
			edndate = "00/00/0000"
		end if 
		rsdate.close
		set rsdate = Nothing

		'tahapan pengecekan status closing ---
		'1. cek dahulu status closing untuk modul WSVRA jika false lnjut ke point 2
		'2. cek dahulu status closing pada modul WSVRA, jika status false proses closing PAYROLL tidak bisa dilanjut, silahkan closing dahulu modul WSVRA, jika true lanjut point 3
		'3. cek dahulu inputan BKM PERAWATAN DAN PANEN jika masih ada inputan belum approved closing tidak bisa dilanjut, jika iya lanjut point 4
		'4. jika semua true pada point 1-3 lnjut closing, 
		'Proses closing melakukan:
		'insert ke tabel EMPALLDED_UNIT_CLOSING sesuai module masing-masing
		'hitung skor sesuai tgl akhir di empalldedstatus
		' update skor, update tgl proses, update status true di modul PAYROLL_CLOSING


		IF txtClosingStatus = "FALSE" THEN  


				'cek dahulu jumlah BKM yang belum di approved
				qpv = "select COUNT(BKMNO) AS total " & _
				"	from  BKMHDR  " & _
				"	WHERE EXTRACT(YEAR FROM BKMDATE) = '" & tahun & "' " & _
				"	AND EXTRACT(MONTH FROM BKMDATE) = '" & bulan & "' " & _
				"	AND ROWSTATE<>'Approved' " & _
				"   AND DOCUMENT_CLASSIFICATION IN('501','503') " & _
				"	AND FCBA='" & fcba & "'" 
				set rspv = Server.CreateObject("ADODB.Recordset")
				rspv.Open qpv, connstr
				if not rspv.eof then
					statuspv = cdbl(rspv("total"))

				end if 

				'response.write statuspv
				'response.end

				if statuspv <= 0 THEN

				
					' cek dahulu di tabel closing unit untuk module yang di process
					qLineMax = "select count(SKOR) as total from EMPALLDED_UNIT_CLOSING where fcba = '" & session("fcBA") & "' AND YEAR='" & tahun & "' AND MONTH='" & bulan & "' AND  MODULENAME = 'PAYROLL_CLOSING'" 	
					'response.write qLineMax		
					set rsqLineMax = Server.CreateObject("ADODB.Recordset")
					rsqLineMax.Open qLineMax, connstr
					if not rsqLineMax.eof then
						tempMaxLine = cdbl(rsqLineMax("total"))
					else
						tempMaxLine = 0
					end if 

					IF tempMaxLine= "0" THEN

						'INSERT CLOSING LOG
						strClosingLog = " INSERT INTO EMPALLDED_UNIT_CLOSING(YEAR, MONTH, END_DATE_CLOSING, END_DATE_PROCESS,SKOR, FCBA, FCENTRY, FCIP, LASTUPDATE,MODULENAME,LASTTIME) VALUES (" & _
									" '" & tahun & "', " & _
									" '" & bulan & "'," & _
									" TO_DATE('" & edndate & "', 'mm/dd/yyyy')," & _
									" TO_DATE('" & Endingtime & "', 'dd/mm/yyyy hh24:mi:ss')," & _
									" '0' " & _
									" ,'" & fcba & "' " & _
									" ,'" & Session("CodeName") & "' " & _
									" ,'" & Request.ServerVariables("REMOTE_ADDR") & "' " & _
									" , sysdate,'PAYROLL_CLOSING' " & _
									" ,to_char(sysdate, 'HH:MI'))"  
						
						connStr.execute(strClosingLog)


						' HITUNG SKOR SESUAI SELISIH HARI
						q = "SELECT " & _
						"	end_date_closing, " & _
						"	end_date_process, " & _
						"	modulename, " & _
						"	NVL(CASE  " & _
						"		WHEN (end_date_closing - end_date_process + 3) < 0 THEN 100 " & _
						"		WHEN (end_date_closing - end_date_process + 3) BETWEEN 1 AND 2 THEN 80 " & _
						"		WHEN (end_date_closing - end_date_process + 3) BETWEEN 3 AND 5 THEN 60 " & _
						"		WHEN (end_date_closing - end_date_process + 3) > 5 THEN 40  " & _
						"		ELSE 0 " & _
						"	END, 0) AS skor " & _ 
						" FROM EMPALLDED_UNIT_CLOSING " & _
						" WHERE FCBA='" & fcba & "' AND YEAR='" & tahun & "' AND MONTH='" & bulan & "' AND MODULENAME='PAYROLL_CLOSING'"

						set rsq = Server.CreateObject("ADODB.Recordset")
						rsq.Open q, connstr
						if not rsq.eof then
							skordata = cdbl(rsq("skor"))
						else
							skordata = 0
						end if 

						qProcessPayroll = "UPDATE EMPALLDED_UNIT_CLOSING SET SKOR='" & skordata & "',lasttime=to_char(sysdate, 'HH:MI') WHERE YEAR='" & tahun & "' AND MONTH='" & bulan & "' AND FCBA='" & fcba & "' AND MODULENAME='PAYROLL_CLOSING'"	
						connStr.execute(qProcessPayroll)


					ELSE
						'perbaharui tgl proses closing
						qProcessPayroll = "UPDATE EMPALLDED_UNIT_CLOSING SET END_DATE_PROCESS=TO_DATE('" & Endingtime & "', 'dd/mm/yyyy hh24:mi:ss') WHERE YEAR='" & tahun & "' AND MONTH='" & bulan & "' AND FCBA='" & fcba & "' AND MODULENAME='PAYROLL_CLOSING'"	
						connStr.execute(qProcessPayroll)
						
						' HITUNG SKOR SESUAI SELISIH HARI
						q = "SELECT " & _
						"	end_date_closing, " & _
						"	end_date_process, " & _
						"	modulename, " & _
						"	NVL(CASE  " & _
						"		WHEN (end_date_closing - end_date_process + 3) < 0 THEN 100 " & _
						"		WHEN (end_date_closing - end_date_process + 3) BETWEEN 1 AND 2 THEN 80 " & _
						"		WHEN (end_date_closing - end_date_process + 3) BETWEEN 3 AND 5 THEN 60 " & _
						"		WHEN (end_date_closing - end_date_process + 3) > 5 THEN 40  " & _
						"		ELSE 0 " & _
						"	END, 0) AS skor " & _ 
						" FROM EMPALLDED_UNIT_CLOSING " & _
						" WHERE FCBA='" & fcba & "' AND YEAR='" & tahun & "' AND MONTH='" & bulan & "' AND MODULENAME='PAYROLL_CLOSING'"

						set rsq = Server.CreateObject("ADODB.Recordset")
						rsq.Open q, connstr
						if not rsq.eof then
							skordata = cdbl(rsq("skor"))
						else
							skordata = 0
						end if 

						'perbaharui skor
						qProcessPayroll = "UPDATE EMPALLDED_UNIT_CLOSING SET SKOR='" & skordata & "',lasttime=to_char(sysdate, 'HH:MI'),lastupdate=TO_DATE('" & Endingtime & "', 'dd/mm/yyyy hh24:mi:ss') WHERE YEAR='" & tahun & "' AND MONTH='" & bulan & "' AND FCBA='" & fcba & "' AND MODULENAME='PAYROLL_CLOSING'"	
						'response.write qProcessPayroll
						connStr.execute(qProcessPayroll)
					END IF 

					'INSERT CLOSING LOG
					strClosingLog = " INSERT INTO EMPALLDED_UNIT_CLOSING_LOG(YEAR, MONTH, END_DATE_CLOSING, END_DATE_PROCESS, FCBA, FCENTRY, FCIP, LASTUPDATE,LASTTIME, MODULNAME) VALUES (" & _
								" '" & tahun & "', " & _
								" '" & bulan & "'," & _
								" TO_DATE('" & edndate & "', 'mm/dd/yyyy')," & _
								" TO_DATE('" & Endingtime & "', 'dd/mm/yyyy hh24:mi:ss')," & _
								" '" & fcba & "' " & _
								" ,'" & Session("CodeName") & "' " & _
								" ,'" & Request.ServerVariables("REMOTE_ADDR") & "' " & _
								" , sysdate, to_char(sysdate, 'HH:MI'),'PAYROLL_CLOSING')"  
					
					connStr.execute(strClosingLog)

					q = "UPDATE  bkmjournalhdr SET postingflg=1 WHERE accountingperiod=" & bulan & " AND accountingyear=" & tahun & " AND fcba='" & fcba & "' AND postingflg <> 1 " & _ 
								" AND SOURCECODE in('OPR-BKM') "
								connStr.execute(q)



					' hitung closing payroll 
					sqlexec = "BEGIN " & _ 
					"bkmcalculate_payroll_api.proses('" & bulan & "','" & tahun & "','" & fcba & "','" & Session("CodeName") & "');" & _
							"END; " 
				''	On Error Resume Next
					connStr.Execute(sqlexec)

					qProcessPayroll = "BEGIN " & _ 
									"EMPALLDEDSTATUS_API.FinishProcessPayroll (" & tahun & "," & bulan & ",'" & fcba & "','" & Session("CodeName") & "','" & Session("CodeName") & "','" & Request.ServerVariables("REMOTE_ADDR") & "'); " & _ 
									"END;"		
				''	On Error Resume Next
					connstr.execute(qProcessPayroll)


					' hitung closing acounting 
					ProcessPayroll = "BEGIN " & _ 
								"fac_closing_accounting_api.ProsesClosingMonth (" & tahun & "," & bulan & ",'" & fcba & "','" & Session("CodeName") & "'); " & _ 
								"END;"		
				
					connstr.execute(qProcessPayroll)


					' hitung LBM 
					qProcessPayroll = "BEGIN " & _ 
								"opr_accountreport_bulanan_api.Proses (" & bulan & "," & tahun & ",'" & fcba & "'); " & _ 
								"END;"	
					connstr.execute(qProcessPayroll)

					' hitung PPH 
					If Session("COMPANYCODE") = "PT.ARJ" then
						sqlexec = "BEGIN " & _ 
								"bkmcalculate_pph21_api.proses('" & bulan & "','" & tahun & "','" & fcba & "');" & _
								"END; " 
						connStr.Execute(sqlexec)
					End If
					
							
					TimeAkhir=Time()
					Lama =TimeAkhir-TimeAwal

						If Err.Number = 0 Then
						msg = msg &   "<font face=Verdana size=1 color=green>" &  " End Proses Closing " & fcba & " Success,  Time: " & TimeAkhir & " Elapsed " & Hour(Lama) & " Hours " & Minute(Lama) & " Minutes " &  Second(Lama) & " Seconds." & "</font><br>" 
						Else
							msgerror =  "<font face=Verdana size=1 color=red>" & " Proses Error !!!" & "<br>" & Err.Description & "</font><br>" 
							msg = msg & "<font face=Verdana size=1 color=red>" & " Proses Error !!!" & "<br>" & Err.Description & "</font><br>" 
						End If
					On Error GoTo 0
				ELSE 
					msg = msg &  "<font face=Verdana size=1 color=red>" & "PROSES CLOSING PAYROLL BELUM BISA DI LANJUTKAN, DIKARENAKAN TRANSAKSI BKM PERAWATAN DAN PANEN MASIH ADA YANG BELUM DI APPROVED, SILAHKAN APPROVAL DAHULU..!! " & "</font><br>" 
				END IF

			
			


		ELSE
						msgerror =  "<font face=Verdana size=1 color=red>" & "BUSINESSUNIT : " & fcba & " Accounting Period " & bulan & "/" & tahun & " TERKUNCI !!! " & "</font><br>" 
						msg = msg &  "<font face=Verdana size=1 color=red>" & "BUSINESSUNIT : " & fcba & " Accounting Period " & bulan & "/" & tahun & "TERKUNCI !!! " & "</font><br>" 
						msg = msg &  " End Proses Closing " & fcba & " ,Time: " & TimeAkhir & " Elapsed " & Hour(Lama) & " Hours " & Minute(Lama) & " Minutes " &  Second(Lama) & " Seconds." & "<br>" 
		END IF

ELSE 
			'UPDATE STATUS TRUE DI PAYROLL_CLOSING'
			qProcessPayroll = "UPDATE EMPALLDEDSTATUS SET payroll_closing='TRUE' WHERE YEAR='" & tahun & "' AND MONTH='" & bulan & "' AND FCBA='" & fcba & "' AND payroll_closing='FALSE' OR payroll_closing IS NULL"
			connstr.execute(qProcessPayroll)
			msg = msg &  "<font face=Verdana size=4 color=blue>" & "PROSES KUNCI PAYROLL SUKSES...!!" & "</font><br>" 

end if	
		
Response.write msg
	
   
     
%>