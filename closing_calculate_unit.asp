
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
Dim qart,rsart,statusaresta
Dim qpv,rspv,statuspv


ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

fcba  = Session("fcBA")
tahun = Request("tahun")
bulan = Request("bulan")
msgerror = ""
msg = ""


TimeAwal=Time()

Endingtime = FormatDateTime(CDate(Now))
'response.write Endingtime

qCheckStatus = "select UNIT_CLOSING closingStatus from EMPALLDEDSTATUS WHERE YEAR='" & tahun & "' AND MONTH='" & bulan & "' AND FCBA='" & fcba & "' "
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

'tahapan pengecekan status closing
'1. cek dahulu status closing untuk modul PAYROLL jika false lnjut ke point 2
'2. cek dahulu status closing pada modul PAYROLL, jika status false proses closing UNIT tidak bisa dilanjut, silahkan closing dahulu modul PAYROLL, jika true lanjut point 3
'3. jika semua true pada point 1-3 lnjut closing, 
'Proses closing melakukan:
'insert ke tabel EMPALLDED_UNIT_CLOSING sesuai module masing-masing
'hitung skor sesuai tgl akhir di empalldedstatus
' update skor, update tgl proses, update status true di modul PAYROLL_CLOSING


IF txtClosingStatus = "FALSE" THEN  

		
			' cek dahulu di tabel closing unit untuk module yang di process
			qLineMax = "select count(SKOR) as total from EMPALLDED_UNIT_CLOSING where fcba = '" & session("fcBA") & "' AND YEAR='" & tahun & "' AND MONTH='" & bulan & "' AND  MODULENAME = 'UNIT_CLOSING'" 	
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
							" , sysdate,'UNIT_CLOSING' " & _
							" ,to_char(sysdate, 'HH:MI'))"  
				
				connStr.execute(strClosingLog)


				' HITUNG SKOR SESUAI SELISIH HARI
				q = "SELECT " & _
				"	end_date_closing, " & _
				"	end_date_process, " & _
				"	modulename, " & _
				"	CASE  " & _
				"		WHEN (end_date_closing - end_date_process + 3) < 0 THEN 100 " & _
				"		WHEN (end_date_closing - end_date_process + 3) BETWEEN 1 AND 2 THEN 80 " & _
				"		WHEN (end_date_closing - end_date_process + 3) BETWEEN 3 AND 5 THEN 60 " & _
				"		WHEN (end_date_closing - end_date_process + 3) > 5 THEN 40  " & _
				"	END AS skor " & _ 
				" FROM EMPALLDED_UNIT_CLOSING " & _
				" WHERE FCBA='" & fcba & "' AND YEAR='" & tahun & "' AND MONTH='" & bulan & "' AND MODULENAME='UNIT_CLOSING'"

				set rsq = Server.CreateObject("ADODB.Recordset")
				rsq.Open q, connstr
				if not rsq.eof then
					skordata = rsq("skor")
				else
					skordata = 0
				end if 

				qProcessPayroll = "UPDATE EMPALLDED_UNIT_CLOSING SET SKOR='" & skordata & "',lasttime=to_char(sysdate, 'HH:MI') WHERE YEAR='" & tahun & "' AND MONTH='" & bulan & "' AND FCBA='" & fcba & "' AND MODULENAME='UNIT_CLOSING'"	
				connStr.execute(qProcessPayroll)


			ELSE
				'perbaharui tgl proses closing
				qProcessPayroll = "UPDATE EMPALLDED_UNIT_CLOSING SET END_DATE_PROCESS=TO_DATE('" & Endingtime & "', 'dd/mm/yyyy hh24:mi:ss') WHERE YEAR='" & tahun & "' AND MONTH='" & bulan & "' AND FCBA='" & fcba & "' AND MODULENAME='UNIT_CLOSING'"	
				connStr.execute(qProcessPayroll)
				
				' HITUNG SKOR SESUAI SELISIH HARI
				q = "SELECT " & _
				"	end_date_closing, " & _
				"	end_date_process, " & _
				"	modulename, " & _
				"	CASE  " & _
				"		WHEN (end_date_closing - end_date_process + 3) < 0 THEN 100 " & _
				"		WHEN (end_date_closing - end_date_process + 3) BETWEEN 1 AND 2 THEN 80 " & _
				"		WHEN (end_date_closing - end_date_process + 3) BETWEEN 3 AND 5 THEN 60 " & _
				"		WHEN (end_date_closing - end_date_process + 3) > 5 THEN 40  " & _
				"	END AS skor " & _ 
				" FROM EMPALLDED_UNIT_CLOSING " & _
				" WHERE FCBA='" & fcba & "' AND YEAR='" & tahun & "' AND MONTH='" & bulan & "' AND MODULENAME='UNIT_CLOSING'"

				set rsq = Server.CreateObject("ADODB.Recordset")
				rsq.Open q, connstr
				if not rsq.eof then
					skordata = rsq("skor")
				else
					skordata = 0
				end if 

				'perbaharui skor
				qProcessPayroll = "UPDATE EMPALLDED_UNIT_CLOSING SET SKOR='" & skordata & "',lasttime=to_char(sysdate, 'HH:MI'),lastupdate=TO_DATE('" & Endingtime & "', 'dd/mm/yyyy hh24:mi:ss') WHERE YEAR='" & tahun & "' AND MONTH='" & bulan & "' AND FCBA='" & fcba & "' AND MODULENAME='UNIT_CLOSING'"	
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
						" , sysdate, to_char(sysdate, 'HH:MI'),'UNIT_CLOSING')"  
			
			connStr.execute(strClosingLog)
			

			'UPDATE STATUS TRUE DI STOCK_CLOSING'
			qProcessPayroll = "UPDATE EMPALLDEDSTATUS SET UNIT_CLOSING='TRUE',ARESTA_CLOSING='TRUE',KASBANK_CLOSING='TRUE',STOCK_CLOSING='TRUE',WIP_CLOSING='TRUE',WSVRA_CLOSING='TRUE',PAYROLL_CLOSING='TRUE' WHERE YEAR='" & tahun & "' AND MONTH='" & bulan & "' AND FCBA='" & fcba & "' AND UNIT_CLOSING='FALSE' OR UNIT_CLOSING IS NULL"
			''	response.write(qProcessPayroll)		
			connstr.execute(qProcessPayroll)
			
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
			msgerror =  "<font face=Verdana size=1 color=red>" & "BUSINESSUNIT : " & fcba & " Accounting Period " & bulan & "/" & tahun & " LOCKED !!! " & "</font><br>" 
			msg = msg &  "<font face=Verdana size=1 color=red>" & "BUSINESSUNIT : " & fcba & " Accounting Period " & bulan & "/" & tahun & "LOCKED !!! " & "</font><br>" 
			msg = msg &  " End Proses Closing " & fcba & " ,Time: " & TimeAkhir & " Elapsed " & Hour(Lama) & " Hours " & Minute(Lama) & " Minutes " &  Second(Lama) & " Seconds." & "<br>" 
END IF
		
	
	Response.write msg
	
   
     
%>