
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


dim txtcari,fcba,tahun,bulan,msg,msgerror,company,filtercompany1,fcbaloop,ScriptName
Dim qABC, qLineMax, rsqLineMax, tempMaxLine
Dim TimeAwal,TimeAkhir,Hitung,Lama, sqlpt
Dim sqldate,rsdate,edndate
Dim q,rsq,skordata,lockdata,txtClosingStatus
Dim v_bulan_after,v_tahun_after
ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

fcba  		= Session("fcBA")
tahun 		= Request("tahun")
bulan 		= Request("bulan")
lockdata  	= Request("lock")
msgerror = ""
msg = ""


IF  bulan = 12 THEN
    v_bulan_after = 1
    v_tahun_after = tahun + 1
ELSE
    v_bulan_after = bulan + 1
    v_tahun_after = tahun 
END IF

if lockdata = 1 then

	TimeAwal=Time()

	Endingtime = FormatDateTime(CDate(Now))
	'response.write Endingtime

	qCheckStatus = "select ARESTA_CLOSING closingStatus from EMPALLDEDSTATUS WHERE YEAR='" & tahun & "' AND MONTH='" & bulan & "' AND FCBA='" & fcba & "'"
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
	msg = msg &   "<font face=Verdana size=1 color=green>" & " Begin Proses Closing " & fcba & " Accounting Period: " & bulan & "/" & tahun & " Time: " & TimeAwal & "</font><br><br>" 



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

	
	IF UCASE(TRIM(txtClosingStatus & "")) = "FALSE" THEN
		
		' proses closing
		' cek dahulu di tabel closing unit untuk module yang di process
		qLineMax = "select count(SKOR) as total from EMPALLDED_UNIT_CLOSING where fcba = '" & session("fcBA") & "' AND YEAR='" & tahun & "' AND MONTH='" & bulan & "' AND  MODULENAME = 'ARESTA_CLOSING'" 	
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
						" , sysdate,'ARESTA_CLOSING' " & _
						" ,to_char(sysdate, 'HH:MI'))"  
			
			connStr.execute(strClosingLog)


			' HITUNG SKOR SESUAI SELISIH HARI
			q = "SELECT " & _
			"	end_date_closing, " & _
			"	end_date_process, " & _
			"	modulename, " & _
			"	NVL(CASE  " & _
			"		WHEN (end_date_closing - end_date_process + 1) < 0 THEN 100 " & _
			"		WHEN (end_date_closing - end_date_process + 1) BETWEEN 1 AND 2 THEN 80 " & _
			"		WHEN (end_date_closing - end_date_process + 1) BETWEEN 3 AND 5 THEN 60 " & _
			"		WHEN (end_date_closing - end_date_process + 1) > 5 THEN 40  " & _
			"		ELSE 0 " & _
			"	END, 0) AS skor " & _ 
			" FROM EMPALLDED_UNIT_CLOSING " & _
			" WHERE FCBA='" & fcba & "' AND YEAR='" & tahun & "' AND MONTH='" & bulan & "' AND MODULENAME='ARESTA_CLOSING'"

			set rsq = Server.CreateObject("ADODB.Recordset")
			rsq.Open q, connstr
			if not rsq.eof then
				skordata = cdbl(rsq("skor"))
			else
				skordata = 0
			end if 

			'Hapus data di aresta FIELD_HIST_TAB
			q = "DELETE FROM FIELD_HIST_TAB WHERE YEAR='" & v_tahun_after & "' AND MONTH='" & v_bulan_after & "' AND FCBA='" & fcba & "'"
			connStr.execute(q)

			'Insert Aresta untuk periode bulan depan
			q = "INSERT INTO FIELD_HIST_TAB (YEAR, MONTH, FCCODE, FCNAME, DIVISION, " & _ 
			"HECTARAGEPLANTED, TERRAINTYPE, SOILTYPE, OWNERSHIP, ACTIVATION, " & _ 
			"CROPTYPE, PLANTINGDATE, PLANTINGMATERIAL, TOTALSTAND, STATUS, " & _ 
			"PREVSEMESTER_ABW, PRESSEMESTER_ABW, HARVESTINGBASED_ABW, MAIN_ROAD, COLLECTIONROAD, " & _ 
			"MAINDRAIN, COLLECTIONDRAIN, SUBSIDIARYDRAIN, COLLECTIONPOINT, HARVESTINGBRIDGE, " & _ 
			"PERMANENTBRIDGE, WOODBRIDGE, BUIST, OWLNEST, FCENTRY, " & _ 
			"FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME, " & _ 
			"HARVESTINGSECTION, FIELD_INFO_01, FIELD_INFO_02, FIELD_INFO_03, FIELD_INFO_04, " & _ 
			"FIELD_INFO_05, ISACTIVE_HARVESTED, VALID_FROM, LINENO, ACTION_STATE, " & _ 
			"AREAL_LC, CADANGAN, OKUPASI, BIBITAN, PABRIK, " & _ 
			"EMPLASMEN, PARIT, JALAN, PRASARANA_LAINNYA, QUARRY, " & _ 
			"TANGGUL, ENCLAVE, TAMBANG, KONSERVASI, BUKIT_SUNGAI, " & _ 
			"NO_IMAGERY, LUAS_GROSS, HECTARAGEHARVESTED, PRODUCTIVESTAND, HARVESTINGSECTIONNUMBER, " & _ 
			"PLANTINGPOINT, FULLSTAND,DX_BLOCK_STATUS)" & _ 
			"	Select '" & v_tahun_after & "','" & v_bulan_after & "', FCCODE, FCNAME, DIVISION, " & _ 
			"		HECTARAGEPLANTED, TERRAINTYPE, SOILTYPE, OWNERSHIP, ACTIVATION, " & _ 
			"		CROPTYPE, PLANTINGDATE, PLANTINGMATERIAL, TOTALSTAND, STATUS, " & _ 
			"		PREVSEMESTER_ABW, PRESSEMESTER_ABW, HARVESTINGBASED_ABW, MAIN_ROAD, COLLECTIONROAD, " & _ 
			"		MAINDRAIN, COLLECTIONDRAIN, SUBSIDIARYDRAIN, COLLECTIONPOINT, HARVESTINGBRIDGE, " & _ 
			"		PERMANENTBRIDGE, WOODBRIDGE, BUIST, OWLNEST, FCENTRY, " & _ 
			"		FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME, " & _ 
			"		HARVESTINGSECTION, FIELD_INFO_01, FIELD_INFO_02, FIELD_INFO_03, FIELD_INFO_04, " & _ 
			"		FIELD_INFO_05, ISACTIVE_HARVESTED, VALID_FROM, LINENO, ACTION_STATE, " & _ 
			"		AREAL_LC, CADANGAN, OKUPASI, BIBITAN, PABRIK, " & _ 
			"		EMPLASMEN, PARIT, JALAN, PRASARANA_LAINNYA, QUARRY, " & _ 
			"		TANGGUL, ENCLAVE, TAMBANG, KONSERVASI, BUKIT_SUNGAI, " & _ 
			"		NO_IMAGERY, LUAS_GROSS, HECTARAGEHARVESTED, PRODUCTIVESTAND, HARVESTINGSECTIONNUMBER, " & _ 
			"		PLANTINGPOINT, FULLSTAND,DX_BLOCK_STATUS " & _ 
			"	From   FIELD_HIST_TAB WHERE FCBA='" & fcba & "' AND YEAR='" & tahun & "' AND MONTH='" & bulan & "'"
			connStr.execute(q)

			qProcessPayroll = "UPDATE EMPALLDED_UNIT_CLOSING SET SKOR='" & skordata & "' WHERE YEAR='" & tahun & "' AND MONTH='" & bulan & "' AND FCBA='" & fcba & "' AND MODULENAME='ARESTA_CLOSING'"	
			connStr.execute(qProcessPayroll)


		ELSE

			qProcessPayroll = "UPDATE EMPALLDED_UNIT_CLOSING SET END_DATE_PROCESS=TO_DATE('" & Endingtime & "', 'dd/mm/yyyy hh24:mi:ss') WHERE YEAR='" & tahun & "' AND MONTH='" & bulan & "' AND FCBA='" & fcba & "' AND MODULENAME='ARESTA_CLOSING'"	
			connStr.execute(qProcessPayroll)
			
			' HITUNG SKOR SESUAI SELISIH HARI
			q = "SELECT " & _
			"	end_date_closing, " & _
			"	end_date_process, " & _
			"	modulename, " & _
			"	NVL(CASE  " & _
			"		WHEN (end_date_closing - end_date_process + 1) < 0 THEN 100 " & _
			"		WHEN (end_date_closing - end_date_process + 1) BETWEEN 1 AND 2 THEN 80 " & _
			"		WHEN (end_date_closing - end_date_process + 1) BETWEEN 3 AND 5 THEN 60 " & _
			"		WHEN (end_date_closing - end_date_process + 1) > 5 THEN 40  " & _
			"		ELSE 0 " & _
			"	END, 0) AS skor " & _ 
			" FROM EMPALLDED_UNIT_CLOSING " & _
			" WHERE FCBA='" & fcba & "' AND YEAR='" & tahun & "' AND MONTH='" & bulan & "' AND MODULENAME='ARESTA_CLOSING'"

			set rsq = Server.CreateObject("ADODB.Recordset")
			rsq.Open q, connstr
			if not rsq.eof then
				If IsNull(rsq("skor")) Then
					skordata = 0
				Else
					skordata = cdbl(rsq("skor"))
				End If
			else
				skordata = 0
			end if 
			rsq.close
			set rsq = Nothing

			qProcessPayroll = "UPDATE EMPALLDED_UNIT_CLOSING SET SKOR='" & skordata & "',lastupdate=TO_DATE('" & Endingtime & "', 'dd/mm/yyyy hh24:mi:ss') WHERE YEAR='" & tahun & "' AND MONTH='" & bulan & "' AND FCBA='" & fcba & "' AND MODULENAME='ARESTA_CLOSING'"	
			connStr.execute(qProcessPayroll)

			'Hapus data di aresta FIELD_HIST_TAB
			q = "DELETE FROM FIELD_HIST_TAB WHERE YEAR='" & v_tahun_after & "' AND MONTH='" & v_bulan_after & "' AND FCBA='" & fcba & "'"
			connStr.execute(q)

			'Insert Aresta untuk periode bulan depan
			q = "INSERT INTO FIELD_HIST_TAB (YEAR, MONTH, FCCODE, FCNAME, DIVISION, " & _ 
			"HECTARAGEPLANTED, TERRAINTYPE, SOILTYPE, OWNERSHIP, ACTIVATION, " & _ 
			"CROPTYPE, PLANTINGDATE, PLANTINGMATERIAL, TOTALSTAND, STATUS, " & _ 
			"PREVSEMESTER_ABW, PRESSEMESTER_ABW, HARVESTINGBASED_ABW, MAIN_ROAD, COLLECTIONROAD, " & _ 
			"MAINDRAIN, COLLECTIONDRAIN, SUBSIDIARYDRAIN, COLLECTIONPOINT, HARVESTINGBRIDGE, " & _ 
			"PERMANENTBRIDGE, WOODBRIDGE, BUIST, OWLNEST, FCENTRY, " & _ 
			"FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME, " & _ 
			"HARVESTINGSECTION, FIELD_INFO_01, FIELD_INFO_02, FIELD_INFO_03, FIELD_INFO_04, " & _ 
			"FIELD_INFO_05, ISACTIVE_HARVESTED, VALID_FROM, LINENO, ACTION_STATE, " & _ 
			"AREAL_LC, CADANGAN, OKUPASI, BIBITAN, PABRIK, " & _ 
			"EMPLASMEN, PARIT, JALAN, PRASARANA_LAINNYA, QUARRY, " & _ 
			"TANGGUL, ENCLAVE, TAMBANG, KONSERVASI, BUKIT_SUNGAI, " & _ 
			"NO_IMAGERY, LUAS_GROSS, HECTARAGEHARVESTED, PRODUCTIVESTAND, HARVESTINGSECTIONNUMBER, " & _ 
			"PLANTINGPOINT, FULLSTAND,DX_BLOCK_STATUS)" & _ 
			"	Select '" & v_tahun_after & "','" & v_bulan_after & "', FCCODE, FCNAME, DIVISION, " & _ 
			"		HECTARAGEPLANTED, TERRAINTYPE, SOILTYPE, OWNERSHIP, ACTIVATION, " & _ 
			"		CROPTYPE, PLANTINGDATE, PLANTINGMATERIAL, TOTALSTAND, STATUS, " & _ 
			"		PREVSEMESTER_ABW, PRESSEMESTER_ABW, HARVESTINGBASED_ABW, MAIN_ROAD, COLLECTIONROAD, " & _ 
			"		MAINDRAIN, COLLECTIONDRAIN, SUBSIDIARYDRAIN, COLLECTIONPOINT, HARVESTINGBRIDGE, " & _ 
			"		PERMANENTBRIDGE, WOODBRIDGE, BUIST, OWLNEST, FCENTRY, " & _ 
			"		FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME, " & _ 
			"		HARVESTINGSECTION, FIELD_INFO_01, FIELD_INFO_02, FIELD_INFO_03, FIELD_INFO_04, " & _ 
			"		FIELD_INFO_05, ISACTIVE_HARVESTED, VALID_FROM, LINENO, ACTION_STATE, " & _ 
			"		AREAL_LC, CADANGAN, OKUPASI, BIBITAN, PABRIK, " & _ 
			"		EMPLASMEN, PARIT, JALAN, PRASARANA_LAINNYA, QUARRY, " & _ 
			"		TANGGUL, ENCLAVE, TAMBANG, KONSERVASI, BUKIT_SUNGAI, " & _ 
			"		NO_IMAGERY, LUAS_GROSS, HECTARAGEHARVESTED, PRODUCTIVESTAND, HARVESTINGSECTIONNUMBER, " & _ 
			"		PLANTINGPOINT, FULLSTAND,DX_BLOCK_STATUS " & _ 
			"	From   FIELD_HIST_TAB WHERE FCBA='" & fcba & "' AND YEAR='" & tahun & "' AND MONTH='" & bulan & "'"
			connStr.execute(q)




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
					" , sysdate, to_char(sysdate, 'HH:MI'),'ARESTA_CLOSING')"  
		
		connStr.execute(strClosingLog)
		
		msg = msg &  "<font face=Verdana size=4 color=blue>" & "PROSES CLOSING ARESTA SUKSES...!!" & "</font><br>" 
		

	ELSE
		' sudah terkunci
		msg = msg &  "<font face=Verdana size=4 color=blue>" & "CLOSING ARESTA TERKUNCI...!!" & "</font><br>" 
	
	END IF

ELSE 
	'response.write "tes lock"
	'Proses lock aresta
	qProcessPayroll = "UPDATE EMPALLDEDSTATUS SET ARESTA_CLOSING='TRUE' WHERE YEAR='" & tahun & "' AND MONTH='" & bulan & "' AND FCBA='" & fcba & "' AND ARESTA_CLOSING='FALSE' OR ARESTA_CLOSING IS NULL"
	connstr.execute(qProcessPayroll)
	msg = msg &  "<font face=Verdana size=4 color=blue>" & "PROSES KUNCI ARESTA SUKSES...!!" & "</font><br>" 

end if

	msg = msg &   "<br>" 
	msg = msg &   "<font face=Verdana size=1 color=green>" & " End Proses Closing " & fcba & " Accounting Period: " & bulan & "/" & tahun & " Time: " & Endingtime & "</font><br><br>"
	Response.write msg
	
   
     
%>