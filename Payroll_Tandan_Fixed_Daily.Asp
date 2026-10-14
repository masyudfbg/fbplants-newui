<%
	

	'Amandemen Premi Pandan
	'1. Code Allowance Panen = A24
	'   Qty Basis = Rate1 
	'   Premi Pertandan = Rate 2
	'2. Potongan Buah Mentah Code : D09 Rate1 = Rate
	'2. Potongan Buah Terlalu Masak Code : D10 Rate1 = Rate
	Dim nQtyBasis,nQtyPremiTandan,nRpPremiTandan,nRatePremiTandan
	Dim qty_buah_mentah,qty_tangkai_panjang,nTangkaiPanjang,nBuahTinggal,nTdkDisusun,nPelepah,nPremiPanen
	Dim qEmployee, rsEmployee, sTempPresentGrade
	
	Dim rsharijumat, sqlharijumat, isharijumat
	Dim rsconstantarate, sqlconstantarate, valconstantarate

	
	' Take the Reference of Cost Activity, as per New Concept, by using this reference, we can directly refer to Original Table
	' by NAF, 12 May 2013
	' First, take the Original Table of RECEIVENOTE.

	tempProcessCode = "PAYROLL"
	tempProcessNameCode = "Payroll_Tandan_Fixed_Daily.asp"
	tempRefCodeReceive = "ATTENDANCE_GAD"

	qReference = "select REF_NUMBERKEYS, REF_STRINGSEPARATOR, REF_STRINGKEYS " & _
				  " from process_tableref_tab " & _
				  " where " & _
				  " processcode = '" & tempProcessCode & "' and " & _
				  " upper(processnamecode) =  upper('" & tempProcessNameCode & "') and " & _
				  " fccode =  '" & tempRefCodeReceive & "'"

	Set rsReferenceReceive = Server.CreateObject("ADODB.Recordset")
	rsReferenceReceive.CursorType = 3
	rsReferenceReceive.CursorLocation = 3
	rsReferenceReceive.Open qReference, connstr

	doneReceive = 1
	if not rsReferenceReceive.eof then
		txtRefNumberKeys_Receive = rsReferenceReceive("REF_NUMBERKEYS")
		txtRefStringSeparator_Receive = rsReferenceReceive("REF_STRINGSEPARATOR")
		txtRefStringKeys_Receive = rsReferenceReceive("REF_STRINGKEYS")
		doneReceive = 1
	else
		doneReceive = 0
	end if
	rsReferenceReceive.close
	set rsReferenceReceive = Nothing
	' End of Picking Referential Integrity
	
	
	qReference = " SELECT NVL( PARAMETER_API.GetParam1('PROCESS', 'PAYROLL_PANEN', '" & Session("fcBA") & "' ), 'DEFAULT') OPTIONS FROM DUAL"

	Set rsReferenceReceive = Server.CreateObject("ADODB.Recordset")
	rsReferenceReceive.CursorType = 3
	rsReferenceReceive.CursorLocation = 3
	rsReferenceReceive.Open qReference, connstr

	if not rsReferenceReceive.eof then
		Defaul_Process = rsReferenceReceive("OPTIONS")
	else
		Defaul_Process = "DEFAULT"
	end if
	rsReferenceReceive.close
	set rsReferenceReceive = Nothing	
	
	'response.write "'" & Defaul_Process & "'"

	IF trim(Defaul_Process) = "PANEN_BORONGAN" THEN
%>
	<!--#include virtual="/payroll_kerja_panen_daily.asp"-->
<%	elseif trim(Defaul_Process) = "MULTILEVEL_HARVESTING_01" THEN%>
<!--#include virtual="/payroll_multilevel_harvesting_01_daily.asp"-->
<!--#include virtual="/payroll_potongan_panen_daily.asp"-->


	
<%	
    Else
	response.write "Not Calculate Premi / Potongan Panen ..." & "<br>"
	END IF ' **** Pilihan Opsi Cara hitung Premi Panen **** '
%>