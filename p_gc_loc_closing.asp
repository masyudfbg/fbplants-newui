<%
	Option Explicit
	response.buffer=True
%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>ePlaS - General_Charges Closing</title>
</head>
<body>
<%
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
Dim TimeAwal,TimeAkhir,Hitung,Lama
TimeAwal=Time()

Dim rsVehicle, qVehicle, nxPositive, nxNegative, nCostTotal
Dim nTotalAlokasi, rsCostVehicle, qCostVehicle, nBasis, rsBasis, qBasis, qInsertCR
Dim rsVHActivity, qVHActivity, nWSAlokasi, qInsertDB, LT, LC, FDDATE, FCJOB, nSelisih
Dim qABC, q, qLineMax, rsqLineMax, tempMaxLine
Dim strClosingLog, strLenghtime, rsLenghtime, upClosingLog, txtLastDay, txtNextMonth
Dim maxLineNoVehicleCB, tempProcessCode, tempProcessNameCode, tempRefCodePVDetail, qReference
Dim rsReferencePVDetail, donePVDetail, qNormalisasiDetail, qVoucher, rsVoucher, qCost_Activity
Dim qVoucherDetail, rsVoucherDetail, maxLineNoDetail
Dim txtRefStringSeparator_PVDetail, txtRefStringValues_PVDetail, nQuantity, nOutput
Dim txtRefNumberKeys_PVDetail, txtRefStringKeys_PVDetail, qCost_ActivityHistory, maxLineNoContractorCB
Dim rsContract, qContract, rsAggrement, qCostContract, nRate, maxLineNoBuildingCB, rsCostActivityMax, qCostActivityMax
Dim rsBuilding, qBuilding, rsCostBuilding, qCostBuilding, rsFieldAlokasi, qFieldAlokasi
Dim nTMorTBM, nFFAlokasi, maxLineNo, maxLineNoGcCB
Dim temp_DebetCredit

Dim arrRS, objRSArray, ExecuteSQL, qDebetAlokasi, qCreditAlokasi, i, jobcodeNoBasis
Dim qSelisih, objectSQL, objectRS, FirstRecord, UpdateAmoutSelisih
Dim Cost_In, Alloc_Db, Alloc_Cr

Dim qnon_upr, val_hasil, qborongan

Function GetRSArray(strSQL)
	set objRSArray = connStr.Execute(strSQL)
	If Not objRSArray.EOF Then
		arrRS = objRSArray.GetRows()
		GetRSArray = arrRS
	Else
		Response.Write("There are currently no records returned.<br/>")
		Exit Function
	End If
	objRSArray.Close()
	Set objRSArray = Nothing
End Function


'--Added by Ukie at 25-Oct-2012--	
Dim Startingtime, Endingtime
Startingtime = FormatDateTime(CDate(Now))
'--End Added--	
%>
<div id="status_running"><img src="ICON/working.gif"></div><br>
<font face="Verdana" size="1">
General Charges Last Period Close <%=FormatDateTime(Session("cls_General_Charges"),1)%> <br> 
General Charges Current Period is <%=FormatDateTime(Session("v_General_Charges"),1)%> <br><br>
Please Wait...<br><br>
<%
 Response.Write "Start ..." & Time() & "<br><br>" 

dim rsEmpAllDed,rsPO,qReceive,qPO,qControlJob,nPurchaseControl,rsControlJob,txtControlJob,qAllowance,nQtyReceived,qNormalisasi,nPresentRate,nDays,nAmount
dim v_AllowanceCode

' XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' Add by NAF to select lastperiodposted
Dim qPeriodControl, rsPeriodControl, temp_datePeriodControl

' Water Flow Check, By NAF
' 2 December 2013

Dim qWaterFlowCheck, rsWaterFlowCheck, CheckWaterFlow, tempErrorWaterFlow

qWaterFlowCheck = " select PERIODCONTROL_API.getStatusPrevClosing('General_Charges','" & Session("fcBA")  & "') waterFlowCheck from dual "

Set rsWaterFlowCheck = Server.CreateObject("ADODB.Recordset")
rsWaterFlowCheck.CursorType = 3
rsWaterFlowCheck.CursorLocation = 3
rsWaterFlowCheck.Open qWaterFlowCheck, connstr

if not rsWaterFlowCheck.eof then
    CheckWaterFlow = rsWaterFlowCheck("waterFlowCheck")
else
    Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error, No WATER FLOW CHECK for General_Charges " & "<br>"
end if
rsWaterFlowCheck.close
set rsWaterFlowCheck = Nothing


if CheckWaterFlow = "TRUE" then 

	qPeriodControl=" select  add_months(LASTPERIODPOSTED,1) LastPeriodPosted from PERIODCONTROL  " & _
	  " where SYSTEM='General_Charges' and fcba='" & Session("fcBA") & "'"

	Set rsPeriodControl = Server.CreateObject("ADODB.Recordset")
	rsPeriodControl.CursorType = 3
	rsPeriodControl.CursorLocation = 3
	rsPeriodControl.Open qPeriodControl, connstr

	if not rsPeriodControl.eof then
		temp_datePeriodControl = rsPeriodControl("LASTPERIODPOSTED")
	else
		Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error, No PERIODCONTROL for General_Charges " & "<br>"
	end if
	' End Added
	'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

	' Remarked by NAF, as per New Concept, 7 May 2013
	' For Workshop Closing, for 2nd Book in Accounting

	tempProcessCode = "GENERAL_CHARGES_CLOSING"
	tempProcessNameCode = "p_General_Charges_Closing.asp"
	tempRefCodePVDetail = "GCCB"

	' Take the Reference of Cost Activity, as per New Concept, by using this reference, we can directly refer to Original Table
	' by NAF, 12 May 2013
	' Take the Original Table of PAYMENTVOUCHERDETAIL.
	qReference = "select REF_NUMBERKEYS, REF_STRINGSEPARATOR, REF_STRINGKEYS " & _
				  " from process_tableref_tab " & _
				  " where " & _
				  " processcode = '" & tempProcessCode & "' and " & _
				  " upper(processnamecode) =  upper('" & tempProcessNameCode & "') and " & _
				  " fccode =  '" & tempRefCodePVDetail & "'"

	Set rsReferencePVDetail = Server.CreateObject("ADODB.Recordset")
	rsReferencePVDetail.CursorType = 3
	rsReferencePVDetail.CursorLocation = 3
	rsReferencePVDetail.Open qReference, connstr

	donePVDetail = 1
	if not rsReferencePVDetail.eof then
		txtRefNumberKeys_PVDetail = rsReferencePVDetail("REF_NUMBERKEYS")
		txtRefStringSeparator_PVDetail = rsReferencePVDetail("REF_STRINGSEPARATOR")
		txtRefStringKeys_PVDetail = rsReferencePVDetail("REF_STRINGKEYS")
		donePVDetail = 1
	else
		donePVDetail = 0
	end if
	rsReferencePVDetail.close
	set rsReferencePVDetail = Nothing




	' Get the Max LineNo of Cost Activity
	Set rsCostActivityMax = Server.CreateObject("ADODB.Recordset")
	rsCostActivityMax.CursorType = 3
	rsCostActivityMax.CursorLocation = 3
									
	qCostActivityMax =  " select nvl(max(nvl(lineno,0)),0) maxLineNo from gccb " & _
						" where fcba = '" &  Session("fcBA") & "' " 
	rsCostActivityMax.Open qCostActivityMax, connstr
	
	If Not rsCostActivityMax.EOF Then
		maxLineNoGcCB = cdbl(rsCostActivityMax("maxLineNo"))
	Else
		maxLineNoGcCB = 0
	End If

	rsCostActivityMax.close
	set rsCostActivityMax = Nothing
	' End of Get the Max LineNo of Cost Activity


	'---- Control Job
	qControlJob="select ControlJob from Module where NAME='General_Charges'"
	Set rsControlJob = Server.CreateObject("ADODB.Recordset")
	rsControlJob.CursorType = 3
	rsControlJob.CursorLocation = 3
	rsControlJob.Open qControlJob, connstr
	if not rsControlJob.eof then
	   txtControlJob=rsControlJob("ControlJob")
	else
	   Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error SAC200367 Control Job For General Charges is Empty"
	   txtControlJob=""
	end if
	rsControlJob.close
	set rsControlJob = Nothing

	response.Flush()

	if txtControlJob<>"" then
		'---- Normalisasi
		Response.Write " Normalisasi General Charges Base Costing" & "<br>"
		Response.Flush()
		
		connstr.BeginTrans
		qNormalisasi = " delete" & _
		               " from gccb" & _
		               " where fcba = '" & Session("fcBA") & "'" & _
		               "       and to_number(to_char(transdate,'yyyy')) = " & Year(Session("v_General_Charges")) & "" & _
                       "       and to_number(to_char(transdate,'mm')) = " & Month(Session("v_General_Charges")) & ""
		connstr.execute(qNormalisasi)
		
		If Err.Number = 0 Then
			connstr.CommitTrans
		Else
			connstr.RollbackTrans
			Response.write("Error Save ["&qNormalisasi&"] ("&Err.Description&")")
			Response.Flush
		End If
		
		'---- Normalisasi

		' qNormalisasiDetail = " delete from cost_activity_hist_tab " & _
						 ' " where " & _
						 ' " lineheaderno || fcba in (select lineno||fcba from cost_activity " & _
								' " where  to_number(to_Char(TDATE,'mm'))=" & Month(Session("v_General_Charges")) &  _
								' " and to_number(to_Char(TDATE,'yyyy'))=" & Year(Session("v_General_Charges")) & " and fcBA='" & Session("fcBA") & "' and SOURCETYPE='GC')"
						 
		' connstr.execute(qNormalisasiDetail)			
		' connStr.execute("Commit")
		
		
		' qNormalisasi=" Delete From COST_ACTIVITY where  to_number(to_Char(TDATE,'mm'))=" & Month(Session("v_General_Charges")) &  _
					 ' " and to_number(to_Char(TDATE,'yyyy'))=" & Year(Session("v_General_Charges")) & " and fcBA='" & Session("fcBA") & "'" & _
					 ' " and SOURCETYPE = 'GC'"
		' connstr.execute(qNormalisasi)	
		' connStr.execute("Commit")
		
		connstr.BeginTrans
		qNormalisasiDetail = " update cost_activity_hist_tab set deleteflag = 'REMOVE'" & _
						 " where " & _
						 " lineheaderno || fcba in (select lineno||fcba from cost_activity " & _
								" where  to_number(to_Char(TDATE,'mm'))=" & Month(Session("v_General_Charges")) &  _
								" and to_number(to_Char(TDATE,'yyyy'))=" & Year(Session("v_General_Charges")) & " and fcBA='" & Session("fcBA") & "' and SOURCETYPE='GC')"
						 
		connstr.execute(qNormalisasiDetail)
		
		If Err.Number = 0 Then
			connstr.CommitTrans
		Else
			connstr.RollbackTrans
			Response.write("Error Save ["&qNormalisasiDetail&"] ("&Err.Description&")")
			Response.Flush
		End If
		
		connstr.BeginTrans
		qNormalisasiDetail = " delete from cost_activity_hist_tab where deleteflag = 'REMOVE' and fcba =  '" & Session("fcBA") & "'"

		connstr.execute(qNormalisasiDetail)
		
		If Err.Number = 0 Then
			connstr.CommitTrans
		Else
			connstr.RollbackTrans
			Response.write("Error Save ["&qNormalisasiDetail&"] ("&Err.Description&")")
			Response.Flush
		End If
		
		connstr.BeginTrans
		qNormalisasiDetail="   update Cost_Activity set deleteflag = 'REMOVE' where  to_number(to_Char(TDATE,'mm'))=" & Month(Session("v_General_Charges")) &  _
					 " and to_number(to_Char(TDATE,'yyyy'))=" & Year(Session("v_General_Charges")) & " and fcBA='" & Session("fcBA") & "'" & _
					 " and SOURCETYPE = 'GC'"
		connstr.execute(qNormalisasiDetail)
		
		If Err.Number = 0 Then
			connstr.CommitTrans
		Else
			connstr.RollbackTrans
			Response.write("Error Save ["&qNormalisasiDetail&"] ("&Err.Description&")")
			Response.Flush
		End If
		
		connstr.BeginTrans
		qNormalisasiDetail = " delete cost_activity where deleteflag = 'REMOVE' and fcba =  '" & Session("fcBA") & "'"

		connstr.execute(qNormalisasiDetail)
		
		If Err.Number = 0 Then
			connstr.CommitTrans
		Else
			connstr.RollbackTrans
			Response.write("Error Save ["&qNormalisasiDetail&"] ("&Err.Description&")")
			Response.Flush
		End If
		
        Response.write("<font face=Verdana size=1 color=Green>Normalisasi General Charges Base Costing, Success!.<br/><br/></font>")
        Response.Flush
		
		Response.Write "Generating Data..." & "<br/>"
		Response.Flush
		

		connstr.BeginTrans
		ExecuteSQL = " delete from costtotal_pertarget" & _
		             " where fcba = '" & Session("fcBA") & "'" & _
					 "       and year = " & Year(Session("v_General_Charges")) & "" & _
					 "       and month = " & Month(Session("v_General_Charges")) & "" & _
					 "       and targettype = 'GC'"
					 
		'Response.Write ExecuteSQL
		'Response.Flush
		connstr.execute(ExecuteSQL)
		
		If Err.Number = 0 Then
			connstr.CommitTrans
		Else
			connstr.RollbackTrans
			Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
			Response.Flush
		End If
		
		connstr.BeginTrans
		ExecuteSQL = " delete from btl_allocation" & _
		             " where fcba = '" & Session("fcBA") & "'"
					 
		'Response.Write ExecuteSQL
		'Response.Flush
		connstr.execute(ExecuteSQL)
		
		If Err.Number = 0 Then
			connstr.CommitTrans
		Else
			connstr.RollbackTrans
			Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
			Response.Flush
		End If
		
		connstr.BeginTrans
		ExecuteSQL = " insert into costtotal_pertarget" & _
		             " (FCBA, YEAR, MONTH, TARGETTYPE, TARGETCODE, COSTTOTAL, BASIS)" & _
					 " select fcba" & _
					 "        , year" & _
					 "        , month" & _
					 "        , targettype" & _
					 "        , targetcode" & _
					 "        , costtotal" & _
					 "        , GC_CLOSING_API.GET_BASIS(fcba, year, month) basis" & _
					 " from ( " & _
					 "     select fcba" & _
					 "            , to_number(to_char(tdate,'yyyy')) year" & _
					 "            , to_number(to_char(tdate,'mm')) month" & _
					 "            , targettype" & _
					 "            , targetcode" & _
					 "            , nvl(sum(amount),0) as costtotal" & _
					 "     from cost_activity" & _
					 "     where fcba = '" & Session("fcBA") & "'" & _
					 "           and to_number(to_char(tdate,'yyyy')) = " & Year(Session("v_General_Charges")) & "" & _
					 "           and to_number(to_char(tdate,'mm')) = " & Month(Session("v_General_Charges")) & "" & _
					 "           and targettype = 'GC'" & _
					 "     group by fcba" & _
					 "              , to_number(to_char(tdate,'yyyy'))" & _
					 "              , to_number(to_char(tdate,'mm'))" & _
					 "              , targettype" & _
					 "              , targetcode" & _
					 " )"

		connstr.execute(ExecuteSQL)
		
		If Err.Number = 0 Then
			connstr.CommitTrans
		Else
			connstr.RollbackTrans
			Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
			Response.Flush
		End If
		
		if Session("Estate_Mill") = "E" then
			ExecuteSQL = " insert into btl_allocation" & _
                         " (fcba, jobcode, targettype, targetcode, unit, nfactor, basis)" & _
                         " select fcba" & _
                         "        , decode(status, 'TM', '6509801', 'TBM', '1328901', '1140800') jobcode" & _
                         "        , 'FF' targettype" & _
                         "        , targetcode" & _
                         "        , unit" & _
						 "        , 1 nfactor" & _
						 "        , 0 basis" & _
                         " from (" & _
                         "     select fcba" & _
                         "            , fccode targetcode" & _
                         "            , field_api.getvalid_status (fccode, fcba, to_date(" & Year(Session("v_General_Charges")) & "||lpad(" & Month(Session("v_General_Charges")) & ",2,'0')||'01','yyyymmdd')) status" & _
                         "            , nvl(field_api.getvalid_ha (fccode, fcba, to_date(" & Year(Session("v_General_Charges")) & "||lpad(" & Month(Session("v_General_Charges")) & ",2,'0')||'01','yyyymmdd')), 0) unit" & _
                         "     from field" & _
                         "     where fcba = '" & Session("fcBA") & "'" & _
                         "           and nvl(field_api.getvalid_activation (fccode, fcba, to_date(" & Year(Session("v_General_Charges")) & "||lpad(" & Month(Session("v_General_Charges")) & ",2,'0')||'01','yyyymmdd')),'N') = 'Y'" & _
                         "           and field_api.getvalid_status (fccode, fcba, to_date(" & Year(Session("v_General_Charges")) & "||lpad(" & Month(Session("v_General_Charges")) & ",2,'0')||'01','yyyymmdd')) is not null" & _
                         " )" & _
                         " where unit <> 0"
		Elseif (Session("Estate_Mill") = "M" Or Session("Estate_Mill") = "P") then
			ExecuteSQL = " insert into btl_allocation" & _
                         " (fcba, jobcode, targettype, targetcode, unit, nfactor, basis)" & _
                         " select fcba" & _
                         "        , '6610600' jobcode" & _
                         "        , 'MA' targettype" & _
                         "        , targetcode" & _
                         "        , unit" & _
						 "        , 1 nfactor" & _
						 "        , 0 basis" & _
                         " from (" & _
                         "     select fcba" & _
                         "            , fccode targetcode" & _
                         "            , 1 unit" & _
                         "     from machine" & _
                         "     where fcba = '" & Session("fcBA") & "'" & _
                         "           and nvl(activation, 'N') = 'Y'" & _
                         " )" & _
                         " where unit <> 0"
						 
		Elseif Session("Estate_Mill") = "W" then
			ExecuteSQL = " insert into btl_allocation" & _
                         " (fcba, jobcode, targettype, targetcode, unit, nfactor, basis)" & _
                         " select fcba" & _
						 "        , jobcode" & _
						 "        , targettype" & _
						 "        , targetcode" & _
						 "        , unit" & _
						 "        , nfactor" & _
						 "        , basis" & _
						 " from (" & _
						 "     select aa.fcba" & _
						 "          , aa.fccode jobcode" & _
						 "          , 'CB' targettype" & _
						 "          , 'BS' targetcode" & _
						 "          , decode(businessunit_api.get_fctype(aa.codefcba)" & _
						 "                   , 'M', nvl(aa.millcapacity, 0)" & _
						 "                   , 'E', nvl(field_api.GetValid_HaEstate(aa.codefcba,to_date(" & Year(Session("v_General_Charges")) & "||lpad(" & Month(Session("v_General_Charges")) & ",2,'0')||'01','YYYYMMDD')), 0)" & _
						 "                   , 0" & _
						 "                 ) unit" & _
						 "          , nfactor" & _
						 "          , nvl(totalbasis, 0) basis" & _
						 "     from jobcodeintraco aa" & _
						 "     left join ( Select fcba, fctype, Sum(basis) totalbasis" & _
						 "                 from ( select fcba" & _
						 "                               , 'E' fctype" & _
						 "                               , nvl(field_api.GetValid_HaEstate(a.codefcba,to_date(" & Year(Session("v_General_Charges")) & "||lpad(" & Month(Session("v_General_Charges")) & ",2,'0')||'01','YYYYMMDD')),0) basis" & _
						 "                        from jobcodeintraco a " & _
						 "                        where fcba = '" & Session("fcBA") & "'" & _
						 "                        union all" & _
						 "                        select fcba" & _
						 "                               , 'M' fctype" & _
						 "                               , nvl(a.millcapacity,0) basis" & _
						 "                        from jobcodeintraco a" & _
						 "                        where fcba = '" & Session("fcBA") & "'" & _
						 "                      )" & _
						 "                 group by fcba, fctype" & _
						 "               ) bb" & _
						 "     on (aa.fcba = bb.fcba and businessunit_api.get_fctype(aa.codefcba) = bb.fctype)" & _
						 " ) where unit <> 0"

		End If

		connstr.BeginTrans
		connstr.execute(ExecuteSQL)
		
		If Err.Number = 0 Then
			connstr.CommitTrans
		Else
			connstr.RollbackTrans
			Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
			Response.Flush
		End If
		
        Response.write("<font face=Verdana size=1 color=Green>Generating Data, Success!.<br/><br/></font>")
        Response.Flush
		
		Response.Write "Generate Credit Allocation" & "<br/>"
		Response.Flush
		
		if Session("Estate_Mill") = "E" then
			jobcodeNoBasis = "1330300"
		Else
			jobcodeNoBasis = "1140800"
		End If
		
		qCreditAlokasi = " select fcba" & _
                         "        , targettype" & _
                         "        , targetcode" & _
                         "        , costtotal" & _
						 "        , basis" & _
                         " from costtotal_pertarget" & _
                         " where fcba = '" & Session("fcBA") & "'" & _
                         "       and year = " & Year(Session("v_General_Charges")) & "" & _
                         "       and month = " & Month(Session("v_General_Charges")) & "" & _
                         "       and targettype = 'GC'"
					  
		arrRS = GetRSArray(qCreditAlokasi)
		
		If IsArray(arrRS) Then
			For i = LBound(arrRS, 2) To UBound(arrRS, 2)
				If Cdbl(arrRS(4, i)) = 0 Then
					connstr.BeginTrans
					maxLineNoGcCB = maxLineNoGcCB + 1
					qInsertCR = " insert into gccb (fieldcode" & _
								"                   , targettype" & _
								"                   , targetcode" & _
								"                   , jobcode" & _
								"                   , quantity" & _
								"                   , amount" & _
								"                   , transdate" & _
								"                   , referenceno" & _
								"                   , fcentry" & _
								"                   , fcedit" & _
								"                   , fcip" & _
								"                   , fcba" & _
								"                   , lastupdate" & _
								"                   , lasttime" & _
								"                   , lineno)" & _
								" values ('" & arrRS(2, i) & "'" & _
								"         , 'CB'" & _
								"         , 'BS'" & _
								"         , '" & txtControlJob & "'" & _
								"         , '0'" & _
								"         , " & cdbl(arrRS(3, i)) * -1 & "" & _
								"         , " & " to_date('" & Session("v_General_Charges") & "','mm/dd/yyyy')" & _
								"         , '" & arrRS(2, i) & "-Cr-No Basis" & "'" & _
								"         , '" & Session("CodeName") & "'" & _
								"         , '" & Session("CodeName") & "'" & _
								"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
								"         , '" & Session("fcBA") & "'" & _
								"         , sysdate" & _
								"         , to_char(sysdate, 'HH24:MI')" & _
								"         , " & maxLineNoGcCB & ")"
					 'response.Write(qInsertCR)			  
					connStr.execute(qInsertCR)
					
					maxLineNoGcCB = maxLineNoGcCB + 1
					qInsertDB = " insert into gccb (fieldcode" & _
								"                   , targettype" & _
								"                   , targetcode" & _
								"                   , jobcode" & _
								"                   , quantity" & _
								"                   , amount" & _
								"                   , transdate" & _
								"                   , referenceno" & _
								"                   , fcentry" & _
								"                   , fcedit" & _
								"                   , fcip" & _
								"                   , fcba" & _
								"                   , lastupdate" & _
								"                   , lasttime" & _
								"                   , lineno)" & _
								" values ('" & arrRS(2, i) & "'" & _
								"         , 'CB'" & _
								"         , 'BS'" & _
								"         , '" & jobcodeNoBasis & "'" & _
								"         , '0'" & _
								"         , " & cdbl(arrRS(3, i)) & "" & _
								"         , " & " to_date('" & Session("v_General_Charges") & "','mm/dd/yyyy')" & _
								"         , '" & arrRS(2, i) & "-Db-No Basis" & "'" & _
								"         , '" & Session("CodeName") & "'" & _
								"         , '" & Session("CodeName") & "'" & _
								"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
								"         , '" & Session("fcBA") & "'" & _
								"         , sysdate" & _
								"         , to_char(sysdate, 'HH24:MI')" & _
								"         , " & maxLineNoGcCB & ")"
					 'response.Write(qInsertDB)			  
					connStr.execute(qInsertDB)

					If Err.Number = 0 Then
						connstr.CommitTrans
						Response.Flush
					Else
						connstr.RollbackTrans
						Response.write("Error Save ["&qInsertCR&"] - ["&qInsertDB&"] ("&Err.Description&")")
						Response.Flush
					End If
				Else
					connstr.BeginTrans
					maxLineNoGcCB = maxLineNoGcCB + 1
					qInsertCR = " insert into gccb (fieldcode" & _
								"                   , targettype" & _
								"                   , targetcode" & _
								"                   , jobcode" & _
								"                   , quantity" & _
								"                   , amount" & _
								"                   , transdate" & _
								"                   , referenceno" & _
								"                   , fcentry" & _
								"                   , fcedit" & _
								"                   , fcip" & _
								"                   , fcba" & _
								"                   , lastupdate" & _
								"                   , lasttime" & _
								"                   , lineno)" & _
								" values ('" & arrRS(2, i) & "'" & _
								"         , 'CB'" & _
								"         , 'BS'" & _
								"         , '" & txtControlJob & "'" & _
								"         , '0'" & _
								"         , " & cdbl(arrRS(3, i)) * -1 & "" & _
								"         , " & " to_date('" & Session("v_General_Charges") & "','mm/dd/yyyy')" & _
								"         , '" & arrRS(2, i) & "-Cr" & "'" & _
								"         , '" & Session("CodeName") & "'" & _
								"         , '" & Session("CodeName") & "'" & _
								"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
								"         , '" & Session("fcBA") & "'" & _
								"         , sysdate" & _
								"         , to_char(sysdate, 'HH24:MI')" & _
								"         , " & maxLineNoGcCB & ")"
					 'response.Write(qInsertCR)			  
					connStr.execute(qInsertCR)

					If Err.Number = 0 Then
						connstr.CommitTrans
						Response.Flush
					Else
						connstr.RollbackTrans
						Response.write("Error Save ["&qInsertCR&"] ("&Err.Description&")")
						Response.Flush
					End If
				End If
			Next
			Erase arrRS
		End If
		
		Response.write("<font face=Verdana size=1 color=Green> Generate Credit Allocation, Success!.<br/><br/></font>")
		Response.Flush
		
		Response.Write "Generate Debet Allocation" & "<br/>"
		Response.Flush

		if Session("Estate_Mill") = "E" then
			qDebetAlokasi = " select ca.fcba" & _
							" 	     , ca.year" & _
							" 	     , ca.month" & _
							" 	     , ca.targettype sourcetype" & _
							" 	     , ca.targetcode sourcecode" & _
							" 	     , ca.costtotal" & _
							" 	     , ca.basis" & _
							" 	     , btl.targetcode" & _
							" 	     , btl.nfactor" & _
							" 	     , btl.unit" & _
							"        , btl.jobcode" & _
							"        , btl.targettype" & _
							"        , btl.targetcode" & _
							" from costtotal_pertarget ca, btl_allocation btl" & _
							" where ca.fcba = btl.fcba" & _
							"       and ca.fcba = '" & Session("fcBA") & "'" & _
							"       and ca.year = " & Year(Session("v_General_Charges")) & "" & _
							"       and ca.month = " & Month(Session("v_General_Charges")) & "" & _
							"       and ca.targettype = 'GC'"
							
		ElseIf Session("Estate_Mill") = "W" then
			qDebetAlokasi = " select ca.fcba" & _
							"        , ca.year" & _
							"        , ca.month" & _
							"        , ca.targettype sourcetype" & _
							"        , ca.targetcode sourcecode" & _
							"        , ca.costtotal" & _
							"        , btl.basis" & _
							"        , btl.jobcode" & _
							"        , btl.nfactor" & _
							"        , btl.unit" & _
							"        , btl.jobcode" & _
							"        , btl.targettype" & _
							"        , btl.targetcode" & _
							" from costtotal_pertarget ca, btl_allocation btl" & _
							" where ca.fcba = btl.fcba" & _
							"       and ca.fcba = '" & Session("fcBA") & "'" & _
							"       and ca.year = " & Year(Session("v_General_Charges")) & "" & _
							"       and ca.month = " & Month(Session("v_General_Charges")) & "" & _
							"       and ca.targettype = 'GC'"
		ElseIf (Session("Estate_Mill") = "M" Or Session("Estate_Mill") = "P") then
			qDebetAlokasi = " select ca.fcba" & _
							" 	     , ca.year" & _
							" 	     , ca.month" & _
							" 	     , ca.targettype sourcetype" & _
							" 	     , ca.targetcode sourcecode" & _
							" 	     , ca.costtotal" & _
							" 	     , ca.basis" & _
							" 	     , btl.targetcode" & _
							" 	     , btl.nfactor" & _
							" 	     , btl.unit" & _
							"        , btl.jobcode" & _
							"        , btl.targettype" & _
							"        , btl.targetcode" & _
							" from costtotal_pertarget ca, btl_allocation btl" & _
							" where ca.fcba = btl.fcba" & _
							"       and ca.fcba = '" & Session("fcBA") & "'" & _
							"       and ca.year = " & Year(Session("v_General_Charges")) & "" & _
							"       and ca.month = " & Month(Session("v_General_Charges")) & "" & _
							"       and ca.targettype = 'GC'"
		
		Else
			qDebetAlokasi = " "
		End If
		
		arrRS = GetRSArray(qDebetAlokasi)
		
		If IsArray(arrRS) Then
			For i = LBound(arrRS, 2) To UBound(arrRS, 2)
				If Cdbl(arrRS(6, i)) <> 0 Then
					nFFAlokasi = Round(Cdbl(arrRS(9, i))/Cdbl(arrRS(6, i))*Cdbl(arrRS(5, i))*Cdbl(arrRS(8, i)), 0)
					If Cdbl(nFFAlokasi) <> 0 Then
						connstr.BeginTrans
						maxLineNoGcCB = maxLineNoGcCB + 1
						qInsertDB = " insert into gccb (fieldcode" & _
									"                   , targettype" & _
									"                   , targetcode" & _
									"                   , jobcode" & _
									"                   , quantity" & _
									"                   , amount" & _
									"                   , transdate" & _
									"                   , referenceno" & _
									"                   , fcentry" & _
									"                   , fcedit" & _
									"                   , fcip" & _
									"                   , fcba" & _
									"                   , lastupdate" & _
									"                   , lasttime" & _
									"                   , lineno)" & _
									" values ('" & arrRS(4, i) & "'" & _
									"         , '" & arrRS(11, i) & "'" & _
									"         , '" & arrRS(12, i) & "'" & _
									"         , '" & arrRS(10, i) & "'" & _
									"         , " & cdbl(arrRS(9, i)) & "" & _
									"         , " & nFFAlokasi & "" & _
									"         , " & " to_date('" & Session("v_General_Charges") & "','mm/dd/yyyy')" & _
									"         , '" & arrRS(4, i) & "-Db" & "'" & _
									"         , '" & Session("CodeName") & "'" & _
									"         , '" & Session("CodeName") & "'" & _
									"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
									"         , '" & Session("fcBA") & "'" & _
									"         , sysdate" & _
									"         , to_char(sysdate, 'HH24:MI')" & _
									"         , " & maxLineNoGcCB & ")"
						 'response.Write(qInsertDB)			  
						connStr.execute(qInsertDB)

						If Err.Number = 0 Then
							connstr.CommitTrans
							Response.Flush
						Else
							connstr.RollbackTrans
							Response.write("Error Save ["&qInsertDB&"] ("&Err.Description&")")
							Response.Flush
						End If
					End If
				End If
			Next
			Erase arrRS
		End If
		
		Response.write("<font face=Verdana size=1 color=Green> Generate Debet Allocation, Success!.<br/><br/></font>")
		Response.Flush
		
		Response.Write "Normalisasi Rounding" & "<br/>"
		Response.Flush
		
		qSelisih = " select fcba" & _
				   "        , fieldcode" & _
				   "        , transdate" & _
				   "        , amount" & _
				   " from (" & _
				   "     select fcba" & _
				   "            , fieldcode" & _
				   "            , transdate" & _
				   "            , sum(amount) amount" & _
				   "     from gccb" & _
				   "     where fcba = '" & Session("fcBA") & "'" & _
				   "           and to_number(to_char(transdate,'yyyy')) = " & Year(Session("v_General_Charges")) & "" & _
				   "           and to_number(to_char(transdate,'mm')) = " & Month(Session("v_General_Charges")) & "" & _
				   "     group by fcba" & _
				   "              , fieldcode" & _
				   "              , transdate" & _
				   " )" & _
				   " where amount <> 0"
				   
		arrRS = GetRSArray(qSelisih)
		If IsArray(arrRS) Then
			For i = LBound(arrRS, 2) To UBound(arrRS, 2)
				Set objectRS = Server.CreateObject("ADODB.Recordset")
				objectSQL = " select min (lineno) minlineno" & _
							" from gccb" & _
							" where fcba = '" & Session("fcBA") & "'" & _
							"       and transdate = to_date('" & arrRS(2, i) & "', 'mm/dd/yyyy')" & _
							"       and fieldcode = '" & arrRS(1, i) & "'" & _
							"       and jobcode <> '" & txtControlJob & "'"
				objectRS.Open objectSQL, connstr
				If Not objectRS.EOF Then
					FirstRecord = Cdbl(objectRS("minlineno"))
				Else
					FirstRecord = 0.0000001
				End If
				objectRS.close
				set objectRS = Nothing
				
				UpdateAmoutSelisih = " update gccb" & _
									 " set amount = amount - " & cdbl(arrRS(3, i)) & "" & _
									 " where lineno = '" & FirstRecord & "'" & _
									 "       and fcba = '" & Session("fcBA") & "'"
				connStr.execute(UpdateAmoutSelisih)
			Next
			Erase arrRS
		End If
		
		Response.write("<font face=Verdana size=1 color=Green> Normalisasi Rounding, Success!.<br/><br/></font>")
		Response.Flush
		
		Set objectRS = Server.CreateObject("ADODB.Recordset")
		objectSQL = " select nvl(sum(costtotal), 0) costtotal" & _
					"        , nvl(sum(alloc_db), 0) alloc_db" & _
					"        , nvl(sum(alloc_cr), 0)*-1 alloc_cr" & _
					" from (" & _
					"     select sum(costtotal) costtotal, 0 alloc_db , 0 alloc_cr" & _
					"     from costtotal_pertarget" & _
					"     where fcba = '" & Session("fcBA") & "'" & _
					"           and year = " & Year(Session("v_General_Charges")) & "" & _
					"           and month = " & Month(Session("v_General_Charges")) & "" & _
					"           and targettype = 'GC'" & _
					"     union all" & _
					"     select 0 costtotal, sum(amount) alloc_db , 0 alloc_cr" & _
					"     from gccb" & _
					"     where fcba = '" & Session("fcBA") & "'" & _
					"           and to_number(to_char(transdate, 'yyyy')) = " & Year(Session("v_General_Charges")) & "" & _
					"           and to_number(to_char(transdate, 'mm')) = " & Month(Session("v_General_Charges")) & "" & _
					"           and jobcode <> '" & txtControlJob & "'" & _
					"     union all" & _
					"     select 0 costtotal,0 alloc_db ,  sum(amount) alloc_cr" & _
					"     from gccb" & _
					"     where fcba = '" & Session("fcBA") & "'" & _
					"           and to_number(to_char(transdate, 'yyyy')) = " & Year(Session("v_General_Charges")) & "" & _
					"           and to_number(to_char(transdate, 'mm')) = " & Month(Session("v_General_Charges")) & "" & _
					"           and jobcode = '" & txtControlJob & "'" & _
					" )"

			objectRS.Open objectSQL, connstr
			If Not objectRS.EOF Then
				Cost_In = Cdbl(objectRS("costtotal"))
				Alloc_Db = Cdbl(objectRS("alloc_db"))
				Alloc_Cr = Cdbl(objectRS("alloc_cr"))
			Else
			End If
			objectRS.close
			set objectRS = Nothing

			If Cost_In <> Alloc_Db Then
				err.number = 1001
				err.description = "Alokasi sisi debet tidak sesuai dengan Biaya yang Masuk! Cek Kembali Biaya yang masuk ke target GC"
			ElseIf Cost_In <> Alloc_Cr Then
				err.number = 1001
				err.description = "Alokasi sisi credit tidak sesuai dengan Biaya yang Masuk! Cek Kembali Biaya yang masuk ke target GC"
			ElseIf (Cost_In = Alloc_Db) And (Cost_In = Alloc_Cr) Then
				'----- Tranfer to Activity Costing
				response.Write "Transfer to Activity Based Costing..." & "<br>"
				Response.Flush

				' Get the Max LineNo of Cost Activity
				Set rsCostActivityMax = Server.CreateObject("ADODB.Recordset")
				rsCostActivityMax.CursorType = 3
				rsCostActivityMax.CursorLocation = 3
												
				qCostActivityMax =  " select nvl(max(lineno), 0) maxLineNo from cost_activity " & _
									" where fcba = '" &  Session("fcBA") & "' " 
				rsCostActivityMax.Open qCostActivityMax, connstr
				
				If Not rsCostActivityMax.EOF Then
					maxLineNo = cdbl(rsCostActivityMax("maxLineNo"))
				Else
					maxLineNo = 0
				End If

				rsCostActivityMax.close
				set rsCostActivityMax = Nothing
				' End of Get the Max LineNo of Cost Activity
				
				qVoucher = " select transdate tdate" & _
						   "        , jobcode" & _
						   "        , 'GC' source" & _
						   "        , 'GC' sourcetype" & _
						   "        , fieldcode sourcecode" & _
						   "        , '' sourcecode_status" & _
						   "        , targettype" & _
						   "        , targetcode" & _
						   "        , '' targetcode_status" & _
						   "        , referenceno" & _
						   "        , sum(quantity) quantity" & _
						   "        , 0 output" & _
						   "        , sum(amount) amount" & _
						   "        , fcba" & _
						   "        , to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') closing_date" & _
						   " from gccb" & _
						   " where fcba = '" & Session("fcBA") & "'" & _
						   "       and to_number(to_char(transdate,'yyyy')) = " & Year(Session("v_General_Charges")) & "" & _
						   "       and to_number(to_char(transdate,'mm')) = " & Month(Session("v_General_Charges")) & "" & _
						   " group by transdate" & _
						   "          , jobcode" & _
						   "          , 'GC'" & _
						   "          , 'GC'" & _
						   "          , fieldcode" & _
						   "          , ''" & _
						   "          , targettype" & _
						   "          , targetcode" & _
						   "          , ''" & _
						   "          , referenceno" & _
						   "          , fcba"

				arrRS = GetRSArray(qVoucher)
				
				If IsArray(arrRS) Then
					For i = LBound(arrRS, 2) To UBound(arrRS, 2)
					
						if cdbl(arrRS(12, i)) >= 0 then 
							temp_DebetCredit = "Dr"
						else
							temp_DebetCredit = "Cr"
						end if
						
						connstr.BeginTrans
						maxLineNo = maxLineNo + 1
						qCost_Activity="Insert into Cost_Activity " & _
											" ( TDATE, JOBCODE, SOURCE, SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
											"   TARGETTYPE, TARGETCODE, TARGETCODE_STATUS, REFERENCENO, " & _
											"   QUANTITY, OUTPUT, AMOUNT, FCBA, CLOSING_DATE, PROCESS_DATE," & _
											" TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
											" TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
											" TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
											" TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
											" TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _	
											" ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
											" LASTUPDATE, LASTTIME, LINENO ) " & _
											" values ( to_date('" & arrRS(0, i) & "','mm/dd/yyyy') " & ",'" & arrRS(1, i)  & _ 
											"','" & arrRS(2, i) & "','" & arrRS(3, i) & "','" & arrRS(4, i) & "'" & _
											",'" & arrRS(5, i) & "','" & arrRS(6, i) & "','" & arrRS(7, i) & "'" & _
											",'" & arrRS(8, i) & "','" & arrRS(9, i) & "'," & _
											cdbl(arrRS(10, i)) & "," & _
											cdbl(arrRS(11, i)) & "," & _
											cdbl(arrRS(12, i)) & ",'" & _
											Session("fcBA") & "'," & _
											" to_date('" & arrRS(14, i) & "','mm/dd/yyyy') , sysdate, " & _
											" Cost_Activity_Info_Api.Get_Title_Reference_01 ('" & arrRS(3, i) & "','" & arrRS(6, i) & "','" & temp_DebetCredit & "'), '" & arrRS(7, i) & "', " & _
											" Cost_Activity_Info_Api.Get_Title_Reference_02 ('" & arrRS(3, i) & "','" & arrRS(6, i) & "','" & temp_DebetCredit & "'), '" & arrRS(7, i) & "', " & _
											" '','', " & _
											" '','', " & _
											" '','','" & temp_DebetCredit & "', " & _
											Month(Session("v_General_Charges")) & ", " & _
											Year(Session("v_General_Charges")) & ", " & _
											" sysdate, to_char(sysdate,'HH24:MI'), " & maxLineNo & " ) "
						'response.write 	qCost_Activity				
						connstr.execute(qCost_Activity)
						
						If Err.Number = 0 Then
							connstr.CommitTrans
							Response.Flush
						Else
							connstr.RollbackTrans
							Response.write("Error Save ["&qCost_Activity&"] ("&Err.Description&")")
							Response.Flush
						End If
					Next
					Erase arrRS
				End If
				
				Response.write("<font face=Verdana size=1 color=Green>Transfer to Activity Based Costing..., Success!.<br/><br/></font>")
				Response.Flush
			End If
		
	End If

	%>
	<!--#include virtual="/p_LOC_Closing.Asp"-->

	<%if err.number <> 0 then%>
		<br><img src="images/error.gif">
		<%
		
		connstr.BeginTrans
		ExecuteSQL = " delete from costtotal_pertarget" & _
		             " where fcba = '" & Session("fcBA") & "'" & _
					 "       and year = " & Year(Session("v_General_Charges")) & "" & _
					 "       and month = " & Month(Session("v_General_Charges")) & "" & _
					 "       and targettype = 'GC'"
					 
		'Response.Write ExecuteSQL
		'Response.Flush
		connstr.execute(ExecuteSQL)
		
		If Err.Number = 0 Then
			connstr.CommitTrans
		Else
			connstr.RollbackTrans
			Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
			Response.Flush
		End If
		
		connstr.BeginTrans
		ExecuteSQL = " delete from btl_allocation" & _
		             " where fcba = '" & Session("fcBA") & "'"
					 
		'Response.Write ExecuteSQL
		'Response.Flush
		connstr.execute(ExecuteSQL)
		
		If Err.Number = 0 Then
			connstr.CommitTrans
		Else
			connstr.RollbackTrans
			Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
			Response.Flush
		End If
		
		connstr.BeginTrans
		ExecuteSQL = " delete from costtotal_pertarget" & _
		             " where fcba = '" & Session("fcBA") & "'" & _
					 "       and year = " & Year(Session("v_General_Charges")) & "" & _
					 "       and month = " & Month(Session("v_General_Charges")) & "" & _
					 "       and targettype = 'LOC'"
					 
		'Response.Write ExecuteSQL
		'Response.Flush
		connstr.execute(ExecuteSQL)
		
		If Err.Number = 0 Then
			connstr.CommitTrans
		Else
			connstr.RollbackTrans
			Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
			Response.Flush
		End If
		
		connstr.BeginTrans
		ExecuteSQL = " delete from loc_allocation" & _
					 " where fcba = '" & Session("fcBA") & "'"
					 
		'Response.Write ExecuteSQL
		'Response.Flush
		connstr.execute(ExecuteSQL)
		
		If Err.Number = 0 Then
			connstr.CommitTrans
		Else
			connstr.RollbackTrans
			Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
			Response.Flush
		End If
		
		connstr.BeginTrans
		ExecuteSQL = " delete from attendance_gad_closing" & _
					 " where fcba = '" & Session("fcBA") & "'"
					 
		'Response.Write ExecuteSQL
		'Response.Flush
		connstr.execute(ExecuteSQL)
		
		If Err.Number = 0 Then
			connstr.CommitTrans
		Else
			connstr.RollbackTrans
			Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
			Response.Flush
		End If
		
		response.Write  "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" & err.description & "</font>"
		%>
		<script language="JavaScript" type="text/javascript" >
		document.getElementById('status_running').style.visibility = 'hidden';
		</script>	
	<%Else	
		connstr.BeginTrans
		ExecuteSQL = " delete from costtotal_pertarget" & _
		             " where fcba = '" & Session("fcBA") & "'" & _
					 "       and year = " & Year(Session("v_General_Charges")) & "" & _
					 "       and month = " & Month(Session("v_General_Charges")) & "" & _
					 "       and targettype = 'GC'"
					 
		'Response.Write ExecuteSQL
		'Response.Flush
		connstr.execute(ExecuteSQL)
		
		If Err.Number = 0 Then
			connstr.CommitTrans
		Else
			connstr.RollbackTrans
			Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
			Response.Flush
		End If
		
		connstr.BeginTrans
		ExecuteSQL = " delete from btl_allocation" & _
		             " where fcba = '" & Session("fcBA") & "'"
					 
		'Response.Write ExecuteSQL
		'Response.Flush
		connstr.execute(ExecuteSQL)
		
		If Err.Number = 0 Then
			connstr.CommitTrans
		Else
			connstr.RollbackTrans
			Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
			Response.Flush
		End If
		
		connstr.BeginTrans
		ExecuteSQL = " delete from costtotal_pertarget" & _
		             " where fcba = '" & Session("fcBA") & "'" & _
					 "       and year = " & Year(Session("v_General_Charges")) & "" & _
					 "       and month = " & Month(Session("v_General_Charges")) & "" & _
					 "       and targettype = 'LOC'"
					 
		'Response.Write ExecuteSQL
		'Response.Flush
		connstr.execute(ExecuteSQL)
		
		If Err.Number = 0 Then
			connstr.CommitTrans
		Else
			connstr.RollbackTrans
			Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
			Response.Flush
		End If
		
		connstr.BeginTrans
		ExecuteSQL = " delete from loc_allocation" & _
					 " where fcba = '" & Session("fcBA") & "'"
					 
		'Response.Write ExecuteSQL
		'Response.Flush
		connstr.execute(ExecuteSQL)
		
		If Err.Number = 0 Then
			connstr.CommitTrans
		Else
			connstr.RollbackTrans
			Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
			Response.Flush
		End If
		
		connstr.BeginTrans
		ExecuteSQL = " delete from attendance_gad_closing" & _
					 " where fcba = '" & Session("fcBA") & "'"
					 
		'Response.Write ExecuteSQL
		'Response.Flush
		connstr.execute(ExecuteSQL)
		
		If Err.Number = 0 Then
			connstr.CommitTrans
		Else
			connstr.RollbackTrans
			Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
			Response.Flush
		End If
		
			' Add By Akbar 18 Desemeber 2015
		
			Dim QAccountPeriod,rsAccountPeriod,temp_NextCloseAccountingPeriod, temp_NextCloseAccountingYear, temp_NextCurrentAccountingPeriod, temp_NextCurrentAccountingYear, temp_EndingAccountingPeriod
				
			QAccountPeriod	=	" select periodcontrol_api.GetNextCloseAccountingPeriod(system,fcba) NextCloseAccountingPeriod ,  " & _ 
								" periodcontrol_api.GetNextCloseAccountingYear(system,fcba) NextCloseAccountingYear ,  " & _ 
								" periodcontrol_api.GetNextCurrentAccountingPeriod(system,fcba) NextCurrentAccountingPeriod ,  " & _ 
								" periodcontrol_api.GetNextCurrentAccountingYear(system,fcba) NextCurrentAccountingYear,  " & _ 
								" periodcontrol_api.GetEndingAccountingPeriodTY(system,fcba)  EndingAccountingPeriod " & _ 
								" from periodcontrol p  " & _ 
								" where UPPER(system) ='GENERAL_CHARGES' and fcba='" & Session("fcBA") & "'"
			Set rsAccountPeriod = Server.CreateObject("ADODB.Recordset")
				rsAccountPeriod.CursorType = 3
				rsAccountPeriod.CursorLocation = 3
				rsAccountPeriod.Open QAccountPeriod, connstr
				if not rsAccountPeriod.eof then
					temp_NextCloseAccountingPeriod =  Cint(rsAccountPeriod("NextCloseAccountingPeriod"))
					temp_NextCloseAccountingYear   =  Cint(rsAccountPeriod("NextCloseAccountingYear"))
					temp_NextCurrentAccountingPeriod = Cint(rsAccountPeriod("NextCurrentAccountingPeriod"))
					temp_NextCurrentAccountingYear = Cint(rsAccountPeriod("NextCurrentAccountingYear"))
					temp_EndingAccountingPeriod = Cint(rsAccountPeriod("EndingAccountingPeriod"))
				else
					Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error, No PERIODCONTROL for GENERAL_CHARGES " & "<br>"
				end if
			rsAccountPeriod.close
			set rsAccountPeriod = Nothing		
	
		
		if Session("Estate_Mill")="E" then
		            
					IF temp_NextCurrentAccountingPeriod < 12   THEN
			
						q= " Update PeriodControl Set LASTPERIODPOSTED=last_day(add_months(LASTPERIODPOSTED,1)),CURRENTDATE=add_months(CURRENTDATE,1) " & _ 
						   " ,CloseAccountingPeriod = " & temp_NextCloseAccountingPeriod & " " & _  
						   " ,CloseAccountingYear = " & temp_NextCloseAccountingYear & " " & _  
						   " ,CurrentAccountingPeriod = " & temp_NextCurrentAccountingPeriod & " " & _  
						   " ,CurrentAccountingYear = " & temp_NextCurrentAccountingYear & " " & _    
						   " ,lastupdate = sysdate, lasttime = to_char(sysdate, 'hh24:mi') " & _
						   " where UPPER(system) IN ('GENERAL_CHARGES','WEIGHBRIDGE','MILL_PRODUCTION','DELIVERY') and fcba='" & Session("fcBA") & "'"
					ELSE
						IF temp_NextCloseAccountingPeriod < 12 THEN
							q=" Update PeriodControl Set LASTPERIODPOSTED=last_day(add_months(LASTPERIODPOSTED,1)) ,CURRENTDATE=add_months(CURRENTDATE,1) " & _ 
							  " ,CloseAccountingPeriod = " & temp_NextCloseAccountingPeriod & " " & _  
							  " ,CloseAccountingYear = " & temp_NextCloseAccountingYear & " " & _  
							  " ,CurrentAccountingPeriod = " & temp_NextCurrentAccountingPeriod & " " & _  
							  " ,CurrentAccountingYear = " & temp_NextCurrentAccountingYear & " " & _
							  " ,lastupdate = sysdate, lasttime = to_char(sysdate, 'hh24:mi') " & _
							  " where UPPER(system) IN ('GENERAL_CHARGES','WEIGHBRIDGE','MILL_PRODUCTION','DELIVERY') and fcba='" & Session("fcBA") & "'"
						ELSE
							 q=" Update PeriodControl Set LASTPERIODPOSTED=last_day(add_months(LASTPERIODPOSTED,1)) " & _ 
							   " ,CloseAccountingPeriod = " & temp_NextCloseAccountingPeriod & " " & _  
						  " ,CloseAccountingYear = " & temp_NextCloseAccountingYear & " " & _  
						  " ,CurrentAccountingPeriod = " & temp_NextCurrentAccountingPeriod & " " & _  
						  " ,CurrentAccountingYear = " & temp_NextCurrentAccountingYear & " " & _     
							  " ,lastupdate = sysdate, lasttime = to_char(sysdate, 'hh24:mi') " & _
							  " where  UPPER(system) IN ('GENERAL_CHARGES','WEIGHBRIDGE','MILL_PRODUCTION','DELIVERY')  and fcba='" & Session("fcBA") & "'"
						END IF
					END IF
					
					IF temp_NextCloseAccountingPeriod = temp_EndingAccountingPeriod AND temp_NextCurrentAccountingPeriod = 1  THEN
					
						q=" Update PeriodControl Set CURRENTDATE=add_months(CURRENTDATE,1) " & _ 
						  " ,CloseAccountingPeriod = " & temp_NextCloseAccountingPeriod & " " & _  
						  " ,CloseAccountingYear = " & temp_NextCloseAccountingYear & " " & _  
						  " ,CurrentAccountingPeriod = " & temp_NextCurrentAccountingPeriod & " " & _  
						  " ,CurrentAccountingYear = " & temp_NextCurrentAccountingYear & " " & _  
						  " ,lastupdate = sysdate, lasttime = to_char(sysdate, 'hh24:mi') " & _
						  " where UPPER(system) IN ('GENERAL_CHARGES','WEIGHBRIDGE','MILL_PRODUCTION','DELIVERY')  and and fcba='" & Session("fcBA") & "'"
					END IF	
		
		
		Else
			
					IF temp_NextCurrentAccountingPeriod < 12   THEN
			
						q= " Update PeriodControl Set LASTPERIODPOSTED=last_day(add_months(LASTPERIODPOSTED,1)),CURRENTDATE=add_months(CURRENTDATE,1) " & _ 
						   " ,CloseAccountingPeriod = " & temp_NextCloseAccountingPeriod & " " & _  
						   " ,CloseAccountingYear = " & temp_NextCloseAccountingYear & " " & _  
						   " ,CurrentAccountingPeriod = " & temp_NextCurrentAccountingPeriod & " " & _  
						   " ,CurrentAccountingYear = " & temp_NextCurrentAccountingYear & " " & _    
						   " ,lastupdate = sysdate, lasttime = to_char(sysdate, 'hh24:mi') " & _
						   " where UPPER(system) IN ('GENERAL_CHARGES','HARVESTING')  and fcba='" & Session("fcBA") & "'"
					ELSE
						IF temp_NextCloseAccountingPeriod < 12 THEN
							q=" Update PeriodControl Set LASTPERIODPOSTED=last_day(add_months(LASTPERIODPOSTED,1)) ,CURRENTDATE=add_months(CURRENTDATE,1) " & _ 
							  " ,CloseAccountingPeriod = " & temp_NextCloseAccountingPeriod & " " & _  
							  " ,CloseAccountingYear = " & temp_NextCloseAccountingYear & " " & _  
							  " ,CurrentAccountingPeriod = " & temp_NextCurrentAccountingPeriod & " " & _  
							  " ,CurrentAccountingYear = " & temp_NextCurrentAccountingYear & " " & _
							  " ,lastupdate = sysdate, lasttime = to_char(sysdate, 'hh24:mi') " & _
							  " where UPPER(system) IN ('GENERAL_CHARGES','HARVESTING') and fcba='" & Session("fcBA") & "'"
						ELSE
							 q=" Update PeriodControl Set LASTPERIODPOSTED=last_day(add_months(LASTPERIODPOSTED,1)) " & _ 
							   " ,CloseAccountingPeriod = " & temp_NextCloseAccountingPeriod & " " & _  
						  " ,CloseAccountingYear = " & temp_NextCloseAccountingYear & " " & _  
						  " ,CurrentAccountingPeriod = " & temp_NextCurrentAccountingPeriod & " " & _  
						  " ,CurrentAccountingYear = " & temp_NextCurrentAccountingYear & " " & _     
							  " ,lastupdate = sysdate, lasttime = to_char(sysdate, 'hh24:mi') " & _
							  " where  UPPER(system) IN ('GENERAL_CHARGES','HARVESTING')   and fcba='" & Session("fcBA") & "'"
						END IF
					END IF
					
					IF temp_NextCloseAccountingPeriod = temp_EndingAccountingPeriod AND temp_NextCurrentAccountingPeriod = 1  THEN
					
						q=" Update PeriodControl Set CURRENTDATE=add_months(CURRENTDATE,1) " & _ 
						  " ,CloseAccountingPeriod = " & temp_NextCloseAccountingPeriod & " " & _  
						  " ,CloseAccountingYear = " & temp_NextCloseAccountingYear & " " & _  
						  " ,CurrentAccountingPeriod = " & temp_NextCurrentAccountingPeriod & " " & _  
						  " ,CurrentAccountingYear = " & temp_NextCurrentAccountingYear & " " & _  
						  " ,lastupdate = sysdate, lasttime = to_char(sysdate, 'hh24:mi') " & _
						  " where UPPER(system) IN ('GENERAL_CHARGES','HARVESTING')  and and fcba='" & Session("fcBA") & "'"
					END IF	
					
				
	
		End if
		
		connStr.execute(q)
		TimeAkhir = Time()
		Lama = TimeAkhir - TimeAwal
		Response.Write "<br>" & "Finish ... " & time() & "<br><br>"
		Response.Write " Time Elapsed " & Hour(Lama) & " Hours " & Minute(Lama) & " Minutes " &  Second(Lama) & " Seconds."	
		
		'--Added by Ukie at 25-Oct-2012--	
		Endingtime = FormatDateTime(CDate(Now))

		qLineMax = "select nvl(max(lineno),0) tempMaxLine from CLOSINGLOG where fcba = '" & session("fcBA") & "' AND SYSTEM = 'General_Charges'" 
									
			set rsqLineMax = Server.CreateObject("ADODB.Recordset")
			rsqLineMax.Open qLineMax, connstr
			if not rsqLineMax.eof then
				tempMaxLine = cdbl(rsqLineMax("tempMaxLine"))
			else
				tempMaxLine = 0
			end if 

			tempMaxLine = tempMaxLine + 1

		strClosingLog = " INSERT INTO CLOSINGLOG(SYSTEM, FCBA, STARTINGTIME, ENDINGTIME, LINENO, lastupdate, lasttime,fcentry) VALUES (" & _
						" 'General_Charges', " & _
						" '" & Session("fcBA") & "'," & _
						" TO_DATE('" & Startingtime & "', 'mm/dd/yyyy hh:mi:ssam')," & _
						" TO_DATE('" & Endingtime & "', 'mm/dd/yyyy hh:mi:ssam')," & _
						" " & tempMaxLine & "" & _
						" , sysdate" & _
						" , to_char(sysdate, 'hh24:mi')" & _
						" ,'" & Session("CodeName") & "'" & _ 
						" )"
		connStr.execute(strClosingLog)

		strLenghtime = 	" SELECT" & _
						" FCBA," & _
						" LINENO," & _
						" SYSTEM," & _
						" TO_CHAR(STARTINGTIME, 'DD.MM.YYYY:HH24:MI:SS') STARTINGTIME," & _
						" TO_CHAR(ENDINGTIME, 'DD.MM.YYYY:HH24:MI:SS') ENDINGTIME," & _
						" TRUNC((((86400*(ENDINGTIME-STARTINGTIME))/60)/60)/24) DAYS," & _
						" TRUNC(((86400*(ENDINGTIME-STARTINGTIME))/60)/60)-24*(TRUNC((((86400*(ENDINGTIME-STARTINGTIME))/60)/60)/24)) HRS," & _
						" TRUNC((86400*(ENDINGTIME-STARTINGTIME))/60)-60*(TRUNC(((86400*(ENDINGTIME-STARTINGTIME))/60)/60)) MIN," & _
						" TRUNC(86400*(ENDINGTIME-STARTINGTIME))-60*(TRUNC((86400*(ENDINGTIME-STARTINGTIME))/60)) SEC" & _
						" FROM CLOSINGLOG" & _
						" WHERE FCBA = '" & Session("fcBA") & "' AND LINENO = " & tempMaxLine & " AND SYSTEM = 'General_Charges'"
					
		Set rsLenghtime = Server.CreateObject("ADODB.Recordset")
		rsLenghtime.CursorType = 3
		rsLenghtime.CursorLocation = 3
		rsLenghtime.Open strLenghtime, connstr

		upClosingLog = " UPDATE CLOSINGLOG SET" & _
						" LENGTHTIME = '" & rsLenghtime(5) & " Days " & rsLenghtime(6) & " Hours " & rsLenghtime(7) & " Minutes " & rsLenghtime(8) & " Seconds " & "'" & _
						" WHERE FCBA = '" & Session("fcBA") & "' AND LINENO = " & tempMaxLine & " AND SYSTEM = 'General_Charges'"		
		connStr.execute(upClosingLog)
		'--End Added--
		
		%>
		<br>
		<img src="images/ok.gif"><font color=Blue pointsize=9 face=Verdana>General Charges Closing is succesfully...
		<script language="JavaScript" type="text/javascript" >
		document.getElementById('status_running').style.visibility = 'hidden';
		</script>
		
		<%
		
			txtLastDay  =cmdFindPeriod("General_Charges")
			txtNextMonth=cmdPeriodClose("General_Charges")		
			
			Session("v_General_Charges")=txtLastDay
			Session("cls_General_Charges")= txtNextMonth	
		%>
	<%End if%>
<%
else 
%>
	<script language="JavaScript" type="text/javascript" >
	document.getElementById('status_running').style.visibility = 'hidden';
	</script>
	<br><br><img src="images/error.gif">General Charges Closing is NOT completed...
	<%
    tempErrorWaterFlow = "Please Check Previous Closing Process, its still not DONE yet"
    response.Write  "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" & tempErrorWaterFlow & "</font>" %>
	<script language="JavaScript" type="text/javascript" >
	document.getElementById('status_running').style.visibility = 'hidden';
	</script>	
	<img src="images/error.gif" />
<%
end if
%> 
</body>
</html>



<%
	Response.write "********** LOC Closing ***********<br/><br/>"
	Response.Flush
%>
<%
Dim first_fieldcode, first_transdate, first_jobcode, first_targettype, first_targetcode
Dim temp_referenceno

'---- Control Job
qControlJob="select ControlJob from Module where NAME='LOC'"
Set rsControlJob = Server.CreateObject("ADODB.Recordset")
rsControlJob.CursorType = 3
rsControlJob.CursorLocation = 3
rsControlJob.Open qControlJob, connstr
if not rsControlJob.eof then
   txtControlJob=rsControlJob("ControlJob")
else
   Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error SAC200368 Control Job For LOC is Empty"
   txtControlJob=""
end if
rsControlJob.close
set rsControlJob = Nothing

response.Flush()

if txtControlJob<>"" then
	'---- Normalisasi
	Response.Write " Normalisasi LOC Base Costing" & "<br>"
	Response.Flush()

	connstr.BeginTrans
	qNormalisasi = " delete" & _
				   " from loccb" & _
				   " where fcba = '" & Session("fcBA") & "'" & _
				   "       and to_number(to_char(transdate,'yyyy')) = " & Year(Session("v_General_Charges")) & "" & _
				   "       and to_number(to_char(transdate,'mm')) = " & Month(Session("v_General_Charges")) & ""
	connstr.execute(qNormalisasi)
	
	If Err.Number = 0 Then
		connstr.CommitTrans
	Else
		connstr.RollbackTrans
		Response.write("Error Save ["&qNormalisasi&"] ("&Err.Description&")")
		Response.Flush
	End If
	
		connstr.BeginTrans
		qNormalisasiDetail = " update cost_activity_hist_tab set deleteflag = 'REMOVE'" & _
						 " where " & _
						 " lineheaderno || fcba in (select lineno||fcba from cost_activity " & _
								" where  to_number(to_Char(TDATE,'mm'))=" & Month(Session("v_General_Charges")) &  _
								" and to_number(to_Char(TDATE,'yyyy'))=" & Year(Session("v_General_Charges")) & " and fcBA='" & Session("fcBA") & "' and SOURCETYPE='LOC')"
						 
		connstr.execute(qNormalisasiDetail)
		
		If Err.Number = 0 Then
			connstr.CommitTrans
		Else
			connstr.RollbackTrans
			Response.write("Error Save ["&qNormalisasiDetail&"] ("&Err.Description&")")
			Response.Flush
		End If
		
		connstr.BeginTrans
		qNormalisasiDetail = " delete from cost_activity_hist_tab where deleteflag = 'REMOVE' and fcba =  '" & Session("fcBA") & "'"

		connstr.execute(qNormalisasiDetail)
		
		If Err.Number = 0 Then
			connstr.CommitTrans
		Else
			connstr.RollbackTrans
			Response.write("Error Save ["&qNormalisasiDetail&"] ("&Err.Description&")")
			Response.Flush
		End If

	connstr.BeginTrans
    qNormalisasi=" UPDATE COST_ACTIVITY set deleteflag = 'REMOVE' where  to_number(to_Char(TDATE,'mm'))=" & Month(Session("v_General_Charges")) &  _
				 " and to_number(to_Char(TDATE,'yyyy'))=" & Year(Session("v_General_Charges")) & " and fcBA='" & Session("fcBA") & "'" & _
				 " and SOURCETYPE = 'LOC'"
				 
	connstr.execute(qNormalisasi)
	
	If Err.Number = 0 Then
		connstr.CommitTrans
	Else
		connstr.RollbackTrans
		Response.write("Error Save ["&qNormalisasi&"] ("&Err.Description&")")
		Response.Flush
	End If

	connstr.BeginTrans
	qNormalisasiDetail = " delete cost_activity where deleteflag = 'REMOVE' and fcba =  '" & Session("fcBA") & "'"

	connstr.execute(qNormalisasiDetail)
	
	If Err.Number = 0 Then
		connstr.CommitTrans
	Else
		connstr.RollbackTrans
		Response.write("Error Save ["&qNormalisasiDetail&"] ("&Err.Description&")")
		Response.Flush
	End If
	
	Response.write("<font face=Verdana size=1 color=Green>Normalisasi LOC Base Costing, Success!.<br/><br/></font>")
	Response.Flush
	
	Response.Write "Generating Data..." & "<br/>"
	Response.Flush
	
	connstr.BeginTrans
	ExecuteSQL = " delete from costtotal_pertarget" & _
				 " where fcba = '" & Session("fcBA") & "'" & _
				 "       and year = " & Year(Session("v_General_Charges")) & "" & _
				 "       and month = " & Month(Session("v_General_Charges")) & "" & _
				 "       and targettype in ('LOC', 'SOC')"
				 
	'Response.Write ExecuteSQL
	'Response.Flush
	connstr.execute(ExecuteSQL)
	
	If Err.Number = 0 Then
		connstr.CommitTrans
	Else
		connstr.RollbackTrans
		Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
		Response.Flush
	End If
	
	connstr.BeginTrans
	ExecuteSQL = " delete from loc_allocation" & _
				 " where fcba = '" & Session("fcBA") & "'"
				 
	'Response.Write ExecuteSQL
	'Response.Flush
	connstr.execute(ExecuteSQL)
	
	If Err.Number = 0 Then
		connstr.CommitTrans
	Else
		connstr.RollbackTrans
		Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
		Response.Flush
	End If
	
	connstr.BeginTrans
	ExecuteSQL = " delete from attendance_gad_closing" & _
				 " where fcba = '" & Session("fcBA") & "'"
				 
	'Response.Write ExecuteSQL
	'Response.Flush
	connstr.execute(ExecuteSQL)
	
	If Err.Number = 0 Then
		connstr.CommitTrans
	Else
		connstr.RollbackTrans
		Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
		Response.Flush
	End If
	
	connstr.BeginTrans
	ExecuteSQL = " delete from allowances_loc_tab" & _
				 " where fcba = '" & Session("fcBA") & "'"
				 
	'Response.Write ExecuteSQL
	'Response.Flush
	connstr.execute(ExecuteSQL)
	
	If Err.Number = 0 Then
		connstr.CommitTrans
	Else
		connstr.RollbackTrans
		Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
		Response.Flush
	End If
	
	connstr.BeginTrans
	If Session("Estate_Mill") = "E" then
		ExecuteSQL = " insert into allowances_loc_tab" & _
                     " select fcba" & _
                     "        , to_number(to_char(tdate, 'yyyy')) year" & _
                     "        , to_number(to_char(tdate, 'mm')) month" & _
                     "        , allowdedcode" & _
                     "        , targettype" & _
                     "        , targetcode" & _
                     "        , sum(amount) amount" & _
                     " from ( select fcba" & _
                     "               , trunc(tdate, 'mm') tdate" & _
                     "               , substr(referenceno, 1, 3) allowdedcode" & _
                     "               , targettype" & _
                     "               , targetcode                  " & _
                     "               , amount" & _
                     "        from cost_activity" & _
                     "        where fcba = '" & Session("fcBA") & "'" & _
                     "              and to_number(to_char(tdate,'yyyy')) = " & Year(Session("v_General_Charges")) & "" & _
                     "              and to_number(to_char(tdate,'mm')) = " & Month(Session("v_General_Charges")) & "" & _
                     "              and sourcetype in ('ES', 'EM')" & _
                     "              and targettype in ('SOC', 'LOC')" & _
                     "        union all" & _
                     "        select fcba" & _
                     "               , trunc(tdate, 'mm') tdate" & _
                     "               , null allowdedcode" & _
                     "               , targettype" & _
                     "               , targetcode           " & _
                     "               , amount" & _
                     "        from cost_activity" & _
                     "        where fcba = '" & Session("fcBA") & "'" & _
                     "              and to_number(to_char(tdate,'yyyy')) = " & Year(Session("v_General_Charges")) & "" & _
                     "              and to_number(to_char(tdate,'mm')) = " & Month(Session("v_General_Charges")) & "" & _
                     "              and sourcetype not in ('ES', 'EM')" & _
                     "              and targettype in ('SOC', 'LOC')" & _
                     " )" & _
                     " group by fcba" & _
                     "          , to_number(to_char(tdate, 'yyyy'))" & _
                     "          , to_number(to_char(tdate, 'mm'))" & _
                     "          , allowdedcode" & _
                     "          , targettype" & _
                     "          , targetcode"

	ElseIf (Session("Estate_Mill") = "M" Or Session("Estate_Mill") = "P") Then
		ExecuteSQL = " insert into attendance_gad_closing" & _
					 " select *" & _
					 " from attendance_gad" & _
					 " where fcba = '" & Session("fcBA") & "'" & _
					 "       and to_number(to_char(fddate, 'yyyy')) = " & Year(Session("v_General_Charges")) & "" & _
					 "       and to_number(to_char(fddate, 'mm')) = " & Month(Session("v_General_Charges")) & "" & _
					 "       and attendance not in ('UN', 'KB', 'KL')" & _
					 "       and locationtype = 'MA'" & _
					 "       and mandays <> 0"
	End If

	connstr.execute(ExecuteSQL)
	
	If Err.Number = 0 Then
		connstr.CommitTrans
	Else
		connstr.RollbackTrans
		Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
		Response.Flush
	End If
	
	connstr.BeginTrans
	If Session("Estate_Mill")="E" then
		ExecuteSQL = " select * from dual"
	ElseIf (Session("Estate_Mill")= "M" Or Session("Estate_Mill")= "P") Then
		ExecuteSQL = " insert into loc_allocation" & _
                     " (fcba, fddate, jobcode, locationtype, locationcode, mandays, loc_group)" & _
                     " select fcba" & _
                     "        , fddate" & _
                     "        , jobcode" & _
                     "        , locationtype" & _
                     "        , locationcode" & _
                     "        , mandays" & _
                     "        , 'MILL' loc_group" & _
                     " from (" & _
                     "     select fcba" & _
                     "            , trunc(fddate, 'mm') fddate" & _
                     "            , jobcode" & _
                     "            , locationtype" & _
                     "            , locationcode" & _
                     "            , sum(mandays) mandays" & _
                     "     from attendance_gad_closing" & _
                     "     where fcba = '" & Session("fcBA") & "'" & _
                     "     group by fcba" & _
                     "              , trunc(fddate, 'mm')" & _
                     "              , jobcode" & _
                     "              , locationtype" & _
                     "              , locationcode" & _
                     " )"
	End If


	connstr.execute(ExecuteSQL)
	
	If Err.Number = 0 Then
		connstr.CommitTrans
	Else
		connstr.RollbackTrans
		Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
		Response.Flush
	End If
	
	connstr.BeginTrans
	If Session("Estate_Mill")="E" then
		ExecuteSQL = " insert into costtotal_pertarget" & _
					 " (FCBA, YEAR, MONTH, TARGETTYPE, TARGETCODE, COSTTOTAL, BASIS)" & _
                     " select fcba" & _
                     "        , to_number(to_char(tdate, 'yyyy')) year" & _
                     "        , to_number(to_char(tdate, 'mm')) month" & _
                     "        , targettype" & _
                     "        , targetcode" & _
                     "        , sum(amount) costtotal" & _
                     "        , 0 basis" & _
                     " from (" & _
                     "     select ca.fcba" & _
                     "            , ca.tdate" & _
                     "            , ca.targettype" & _
                     "            , ca.targetcode" & _
                     "            , ca.amount" & _
                     "     from ( select fcba" & _
                     "                   , trunc(tdate, 'mm') tdate" & _
                     "                   , targettype" & _
                     "                   , targetcode" & _
                     "                   , substr(referenceno, 1, 3) allowdedcode" & _
                     "                   , amount" & _
                     "            from cost_activity" & _
                     "            where fcba = '" & Session("fcBA") & "'" & _
                     "                  and to_number(to_char(tdate,'yyyy')) = " & Year(Session("v_General_Charges")) & "" & _
                     "                  and to_number(to_char(tdate,'mm')) = " & Month(Session("v_General_Charges")) & "" & _
                     "                  and sourcetype in ('ES', 'EM')" & _
                     "                  and targettype in ('SOC', 'LOC')" & _
                     "     ) ca inner join upr_component_tab uct" & _
                     "     on (ca.allowdedcode = uct.allowdedcode)" & _
                     "     union all" & _
                     "     select fcba" & _
                     "            , trunc(tdate, 'mm') tdate" & _
                     "            , targettype" & _
                     "            , targetcode" & _
                     "            , amount" & _
                     "     from cost_activity" & _
                     "     where fcba = '" & Session("fcBA") & "'" & _
                     "     and to_number(to_char(tdate,'yyyy')) = " & Year(Session("v_General_Charges")) & "" & _
                     "     and to_number(to_char(tdate,'mm')) = " & Month(Session("v_General_Charges")) & "" & _
                     "     and sourcetype not in ('ES', 'EM')" & _
                     "     and targettype in ('SOC', 'LOC') " & _
                     " )" & _
                     " group by fcba" & _
                     "          , to_number(to_char(tdate, 'yyyy'))" & _
                     "          , to_number(to_char(tdate, 'mm'))" & _
                     "          , targettype" & _
                     "          , targetcode"
	else	
		ExecuteSQL = " insert into costtotal_pertarget" & _
					 " (FCBA, YEAR, MONTH, TARGETTYPE, TARGETCODE, COSTTOTAL, BASIS)" & _
					 " select fcba" & _
					 "        , to_number(to_char(tdate,'yyyy')) year" & _
					 "        , to_number(to_char(tdate,'mm')) month" & _
					 "        , targettype" & _
					 "        , targetcode" & _
					 "        , costtotal" & _
					 "        , loc_closing_api.get_basis(fcba, tdate, gcmaster_api.get_gcgroup(fcba, 'LOC', targetcode)) basis" & _
					 " from (" & _
					 "      select fcba" & _
					 "             , trunc(tdate, 'mm') tdate" & _
					 "             , targettype" & _
					 "             , targetcode" & _
					 "             , nvl(sum(amount),0) as costtotal" & _
					 "      from cost_activity" & _
					 "      where fcba = '" & Session("fcBA") & "'" & _
					 "            and to_number(to_char(tdate,'yyyy')) = " & Year(Session("v_General_Charges")) & "" & _
					 "            and to_number(to_char(tdate,'mm')) = " & Month(Session("v_General_Charges")) & "" & _
					 "            and targettype = 'LOC'" & _
					 "      group by fcba" & _
					 "               , trunc(tdate, 'mm')" & _
					 "               , targettype" & _
					 "               , targetcode" & _
					 " )"
	end if
	
	connstr.execute(ExecuteSQL)
	
	If Err.Number = 0 Then
		connstr.CommitTrans
	Else
		connstr.RollbackTrans
		Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
		Response.Flush
	End If
	
	Response.write("<font face=Verdana size=1 color=Green>Generating Data, Success!.<br/><br/></font>")
	Response.Flush

	Response.Write "Generate Credit Allocation" & "<br/>"
	Response.Flush
	
	jobcodeNoBasis = "1140800"
	
	If Session("Estate_Mill")="E" then
		qCreditAlokasi = " select * from (" & _
                         "     select fcba" & _
                         "            , year" & _
                         "            , month" & _
                         "            , allowdedcode" & _
                         "            , targettype" & _
                         "            , targetcode" & _
                         "            , costtotal" & _
                         "            , new_loc_closing_api.get_allowded_role(allowdedcode) role" & _
                         "            , new_loc_closing_api.get_basisperallowded(fcba, year, month, allowdedcode, targettype, targetcode) basis" & _
                         "     from allowances_loc_tab" & _
                         "     where fcba = '" & Session("fcBA") & "'" & _
                         " )" & _
                         " where role <> 'BORONGAN'" & _
                         "       and basis = 0"
						 
		arrRS = GetRSArray(qCreditAlokasi)
		
		If IsArray(arrRS) Then
			For i = LBound(arrRS, 2) To UBound(arrRS, 2)
				If Cdbl(arrRS(6, i)) <> 0 Then
					temp_referenceno = arrRS(3, i)
					if IsNull(temp_referenceno) then
						temp_referenceno = "New"
					end if
					
					connstr.BeginTrans
					qInsertCR = " insert into loccb (fieldcode" & _
								"                   , targettype" & _
								"                   , targetcode" & _
								"                   , jobcode" & _
								"                   , quantity" & _
								"                   , amount" & _
								"                   , transdate" & _
								"                   , referenceno" & _
								"                   , fcentry" & _
								"                   , fcedit" & _
								"                   , fcip" & _
								"                   , fcba" & _
								"                   , lastupdate" & _
								"                   , lasttime" & _
								"                   )" & _
								" values ('" & arrRS(5, i) & "'" & _
								"         , 'CB'" & _
								"         , 'BS'" & _
								"         , '" & txtControlJob & "'" & _
								"         , '0'" & _
								"         , " & Cdbl(arrRS(6, i)) * -1 & "" & _
								"         , " & " to_date('" & Session("v_General_Charges") & "','mm/dd/yyyy')" & _
								"         , '" & arrRS(5, i) & "-Cr-NoAct-" & temp_referenceno & "'" & _
								"         , '" & Session("CodeName") & "'" & _
								"         , '" & Session("CodeName") & "'" & _
								"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
								"         , '" & Session("fcBA") & "'" & _
								"         , sysdate" & _
								"         , to_char(sysdate, 'HH24:MI')" & _
								"         )"
					 'response.Write(qInsertCR)			  
					connStr.execute(qInsertCR)
					
					qInsertDB = " insert into loccb (fieldcode" & _
								"                   , targettype" & _
								"                   , targetcode" & _
								"                   , jobcode" & _
								"                   , quantity" & _
								"                   , amount" & _
								"                   , transdate" & _
								"                   , referenceno" & _
								"                   , fcentry" & _
								"                   , fcedit" & _
								"                   , fcip" & _
								"                   , fcba" & _
								"                   , lastupdate" & _
								"                   , lasttime" & _
								"                   )" & _
								" values ('" & arrRS(5, i) & "'" & _
								"         , 'CB'" & _
								"         , 'BS'" & _
								"         , '" & jobcodeNoBasis & "'" & _
								"         , '0'" & _
								"         , " & Cdbl(arrRS(6, i)) & "" & _
								"         , " & " to_date('" & Session("v_General_Charges") & "','mm/dd/yyyy')" & _
								"         , '" & arrRS(5, i) & "-Db-NoAct-" & temp_referenceno & "'" & _
								"         , '" & Session("CodeName") & "'" & _
								"         , '" & Session("CodeName") & "'" & _
								"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
								"         , '" & Session("fcBA") & "'" & _
								"         , sysdate" & _
								"         , to_char(sysdate, 'HH24:MI')" & _
								"         )"
					 'response.Write(qInsertDB)			  
					connStr.execute(qInsertDB)

					If Err.Number = 0 Then
						connstr.CommitTrans
						Response.Flush
					Else
						connstr.RollbackTrans
						Response.write("Error Save ["&qInsertCR&"] - ["&qInsertDB&"] ("&Err.Description&")")
						Response.Flush
					End If
				End If
			Next
			Erase arrRS
		End If
	else
		qCreditAlokasi = " select fcba" & _
						 "        , targettype" & _
						 "        , targetcode" & _
						 "        , costtotal" & _
						 "        , basis" & _
						 " from costtotal_pertarget" & _
						 " where fcba = '" & Session("fcBA") & "'" & _
						 "       and year = " & Year(Session("v_General_Charges")) & "" & _
						 "       and month = " & Month(Session("v_General_Charges")) & "" & _
						 "       and targettype = 'LOC'"
						 
		arrRS = GetRSArray(qCreditAlokasi)
		
		If IsArray(arrRS) Then
			For i = LBound(arrRS, 2) To UBound(arrRS, 2)
				If Cdbl(arrRS(4, i)) = 0 Then
					connstr.BeginTrans
					qInsertCR = " insert into loccb (fieldcode" & _
								"                   , targettype" & _
								"                   , targetcode" & _
								"                   , jobcode" & _
								"                   , quantity" & _
								"                   , amount" & _
								"                   , transdate" & _
								"                   , referenceno" & _
								"                   , fcentry" & _
								"                   , fcedit" & _
								"                   , fcip" & _
								"                   , fcba" & _
								"                   , lastupdate" & _
								"                   , lasttime" & _
								"                   )" & _
								" values ('" & arrRS(2, i) & "'" & _
								"         , 'CB'" & _
								"         , 'BS'" & _
								"         , '" & txtControlJob & "'" & _
								"         , '0'" & _
								"         , " & cdbl(arrRS(3, i)) * -1 & "" & _
								"         , " & " to_date('" & Session("v_General_Charges") & "','mm/dd/yyyy')" & _
								"         , '" & arrRS(2, i) & "-Cr-No Basis" & "'" & _
								"         , '" & Session("CodeName") & "'" & _
								"         , '" & Session("CodeName") & "'" & _
								"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
								"         , '" & Session("fcBA") & "'" & _
								"         , sysdate" & _
								"         , to_char(sysdate, 'HH24:MI')" & _
								"         )"
					 'response.Write(qInsertCR)			  
					connStr.execute(qInsertCR)
					
					qInsertDB = " insert into loccb (fieldcode" & _
								"                   , targettype" & _
								"                   , targetcode" & _
								"                   , jobcode" & _
								"                   , quantity" & _
								"                   , amount" & _
								"                   , transdate" & _
								"                   , referenceno" & _
								"                   , fcentry" & _
								"                   , fcedit" & _
								"                   , fcip" & _
								"                   , fcba" & _
								"                   , lastupdate" & _
								"                   , lasttime" & _
								"                   )" & _
								" values ('" & arrRS(2, i) & "'" & _
								"         , 'CB'" & _
								"         , 'BS'" & _
								"         , '" & jobcodeNoBasis & "'" & _
								"         , '0'" & _
								"         , " & cdbl(arrRS(3, i)) & "" & _
								"         , " & " to_date('" & Session("v_General_Charges") & "','mm/dd/yyyy')" & _
								"         , '" & arrRS(2, i) & "-Db-No Basis" & "'" & _
								"         , '" & Session("CodeName") & "'" & _
								"         , '" & Session("CodeName") & "'" & _
								"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
								"         , '" & Session("fcBA") & "'" & _
								"         , sysdate" & _
								"         , to_char(sysdate, 'HH24:MI')" & _
								"         )"
					 'response.Write(qInsertDB)			  
					connStr.execute(qInsertDB)

					If Err.Number = 0 Then
						connstr.CommitTrans
						Response.Flush
					Else
						connstr.RollbackTrans
						Response.write("Error Save ["&qInsertCR&"] - ["&qInsertDB&"] ("&Err.Description&")")
						Response.Flush
					End If
				Else
					connstr.BeginTrans
					qInsertCR = " insert into loccb (fieldcode" & _
								"                   , targettype" & _
								"                   , targetcode" & _
								"                   , jobcode" & _
								"                   , quantity" & _
								"                   , amount" & _
								"                   , transdate" & _
								"                   , referenceno" & _
								"                   , fcentry" & _
								"                   , fcedit" & _
								"                   , fcip" & _
								"                   , fcba" & _
								"                   , lastupdate" & _
								"                   , lasttime" & _
								"                   )" & _
								" values ('" & arrRS(2, i) & "'" & _
								"         , 'CB'" & _
								"         , 'BS'" & _
								"         , '" & txtControlJob & "'" & _
								"         , '0'" & _
								"         , " & cdbl(arrRS(3, i)) * -1 & "" & _
								"         , " & " to_date('" & Session("v_General_Charges") & "','mm/dd/yyyy')" & _
								"         , '" & arrRS(2, i) & "-Cr" & "'" & _
								"         , '" & Session("CodeName") & "'" & _
								"         , '" & Session("CodeName") & "'" & _
								"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
								"         , '" & Session("fcBA") & "'" & _
								"         , sysdate" & _
								"         , to_char(sysdate, 'HH24:MI')" & _
								"         )"
					 'response.Write(qInsertCR)			  
					connStr.execute(qInsertCR)

					If Err.Number = 0 Then
						connstr.CommitTrans
						Response.Flush
					Else
						connstr.RollbackTrans
						Response.write("Error Save ["&qInsertCR&"] ("&Err.Description&")")
						Response.Flush
					End If
				End If
			Next
			Erase arrRS
		End If
						 
    end if
				  

	
	Response.write("<font face=Verdana size=1 color=Green> Generate Credit Allocation, Success!.<br/><br/></font>")
	Response.Flush
	
	Response.Write "Generate Debet Allocation" & "<br/>"
	Response.Flush
	
	
	If Session("Estate_Mill")="E" then
		qDebetAlokasi = " select aa.fcba" & _
						"        , aa.transdate" & _
						"        , aa.targettype source_sourcetype" & _
						"        , aa.targetcode source_sourcecode" & _
						"        , aa.costtotal" & _
						"        , new_loc_closing_api.get_mandays_per_st(aa.fcba, aa.year, aa.month, aa.section, aa.grade, aa.sourcetype) basis" & _
						"        , bb.jobcode" & _
						"        , bb.targettype" & _
						"        , bb.targetcode       " & _
						"        , bb.mandays actual" & _
						"        , aa.section" & _
						" from ( select fcba" & _
						"               , year" & _
						"               , month" & _
						"               , last_day(to_date(month||'/1/'||year, 'mm/dd/yyyy')) transdate" & _
						"               , new_loc_closing_api.get_sourcetype(targettype) sourcetype" & _
						"               , new_loc_closing_api.get_section(targettype, targetcode) section" & _
						"               , new_loc_closing_api.get_grade(targettype, targetcode) grade" & _
						"               , targettype" & _
						"               , targetcode" & _
						"               , costtotal       " & _
						"        from costtotal_pertarget" & _
						"        where fcba = '" & Session("fcBA") & "'" & _
						" ) aa" & _
						" inner join ( select fcba" & _
						"                      , transdate" & _
						"                      , section" & _
						"                      , jobcode" & _
						"                      , targettype" & _
						"                      , targetcode" & _
						"                      , sum(mandays) mandays" & _
						"              from employee_activities_es_tab" & _
						"              where fcba = '" & Session("fcBA") & "'" & _
						"              group by fcba" & _
						"                       , transdate" & _
						"                       , section" & _
						"                       , jobcode" & _
						"                       , targettype" & _
						"                       , targetcode" & _
						" ) bb" & _
						" on (aa.fcba = bb.fcba and aa.transdate = bb.transdate and aa.section = bb.section)" & _
						" where aa.sourcetype = 'ES'" & _
						" union all" & _
						" select aa.fcba" & _
						"        , aa.transdate" & _
						"        , aa.targettype source_sourcetype" & _
						"        , aa.targetcode source_sourcecode" & _
						"        , aa.costtotal" & _
						"        , new_loc_closing_api.get_mandays_per_st(aa.fcba, aa.year, aa.month, aa.section, aa.grade, aa.sourcetype) basis" & _
						"        , bb.jobcode" & _
						"        , bb.targettype" & _
						"        , bb.targetcode       " & _
						"        , bb.mandays actual" & _
						"        , aa.section" & _
						" from ( select fcba" & _
						"               , year" & _
						"               , month" & _
						"               , last_day(to_date(month||'/1/'||year, 'mm/dd/yyyy')) transdate" & _
						"               , new_loc_closing_api.get_sourcetype(targettype) sourcetype" & _
						"               , new_loc_closing_api.get_section(targettype, targetcode) section" & _
						"               , new_loc_closing_api.get_grade(targettype, targetcode) grade" & _
						"               , targettype" & _
						"               , targetcode" & _
						"               , costtotal       " & _
						"        from costtotal_pertarget" & _
						"        where fcba = '" & Session("fcBA") & "'" & _
						" ) aa" & _
						" inner join ( select fcba" & _
						"                      , transdate" & _
						"                      , section" & _
						"                      , grade" & _
						"                      , jobcode" & _
						"                      , targettype" & _
						"                      , targetcode" & _
						"                      , sum(mandays) mandays" & _
						"              from employee_activities_em_tab" & _
						"              where fcba = '" & Session("fcBA") & "'" & _
						"              group by fcba" & _
						"                       , transdate" & _
						"                       , section" & _
						"                       , grade" & _
						"                       , jobcode" & _
						"                       , targettype" & _
						"                       , targetcode" & _
						" ) bb" & _
						" on (aa.fcba = bb.fcba and aa.transdate = bb.transdate and aa.section = bb.section and aa.grade = bb.grade)" & _
						" where aa.sourcetype = 'EM'"
		
		arrRS = GetRSArray(qDebetAlokasi)

		If IsArray(arrRS) Then
			For i = LBound(arrRS, 2) To UBound(arrRS, 2)
				If Cdbl(arrRS(5, i)) <> 0 Then
					nFFAlokasi = 0
					nFFAlokasi = Round(Cdbl(arrRS(4, i))/Cdbl(arrRS(5, i))*Cdbl(arrRS(9, i)), 0)
					If Cdbl(nFFAlokasi) <> 0 Then
						connstr.BeginTrans
						
						qInsertCR = " insert into loccb (fieldcode" & _
									"                   , targettype" & _
									"                   , targetcode" & _
									"                   , jobcode" & _
									"                   , quantity" & _
									"                   , amount" & _
									"                   , transdate" & _
									"                   , referenceno" & _
									"                   , fcentry" & _
									"                   , fcedit" & _
									"                   , fcip" & _
									"                   , fcba" & _
									"                   , lastupdate" & _
									"                   , lasttime" & _
									"                   )" & _
									" values ('" & arrRS(3, i) & "'" & _
									"         , 'CB'" & _
									"         , 'BS'" & _
									"         , '" & txtControlJob & "'" & _
									"         , '" & arrRS(9, i) & "'" & _
									"         , " & nFFAlokasi * -1 & "" & _
									"         , " & " to_date('" & arrRS(1, i) & "','mm/dd/yyyy')" & _
									"         , '" & arrRS(3, i) & "-Cr-New'" & _
									"         , '" & Session("CodeName") & "'" & _
									"         , '" & Session("CodeName") & "'" & _
									"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
									"         , '" & Session("fcBA") & "'" & _
									"         , sysdate" & _
									"         , to_char(sysdate, 'HH24:MI')" & _
									"         )"
						 'response.Write(qInsertCR)			  
						connStr.execute(qInsertCR)
						
						qInsertDB = " insert into loccb (fieldcode" & _
									"                   , targettype" & _
									"                   , targetcode" & _
									"                   , jobcode" & _
									"                   , quantity" & _
									"                   , amount" & _
									"                   , transdate" & _
									"                   , referenceno" & _
									"                   , fcentry" & _
									"                   , fcedit" & _
									"                   , fcip" & _
									"                   , fcba" & _
									"                   , lastupdate" & _
									"                   , lasttime" & _
									"                   )" & _
									" values ('" & arrRS(3, i) & "'" & _
									"         , '" & arrRS(7, i) & "'" & _
									"         , '" & arrRS(8, i) & "'" & _
									"         , '" & arrRS(6, i) & "'" & _
									"         , " & cdbl(arrRS(9, i)) & "" & _
									"         , " & nFFAlokasi & "" & _
									"         , " & " to_date('" & arrRS(1, i) & "','mm/dd/yyyy')" & _
									"         , '" & arrRS(3, i) & "-Db-New" & "'" & _
									"         , '" & Session("CodeName") & "'" & _
									"         , '" & Session("CodeName") & "'" & _
									"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
									"         , '" & Session("fcBA") & "'" & _
									"         , sysdate" & _
									"         , to_char(sysdate, 'HH24:MI')" & _
									"         )"
						 'response.Write(qInsertDB)			  
						connStr.execute(qInsertDB)

						If Err.Number = 0 Then
							connstr.CommitTrans
							Response.Flush
						Else
							connstr.RollbackTrans
							Response.write("Error Save ["&qInsertDB&"] ("&Err.Description&")")
							Response.Flush
						End If
					End If
				End If
			Next
			Erase arrRS
		End If
	else
		qDebetAlokasi = " select ca.fcba" & _
						"        , ca.fddate" & _
						"        , targettype" & _
						"        , targetcode" & _
						"        , costtotal" & _
						"        , basis" & _
						"        , jobcode" & _
						"        , locationtype" & _
						"        , locationcode" & _
						"        , mandays" & _
						"        , loc_group" & _
						" from (" & _
						"     select fcba" & _
						"            , to_date(lpad(month, 2, 0)||'/01/'||year,'mm/dd/yyyy') fddate" & _
						"            , targettype" & _
						"            , targetcode" & _
						"            , costtotal" & _
						"            , basis" & _
						"     from costtotal_pertarget" & _
						"     where fcba = '" & Session("fcBA") & "'" & _
						"           and year = " & Year(Session("v_General_Charges")) & "" & _
						"           and month = " & Month(Session("v_General_Charges")) & "" & _
						"           and targettype = 'LOC'" & _
						" ) ca" & _
						" inner join loc_allocation lo" & _
						" on (ca.fcba = lo.fcba and ca.fddate = lo.fddate and gcmaster_api.get_gcgroup(ca.fcba, 'LOC', ca.targetcode) = lo.loc_group)"
		
		arrRS = GetRSArray(qDebetAlokasi)

		If IsArray(arrRS) Then
			For i = LBound(arrRS, 2) To UBound(arrRS, 2)
				If Cdbl(arrRS(5, i)) <> 0 Then
					nFFAlokasi = Round(Cdbl(arrRS(9, i))/Cdbl(arrRS(5, i))*Cdbl(arrRS(4, i)), 0)
					If Cdbl(nFFAlokasi) <> 0 Then
						connstr.BeginTrans
						qInsertDB = " insert into loccb (fieldcode" & _
									"                   , targettype" & _
									"                   , targetcode" & _
									"                   , jobcode" & _
									"                   , quantity" & _
									"                   , amount" & _
									"                   , transdate" & _
									"                   , referenceno" & _
									"                   , fcentry" & _
									"                   , fcedit" & _
									"                   , fcip" & _
									"                   , fcba" & _
									"                   , lastupdate" & _
									"                   , lasttime" & _
									"                   )" & _
									" values ('" & arrRS(3, i) & "'" & _
									"         , '" & arrRS(7, i) & "'" & _
									"         , '" & arrRS(8, i) & "'" & _
									"         , '" & arrRS(6, i) & "'" & _
									"         , " & cdbl(arrRS(9, i)) & "" & _
									"         , " & nFFAlokasi & "" & _
									"         , " & " to_date('" & Session("v_General_Charges") & "','mm/dd/yyyy')" & _
									"         , '" & arrRS(3, i) & "-Db" & "'" & _
									"         , '" & Session("CodeName") & "'" & _
									"         , '" & Session("CodeName") & "'" & _
									"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
									"         , '" & Session("fcBA") & "'" & _
									"         , sysdate" & _
									"         , to_char(sysdate, 'HH24:MI')" & _
									"         )"
						 'response.Write(qInsertDB)			  
						connStr.execute(qInsertDB)

						If Err.Number = 0 Then
							connstr.CommitTrans
							Response.Flush
						Else
							connstr.RollbackTrans
							Response.write("Error Save ["&qInsertDB&"] ("&Err.Description&")")
							Response.Flush
						End If
					End If
				End If
			Next
			Erase arrRS
		End If
	end if

	Response.write("<font face=Verdana size=1 color=Green> Generate Debet Allocation, Success!.<br/><br/></font>")
	Response.Flush
	
	Response.Write "Generate Non UPR Allocation" & "<br/>"
	Response.Flush
	
	If Session("Estate_Mill")="E" then
		qnon_upr = " select aaa.fcba" & _
                   "        , aaa.year" & _
                   "        , aaa.month" & _
                   "        , aaa.allowdedcode" & _
                   "        , aaa.targettype" & _
                   "        , aaa.targetcode" & _
                   "        , aaa.sourcetype" & _
                   "        , aaa.section" & _
                   "        , aaa.costtotal" & _
                   "        , aaa.basis" & _
                   "        , bbb.jobcode" & _
                   "        , bbb.tt" & _
                   "        , bbb.tc" & _
                   "        , NEW_LOC_CLOSING_API.get_mandays_per_allowded_job(aaa.fcba" & _
                   "                                                           , aaa.year" & _
                   "                                                           , aaa.month" & _
                   "                                                           , aaa.section" & _
                   "                                                           , aaa.sourcetype" & _
                   "                                                           , aaa.allowdedcode" & _
                   "                                                           , bbb.jobcode" & _
                   "                                                           , bbb.tt" & _
                   "                                                           , bbb.tc) actual" & _
                   " from (" & _
                   " select fcba" & _
                   "        , year" & _
                   "        , month" & _
                   "        , allowdedcode" & _
                   "        , targettype" & _
                   "        , targetcode" & _
                   "        , sourcetype" & _
                   "        , section" & _
                   "        , costtotal" & _
                   "        , NEW_LOC_CLOSING_API.get_mandays_per_allowded(fcba" & _
                   "                                                       , year" & _
                   "                                                       , month" & _
                   "                                                       , section" & _
                   "                                                       , sourcetype" & _
                   "                                                       , allowdedcode) basis" & _
                   " from (" & _
                   "     select aa.fcba" & _
                   "            , aa.year" & _
                   "            , aa.month" & _
                   "            , aa.allowdedcode" & _
                   "            , aa.targettype" & _
                   "            , aa.targetcode" & _
                   "            , new_loc_closing_api.get_sourcetype(aa.targettype) sourcetype" & _
                   "            , new_loc_closing_api.get_section(aa.targettype, aa.targetcode) section" & _
                   "            , aa.costtotal" & _
                   "     from allowances_loc_tab aa" & _
                   "     inner join ( select allowdedcode " & _
                   "                  from allowances_loc_tab" & _
                   "                  where fcba = '" & Session("fcBA") & "'" & _
                   "                        and allowdedcode is not null" & _
                   "                        and allowdedcode not in ('A45', 'A26')" & _
                   "                  minus" & _
                   "                  select allowdedcode" & _
                   "                  from upr_component_tab" & _
                   "     ) bb" & _
                   "     on (aa.allowdedcode = bb.allowdedcode)" & _
                   "     where aa.fcba = '" & Session("fcBA") & "'" & _
                   " )" & _
                   " ) aaa" & _
                   " inner join (select distinct fcba" & _
                   "                    , sourcetype" & _
                   "                    , section" & _
                   "                    , jobcode" & _
                   "                    , targettype tt" & _
                   "                    , targetcode tc " & _
                   "             from employee_activities_em_tab" & _
                   "             where fcba = '" & Session("fcBA") & "') bbb" & _
                   " on (aaa.fcba = bbb.fcba and aaa.sourcetype = bbb.sourcetype and aaa.section = bbb.section)"
				  
		arrRS = GetRSArray(qnon_upr)
	
		If IsArray(arrRS) Then
			For i = LBound(arrRS, 2) To UBound(arrRS, 2)
				If Cdbl(arrRS(9, i)) <> 0 Then 'Jika Basis <> 0
					val_hasil = 0
					val_hasil = Round( (Cdbl(arrRS(8, i))/Cdbl(arrRS(9, i))) * Cdbl(arrRS(13, i)) , 0)
					
					if val_hasil <> 0 then 'Jika hasil proporsi <> 0
						connstr.BeginTrans
						qInsertCR = " insert into loccb (fieldcode" & _
									"                   , targettype" & _
									"                   , targetcode" & _
									"                   , jobcode" & _
									"                   , quantity" & _
									"                   , amount" & _
									"                   , transdate" & _
									"                   , referenceno" & _
									"                   , fcentry" & _
									"                   , fcedit" & _
									"                   , fcip" & _
									"                   , fcba" & _
									"                   , lastupdate" & _
									"                   , lasttime" & _
									"                   )" & _
									" values ('" & arrRS(5, i) & "'" & _
									"         , 'CB'" & _
									"         , 'BS'" & _
									"         , '" & txtControlJob & "'" & _
									"         , '" & arrRS(13, i) & "'" & _
									"         , " & val_hasil * -1 & "" & _
									"         , " & " to_date('" & Session("v_General_Charges") & "','mm/dd/yyyy')" & _
									"         , '" & arrRS(5, i) & "-Cr-" & arrRS(3, i) & "'" & _
									"         , '" & Session("CodeName") & "'" & _
									"         , '" & Session("CodeName") & "'" & _
									"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
									"         , '" & Session("fcBA") & "'" & _
									"         , sysdate" & _
									"         , to_char(sysdate, 'HH24:MI')" & _
									"         )"
						 'response.Write(qInsertCR)			  
						connStr.execute(qInsertCR)
						
						qInsertDB = " insert into loccb (fieldcode" & _
									"                   , targettype" & _
									"                   , targetcode" & _
									"                   , jobcode" & _
									"                   , quantity" & _
									"                   , amount" & _
									"                   , transdate" & _
									"                   , referenceno" & _
									"                   , fcentry" & _
									"                   , fcedit" & _
									"                   , fcip" & _
									"                   , fcba" & _
									"                   , lastupdate" & _
									"                   , lasttime" & _
									"                   )" & _
									" values ('" & arrRS(5, i) & "'" & _
									"         , '" & arrRS(11, i) & "'" & _
									"         , '" & arrRS(12, i) & "'" & _
									"         , '" & arrRS(10, i) & "'" & _
									"         , '" & arrRS(13, i) & "'" & _
									"         , " & val_hasil & "" & _
									"         , " & " to_date('" & Session("v_General_Charges") & "','mm/dd/yyyy')" & _
									"         , '" & arrRS(5, i) & "-Db-" & arrRS(3, i) & "'" & _
									"         , '" & Session("CodeName") & "'" & _
									"         , '" & Session("CodeName") & "'" & _
									"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
									"         , '" & Session("fcBA") & "'" & _
									"         , sysdate" & _
									"         , to_char(sysdate, 'HH24:MI')" & _
									"         )"
						 'response.Write(qInsertDB)			  
						connStr.execute(qInsertDB)

						If Err.Number = 0 Then
							connstr.CommitTrans
							Response.Flush
						Else
							connstr.RollbackTrans
							Response.write("Error Save ["&qInsertCR&"] - ["&qInsertDB&"] ("&Err.Description&")")
							Response.Flush
						End If
					end if
				End If
			Next
			Erase arrRS
		End If
		
		'ES
		qnon_upr = " select aaa.fcba" & _
                   "        , aaa.year" & _
                   "        , aaa.month" & _
                   "        , aaa.allowdedcode" & _
                   "        , aaa.targettype" & _
                   "        , aaa.targetcode" & _
                   "        , aaa.sourcetype" & _
                   "        , aaa.section" & _
                   "        , aaa.costtotal" & _
                   "        , aaa.basis" & _
                   "        , bbb.jobcode" & _
                   "        , bbb.tt" & _
                   "        , bbb.tc" & _
                   "        , NEW_LOC_CLOSING_API.get_mandays_per_allowded_job(aaa.fcba" & _
                   "                                                           , aaa.year" & _
                   "                                                           , aaa.month" & _
                   "                                                           , aaa.section" & _
                   "                                                           , aaa.sourcetype" & _
                   "                                                           , aaa.allowdedcode" & _
                   "                                                           , bbb.jobcode" & _
                   "                                                           , bbb.tt" & _
                   "                                                           , bbb.tc) actual" & _
                   " from (" & _
                   " select fcba" & _
                   "        , year" & _
                   "        , month" & _
                   "        , allowdedcode" & _
                   "        , targettype" & _
                   "        , targetcode" & _
                   "        , sourcetype" & _
                   "        , section" & _
                   "        , costtotal" & _
                   "        , NEW_LOC_CLOSING_API.get_mandays_per_allowded(fcba" & _
                   "                                                       , year" & _
                   "                                                       , month" & _
                   "                                                       , section" & _
                   "                                                       , sourcetype" & _
                   "                                                       , allowdedcode) basis" & _
                   " from (" & _
                   "     select aa.fcba" & _
                   "            , aa.year" & _
                   "            , aa.month" & _
                   "            , aa.allowdedcode" & _
                   "            , aa.targettype" & _
                   "            , aa.targetcode" & _
                   "            , new_loc_closing_api.get_sourcetype(aa.targettype) sourcetype" & _
                   "            , new_loc_closing_api.get_section(aa.targettype, aa.targetcode) section" & _
                   "            , aa.costtotal" & _
                   "     from allowances_loc_tab aa" & _
                   "     inner join ( select allowdedcode " & _
                   "                  from allowances_loc_tab" & _
                   "                  where fcba = '" & Session("fcBA") & "'" & _
                   "                        and allowdedcode is not null" & _
                   "                        and allowdedcode not in ('A45', 'A26')" & _
                   "                  minus" & _
                   "                  select allowdedcode" & _
                   "                  from upr_component_tab" & _
                   "     ) bb" & _
                   "     on (aa.allowdedcode = bb.allowdedcode)" & _
                   "     where aa.fcba = '" & Session("fcBA") & "'" & _
                   " )" & _
                   " ) aaa" & _
                   " inner join (select distinct fcba" & _
                   "                    , sourcetype" & _
                   "                    , section" & _
                   "                    , jobcode" & _
                   "                    , targettype tt" & _
                   "                    , targetcode tc " & _
                   "             from employee_activities_es_tab" & _
                   "             where fcba = '" & Session("fcBA") & "') bbb" & _
                   " on (aaa.fcba = bbb.fcba and aaa.sourcetype = bbb.sourcetype and aaa.section = bbb.section)"
				  
		arrRS = GetRSArray(qnon_upr)
	
		If IsArray(arrRS) Then
			For i = LBound(arrRS, 2) To UBound(arrRS, 2)
				If Cdbl(arrRS(9, i)) <> 0 Then 'Jika Basis <> 0
					val_hasil = 0
					val_hasil = Round( (Cdbl(arrRS(8, i))/Cdbl(arrRS(9, i))) * Cdbl(arrRS(13, i)) , 0)
					
					if val_hasil <> 0 then 'Jika hasil proporsi <> 0
						connstr.BeginTrans
						qInsertCR = " insert into loccb (fieldcode" & _
									"                   , targettype" & _
									"                   , targetcode" & _
									"                   , jobcode" & _
									"                   , quantity" & _
									"                   , amount" & _
									"                   , transdate" & _
									"                   , referenceno" & _
									"                   , fcentry" & _
									"                   , fcedit" & _
									"                   , fcip" & _
									"                   , fcba" & _
									"                   , lastupdate" & _
									"                   , lasttime" & _
									"                   )" & _
									" values ('" & arrRS(5, i) & "'" & _
									"         , 'CB'" & _
									"         , 'BS'" & _
									"         , '" & txtControlJob & "'" & _
									"         , '" & arrRS(13, i) & "'" & _
									"         , " & val_hasil * -1 & "" & _
									"         , " & " to_date('" & Session("v_General_Charges") & "','mm/dd/yyyy')" & _
									"         , '" & arrRS(5, i) & "-Cr-" & arrRS(3, i) & "'" & _
									"         , '" & Session("CodeName") & "'" & _
									"         , '" & Session("CodeName") & "'" & _
									"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
									"         , '" & Session("fcBA") & "'" & _
									"         , sysdate" & _
									"         , to_char(sysdate, 'HH24:MI')" & _
									"         )"
						 'response.Write(qInsertCR)			  
						connStr.execute(qInsertCR)
						
						qInsertDB = " insert into loccb (fieldcode" & _
									"                   , targettype" & _
									"                   , targetcode" & _
									"                   , jobcode" & _
									"                   , quantity" & _
									"                   , amount" & _
									"                   , transdate" & _
									"                   , referenceno" & _
									"                   , fcentry" & _
									"                   , fcedit" & _
									"                   , fcip" & _
									"                   , fcba" & _
									"                   , lastupdate" & _
									"                   , lasttime" & _
									"                   )" & _
									" values ('" & arrRS(5, i) & "'" & _
									"         , '" & arrRS(11, i) & "'" & _
									"         , '" & arrRS(12, i) & "'" & _
									"         , '" & arrRS(10, i) & "'" & _
									"         , '" & arrRS(13, i) & "'" & _
									"         , " & val_hasil & "" & _
									"         , " & " to_date('" & Session("v_General_Charges") & "','mm/dd/yyyy')" & _
									"         , '" & arrRS(5, i) & "-Db-" & arrRS(3, i) & "'" & _
									"         , '" & Session("CodeName") & "'" & _
									"         , '" & Session("CodeName") & "'" & _
									"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
									"         , '" & Session("fcBA") & "'" & _
									"         , sysdate" & _
									"         , to_char(sysdate, 'HH24:MI')" & _
									"         )"
						 'response.Write(qInsertDB)			  
						connStr.execute(qInsertDB)

						If Err.Number = 0 Then
							connstr.CommitTrans
							Response.Flush
						Else
							connstr.RollbackTrans
							Response.write("Error Save ["&qInsertCR&"] - ["&qInsertDB&"] ("&Err.Description&")")
							Response.Flush
						End If
					end if
				End If
			Next
			Erase arrRS
		End If
		
	end if
	
	Response.write("<font face=Verdana size=1 color=Green>Generate Non UPR Allocation, Success!.<br/><br/></font>")
	Response.Flush
	
	Response.Write "Generate Borongan Allocation" & "<br/>"
	Response.Flush
	
	If Session("Estate_Mill")="E" then
		qborongan = " select fcba" & _
					"        , transdate" & _
					"        , allowdedcode" & _
					"        , alokasi_tt targettype" & _
					"        , alokasi_tc targetcode" & _
					"        , jobcode" & _
					"        , targettype tt" & _
					"        , targetcode tc" & _
					"        , sum(amount) amount" & _
					" from (" & _
					"     select fcba" & _
					"            , trunc(transdate, 'mm') transdate" & _
					"            , allowdedcode" & _
					"            , employeecode" & _
					"            , section" & _
					"            , grade" & _
					"            , amount" & _
					"            , sourcetype" & _
					"            , jobcode" & _
					"            , targettype" & _
					"            , targetcode" & _
					"            , alokasi_jobcode" & _
					"            , alokasi_tt" & _
					"            , alokasi_tc" & _
					"     from employee_borongan_tab" & _
					"     where fcba = '" & Session("fcBA") & "'" & _
					"           and alokasi_tt = 'LOC'" & _
					" )" & _
					" group by fcba" & _
					"          , transdate" & _
					"          , allowdedcode" & _
					"          , alokasi_tt" & _
					"          , alokasi_tc" & _
					"          , jobcode" & _
					"          , targettype" & _
					"          , targetcode"
				  
		arrRS = GetRSArray(qborongan)
	
		If IsArray(arrRS) Then
			For i = LBound(arrRS, 2) To UBound(arrRS, 2)
				If Cdbl(arrRS(8, i)) <> 0 Then 'Jika Basis <> 0
					connstr.BeginTrans
					val_hasil = 0
					val_hasil = Round( Cdbl(arrRS(8, i)), 0)
					qInsertCR = " insert into loccb (fieldcode" & _
								"                   , targettype" & _
								"                   , targetcode" & _
								"                   , jobcode" & _
								"                   , quantity" & _
								"                   , amount" & _
								"                   , transdate" & _
								"                   , referenceno" & _
								"                   , fcentry" & _
								"                   , fcedit" & _
								"                   , fcip" & _
								"                   , fcba" & _
								"                   , lastupdate" & _
								"                   , lasttime" & _
								"                   )" & _
								" values ('" & arrRS(4, i) & "'" & _
								"         , 'CB'" & _
								"         , 'BS'" & _
								"         , '" & txtControlJob & "'" & _
								"         , '0'" & _
								"         , " & val_hasil * -1 & "" & _
								"         , " & " to_date('" & arrRS(1, i) & "','mm/dd/yyyy')" & _
								"         , '" & arrRS(4, i) & "-Cr-" & arrRS(2, i) & "'" & _
								"         , '" & Session("CodeName") & "'" & _
								"         , '" & Session("CodeName") & "'" & _
								"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
								"         , '" & Session("fcBA") & "'" & _
								"         , sysdate" & _
								"         , to_char(sysdate, 'HH24:MI')" & _
								"         )"
					 'response.Write(qInsertCR)			  
					connStr.execute(qInsertCR)
					
					qInsertDB = " insert into loccb (fieldcode" & _
								"                   , targettype" & _
								"                   , targetcode" & _
								"                   , jobcode" & _
								"                   , quantity" & _
								"                   , amount" & _
								"                   , transdate" & _
								"                   , referenceno" & _
								"                   , fcentry" & _
								"                   , fcedit" & _
								"                   , fcip" & _
								"                   , fcba" & _
								"                   , lastupdate" & _
								"                   , lasttime" & _
								"                   )" & _
								" values ('" & arrRS(4, i) & "'" & _
								"         , '" & arrRS(6, i) & "'" & _
								"         , '" & arrRS(7, i) & "'" & _
								"         , '" & arrRS(5, i) & "'" & _
								"         , '0'" & _
								"         , " & val_hasil & "" & _
								"         , " & " to_date('" & arrRS(1, i) & "','mm/dd/yyyy')" & _
								"         , '" & arrRS(4, i) & "-Db-" & arrRS(2, i) & "'" & _
								"         , '" & Session("CodeName") & "'" & _
								"         , '" & Session("CodeName") & "'" & _
								"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
								"         , '" & Session("fcBA") & "'" & _
								"         , sysdate" & _
								"         , to_char(sysdate, 'HH24:MI')" & _
								"         )"
					 'response.Write(qInsertDB)			  
					connStr.execute(qInsertDB)

					If Err.Number = 0 Then
						connstr.CommitTrans
						Response.Flush
					Else
						connstr.RollbackTrans
						Response.write("Error Save ["&qInsertCR&"] - ["&qInsertDB&"] ("&Err.Description&")")
						Response.Flush
					End If
				End If
			Next
			Erase arrRS
		End If
	end if
	
	Response.write("<font face=Verdana size=1 color=Green>Generate Borongan Allocation, Success!.<br/><br/></font>")
	Response.Flush
	
	Response.Write "Normalisasi Rounding" & "<br/>"
	Response.Flush
	
	If Session("Estate_Mill")="E" then
		qSelisih = " select fcba" & _
                   "        , year" & _
                   "        , month" & _
                   "        , targettype" & _
                   "        , targetcode" & _
                   "        , referenceno" & _
                   "        , costtotal" & _
                   "        , alokasi" & _
                   "        , costtotal - alokasi var" & _
                   " from (" & _
                   "     select fcba" & _
                   "            , year" & _
                   "            , month" & _
                   "            , targettype" & _
                   "            , targetcode" & _
                   "            , referenceno" & _
                   "            , sum(costtotal) costtotal" & _
                   "            , NEW_LOC_CLOSING_API.get_totalalokasi(fcba, year, month, targetcode, referenceno) alokasi" & _
                   "     from (" & _
                   "         select fcba" & _
                   "                , year" & _
                   "                , month" & _
                   "                , allowdedcode" & _
                   "                , targettype" & _
                   "                , targetcode" & _
                   "                , new_loc_closing_api.get_reference(allowdedcode) referenceno" & _
                   "                , costtotal       " & _
                   "         from allowances_loc_tab " & _
                   "         where fcba = '" & Session("fcBA") & "'" & _
                   "     )" & _
                   "     group by fcba" & _
                   "              , year" & _
                   "              , month" & _
                   "              , targettype" & _
                   "              , targetcode" & _
                   "              , referenceno" & _
                   "              , NEW_LOC_CLOSING_API.get_totalalokasi(fcba, year, month, targetcode, referenceno)" & _
                   " )" & _
                   " where costtotal - alokasi <> 0"
				   
		arrRS = GetRSArray(qSelisih)
		
		If IsArray(arrRS) Then
			For i = LBound(arrRS, 2) To UBound(arrRS, 2)
			
				Set objectRS = Server.CreateObject("ADODB.Recordset")
				objectSQL = " select fieldcode, transdate, jobcode, targettype, targetcode" & _
							" from (" & _
							"     select *" & _
							"     from loccb" & _
							"     where fcba = '" & Session("fcBA") & "'" & _
							"           and trunc(transdate, 'mm') = to_date('" & arrRS(2, i) & "/01/" & arrRS(1, i) & "', 'mm/dd/yyyy')" & _
							"           and jobcode <> '" & txtControlJob & "'" & _
							"           and fieldcode = '" & arrRS(4, i) & "'" & _
							"           and substr(referenceno, -3) = '" & arrRS(5, i) & "'" & _
							"           and abs(amount) > abs(" & Cdbl(arrRS(8, i)) & ")" & _
							"     order by lastupdate" & _
							" ) where rownum = 1"
				
				objectRS.Open objectSQL, connstr
				If Not objectRS.EOF Then
					first_fieldcode = objectRS("fieldcode")
					first_transdate = objectRS("transdate")
					first_jobcode = objectRS("jobcode")
					first_targettype = objectRS("targettype")
					first_targetcode = objectRS("targetcode")
				Else
					first_fieldcode = ""
					first_transdate = ""
					first_jobcode = ""
					first_targettype = ""
					first_targetcode = ""
				End If
				objectRS.close
				set objectRS = Nothing
				
				connstr.BeginTrans
				qInsertCR = " insert into loccb (fieldcode" & _
							"                   , targettype" & _
							"                   , targetcode" & _
							"                   , jobcode" & _
							"                   , quantity" & _
							"                   , amount" & _
							"                   , transdate" & _
							"                   , referenceno" & _
							"                   , fcentry" & _
							"                   , fcedit" & _
							"                   , fcip" & _
							"                   , fcba" & _
							"                   , lastupdate" & _
							"                   , lasttime" & _
							"                   )" & _
							" values ('" & first_fieldcode & "'" & _
							"         , 'CB'" & _
							"         , 'BS'" & _
							"         , '" & txtControlJob & "'" & _
							"         , '0'" & _
							"         , " & Cdbl(arrRS(8, i)) * -1 & "" & _
							"         , " & " to_date('" & first_transdate & "','mm/dd/yyyy')" & _
							"         , '" & first_fieldcode & "-Cr-" & arrRS(5, i) & "'" & _
							"         , '" & Session("CodeName") & "'" & _
							"         , '" & Session("CodeName") & "'" & _
							"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
							"         , '" & Session("fcBA") & "'" & _
							"         , sysdate" & _
							"         , to_char(sysdate, 'HH24:MI')" & _
							"         )"
				 'response.Write(qInsertCR)			  
				connStr.execute(qInsertCR)
				
				qInsertDB = " insert into loccb (fieldcode" & _
							"                   , targettype" & _
							"                   , targetcode" & _
							"                   , jobcode" & _
							"                   , quantity" & _
							"                   , amount" & _
							"                   , transdate" & _
							"                   , referenceno" & _
							"                   , fcentry" & _
							"                   , fcedit" & _
							"                   , fcip" & _
							"                   , fcba" & _
							"                   , lastupdate" & _
							"                   , lasttime" & _
							"                   )" & _
							" values ('" & first_fieldcode & "'" & _
							"         , '" & first_targettype & "'" & _
							"         , '" & first_targetcode & "'" & _
							"         , '" & first_jobcode & "'" & _
							"         , '0'" & _
							"         , " & Cdbl(arrRS(8, i)) & "" & _
							"         , " & " to_date('" & first_transdate & "','mm/dd/yyyy')" & _
							"         , '" & first_fieldcode & "-Db-" & arrRS(5, i) & "'" & _
							"         , '" & Session("CodeName") & "'" & _
							"         , '" & Session("CodeName") & "'" & _
							"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
							"         , '" & Session("fcBA") & "'" & _
							"         , sysdate" & _
							"         , to_char(sysdate, 'HH24:MI')" & _
							"         )"
				 'response.Write(qInsertDB)			  
				connStr.execute(qInsertDB)

				If Err.Number = 0 Then
					connstr.CommitTrans
					Response.Flush
				Else
					connstr.RollbackTrans
					Response.write("Error Save ["&qInsertCR&"] - ["&qInsertDB&"] ("&Err.Description&")")
					Response.Flush
				End If
			
			Next
			Erase arrRS
		End If

	else
		qSelisih = " select fcba" & _
				   "        , fieldcode" & _
				   "        , transdate" & _
				   "        , amount" & _
				   " from (" & _
				   "     select fcba" & _
				   "            , fieldcode" & _
				   "            , transdate" & _
				   "            , sum(amount) amount" & _
				   "     from loccb" & _
				   "     where fcba = '" & Session("fcBA") & "'" & _
				   "           and to_number(to_char(transdate,'yyyy')) = " & Year(Session("v_General_Charges")) & "" & _
				   "           and to_number(to_char(transdate,'mm')) = " & Month(Session("v_General_Charges")) & "" & _
				   "     group by fcba" & _
				   "              , fieldcode" & _
				   "              , transdate" & _
				   " )" & _
				   " where amount <> 0"
				   
		arrRS = GetRSArray(qSelisih)
		
		If IsArray(arrRS) Then
			For i = LBound(arrRS, 2) To UBound(arrRS, 2)
				Set objectRS = Server.CreateObject("ADODB.Recordset")
				objectSQL = " select fieldcode, transdate, jobcode, targettype, targetcode" & _
							" from loccb" & _
							" where fcba = '" & Session("fcBA") & "'" & _
							"       and transdate = to_date('" & arrRS(2, i) & "', 'mm/dd/yyyy')" & _
							"       and fieldcode = '" & arrRS(1, i) & "'" & _
							"       and jobcode <> '" & txtControlJob & "'" & _
							"       and rownum = 1"
				objectRS.Open objectSQL, connstr
				If Not objectRS.EOF Then
					first_fieldcode = objectRS("fieldcode")
					first_transdate = objectRS("transdate")
					first_jobcode = objectRS("jobcode")
					first_targettype = objectRS("targettype")
					first_targetcode = objectRS("targetcode")
				Else
					first_fieldcode = ""
					first_transdate = ""
					first_jobcode = ""
					first_targettype = ""
					first_targetcode = ""
				End If
				objectRS.close
				set objectRS = Nothing
				
				UpdateAmoutSelisih = " update loccb" & _
									 " set amount = amount - " & cdbl(arrRS(3, i)) & "" & _
									 " where fcba = '" & Session("fcBA") & "'" & _
									 "       and transdate = to_date('" & first_transdate & "', 'mm/dd/yyyy')" & _
									 "       and fieldcode = '" & first_fieldcode & "'" & _
									 "       and jobcode = '" & first_jobcode & "'" & _
									 "       and targettype = '" & first_targettype & "'" & _
									 "       and targetcode = '" & first_targetcode & "'"
									 
				connStr.execute(UpdateAmoutSelisih)
			Next
			Erase arrRS
		End If
	end if
	
	Response.write("<font face=Verdana size=1 color=Green> Normalisasi Rounding, Success!.<br/><br/></font>")
	Response.Flush
	
	
	If Session("Estate_Mill")="E" then
		'----- Tranfer to Activity Costing
		response.Write "Transfer to Activity Based Costing..." & "<br>"
		Response.Flush

		' Get the Max LineNo of Cost Activity
		Set rsCostActivityMax = Server.CreateObject("ADODB.Recordset")
		rsCostActivityMax.CursorType = 3
		rsCostActivityMax.CursorLocation = 3
										
		qCostActivityMax =  " select nvl(max(lineno), 0) maxLineNo from cost_activity " & _
							" where fcba = '" &  Session("fcBA") & "' " 
		rsCostActivityMax.Open qCostActivityMax, connstr
		
		If Not rsCostActivityMax.EOF Then
			maxLineNo = cdbl(rsCostActivityMax("maxLineNo"))
		Else
			maxLineNo = 0
		End If

		rsCostActivityMax.close
		set rsCostActivityMax = Nothing
		' End of Get the Max LineNo of Cost Activity
		
		qABC = " INSERT INTO COST_ACTIVITY (" & _
			   " TDATE, JOBCODE, SOURCE, SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, TARGETTYPE, TARGETCODE, TARGETCODE_STATUS, " & _
			   " REFERENCENO, QUANTITY, OUTPUT, AMOUNT, FCBA, CLOSING_DATE, PROCESS_DATE,  " & _
										" TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
										" TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
										" TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
										" TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
										" TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _	
										" LASTUPDATE, LASTTIME, LINENO ) " & _
			   " select TDATE, JOBCODE, SOURCE, SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, TARGETTYPE, TARGETCODE, TARGETCODE_STATUS, " & _
				  " REFERENCENO, QUANTITY, OUTPUT, AMOUNT, FCBA, CLOSING_DATE, sysdate PROCESS_DATE, " & _
						  " Cost_Activity_Info_Api.Get_Title_Reference_01 (SOURCETYPE,TARGETTYPE,decode ( GreaterOrEqualThan ( amount, 0), 'TRUE','Dr','Cr')),TARGETCODE, " & _
						  " Cost_Activity_Info_Api.Get_Title_Reference_02 (SOURCETYPE,TARGETTYPE,decode ( GreaterOrEqualThan ( amount, 0), 'TRUE','Dr','Cr')),referenceno, " & _
						  " '','', " & _
						  " '','', " & _
						  " '','', " & _
						  " decode ( GreaterOrEqualThan ( amount, 0), 'TRUE','Dr','Cr'), " & _
				  " sysdate LASTUPDATE, to_char(sysdate,'HH24:MI') LASTTIME, rownum + " & maxLineNo & " LINENO " & _
				  " FROM ( " & _
				  " SELECT TRANSDATE tdate, " & _
						   "JOBCODE jobcode, " & _ 
						   "'LOC' source, " & _
						   "'LOC' sourcetype, " & _
						   "FIELDCODE sourcecode, " & _
						   "'' sourcecode_status, " & _
						   "TARGETTYPE targettype, " & _
						   "TARGETCODE targetcode, " & _
						   "'' targetcode_status, " & _
						   "REFERENCENO referenceno, " & _
						   "SUM(QUANTITY) quantity, " & _
						   "0 output, " & _
						   "SUM(AMOUNT) amount, " & _ 
						   "FCBA," & _
						   "to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') closing_date " & _
				  " FROM LOCCB" & _
				  " WHERE fcBA='" & Session("fcBA") & "'" & _
				  "       and to_number(to_Char(TRANSDATE,'yyyy')) = " & Year(Session("v_General_Charges")) & "" & _
				  "       and to_number(to_Char(TRANSDATE,'mm')) = " & Month(Session("v_General_Charges")) & "" & _
				  " group by TRANSDATE, JOBCODE, 'LOC', 'LOC', FIELDCODE, '', TARGETTYPE, TARGETCODE, '', REFERENCENO,  FCBA)"
		'response.Write " qABC "& qABC
		response.Flush()		  
		connstr.Execute(qABC)
		
		Response.write("<font face=Verdana size=1 color=Green>Transfer to Activity Based Costing..., Success!.<br/><br/></font>")
		Response.Flush
	
	else
	
		Set objectRS = Server.CreateObject("ADODB.Recordset")
		objectSQL = " select nvl(sum(costtotal), 0) costtotal" & _
					"        , nvl(sum(alloc_db), 0) alloc_db" & _
					"        , nvl(sum(alloc_cr), 0)*-1 alloc_cr" & _
					" from (" & _
					"     select sum(costtotal) costtotal, 0 alloc_db , 0 alloc_cr" & _
					"     from costtotal_pertarget" & _
					"     where fcba = '" & Session("fcBA") & "'" & _
					"           and year = " & Year(Session("v_General_Charges")) & "" & _
					"           and month = " & Month(Session("v_General_Charges")) & "" & _
					"           and targettype = 'LOC'" & _
					"     union all" & _
					"     select 0 costtotal, sum(amount) alloc_db , 0 alloc_cr" & _
					"     from loccb" & _
					"     where fcba = '" & Session("fcBA") & "'" & _
					"           and to_number(to_char(transdate, 'yyyy')) = " & Year(Session("v_General_Charges")) & "" & _
					"           and to_number(to_char(transdate, 'mm')) = " & Month(Session("v_General_Charges")) & "" & _
					"           and jobcode <> '" & txtControlJob & "'" & _
					"     union all" & _
					"     select 0 costtotal,0 alloc_db ,  sum(amount) alloc_cr" & _
					"     from loccb" & _
					"     where fcba = '" & Session("fcBA") & "'" & _
					"           and to_number(to_char(transdate, 'yyyy')) = " & Year(Session("v_General_Charges")) & "" & _
					"           and to_number(to_char(transdate, 'mm')) = " & Month(Session("v_General_Charges")) & "" & _
					"           and jobcode = '" & txtControlJob & "'" & _
					" )"

			objectRS.Open objectSQL, connstr
			If Not objectRS.EOF Then
				Cost_In = Cdbl(objectRS("costtotal"))
				Alloc_Db = Cdbl(objectRS("alloc_db"))
				Alloc_Cr = Cdbl(objectRS("alloc_cr"))
			Else
			End If
			objectRS.close
			set objectRS = Nothing

			If Cost_In <> Alloc_Db Then
				err.number = 1001
				err.description = "Alokasi sisi debet tidak sesuai dengan Biaya yang Masuk! Cek Kembali Biaya yang masuk ke target LOC"
			ElseIf Cost_In <> Alloc_Cr Then
				err.number = 1001
				err.description = "Alokasi sisi credit tidak sesuai dengan Biaya yang Masuk! Cek Kembali Biaya yang masuk ke target LOC"
			ElseIf (Cost_In = Alloc_Db) And (Cost_In = Alloc_Cr) Then
				'----- Tranfer to Activity Costing
				response.Write "Transfer to Activity Based Costing..." & "<br>"
				Response.Flush

				' Get the Max LineNo of Cost Activity
				Set rsCostActivityMax = Server.CreateObject("ADODB.Recordset")
				rsCostActivityMax.CursorType = 3
				rsCostActivityMax.CursorLocation = 3
												
				qCostActivityMax =  " select nvl(max(lineno), 0) maxLineNo from cost_activity " & _
									" where fcba = '" &  Session("fcBA") & "' " 
				rsCostActivityMax.Open qCostActivityMax, connstr
				
				If Not rsCostActivityMax.EOF Then
					maxLineNo = cdbl(rsCostActivityMax("maxLineNo"))
				Else
					maxLineNo = 0
				End If

				rsCostActivityMax.close
				set rsCostActivityMax = Nothing
				' End of Get the Max LineNo of Cost Activity
				
				qABC = " INSERT INTO COST_ACTIVITY (" & _
					   " TDATE, JOBCODE, SOURCE, SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, TARGETTYPE, TARGETCODE, TARGETCODE_STATUS, " & _
					   " REFERENCENO, QUANTITY, OUTPUT, AMOUNT, FCBA, CLOSING_DATE, PROCESS_DATE,  " & _
												" TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
												" TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
												" TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
												" TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
												" TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _	
												" LASTUPDATE, LASTTIME, LINENO ) " & _
					   " select TDATE, JOBCODE, SOURCE, SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, TARGETTYPE, TARGETCODE, TARGETCODE_STATUS, " & _
						  " REFERENCENO, QUANTITY, OUTPUT, AMOUNT, FCBA, CLOSING_DATE, sysdate PROCESS_DATE, " & _
								  " Cost_Activity_Info_Api.Get_Title_Reference_01 (SOURCETYPE,TARGETTYPE,decode ( GreaterOrEqualThan ( amount, 0), 'TRUE','Dr','Cr')),TARGETCODE, " & _
								  " Cost_Activity_Info_Api.Get_Title_Reference_02 (SOURCETYPE,TARGETTYPE,decode ( GreaterOrEqualThan ( amount, 0), 'TRUE','Dr','Cr')),referenceno, " & _
								  " '','', " & _
								  " '','', " & _
								  " '','', " & _
								  " decode ( GreaterOrEqualThan ( amount, 0), 'TRUE','Dr','Cr'), " & _
						  " sysdate LASTUPDATE, to_char(sysdate,'HH24:MI') LASTTIME, rownum + " & maxLineNo & " LINENO " & _
						  " FROM ( " & _
						  " SELECT TRANSDATE tdate, " & _
								   "JOBCODE jobcode, " & _ 
								   "'LOC' source, " & _
								   "'LOC' sourcetype, " & _
								   "FIELDCODE sourcecode, " & _
								   "'' sourcecode_status, " & _
								   "TARGETTYPE targettype, " & _
								   "TARGETCODE targetcode, " & _
								   "'' targetcode_status, " & _
								   "REFERENCENO referenceno, " & _
								   "SUM(QUANTITY) quantity, " & _
								   "0 output, " & _
								   "SUM(AMOUNT) amount, " & _ 
								   "FCBA," & _
								   "to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') closing_date " & _
						  " FROM LOCCB" & _
						  " WHERE fcBA='" & Session("fcBA") & "'" & _
						  "       and to_number(to_Char(TRANSDATE,'yyyy')) = " & Year(Session("v_General_Charges")) & "" & _
						  "       and to_number(to_Char(TRANSDATE,'mm')) = " & Month(Session("v_General_Charges")) & "" & _
						  " group by TRANSDATE, JOBCODE, 'LOC', 'LOC', FIELDCODE, '', TARGETTYPE, TARGETCODE, '', REFERENCENO,  FCBA)"
				'response.Write " qABC "& qABC
				response.Flush()		  
				connstr.Execute(qABC)
				
				Response.write("<font face=Verdana size=1 color=Green>Transfer to Activity Based Costing..., Success!.<br/><br/></font>")
				Response.Flush
			End If
		end if

End If
%>
<%
	Response.write "********** LOC Closing ***********<br/><br/>"
	Response.Flush
%>

