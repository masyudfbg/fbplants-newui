<%
	Option Explicit
	response.buffer = True
%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>ePlaS - Building Closing</title>
</head>
<body>
<%
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
Dim TimeAwal,TimeAkhir,Hitung,Lama
TimeAwal=Time()


Dim rsVehicle, qVehicle, nxActivity, nxGaji, nCostTotal
Dim nTotalAlokasi, rsCostVehicle, qCostVehicle, nBasis, rsBasis, qBasis, qInsertCR
Dim rsVHActivity, qVHActivity, nWSAlokasi, qInsertDB, LT, LC, FDDATE, FCJOB, nSelisih
Dim qABC, q, qLineMax, rsqLineMax, tempMaxLine
Dim strClosingLog, strLenghtime, rsLenghtime, upClosingLog, txtLastDay, txtNextMonth
Dim maxLineNoVehicleCB, tempProcessCode, tempProcessNameCode, tempRefCodePVDetail, qReference
Dim rsReferencePVDetail, donePVDetail, qNormalisasiDetail, qVoucher, rsVoucher, qCost_Activity,qjournal
Dim qVoucherDetail, rsVoucherDetail, maxLineNoDetail
Dim txtRefStringSeparator_PVDetail, txtRefStringValues_PVDetail, nQuantity, nOutput
Dim txtRefNumberKeys_PVDetail, txtRefStringKeys_PVDetail, qCost_ActivityHistory, maxLineNoContractorCB
Dim rsContract, qContract, rsAggrement, qCostContract, nRate, maxLineNoBuildingCB, rsCostActivityMax, qCostActivityMax
Dim rsBuilding, qBuilding, rsCostBuilding, qCostBuilding, rsFieldAlokasi, qFieldAlokasi
Dim nTMorTBM, nFFAlokasi, maxLineNo
Dim temp_DebetCredit

Dim arrRS, objRSArray, ExecuteSQL, ExecuteSQLCost, qDebetAlokasi, qCreditAlokasi, i, jobcodeNoBasis
Dim qSelisih, objectSQL, objectRS, FirstRecord, UpdateAmoutSelisih
Dim Cost_In, Alloc_Db, Alloc_Cr

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
Building Last Period Close <%=FormatDateTime(Session("cls_Building"),1)%> <br> 
Building Current Period is <%=FormatDateTime(Session("v_Building"),1)%> <br><br>
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

qWaterFlowCheck = " select PERIODCONTROL_API.getStatusPrevClosing('Building','" & Session("fcBA")  & "') waterFlowCheck from dual "

Set rsWaterFlowCheck = Server.CreateObject("ADODB.Recordset")
rsWaterFlowCheck.CursorType = 3
rsWaterFlowCheck.CursorLocation = 3
rsWaterFlowCheck.Open qWaterFlowCheck, connstr

if not rsWaterFlowCheck.eof then
    CheckWaterFlow = rsWaterFlowCheck("waterFlowCheck")
else
    Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error, No WATER FLOW CHECK for Building " & "<br>"
end if
rsWaterFlowCheck.close
set rsWaterFlowCheck = Nothing


if CheckWaterFlow = "TRUE" then 

	

	qPeriodControl=" select  add_months(LASTPERIODPOSTED,1) LastPeriodPosted from PERIODCONTROL  " & _
	  " where SYSTEM='Building' and fcba='" & Session("fcBA") & "'"

	Set rsPeriodControl = Server.CreateObject("ADODB.Recordset")
	rsPeriodControl.CursorType = 3
	rsPeriodControl.CursorLocation = 3
	rsPeriodControl.Open qPeriodControl, connstr

	if not rsPeriodControl.eof then
		temp_datePeriodControl = rsPeriodControl("LASTPERIODPOSTED")
	else
		Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error, No PERIODCONTROL for Building " & "<br>"
	end if
	' End Added
	'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
	

	' Remarked by NAF, as per New Concept, 7 May 2013
	' For Workshop Closing, for 2nd Book in Accounting

	tempProcessCode = "BUILDING_CLOSING"
	tempProcessNameCode = "p_Building_Closing.asp"
	tempRefCodePVDetail = "BUILDINGCB"

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
									
	qCostActivityMax =  " select nvl(max(nvl(lineno, 0)), 0) maxLineNo from BuildingCB " & _
						" where fcba = '" &  Session("fcBA") & "' " 
	rsCostActivityMax.Open qCostActivityMax, connstr
	
	If Not rsCostActivityMax.EOF Then
		maxLineNoBuildingCB = cdbl(rsCostActivityMax("maxLineNo"))
	Else
		maxLineNoBuildingCB = 0
	End If
	
	rsCostActivityMax.close
	set rsCostActivityMax = Nothing
	' End of Get the Max LineNo of Cost Activity	
	

	'---- Control Job
	If Session("Estate_Mill") <> "HO" Then 
		qControlJob="select ControlJob from Module where NAME='Building'"
		Set rsControlJob = Server.CreateObject("ADODB.Recordset")
		rsControlJob.CursorType = 3
		rsControlJob.CursorLocation = 3
		rsControlJob.Open qControlJob, connstr
		if not rsControlJob.eof then
		   txtControlJob=rsControlJob("ControlJob")
		else
		   Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error SAC200366 Control Job For Building is Empty"
		   txtControlJob=""
		end if
		rsControlJob.close
		set rsControlJob = Nothing
	else
		qControlJob="select ControlJob from Module where NAME='Building_HO'"
		Set rsControlJob = Server.CreateObject("ADODB.Recordset")
		rsControlJob.CursorType = 3
		rsControlJob.CursorLocation = 3
		rsControlJob.Open qControlJob, connstr
		if not rsControlJob.eof then
		   txtControlJob=rsControlJob("ControlJob")
		else
		   Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error SAC200366 Control Job For Building is Empty"
		   txtControlJob=""
		end if
		rsControlJob.close
		set rsControlJob = Nothing
	End If
	
	'---- End Control Job
	'response.write txtControlJob
	
	response.Flush()

	if txtControlJob<>"" then
		'---- Normalisasi
		Response.Write " Normalisasi Building Base Costing" & "<br>"
		Response.Flush()

		connstr.BeginTrans
		qNormalisasi = " delete" & _
		               " from buildingcb" & _
		               " where fcba = '" & Session("fcBA") & "'" & _
		               "       and to_number(to_char(transdate,'yyyy')) = " & Year(Session("v_Building")) & "" & _
                       "       and to_number(to_char(transdate,'mm')) = " & Month(Session("v_Building")) & ""
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
									' " where  to_number(to_Char(TDATE,'mm'))=" & Month(Session("v_Building")) &  _
									' " and to_number(to_Char(TDATE,'yyyy'))=" & Year(Session("v_Building")) & " and fcBA='" & Session("fcBA") & "' and SOURCETYPE='BM')"
							 
		' connstr.execute(qNormalisasiDetail)			
		
		' connStr.execute("Commit")
		
		' qNormalisasi=" Delete From COST_ACTIVITY where  to_number(to_Char(TDATE,'mm'))=" & Month(Session("v_Building")) &  _
					 ' " and to_number(to_Char(TDATE,'yyyy'))=" & Year(Session("v_Building")) & " and fcBA='" & Session("fcBA") & "'" & _
					 ' " and SOURCETYPE = 'BM'"
		' connstr.execute(qNormalisasi)	
		
		' connStr.execute("Commit")
		
		connstr.BeginTrans
		qNormalisasiDetail = " update cost_activity_hist_tab set deleteflag = 'REMOVE'" & _
							 " where " & _
							 " lineheaderno || fcba in (select lineno||fcba from cost_activity " & _
									" where  to_number(to_Char(TDATE,'mm'))=" & Month(Session("v_Building")) &  _
									" and to_number(to_Char(TDATE,'yyyy'))=" & Year(Session("v_Building")) & " and fcBA='" & Session("fcBA") & "' and SOURCETYPE='BM')"
							 
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
		qNormalisasiDetail="   update Cost_Activity set deleteflag = 'REMOVE' where  to_number(to_Char(TDATE,'mm'))=" & Month(Session("v_Building")) &  _
					 " and to_number(to_Char(TDATE,'yyyy'))=" & Year(Session("v_Building")) & " and fcBA='" & Session("fcBA") & "'" & _
					 " and SOURCETYPE = 'BM'"
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
		
        Response.write("<font face=Verdana size=1 color=Green>Normalisasi Building Base Costing, Success!.<br/><br/></font>")
        Response.Flush
		
		Response.Write "Generating Data..." & "<br/>"
		Response.Flush
		
		connstr.BeginTrans
		ExecuteSQL = " delete from costtotal_pertarget" & _
		             " where fcba = '" & Session("fcBA") & "'" & _
					 "       and year = " & Year(Session("v_Building")) & "" & _
					 "       and month = " & Month(Session("v_Building")) & "" & _
					 "       and targettype = 'BM'"
					 
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
		ExecuteSQLCost = " insert into costtotal_pertarget" & _
		             " (FCBA, YEAR, MONTH, TARGETTYPE, TARGETCODE, COSTTOTAL, BASIS)" & _
					 " select fcba" & _
					 "        , year" & _
					 "        , month" & _
					 "        , targettype" & _
					 "        , targetcode" & _
					 "        , costtotal" & _
					 "        , GC_CLOSING_API.GET_BASIS(fcba, year, month) basis" & _
					 " from (" & _
					 "     select fcba" & _
					 "            , to_number(to_char(tdate,'yyyy')) year" & _
					 "            , to_number(to_char(tdate,'mm')) month" & _
					 "            , targettype" & _
					 "            , targetcode" & _
					 "            , nvl(sum(amount),0) as costtotal" & _
					 "     from cost_activity" & _
					 "     where fcba = '" & Session("fcBA") & "'" & _
					 "           and to_number(to_char(tdate,'yyyy')) = " & Year(Session("v_Building")) & "" & _
					 "           and to_number(to_char(tdate,'mm')) = " & Month(Session("v_Building")) & "" & _
					 "           and targettype = 'BM'" & _
					 "     group by fcba" & _
					 "              , to_number(to_char(tdate,'yyyy'))" & _
					 "              , to_number(to_char(tdate,'mm'))" & _
					 "              , targettype" & _
					 "              , targetcode" & _
					 " )"
		
		'response.write ExecuteSQLCost
		connstr.execute(ExecuteSQLCost)
		
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
                         "        , decode(status, 'TM', '6510101', 'TBM', '1322109', '1150901') jobcode" & _
                         "        , 'FF' targettype" & _
                         "        , targetcode" & _
                         "        , unit" & _
						 "        , 1 nfactor" & _
						 "        , 0 basis" & _
                         " from (" & _
                         "     select fcba" & _
                         "            , fccode targetcode" & _
                         "            , field_api.getvalid_status (fccode, fcba, to_date(" & Year(Session("v_Building")) & "||lpad(" & Month(Session("v_Building")) & ",2,'0')||'01','yyyymmdd')) status" & _
                         "            , nvl(field_api.getvalid_ha (fccode, fcba, to_date(" & Year(Session("v_Building")) & "||lpad(" & Month(Session("v_Building")) & ",2,'0')||'01','yyyymmdd')), 0) unit" & _
                         "     from field" & _
                         "     where fcba = '" & Session("fcBA") & "'" & _
                         "           and nvl(field_api.getvalid_activation (fccode, fcba, to_date(" & Year(Session("v_Building")) & "||lpad(" & Month(Session("v_Building")) & ",2,'0')||'01','yyyymmdd')),'N') = 'Y'" & _
                         "           and field_api.getvalid_status (fccode, fcba, to_date(" & Year(Session("v_Building")) & "||lpad(" & Month(Session("v_Building")) & ",2,'0')||'01','yyyymmdd')) is not null" & _
                         " )" & _
                         " where unit <> 0"
		Elseif (Session("Estate_Mill") = "M" Or Session("Estate_Mill") = "P") then
			ExecuteSQL = " insert into btl_allocation" & _
                         " (fcba, jobcode, targettype, targetcode, unit, nfactor, basis)" & _
                         " select fcba" & _
                         "        , '6620101' jobcode" & _
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
						 "                   , 'E', nvl(field_api.GetValid_HaEstate(aa.codefcba,to_date(" & Year(Session("v_Building")) & "||lpad(" & Month(Session("v_Building")) & ",2,'0')||'01','YYYYMMDD')), 0)" & _
						 "                   , 0" & _
						 "                 ) unit" & _
						 "          , nfactor" & _
						 "          , nvl(totalbasis, 0) basis" & _
						 "     from jobcodeintraco aa" & _
						 "     left join ( Select fcba, fctype, Sum(basis) totalbasis" & _
						 "                 from ( select fcba" & _
						 "                               , 'E' fctype" & _
						 "                               , nvl(field_api.GetValid_HaEstate(a.codefcba,to_date(" & Year(Session("v_Building")) & "||lpad(" & Month(Session("v_Building")) & ",2,'0')||'01','YYYYMMDD')),0) basis" & _
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
			jobcodeNoBasis = "1150901"
		elseif Session("Estate_Mill") = "HO" then
			jobcodeNoBasis = "3210101"
		Else
			jobcodeNoBasis = "1150901"
		
		End If
		
		qCreditAlokasi = " select fcba" & _
                         "        , targettype" & _
                         "        , targetcode" & _
                         "        , costtotal" & _
						 "        , basis" & _
                         " from costtotal_pertarget" & _
                         " where fcba = '" & Session("fcBA") & "'" & _
                         "       and year = " & Year(Session("v_Building")) & "" & _
                         "       and month = " & Month(Session("v_Building")) & "" & _
                         "       and targettype = 'BM'"
		
		Set arrRS = Server.CreateObject("ADODB.Recordset")
		arrRS.Open qCreditAlokasi, connstr
		if not arrRS.eof then
			Do While Not arrRS.eof
				'response.write "masuk"&"<br>"
				If Cdbl(arrRS("BASIS")) = 0 Then
					connstr.BeginTrans
					maxLineNobuildingcb = maxLineNobuildingcb + 1
					qInsertCR = " insert into buildingcb (fieldcode" & _
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
								" values ('" & arrRS("TARGETCODE") & "'" & _
								"         , 'CB'" & _
								"         , 'BS'" & _
								"         , '" & txtControlJob & "'" & _
								"         , '0'" & _
								"         , " & cdbl(arrRS("costtotal")) * -1 & "" & _
								"         , " & " to_date('" & Session("v_Building") & "','mm/dd/yyyy')" & _
								"         , '" & arrRS("TARGETCODE") & "-Cr-No Basis" & "'" & _
								"         , '" & Session("CodeName") & "'" & _
								"         , '" & Session("CodeName") & "'" & _
								"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
								"         , '" & Session("fcBA") & "'" & _
								"         , sysdate" & _
								"         , to_char(sysdate, 'HH24:MI')" & _
								"         , " & maxLineNobuildingcb & ")"
					'response.Write cdbl(arrRS(3, i))&"<br>"			  
					connStr.execute(qInsertCR)
					
					maxLineNobuildingcb = maxLineNobuildingcb + 1
					qInsertDB = " insert into buildingcb (fieldcode" & _
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
								" values ('" & arrRS("TARGETCODE") & "'" & _
								"         , 'CB'" & _
								"         , 'BS'" & _
								"         , '" & jobcodeNoBasis & "'" & _
								"         , '0'" & _
								"         , " & cdbl(arrRS("COSTTOTAL")) & "" & _
								"         , " & " to_date('" & Session("v_Building") & "','mm/dd/yyyy')" & _
								"         , '" & arrRS("TARGETCODE") & "-Db-No Basis" & "'" & _
								"         , '" & Session("CodeName") & "'" & _
								"         , '" & Session("CodeName") & "'" & _
								"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
								"         , '" & Session("fcBA") & "'" & _
								"         , sysdate" & _
								"         , to_char(sysdate, 'HH24:MI')" & _
								"         , " & maxLineNobuildingcb & ")"
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
					maxLineNobuildingcb = maxLineNobuildingcb + 1
					qInsertCR = " insert into buildingcb (fieldcode" & _
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
								" values ('" & arrRS("TARGETCODE") & "'" & _
								"         , 'CB'" & _
								"         , 'BS'" & _
								"         , '" & txtControlJob & "'" & _
								"         , '0'" & _
								"         , " & cdbl(arrRS("COSTTOTAL")) * -1 & "" & _
								"         , " & " to_date('" & Session("v_Building") & "','mm/dd/yyyy')" & _
								"         , '" & arrRS("TARGETCODE") & "-Cr" & "'" & _
								"         , '" & Session("CodeName") & "'" & _
								"         , '" & Session("CodeName") & "'" & _
								"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
								"         , '" & Session("fcBA") & "'" & _
								"         , sysdate" & _
								"         , to_char(sysdate, 'HH24:MI')" & _
								"         , " & maxLineNobuildingcb & ")"
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
				arrRS.MoveNext
			loop
		end if
		
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
							"       and ca.year = " & Year(Session("v_Building")) & "" & _
							"       and ca.month = " & Month(Session("v_Building")) & "" & _
							"       and ca.targettype = 'BM'"
							
		ElseIf Session("Estate_Mill") = "HO" then
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
							"       and ca.year = " & Year(Session("v_Building")) & "" & _
							"       and ca.month = " & Month(Session("v_Building")) & "" & _
							"       and ca.targettype = 'BM'"
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
							"       and ca.year = " & Year(Session("v_Building")) & "" & _
							"       and ca.month = " & Month(Session("v_Building")) & "" & _
							"       and ca.targettype = 'BM'"
		
		Else
			qDebetAlokasi = " "
		End If
		
		arrRS = GetRSArray(qDebetAlokasi)
		
		If IsArray(arrRS) Then
			For i = LBound(arrRS, 2) To UBound(arrRS, 2)
				If Cdbl(arrRS(6, i)) <> 0 Then
					nFFAlokasi = Round(Cdbl(arrRS(9, i))/Cdbl(arrRS(6, i))*Cdbl(arrRS(5, i))*Cdbl(arrRS(8, i)), 0)
					If Cdbl(nFFAlokasi) <> 0 then
						connstr.BeginTrans
						maxLineNobuildingcb = maxLineNobuildingcb + 1
						qInsertDB = " insert into buildingcb (fieldcode" & _
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
									"         , " & " to_date('" & Session("v_Building") & "','mm/dd/yyyy')" & _
									"         , '" & arrRS(4, i) & "-Db" & "'" & _
									"         , '" & Session("CodeName") & "'" & _
									"         , '" & Session("CodeName") & "'" & _
									"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
									"         , '" & Session("fcBA") & "'" & _
									"         , sysdate" & _
									"         , to_char(sysdate, 'HH24:MI')" & _
									"         , " & maxLineNobuildingcb & ")"
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
				   "     from buildingcb" & _
				   "     where fcba = '" & Session("fcBA") & "'" & _
				   "           and to_number(to_char(transdate,'yyyy')) = " & Year(Session("v_Building")) & "" & _
				   "           and to_number(to_char(transdate,'mm')) = " & Month(Session("v_Building")) & "" & _
				   "     group by fcba" & _
				   "              , fieldcode" & _
				   "              , transdate" & _
				   " )" & _
				   " where amount <> 0"
				   
		arrRS = GetRSArray(qSelisih)
		If IsArray(arrRS) Then
			For i = LBound(arrRS, 2) To UBound(arrRS, 2)
				Set objectRS = Server.CreateObject("ADODB.Recordset")
				objectSQL = " select nvl(min (lineno), 0.0001) minlineno" & _
							" from buildingcb" & _
							" where fcba = '" & Session("fcBA") & "'" & _
							"       and transdate = to_date('" & arrRS(2, i) & "', 'mm/dd/yyyy')" & _
							"       and fieldcode = '" & arrRS(1, i) & "'" & _
							"       and jobcode <> '" & txtControlJob & "'"
				objectRS.Open objectSQL, connstr
				If Not objectRS.EOF Then
					FirstRecord = Cdbl(objectRS("minlineno"))
				Else
					FirstRecord = 0.0001
				End If
				objectRS.close
				set objectRS = Nothing
				
				if FirstRecord = 0.0001 then
					maxLineNobuildingcb = maxLineNobuildingcb + 1
					qInsertDB = " INSERT INTO buildingcb (fieldcode" & _
								"                         , targettype" & _
								"                         , targetcode" & _
								"                         , jobcode" & _
								"                         , quantity" & _
								"                         , amount" & _
								"                         , transdate" & _
								"                         , referenceno" & _
								"                         , fcentry" & _
								"                         , fcedit" & _
								"                         , fcip" & _
								"                         , fcba" & _
								"                         , lastupdate" & _
								"                         , lasttime" & _
								"                         , lineno)" & _
								" VALUES ('" & arrRS(1, i) & "'" & _
								"         , 'CB'" & _
								"         , 'BS'" & _
								"         , '1150901'" & _
								"         , '0'" & _
								"         , ABS(" & cdbl(arrRS(3, i)) & ")" & _
								"         , " & " TO_DATE('" & arrRS(2, i) & "','mm/dd/yyyy')" & _
								"         , '" & arrRS(1, i) & "-Db" & "'" & _
								"         , '" & Session("CodeName") & "'" & _
								"         , '" & Session("CodeName") & "'" & _
								"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
								"         , '" & Session("fcBA") & "'" & _
								"         , SYSDATE" & _
								"         , TO_CHAR(SYSDATE, 'HH24:MI')" & _
								"         , " & maxLineNobuildingcb & ")"
					 'response.Write(qInsertDB)			  
					connStr.execute(qInsertDB)
				else
				
					UpdateAmoutSelisih = " update buildingcb" & _
										 " set amount = amount - " & cdbl(arrRS(3, i)) & "" & _
										 " where lineno = '" & FirstRecord & "'" & _
										 "       and fcba = '" & Session("fcBA") & "'"
					connStr.execute(UpdateAmoutSelisih)
				end if
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
					"           and year = " & Year(Session("v_Building")) & "" & _
					"           and month = " & Month(Session("v_Building")) & "" & _
					"           and targettype = 'BM'" & _
					"     union all" & _
					"     select 0 costtotal, sum(amount) alloc_db , 0 alloc_cr" & _
					"     from buildingcb" & _
					"     where fcba = '" & Session("fcBA") & "'" & _
					"           and to_number(to_char(transdate, 'yyyy')) = " & Year(Session("v_Building")) & "" & _
					"           and to_number(to_char(transdate, 'mm')) = " & Month(Session("v_Building")) & "" & _
					"           and jobcode <> '" & txtControlJob & "'" & _
					"     union all" & _
					"     select 0 costtotal,0 alloc_db ,  sum(amount) alloc_cr" & _
					"     from buildingcb" & _
					"     where fcba = '" & Session("fcBA") & "'" & _
					"           and to_number(to_char(transdate, 'yyyy')) = " & Year(Session("v_Building")) & "" & _
					"           and to_number(to_char(transdate, 'mm')) = " & Month(Session("v_Building")) & "" & _
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
				err.description = "Alokasi sisi debet tidak sesuai dengan Biaya yang Masuk! Cek Kembali Biaya yang masuk ke target BM"
			ElseIf Cost_In <> Alloc_Cr Then
				err.number = 1001
				err.description = "Alokasi sisi credit tidak sesuai dengan Biaya yang Masuk! Cek Kembali Biaya yang masuk ke target BM"
			ElseIf (Cost_In = Alloc_Db) And (Cost_In = Alloc_Cr) Then
				'----- Tranfer to Activity Costing
				response.Write "Transfer to Activity Based Costing..." & "<br>"
				Response.Flush()

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
							   "        , 'BM' source" & _
							   "        , 'BM' sourcetype" & _
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
							   " from buildingcb" & _
							   " where fcba = '" & Session("fcBA") & "'" & _
							   "       and to_number(to_char(transdate,'yyyy')) = " & Year(Session("v_Building")) & "" & _
							   "       and to_number(to_char(transdate,'mm')) = " & Month(Session("v_Building")) & "" & _
							   " group by transdate" & _
							   "          , jobcode" & _
							   "          , 'BM'" & _
							   "          , 'BM'" & _
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
												" LASTUPDATE, LASTTIME, LINENO,journalprefix ) " & _
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
												" '','', " & _
												" '','', " & _
												" '','', " & _
												" '','','" & temp_DebetCredit & "', " & _
												Month(Session("v_Building")) & ",  " & _
												Year(Session("v_Building")) & " ,  " & _
												" sysdate, to_char(sysdate,'HH24:MI'), " & maxLineNo & " ,'BLD') "
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
					
					qjournal=	" MERGE INTO cost_activity x " & _
					" USING " & _
					" ( " & _
					"  SELECT referenceno,journalprefix,idx||LPAD(ROWNUM,4,0) seq from " & _
					"  ( " & _
					"  SELECT DISTINCT referenceno,journalprefix,'" & session("COMPANYNO") & "/'||fcba||'/'||journalprefix||'/'||LPad(accountingperiod,2,0)||'/'||accountingyear||'/' as idx  " & _
					"  FROM cost_activity WHERE fcba='" & Session("FCBA") & "' AND accountingperiod="& Month(Session("v_Building")) & " AND accountingyear=" & Year(Session("v_Building")) & " AND journalprefix='BLD' " & _
					"  Order BY referenceno " & _
					"  ) " & _
					" )y ON (x.referenceno =  y.referenceno and x.journalprefix=y.journalprefix) " & _
					" WHEN MATCHED THEN " & _
					" UPDATE SET x.journalno = y.seq " 
					connstr.execute(qjournal)
		
					
				Response.write("<font face=Verdana size=1 color=Green>Transfer to Activity Based Costing..., Success!.<br/><br/></font>")
				Response.Flush
			End If



	end if
				  
	%>

	<%if err.number <> 0 then%>
		<br><img src="images/error.gif">
		<%
		
		connstr.BeginTrans
		ExecuteSQL = " delete from costtotal_pertarget" & _
		             " where fcba = '" & Session("fcBA") & "'" & _
					 "       and year = " & Year(Session("v_Building")) & "" & _
					 "       and month = " & Month(Session("v_Building")) & "" & _
					 "       and targettype = 'BM'"
					 
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
		
		response.Write  "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" & err.description & "</font>"
		%>
		<script language="JavaScript" type="text/javascript" >
		document.getElementById('status_running').style.visibility = 'hidden';
		</script>	
	<%Else
	
		connstr.BeginTrans
		ExecuteSQL = " delete from costtotal_pertarget" & _
		             " where fcba = '" & Session("fcBA") & "'" & _
					 "       and year = " & Year(Session("v_Building")) & "" & _
					 "       and month = " & Month(Session("v_Building")) & "" & _
					 "       and targettype = 'BM'"
					 
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

		
		' Add By Akbar 18 Desemeber 2015
		
			Dim QAccountPeriod,rsAccountPeriod,temp_NextCloseAccountingPeriod, temp_NextCloseAccountingYear, temp_NextCurrentAccountingPeriod, temp_NextCurrentAccountingYear
				
			QAccountPeriod	=	" select periodcontrol_api.GetNextCloseAccountingPeriod(system,fcba) NextCloseAccountingPeriod ,  " & _ 
								" periodcontrol_api.GetNextCloseAccountingYear(system,fcba) NextCloseAccountingYear ,  " & _ 
								" periodcontrol_api.GetNextCurrentAccountingPeriod(system,fcba) NextCurrentAccountingPeriod ,  " & _ 
								" periodcontrol_api.GetNextCurrentAccountingYear(system,fcba) NextCurrentAccountingYear  " & _ 
								" from periodcontrol p  " & _ 
								" where UPPER(system) ='BUILDING' and fcba='" & Session("fcBA") & "'"
			Set rsAccountPeriod = Server.CreateObject("ADODB.Recordset")
				rsAccountPeriod.CursorType = 3
				rsAccountPeriod.CursorLocation = 3
				rsAccountPeriod.Open QAccountPeriod, connstr
				if not rsAccountPeriod.eof then
					temp_NextCloseAccountingPeriod =  Cint(rsAccountPeriod("NextCloseAccountingPeriod"))
					temp_NextCloseAccountingYear   =  Cint(rsAccountPeriod("NextCloseAccountingYear"))
					temp_NextCurrentAccountingPeriod = Cint(rsAccountPeriod("NextCurrentAccountingPeriod"))
					temp_NextCurrentAccountingYear = Cint(rsAccountPeriod("NextCurrentAccountingYear"))
				else
					Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error, No PERIODCONTROL for BUILDING " & "<br>"
				end if
			rsAccountPeriod.close
			set rsAccountPeriod = Nothing		
			
			IF temp_NextCurrentAccountingPeriod < 12   THEN
			
				q= " Update PeriodControl Set LASTPERIODPOSTED=last_day(add_months(LASTPERIODPOSTED,1)),CURRENTDATE=add_months(CURRENTDATE,1) " & _ 
				   " ,CloseAccountingPeriod = " & temp_NextCloseAccountingPeriod & " " & _  
				   " ,CloseAccountingYear = " & temp_NextCloseAccountingYear & " " & _  
				   " ,CurrentAccountingPeriod = " & temp_NextCurrentAccountingPeriod & " " & _  
				   " ,CurrentAccountingYear = " & temp_NextCurrentAccountingYear & " " & _    
				   " ,lastupdate = sysdate, lasttime = to_char(sysdate, 'hh24:mi') " & _
				   " where UPPER(system) ='BUILDING' and fcba='" & Session("fcBA") & "'"
			ELSE
			    IF temp_NextCloseAccountingPeriod < 12 THEN
				    q=" Update PeriodControl Set LASTPERIODPOSTED=last_day(add_months(LASTPERIODPOSTED,1)) ,CURRENTDATE=add_months(CURRENTDATE,1) " & _ 
					  " ,CloseAccountingPeriod = " & temp_NextCloseAccountingPeriod & " " & _  
					  " ,CloseAccountingYear = " & temp_NextCloseAccountingYear & " " & _  
					  " ,CurrentAccountingPeriod = " & temp_NextCurrentAccountingPeriod & " " & _  
					  " ,CurrentAccountingYear = " & temp_NextCurrentAccountingYear & " " & _  
					  " ,lastupdate = sysdate, lasttime = to_char(sysdate, 'hh24:mi') " & _
					  " where UPPER(system) ='BUILDING' and fcba='" & Session("fcBA") & "'"
				ELSE
				     q=" Update PeriodControl Set LASTPERIODPOSTED=last_day(add_months(LASTPERIODPOSTED,1)) " & _ 
					   " ,CloseAccountingPeriod = " & temp_NextCloseAccountingPeriod & " " & _  
				  " ,CloseAccountingYear = " & temp_NextCloseAccountingYear & " " & _  
				  " ,CurrentAccountingPeriod = " & temp_NextCurrentAccountingPeriod & " " & _  
				  " ,CurrentAccountingYear = " & temp_NextCurrentAccountingYear & " " & _     
					  " ,lastupdate = sysdate, lasttime = to_char(sysdate, 'hh24:mi') " & _
					  " where  UPPER(system) ='BUILDING' and fcba='" & Session("fcBA") & "'"
			    END IF
			END IF
			
			IF temp_NextCloseAccountingPeriod = 13 AND temp_NextCurrentAccountingPeriod = 1  THEN
			
				q=" Update PeriodControl Set CURRENTDATE=add_months(CURRENTDATE,1) " & _ 
				  " ,CloseAccountingPeriod = " & temp_NextCloseAccountingPeriod & " " & _  
				  " ,CloseAccountingYear = " & temp_NextCloseAccountingYear & " " & _  
				  " ,CurrentAccountingPeriod = " & temp_NextCurrentAccountingPeriod & " " & _  
				  " ,CurrentAccountingYear = " & temp_NextCurrentAccountingYear & " " & _  
				  " ,lastupdate = sysdate, lasttime = to_char(sysdate, 'hh24:mi') " & _
				  " where UPPER(system) ='BUILDING' and and fcba='" & Session("fcBA") & "'"
			END IF
			
		connStr.execute(q)
		
		
		TimeAkhir=Time()
		Lama =TimeAkhir-TimeAwal
		Response.Write "<br>" & "Finish ... " & time() & "<br><br>"
		Response.Write " Time Elapsed " & Hour(Lama) & " Hours " & Minute(Lama) & " Minutes " &  Second(Lama) & " Seconds."	
		
		'--Added by Ukie at 25-Oct-2012--	
		Endingtime = FormatDateTime(CDate(Now))

		qLineMax = "select nvl(max(lineno),0) tempMaxLine from CLOSINGLOG where fcba = '" & session("fcBA") & "' AND SYSTEM = 'Building'" 
								
		set rsqLineMax = Server.CreateObject("ADODB.Recordset")
		rsqLineMax.Open qLineMax, connstr
		if not rsqLineMax.eof then
			tempMaxLine = cdbl(rsqLineMax("tempMaxLine"))
		else
			tempMaxLine = 0
		end if 

		tempMaxLine = tempMaxLine + 1

		strClosingLog = " INSERT INTO CLOSINGLOG(SYSTEM, FCBA, STARTINGTIME, ENDINGTIME, LINENO, lastupdate, lasttime) VALUES (" & _
						" 'Building', " & _
						" '" & Session("fcBA") & "'," & _
						" TO_DATE('" & Startingtime & "', 'mm/dd/yyyy hh:mi:ssam')," & _
						" TO_DATE('" & Endingtime & "', 'mm/dd/yyyy hh:mi:ssam')," & _
						" " & tempMaxLine & "" & _
						" , sysdate" & _
						" , to_char(sysdate, 'hh24:mi')" & _
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
						" WHERE FCBA = '" & Session("fcBA") & "' AND LINENO = " & tempMaxLine & " AND SYSTEM = 'Building'"
				
		Set rsLenghtime = Server.CreateObject("ADODB.Recordset")
		rsLenghtime.CursorType = 3
		rsLenghtime.CursorLocation = 3
		rsLenghtime.Open strLenghtime, connstr

		upClosingLog = " UPDATE CLOSINGLOG SET" & _
						" LENGTHTIME = '" & rsLenghtime(5) & " Days " & rsLenghtime(6) & " Hours " & rsLenghtime(7) & " Minutes " & rsLenghtime(8) & " Seconds " & "'" & _
						" WHERE FCBA = '" & Session("fcBA") & "' AND LINENO = " & tempMaxLine & " AND SYSTEM = 'Building'"		
		connStr.execute(upClosingLog)
		'--End Added--
		
		%>
		<br>
		<img src="images/ok.gif"><font color=Blue pointsize=9 face=Verdana>Building Closing is succesfully...
		<script language="JavaScript" type="text/javascript" >
		document.getElementById('status_running').style.visibility = 'hidden';
		</script>
		
		<%
		
			txtLastDay  =cmdFindPeriod("Building")
			txtNextMonth=cmdPeriodClose("Building")		
		
			Session("v_Building")=txtLastDay
			Session("cls_Building")= txtNextMonth	
		%>
		
		
	<%End if%>

<%
else 
%>
	<script language="JavaScript" type="text/javascript" >
	document.getElementById('status_running').style.visibility = 'hidden';
	</script>
	<br><br><img src="images/error.gif">Building Closing is NOT completed...
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