<%@Language=VBScript%>
<%Response.Buffer=true%>
<%Response.Clear%>
<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>ePlaS - Generate Working Time</title>
</head>
<body>
<%

Dim TimeAwal,TimeAkhir,Hitung,Lama,TxtYear

Dim rsVehicle, qVehicle, nxActivity, nxGaji, nCostTotal, qSelisih, objectRS, objectSQL,nTotalUnit,nProporsi,nTotalProporsi
Dim nTotalAlokasi, rsCostVehicle, qCostVehicle, nBasis, rsBasis, qBasis, qInsertCR
Dim rsVHActivity, qVHActivity, nWSAlokasi, qInsertDB, TD, TD1, RF, RFG, nSelisih, nSub, nSubs
Dim qABC, q, qLineMax, rsqLineMax, tempMaxLine
Dim strClosingLog, strLenghtime, rsLenghtime, upClosingLog, txtLastDay, txtNextMonth
Dim maxLineNoVehicleCB, tempProcessCode, tempProcessNameCode, tempRefCodePVDetail, qReference
Dim rsReferencePVDetail, donePVDetail, qNormalisasiDetail, qVoucher, rsVoucher, qCost_Activity
Dim qVoucherDetail, rsVoucherDetail, maxLineNoDetail
Dim txtRefStringSeparator_PVDetail, txtRefStringValues_PVDetail, nQuantity, nOutput
Dim txtRefNumberKeys_PVDetail, txtRefStringKeys_PVDetail, qCost_ActivityHistory
Dim temp_DebetCredit
'--Added by Ukie at 25-Oct-2012--	
Dim Startingtime, Endingtime
TimeAwal=Time()
Startingtime = FormatDateTime(CDate(Now))
TxtYear=Request("txtYear")
'--End Added--	
%>
<div id="status_running"><img src="ICON/working.gif"></div><br>
<font face="Verdana" size="1">
Generate Employee WorkingTime  <%=TxtYear%> <br> 
Please Wait...<br><br>
<%
 Response.Write "Start ..." & Time() & "<br><br>" 

dim rsEmpAllDed,rsPO,qReceive,qPO,qControlJob,nPurchaseControl,rsControlJob,txtControlJob,qAllowance,nQtyReceived,qNormalisasi,nPresentRate,nDays,nAmount
dim v_AllowanceCode

' XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' Add by NAF to select lastperiodposted
Dim qPeriodControl, rsPeriodControl, temp_datePeriodControl
Dim qCostActivityMax, rsCostActivityMax, maxLineNo


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
		
Function GetSSArray(strSQL)
			set objSSArray = connStr.Execute(strSQL)
			If Not objSSArray.EOF Then
				arrSS = objSSArray.GetRows()
				GetSSArray = arrSS
			Else
				Response.Write("There are currently no records returned.<br/>")
                Exit Function
			End If
			objSSArray.Close()
			Set objSSArray = Nothing
End Function
		
Dim qListDate
        connstr.BeginTrans
		qListDate = "INSERT INTO employee_workingtime( " & _
		"EMPCODE,PERIOD_YEAR,PERIOD_MONTH," & _
		"D_1,D_2,D_3,D_4,D_5," & _                 
		"D_6,D_7,D_8,D_9,D_10," & _                
		"D_11,D_12,D_13,D_14,D_15," & _                
		"D_16,D_17,D_18,D_19,D_20," & _
		"D_21,D_22,D_23,D_24,D_25," & _                 
		"D_26,D_27,D_28,D_29,D_30,D_31," & _
		"ACTIVATION,FCENTRY,FCEDIT,FCIP,FCBA,LASTUPDATE,LASTTIME)" & _     
        "SELECT FCCODE,YDAY,MDAY,"	& _ 
		"D_1,D_2,D_3,D_4,D_5," & _                 
		"D_6,D_7,D_8,D_9,D_10," & _                
		"D_11,D_12,D_13,D_14,D_15," & _                
		"D_16,D_17,D_18,D_19,D_20," & _
		"D_21,D_22,D_23,D_24,D_25," & _                 
		"D_26,D_27,D_28,D_29,D_30,D_31," & _		
		"'Y' ACTIVATION,'" & session("CodeName") & "' FCENTRY,'" & session("CodeName") & "' FCEDIT,'' FCIP,'" & session("fcba") & "' FCBA,To_Date('01/01/"& TxtYear & "','MM/DD/YYYY') LASTUPDATE,'00:00' LASTTIME " & _            
		"FROM (" & _
		"SELECT Emp.FCBA,Emp.fccode,YDAY,MDAY," & _
		"MAX(CASE WHEN DDAY =  1 THEN jk ELSE '' END) AS D_1," & _
		"MAX(CASE WHEN DDAY =  2 THEN jk ELSE '' END) AS D_2," & _
		"MAX(CASE WHEN DDAY =  3 THEN jk ELSE '' END) AS D_3," & _
		"MAX(CASE WHEN DDAY =  4 THEN jk ELSE '' END) AS D_4," & _
		"MAX(CASE WHEN DDAY =  5 THEN jk ELSE '' END) AS D_5," & _
		"MAX(CASE WHEN DDAY =  6 THEN jk ELSE '' END) AS D_6," & _
		"MAX(CASE WHEN DDAY =  7 THEN jk ELSE '' END) AS D_7," & _
		"MAX(CASE WHEN DDAY =  8 THEN jk ELSE '' END) AS D_8," & _
		"MAX(CASE WHEN DDAY =  9 THEN jk ELSE '' END) AS D_9," & _
		"MAX(CASE WHEN DDAY =  10 THEN jk ELSE '' END) AS D_10," & _
		"MAX(CASE WHEN DDAY =  11 THEN jk ELSE '' END) AS D_11," & _
		"MAX(CASE WHEN DDAY =  12 THEN jk ELSE '' END) AS D_12," & _
		"MAX(CASE WHEN DDAY =  13 THEN jk ELSE '' END) AS D_13," & _
		"MAX(CASE WHEN DDAY =  14 THEN jk ELSE '' END) AS D_14," & _
		"MAX(CASE WHEN DDAY =  15 THEN jk ELSE '' END) AS D_15," & _
		"MAX(CASE WHEN DDAY =  16 THEN jk ELSE '' END) AS D_16," & _
		"MAX(CASE WHEN DDAY =  17 THEN jk ELSE '' END) AS D_17," & _
		"MAX(CASE WHEN DDAY =  18 THEN jk ELSE '' END) AS D_18," & _
		"MAX(CASE WHEN DDAY =  19 THEN jk ELSE '' END) AS D_19," & _
		"MAX(CASE WHEN DDAY =  20 THEN jk ELSE '' END) AS D_20," & _
		"MAX(CASE WHEN DDAY =  21 THEN jk ELSE '' END) AS D_21," & _
		"MAX(CASE WHEN DDAY =  22 THEN jk ELSE '' END) AS D_22," & _
		"MAX(CASE WHEN DDAY =  23 THEN jk ELSE '' END) AS D_23," & _
		"MAX(CASE WHEN DDAY =  24 THEN jk ELSE '' END) AS D_24," & _
		"MAX(CASE WHEN DDAY =  25 THEN jk ELSE '' END) AS D_25," & _
		"MAX(CASE WHEN DDAY =  26 THEN jk ELSE '' END) AS D_26," & _
		"MAX(CASE WHEN DDAY =  27 THEN jk ELSE '' END) AS D_27," & _
		"MAX(CASE WHEN DDAY =  28 THEN jk ELSE '' END) AS D_28," & _
		"MAX(CASE WHEN DDAY =  29 THEN jk ELSE '' END) AS D_29," & _
		"MAX(CASE WHEN DDAY =  30 THEN jk ELSE '' END) AS D_30," & _
	    "MAX(CASE WHEN DDAY =  31 THEN jk ELSE '' END) AS D_31 " & _ 		
	  "FROM " & _
	  "( " & _
	  "      SELECT DAY,DDAY,MDAY,YDAY,WDAY,GANGCODE,PERIOD,DTPART,JK FROM " & _
	  "      ( " & _
	  "        SELECT " & _
	  "        TRUNC(To_Date('01/01/"& TxtYear & "','MM/DD/YYYY'), 'MM') + LEVEL - 1 DAY, " & _
	  "        To_Number(To_Char(TRUNC(To_Date('01/01/"& TxtYear & "','MM/DD/YYYY'), 'MM') + LEVEL - 1,'DD')) DDAY, " & _
	  "        To_Number(To_Char(TRUNC(To_Date('01/01/"& TxtYear & "','MM/DD/YYYY'), 'MM') + LEVEL - 1,'MM')) MDAY, " & _
	  "        To_Char(TRUNC(To_Date('01/01/"& TxtYear & "','MM/DD/YYYY'), 'MM') + LEVEL - 1,'YYYY') YDAY, " & _
	  "        To_Char(TRUNC(To_Date('01/01/"& TxtYear & "','MM/DD/YYYY'), 'MM') + LEVEL - 1,'D') WDAY " & _
	  "        FROM dual CONNECT BY TRUNC(TRUNC(To_Date('01/01/"& TxtYear & "','MM/DD/YYYY'), 'MM') + LEVEL - 1, 'MM') <= TRUNC(To_Date('12/31/"& TxtYear & "','MM/DD/YYYY'), 'MM') " & _
	  "      )tmp " & _
	  "      INNER JOIN " & _
	  "      ( " & _
	  "        select distinct GangCode, period,dtpart,jk " & _
	  "        from " & _
	  "        ( " & _
	  "        select GangCode,period,2 dtpart,D_2 jk from WorkingTime_Template  where fcba='"& session("FCBA") &"' and Activation='Y' and period='" & TxtYear & "' " & _
	  "        union all " & _
	  "        select GangCode,period,3 dtpart,D_3 jk from WorkingTime_Template  where fcba='"& session("FCBA") &"' and Activation='Y' and period='" & TxtYear & "' " & _
	  "        union all " & _
	  "        select GangCode,period,4 dtpart,D_4 jk from WorkingTime_Template  where fcba='"& session("FCBA") &"' and Activation='Y' and period='" & TxtYear & "' " & _
	  "        union all " & _
	  "        select GangCode,period,5 dtpart,D_5 jk from WorkingTime_Template  where fcba='"& session("FCBA") &"' and Activation='Y' and period='" & TxtYear & "' " & _
	  "        union all " & _
	  "        select GangCode,period,6 dtpart,D_6 jk from WorkingTime_Template  where fcba='"& session("FCBA") &"' and Activation='Y' and period='" & TxtYear & "' " & _
	  "        union all " & _
	  "        select GangCode,period,7 dtpart,D_7 jk from WorkingTime_Template  where fcba='"& session("FCBA") &"' and Activation='Y' and period='" & TxtYear & "' " & _
	  "        union all " & _
	  "        select GangCode,period,1 dtpart,D_1 jk from WorkingTime_Template  where fcba='"& session("FCBA") &"' and Activation='Y' and period='" & TxtYear & "' " & _
	  "        ) " & _
	  "      )tmpjamkerja ON  tmp.YDAY=tmpjamkerja.period  and tmp.WDAY=tmpjamkerja.dtpart " & _
	  ")tgljamkerja " & _
	  "INNER JOIN " & _ 
	  "(SELECT FcCode,GangCode,dateterminate,FCBA FROM EMPlOYEE WHERE FCBA='"& session("FCBA") &"') Emp ON tgljamkerja.gangcode=Emp.GangCode " & _
	  "WHERE (dateterminate IS NULL OR DAY < dateterminate) " & _ 
	  "GROUP BY Emp.FCBA,Emp.FcCode,YDAY,MDAY " & _
	  ")upload " & _
	  "WHERE NOT EXISTS " & _
	  "( " & _
	  "  SELECT 1 FROM  employee_workingtime x WHERE x.Period_Year=upload.YDAY AND x.Period_month=upload.Mday " & _
	  "	AND x.EmpCode=upload.fccode AND x.FCBA=upload.fcba " & _
	  ") " & _
	  "ORDER BY FcCode,YDAY,MDAY " 

	  connStr.execute(qListDate)

     If Err.Number = 0 Then
		connstr.CommitTrans
		Response.Flush
	 Else
		connstr.RollbackTrans
		Response.write("Error Save Query ("&Err.Description&")")
		Response.Flush
	 End If
	 
	 TimeAkhir=Time()
	 Lama =TimeAkhir-TimeAwal
	 Response.Write "<br>" & "Finish ... " & time() & "<br><br>"
	 Response.Write " Time Elapsed " & Hour(Lama) & " Hours " & Minute(Lama) & " Minutes " &  Second(Lama) & " Seconds."	
		
	
%>
		<br>
		<img src="images/ok.gif"><font color=Blue pointsize=9 face=Verdana>Generate Employee WorkingTimeg is Finish...

</body>
</html>