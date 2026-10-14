<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<%
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"

Dim TimeAwal,TimeAkhir,Hitung,Lama, qOpenEmpalldedStatus,TypeUnclose,TypeFilter 
TimeAwal=Time()
TypeUnclose = request("TypeUnclosing")

if TypeUnclose = "A" THEN	
	TypeFilter = "AND system like '%' "
end if

if TypeUnclose = "B" THEN	
	TypeFilter = "AND system in ('Bank','Costbook') "
end if

if TypeUnclose = "C" THEN	
	TypeFilter = "AND system in ('Bank','Contract','Costbook') "
end if

if TypeUnclose = "D" THEN	
	TypeFilter = "AND system in ('Costbook') "
end if

%>
<div id="status_running"><img src="ICON/working.gif"></div><br/>
<font face="Verdana" size="1">
UNIT <%=Session("fcba")%> <br/> 
GL Last Period Close <%=FormatDateTime(Session("cls_Costbook"),1)%> <br/> 
GL Current Period is <%=FormatDateTime(Session("v_Costbook"),1)%> <br/><br/>
Please Wait...<br/>
<%
 Response.Write "Start ..." & Time() & "<br/>" 
dim rsVoucher,rsBank,qVoucher,qPO,nPurchaseControl,txtControlJob,qCost_ActivityDB,qCost_ActivityCR,nQtyReceived,qNormalisasi,txtOriginalST
'---- Normalisasi

Response.Write "to unPost  All Module, Please wait, this action take several minute.." & "<br/>"

qOpenEmpalldedStatus = "BEGIN EMPALLDEDSTATUS_API.UnCloseFinancial (" & Mid(Session("CurP_Costbook"),1,4) & "," & _
		                                                                Mid(Session("CurP_Costbook"),5,2) & ",'" & _
																		Session("fcBA") & "','" & _
																		Session("CodeName") & "','" & _
																		Session("CodeName") & "','" & _
																		Request.ServerVariables("REMOTE_ADDR") & "'); END;"
response.write qOpenEmpalldedStatus & "<br>" 

connstr.execute(qOpenEmpalldedStatus)																		



' Start Remarks By Iwa
' Start Codes for Update - Remove Columns, By NAF, 3 Nov 2014
'qNormalisasiDetail_UPDATE = " update cost_activity_hist_tab set deleteflag = 'REMOVE' " & _
'					 " where " & _
'					 " lineheaderno || fcba in (select lineno||fcba from cost_activity " & _
'							" where  ACCOUNTINGPERIOD=" & Mid(Session("CurP_Costbook"),5,2) &  _
'							" and ACCOUNTINGYEAR=" & Mid(Session("CurP_Costbook"),1,4) & " and fcBA='" & Session("fcBA") & "')"

'connstr.execute(qNormalisasiDetail_UPDATE)
'connStr.execute("Commit")

'qNormalisasiDetail_UPDATE = " delete from cost_activity_hist_tab where deleteflag = 'REMOVE' and fcba =  '" & Session("fcBA") & "'"

'connstr.execute(qNormalisasiDetail_UPDATE)
' connStr.execute("Commit")
' End Remarks By Iwa

' Start Remarked by NAF, 3 Nov 2014							
'qNormalisasiDetail = " delete from cost_activity_hist_tab " & _
'					 " where " & _
'					 " lineheaderno || fcba in (select lineno||fcba from cost_activity " & _
'							" where  to_number(to_Char(TDATE,'mm'))=" & Month(Session("v_Costbook")) &  _
'							" and to_number(to_Char(TDATE,'yyyy'))=" & Year(Session("v_Costbook")) & " and fcBA='" & Session("fcBA") & "')"
					 
'connstr.execute(qNormalisasiDetail)
'connStr.execute("Commit")

'qNormalisasi="Delete From Cost_Activity where  to_number(to_Char(TDATE,'mm'))=" & Month(Session("v_Costbook")) &  _
'			 " and to_number(to_Char(TDATE,'yyyy'))=" & Year(Session("v_Costbook")) & " and fcBA='" & Session("fcBA")  & "'"		 
'connstr.execute(qNormalisasi)
'connStr.execute("Commit")
' End Remarked by NAF

' Start Remarks By Iwa
'qNormalisasiDetail_UPDATE ="Update Cost_Activity set deleteflag = 'REMOVE' where   ACCOUNTINGPERIOD=" & Mid(Session("CurP_Costbook"),5,2) &  _
'			 " and  ACCOUNTINGYEAR=" & Mid(Session("CurP_Costbook"),1,4) & " and fcBA='" & Session("fcBA")  & "'"		 
'connstr.execute(qNormalisasiDetail_UPDATE)
'connStr.execute("Commit")
'
'qNormalisasiDetail_UPDATE = " delete from cost_activity where deleteflag = 'REMOVE' and fcba =  '" & Session("fcBA") & "'"
'
'connstr.execute(qNormalisasiDetail_UPDATE)
'connStr.execute("Commit")
' End Remarks By Iwa

	QAccountPeriod	=	" select " & _
							" periodcontrol_api.GetNextCloseAccountingPeriod(system,fcba) NextCloseAccountingPeriod ,  " & _ 
							" periodcontrol_api.GetNextCloseAccountingYear(system,fcba) NextCloseAccountingYear   " & _ 
							" from periodcontrol p  " & _ 
							" where UPPER(system) ='COSTBOOK' and fcba='" & Session("fcBA") & "'"
	Set rsAccountPeriod = Server.CreateObject("ADODB.Recordset")
		rsAccountPeriod.CursorType = 3
		rsAccountPeriod.CursorLocation = 3
		rsAccountPeriod.Open QAccountPeriod, connstr
		if not rsAccountPeriod.eof then
			temp_NextCloseAccountingPeriod =  rsAccountPeriod("NextCloseAccountingPeriod")
			temp_NextCloseAccountingYear   =  rsAccountPeriod("NextCloseAccountingYear")
		else
			Response.Write  "<font face=Verdana size=1 color=red>" & "IPLAS Error, No PERIODCONTROL for COSTBOOK " & "<br>"
		end if
	rsAccountPeriod.close
	set rsAccountPeriod = Nothing	
	
	SELECT CASE Cint(temp_NextCloseAccountingPeriod) 
		Case 2 
			temp_EndingPeriod = 15
			temp_EndingYear = CInt (temp_NextCloseAccountingYear) - 1 
		Case 3,4,5,6,7,8,9,10,11,12,13,14,15
		    temp_EndingPeriod = CInt (temp_NextCloseAccountingPeriod) - 2
			temp_EndingYear = temp_NextCloseAccountingYear
			
	End Select 
	
	IF temp_EndingPeriod > 12 Then
	      TempEndPosted = 12
    ELSE	
		  TempEndPosted = temp_EndingPeriod
    END IF
	
	temp_CurrentAccountingPeriod = CInt (temp_NextCloseAccountingPeriod) - 1 
	temp_CurrentAccountingYear = CInt (temp_NextCloseAccountingYear) - 1 
	
Response.Write "Setting Period..."
IF Cint(temp_NextCloseAccountingPeriod) <> 1 THEN

IF  Cint(temp_NextCloseAccountingPeriod) > 1 AND Cint(temp_NextCloseAccountingPeriod) < 13   Then
    q=" Update PeriodControl Set " & _ 
	  " CURRENTDATE= TRUNC(to_date('" & temp_CurrentAccountingPeriod & "/01/" & temp_NextCloseAccountingYear & "','mm/dd/yyyy'),'MM') " & _
	  " ,CurrentAccountingPeriod = " & temp_CurrentAccountingPeriod & _
	  " ,CurrentAccountingYear = " & temp_NextCloseAccountingYear & _
	  " ,LASTPERIODPOSTED= LAST_DAY(to_date('" & TempEndPosted & "/01/" & temp_EndingYear & "','mm/dd/yyyy'))" & _ 
	  " ,CloseAccountingPeriod =  " & temp_EndingPeriod & _
	  " ,CloseAccountingYear = " &  temp_EndingYear & _
	  " where fcba='" & Session("fcBA") & "' " & TypeFilter 
	  
	' q=" Update PeriodControl Set LASTPERIODPOSTED=" & _
	  ' " to_date('" & FormatDateTime(Session("cls_Costbook"),2) & "','mm/dd/yyyy') ,CURRENTDATE=" & _
	  ' " to_date('" & FormatDateTime(Session("v_Costbook"),2) & "','mm/dd/yyyy')  " & _
	  ' " ,CloseAccountingPeriod = " & MID(Session("ClsP_Costbook"),5,2) & " " & _  
	  ' " ,CloseAccountingYear = " & MID(Session("ClsP_Costbook"),1,4) & " " & _  
	  ' " ,CurrentAccountingPeriod = " & MID(Session("CurP_Costbook"),5,2) & " " & _  
	  ' " ,CurrentAccountingYear = " & MID(Session("CurP_Costbook"),1,4) & " " & _ 
	  ' " where fcba='" & Session("fcBA") & "'"
End IF



IF Cint(temp_NextCloseAccountingPeriod) > 12 Then
	q=" Update PeriodControl Set " & _
	    " CurrentAccountingPeriod = " & temp_CurrentAccountingPeriod  & _  
	    " ,CurrentAccountingYear = " & temp_NextCloseAccountingYear & _
	    " ,CloseAccountingPeriod = " & temp_EndingPeriod & _
	    " ,CloseAccountingYear = " &  temp_EndingYear & _ 
	    " where fcba='" & Session("fcBA") & "' " & TypeFilter 

END IF
response.write q & "<br>"

connStr.execute(q)
Response.Write "New Period : " & temp_NextCloseAccountingYear & "/" & temp_NextCloseAccountingPeriod & "<br/>"

END IF


IF Cint(temp_NextCloseAccountingPeriod) = 1 THEN
     q=" Update PeriodControl Set " & _ 
	  " CURRENTDATE= to_date('12/01/" & temp_CurrentAccountingYear & "','mm/dd/yyyy') " & _
	  " ,CurrentAccountingPeriod = 13 "& _
	  " ,CurrentAccountingYear = " & temp_CurrentAccountingYear & _
	  " ,LASTPERIODPOSTED= to_date('12/31/" & temp_CurrentAccountingYear & "','mm/dd/yyyy')" & _ 
	  " ,CloseAccountingPeriod =  12 " & _
	  " ,CloseAccountingYear = " &  temp_CurrentAccountingYear & _
	  " where fcba='" & Session("fcBA") & "' " 
	 connStr.execute(q)

Response.Write "New Period : " & temp_CurrentAccountingYear & "/13 <br/>"

END IF

'IF  Cint(MID(Session("CurP_Costbook"),5,2)) < 13 Then
	' q=" Update PeriodControl Set " & _ 
	  ' " LASTPERIODPOSTED= last_day(add_months(to_date('" & FormatDateTime(Session("v_Costbook"),2) & "','mm/dd/yyyy'),-2)) " & _
	  ' " ,CURRENTDATE= add_months(to_date('" & FormatDateTime(Session("v_Costbook"),2) & "','mm/dd/yyyy'),-1) " & _
	  ' " ,CloseAccountingPeriod =  EXTRACT(MONTH FROM add_months(to_date('" & FormatDateTime(Session("v_Costbook"),2) & "','mm/dd/yyyy'),-1)) " & _
	  ' " ,CloseAccountingYear = EXTRACT(YEAR FROM add_months(to_date('" & FormatDateTime(Session("v_Costbook"),2) & "','mm/dd/yyyy'),-1)) " & _
	  ' " ,CurrentAccountingPeriod = EXTRACT(MONTH FROM last_day(add_months(to_date('" & FormatDateTime(Session("v_Costbook"),2) & "','mm/dd/yyyy'),-2))) " & _
	  ' " ,CurrentAccountingYear = EXTRACT(YEAR FROM last_day(add_months(to_date('" & FormatDateTime(Session("v_Costbook"),2) & "','mm/dd/yyyy'),-2))) " & _
	  ' " where fcba='" & Session("fcBA") & "'"
' ELSE
	' q=" Update PeriodControl Set " & _
	    ' " CloseAccountingPeriod = " & MID(Session("ClsP_Costbook"),5,2) & " " & _  
	    ' " ,CloseAccountingYear = " & MID(Session("ClsP_Costbook"),1,4) & " " & _  
	    ' " ,CurrentAccountingPeriod = " & MID(Session("CurP_Costbook"),5,2) & " " & _  
	    ' " ,CurrentAccountingYear = " & MID(Session("CurP_Costbook"),1,4) & " " & _ 
	    ' " where fcba='" & Session("fcBA") & "'"

'END IF	  

'response.write q


Session("v_Costbook")=cmdFindPeriod("Costbook")
Session("cls_Costbook")=cmdPeriodClose("Costbook")
Session("CurP_Costbook")=cmdFindAccountPeriod("Costbook")
Session("ClsP_Costbook")=cmdAccountPeriodClose("Costbook")

	Session("v_Bank")=Session("v_Costbook")
	Session("v_Purchase")=Session("v_Costbook")
	Session("v_Stores")=Session("v_Costbook")
	Session("v_Employee")=Session("v_Costbook")
	Session("v_Workshop")=Session("v_Costbook")
	Session("v_Vehicle")=Session("v_Costbook")
	Session("v_Contract")=Session("v_Costbook")
	Session("v_Nursery")=Session("v_Costbook")
	Session("v_Project")=Session("v_Costbook")
	Session("v_Building")=Session("v_Costbook")
	Session("v_General_Charges")=Session("v_Costbook")
	Session("v_Harvesting")=Session("v_Costbook")
	Session("v_WeighBridge")=Session("v_Costbook")
	Session("v_Mill_Production")=Session("v_Costbook")
	Session("v_Delivery")=Session("v_Costbook")
	
    ' To make always LOGIN, by NAF, 12-Sep-2013 	
    'Session("USERNAME") = Session("v_Costbook")
	'Session("CodeName") =Session("v_Costbook")
	' End of Modification.

	Session("cls_Bank")= Session("cls_Costbook")
	Session("cls_Purchase")= Session("cls_Costbook")
	Session("cls_Stores")= Session("cls_Costbook")
	Session("cls_Employee")= Session("cls_Costbook")
	Session("cls_Workshop")= Session("cls_Costbook")
	Session("cls_Vehicle")= Session("cls_Costbook")
	Session("cls_Contract")= Session("cls_Costbook")
	Session("cls_Nursery")= Session("cls_Costbook")
	Session("cls_Project")=Session("cls_Costbook")
	Session("cls_Building")= Session("cls_Costbook")
	Session("cls_General_Charges")= Session("cls_Costbook")
	Session("cls_Harvesting")= Session("cls_Costbook")
	Session("cls_WeighBridge")= Session("cls_Costbook")
	Session("cls_Mill_Production")=Session("cls_Costbook")
	Session("cls_Delivery")= Session("cls_Costbook")
	
	
	 ' Add by Ali Akbar  28 Desember 2015
	 
	   Session("CurP_Bank")=Session("CurP_Costbook")
	   Session("CurP_Purchase")=Session("CurP_Costbook")
	   Session("CurP_Stores")=Session("CurP_Costbook")
	   Session("CurP_Employee")=Session("CurP_Costbook")
	   Session("CurP_Workshop")=Session("CurP_Costbook")
	   Session("CurP_Vehicle")=Session("CurP_Costbook")
	   Session("CurP_Contract")=Session("CurP_Costbook")
	   Session("CurP_Nursery")=Session("CurP_Costbook")
	   Session("CurP_Project")=Session("CurP_Costbook")
	   Session("CurP_Building")=Session("CurP_Costbook")
	   Session("CurP_General_Charges")=Session("CurP_Costbook")
	   Session("CurP_Reversed")=Session("CurP_Costbook")
	   Session("CurP_Costbook")=Session("CurP_Costbook")
	   Session("CurP_Harvesting")=Session("CurP_Costbook")
	   Session("CurP_WeighBridge")=Session("CurP_Costbook") 
	   Session("CurP_Mill_Production")=Session("CurP_Costbook")
	   Session("CurP_Delivery")=Session("CurP_Costbook") 
	   Session("CurP_ERPReport")=Session("CurP_Costbook")

	   Session("ClsP_Bank")=Session("ClsP_Costbook")
	   Session("ClsP_Purchase")=Session("ClsP_Costbook")
	   Session("ClsP_Stores")=Session("ClsP_Costbook")
	   Session("ClsP_Employee")=Session("ClsP_Costbook")
	   Session("ClsP_Workshop")=Session("ClsP_Costbook")
	   Session("ClsP_Vehicle")=Session("ClsP_Costbook")
	   Session("ClsP_Contract")=Session("ClsP_Costbook")
	   Session("ClsP_Nursery")=Session("ClsP_Costbook")
	   Session("ClsP_Project")=Session("ClsP_Costbook")
	   Session("ClsP_Building")=Session("ClsP_Costbook")
	   Session("ClsP_General_Charges")=Session("ClsP_Costbook")
	   Session("ClsP_Reversed")=Session("ClsP_Costbook")
	   Session("ClsP_Costbook")=Session("ClsP_Costbook")
	   Session("ClsP_Harvesting")=Session("ClsP_Costbook")
	   Session("ClsP_WeighBridge")=Session("ClsP_Costbook")
	   Session("ClsP_Mill_Production")=Session("ClsP_Costbook")
	   Session("ClsP_Delivery")=Session("ClsP_Costbook") 
	   Session("ClsP_ERPReport")=Session("ClsP_Costbook")
	
	
	
' UpdateSL = " UPDATE stockledger" & _
           ' " SET qtyreceived = 0, amtreceived = 0" & _
		   ' "     , qtyreturned = 0, amtreturned = 0" & _
		   ' "     , qtywriteoff = 0, amtwriteoff = 0" & _
		   ' "     , qtytransfered = 0, amttransfered = 0" & _
		   ' "     , qtyreceivedtask = 0, amtreceivedtask = 0" & _
		   ' "     , qtytransporttask = 0, amttransporttask = 0" & _
		   ' "     , qtyissued = 0, amtissued = 0" & _
		   ' "     , amtperunit = 0" & _
		   ' " WHERE year = '" & mid(Session("CurP_Stores"),1,4) & "'" & _
		   ' "       AND month = '" & mid(Session("CurP_Stores"),5,2) & "'" & _
		   ' "       AND fcba = '" & Session("fcBA") & "'"
' connStr.execute(UpdateSL)

'Response.Write "<br/> Actual: " & UpdateSL



'Added by Sukri, date 19-Sep-2012
'As per Mr. Nafis Requirements to update Amount to be zero in Issue Details
' UpdateAmountID = "Update ISSUEDETAILS"
' UpdateAmountID = UpdateAmountID + " Set AMOUNT = 0, PRICE = 0 WHERE"
' UpdateAmountID = UpdateAmountID + " SIVCODE IN (SELECT SIVCODE FROM ISSUE I WHERE"
' UpdateAmountID = UpdateAmountID + " TO_NUMBER(TO_CHAR(I.SIVDATE,'YYYY')) =" &  mid(Session("CurP_Stores"),1,4)  & " AND"
' UpdateAmountID = UpdateAmountID + " TO_NUMBER(TO_CHAR(I.SIVDATE,'MM')) =" &  mid(Session("CurP_Stores"),5,2)  & " AND"
' UpdateAmountID = UpdateAmountID + " I.FCBA='" & Session("fcBA") & "') and fcba = '" & Session("fcBA") & "'"


' connStr.execute(UpdateAmountID)

' 'Response.Write "<br/> Actual Issue: " & UpdateAmountID
' 'End of Added

' Dim currentYear, currentMonth
' Dim nextYear, nextMonth 

' currentYear = Year(Session("v_Stores"))
' currentMonth = MONTH(Session("v_Stores"))

' if currentMonth = 12 then 
   ' nextYear = currentYear + 1
   ' nextMonth = 1
' else
   ' nextYear = currentYear 
   ' nextMonth = 1 + currentMonth
' end if

				' Dim QAccountPeriod,rsAccountPeriod,temp_EndingAccountingPeriod,temp_NextCloseAccountingPeriod,temp_NextCloseAccountingYear

				' QAccountPeriod =" select periodcontrol_api.GetNextCurrentAccountingPeriod(system,fcba) NextCurrentAccountingPeriod ,  " & _ 
								' " periodcontrol_api.GetNextCurrentAccountingYear(system,fcba) NextCurrentAccountingYear ,  " & _ 
				                ' " periodcontrol_api.GetEndingAccountingPeriodTY(system,'" & Session("fcBA") & "')  EndingAccountingPeriod  " & _
                                ' " from periodcontrol p  " & _ 
								' " where UPPER(system) = upper('COSTBOOK') and fcba='" & Session("fcBA") & "'"				
				' Set rsAccountPeriod = Server.CreateObject("ADODB.Recordset")
								' rsAccountPeriod.CursorType = 3
								' rsAccountPeriod.CursorLocation = 3
								' rsAccountPeriod.Open QAccountPeriod, connstr
				' if not rsAccountPeriod.eof then
					' temp_NextCloseAccountingPeriod =  Cint(rsAccountPeriod("NextCurrentAccountingPeriod"))
					' temp_NextCloseAccountingYear   =  Cint(rsAccountPeriod("NextCurrentAccountingYear"))
					' temp_EndingAccountingPeriod = Cint(rsAccountPeriod("EndingAccountingPeriod"))
				' else
					' Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error, No PERIODCONTROL for COSTBOOK " & "<br>"
				' end if
			' rsAccountPeriod.close
			' set rsAccountPeriod = Nothing	
			
			
' UpdateSL = "Update StockLedger Set"
' UpdateSL = UpdateSL + " OPENINGBALANCEQTY = 0, OPENINGBALANCEAMT = 0,"
' UpdateSL = UpdateSL + " QTYRECEIVED = 0, AMTRECEIVED = 0,"
' UpdateSL = UpdateSL + " QTYRETURNED = 0, AMTRETURNED = 0,"
' UpdateSL = UpdateSL + " QTYWRITEOFF = 0, AMTWRITEOFF = 0,"
' UpdateSL = UpdateSL + " QTYTRANSFERED = 0, AMTTRANSFERED = 0,"
' UpdateSL = UpdateSL + " QTYreceivedtask = 0, AMTreceivedtask = 0,"
' UpdateSL = UpdateSL + " QTYtransporttask = 0, AMTtransporttask = 0,"
' UpdateSL = UpdateSL + " QTYISSUED = 0, AMTISSUED = 0,"
' UpdateSL = UpdateSL + " AMTPERUNIT = 0 WHERE"
' UpdateSL = UpdateSL + " YEAR = " & temp_NextCloseAccountingYear & " AND"
' UpdateSL = UpdateSL + " MONTH = " & temp_NextCloseAccountingPeriod & " AND"
' UpdateSL = UpdateSL + " FCBA='" & Session("fcBA") & "'"
' connStr.execute(UpdateSL)

' 'Response.Write "<br/> Next: " & UpdateSL


%>
<%

TimeAkhir=Time()
Lama =TimeAkhir-TimeAwal
Response.Write "Finish ... " & time() & "<br/><br/>"
Response.Write " Time Elapsed " & Hour(Lama) & " Hours " & Minute(Lama) & " Minutes " &  Second(Lama) & " Seconds."%>

<script language="JavaScript" type="text/javascript" >
document.getElementById('status_running').style.visibility = 'hidden';
</script>


<br/><br/><img src="images/ok.gif">Process un-closing all modules are completed...

<%
Response.Write Session("USERNAME")
%>

<%
 if err.number <> 0 then
      
	connStr.RollbackTrans
	response.Write  "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif> NAFIS " & err.description & "</font>"
	cmdView	= "True"%>
    <script language="JavaScript" type="text/javascript" >
	document.getElementById('status_running').style.visibility = 'hidden';
	</script>
	<img src="images/error.gif"><%
 End if%>



