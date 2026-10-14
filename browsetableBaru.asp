	<%
	Dim txtTable,txtModule,q,nCountofField,NO,txtSearch,txtField,txtColom,txtWhere,txtPeriod,txtOrder
	Dim i
	txtTable =Request("txtTable")
	txtModule=Request("txtModule")
	txtSearch=Request("txtSearch")
	txtField =Request("txtField")
    txtWhere=""
    txtColom=""

		 
		If len(trim(txtSearch))>0 and len(trim(txtField)) > 0 then
			if instr(txtSearch,"/") <> 0  AND len(txtSearch)=10  then 'Date Type
			   txtWhere = txtField & " like TO_DATE('" & txtSearch & "','DD/MM/YYYY')"
			else
			   if instr(txtSearch,"=") > 0 then 'Data Fix
				  txtWhere = txtField & (txtSearch) 
			   elseif instr(txtSearch,">") > 0 then 'Data Fix
				  txtWhere = txtField & (txtSearch) 
			   elseif instr(txtSearch,"<") > 0 then 'Data Fix
				  txtWhere = txtField & (txtSearch) 
			   elseif instr(txtSearch,"<=") > 0 then 'Data Fix
				  txtWhere = txtField & (txtSearch) 
			   elseif instr(txtSearch,">=") > 0 then 'Data Fix
				  txtWhere = txtField & (txtSearch) 
			   elseif txtField <> "" and txtSearch <> "" then  'Caracter
				  txtWhere = " upper (" & (txtField) & ") like '%" & ucase(txtSearch) & "%'"
			   end if
			end if
			If txtModule="Data Setup" then  'Jika Date Setup Wherenya tidak ada yang lain sudah dibawah di sql
			   txtWhere=" Where " & txtWhere
			Else
			   txtWhere=" and " & txtWhere
			End if
		End if
		
	    
	if (len(trim(txtSearch))=0 OR len(trim(txtField))=0) and (txtModule="Data Master" or txtModule="Data Setup") then
	   '----------- Special for Table Master & Setup -------------------------'
	   q = "Select * from "&txtTable 
	   if txtModule <>  "Data Setup" then
          q = q + " where fcba='" & session("fcBA") & "'"
	   End if
	elseif len(trim(txtSearch))<>0 and len(trim(txtField))<>0 and (txtModule="Data Master" or txtModule="Data Setup") then

	   if txtWhere <> "" and txtModule <>  "Data Setup" then ' Data Master
	      q = "Select * from " & txtTable & " where fcba='" & session("fcBA") & "'" & txtWhere 
       elseif txtWhere <> "" and txtModule =  "Data Setup" then
	      q = "Select * from " & txtTable & txtWhere 'Data Setup Tidak ada BA
	   else
	      q = "Select * from " & txtTable	  
	   end if
	else
	   '----------- Special for Data Transaction -----------------------------'	   
	   if txtTable="PaymentVoucherDetail" then
	      txtColom = "SELECT p.datecreated,pay.vouchercode, pay.jobcode, job.fcname, pay.locationtype," & _
		             " pay.locationcode, pay.amount, pay.REFERENCE, pay.remarks" & _
                     " FROM paymentvoucherdetail pay, job job,paymentvoucher p" & _
                     " WHERE ( ((pay.jobcode = job.fccode) AND (pay.FCBA = job.FCBA)) and " & _
					 "((pay.vouchercode = p.vouchercode) AND (pay.FCBA = p.FCBA)) )" & _
					 " AND pay.fcba='" & session("fcBA") & "'"
		  txtPeriod= " AND to_char(p.datecreated,'mm-yyyy')='" & session("pCashBook") & "'"
		  txtOrder = " ORDER BY pay.vouchercode"					 		  
	   elseif txtTable="PurchaseOrderDetails" then
		  txtColom = "SELECT pur.pocode, pur.podate, sup.fcName as Supplier, ord.itemcode, s.FCNAME,sum(ord.quantity) as Quantity," & _
                     " sum(ord.amount) as Amount,pur.remarks   FROM purchaseorder pur, purchaseorderdetails ord,stock s,Supplier sup" & _
                     " WHERE( ((pur.pocode = ord.pocode) and (pur.fcba = ord.fcba)) " & _
					 " and ((ord.itemcode = s.fcCode) and (ord.fcba = s.fcba))" & _
					 " and ((pur.SUPPLIERCODE=Sup.FcCode) and (pur.fcba=Sup.fcba)) )" & _
					 " and pur.fcba='" & session("fcBA") & "'"
		  txtPeriod= " AND to_char(pur.podate,'mm-yyyy')='" & session("pPurchase") & "'"
		  txtOrder = " group by pur.pocode, pur.podate, sup.fcName, ord.itemcode, s.FCNAME,pur.remarks ORDER BY pur.pocode"	  
	   elseif txtTable="ReceiveNote" then
	      txtColom = "SELECT rec.receivenotecode, rec.pocode, rec.rndate, rd.locationtype," & _
                     "rd.locationcode, stock.fcname, rd.quantity, rd.qtyreturn," & _
                     "rd.amountreceived, rd.amountreturn  FROM receivenote rec, receivenotedetail rd, stock stock" & _
                     " WHERE ( ((rec.receivenotecode = rd.receivenotecode) and (rec.fcba = rd.fcba))" & _
                     " AND ((stock.fccode = rd.purchaseitemcode) and (stock.fcba = rd.fcba))  )" & _
					 " AND rec.fcba='" & session("fcBA") & "'"
		  txtPeriod= " AND to_char(rec.rndate,'mm-yyyy')='" & session("pPurchase") & "'"
		  txtOrder = " ORDER BY rec.receivenotecode"					 		  
	   elseif txtTable="Invoice" then
	      txtColom = "SELECT Inv.INVOICECODE, Inv.INVOICEDATE, Inv.SupplierCode, InvDet.ITEMCODE,Stock.FcName," & _
                     " InvDet.Quantity, InvDet.Amount" & _
                     " FROM Invoice Inv, InvoiceDetail InvDet, stock stock" & _
                     " WHERE ( ((inv.INVOICECODE = invDet.INVOICECODE) and (inv.fcba = invDet.fcba))" & _
                     " AND ((stock.fccode = InvDet.ITEMCODE) and (stock.fcba = InvDet.fcba)) )" & _
					 " AND InvDet.fcba='" & session("fcBA") & "'"
		  txtPeriod= " AND to_char(Inv.INVOICEDATE,'mm-yyyy')='" & session("pPurchase") & "'"
		  txtOrder = " ORDER BY Inv.INVOICECODE"					 		  
	   elseif txtTable="IssueDetails" then
	      txtColom = "SELECT issue.sivdate,iss.sivcode, iss.stockcode, stock.fcname as StockName, iss.locationcode," & _
                     "iss.locationtype, iss.quantity, iss.price, iss.amount, job.fcname" & _
                     " FROM issuedetails iss, issue issue, stock stock, job job" & _
                     " WHERE (    ((iss.sivcode = issue.sivcode) and (iss.fcBA = issue.fcBA))" & _
                     " AND ((iss.stockcode = stock.fccode) and (iss.fcBA= stock.fcBa)) " & _
					 " AND ((iss.jobcode = job.fccode) and (iss.fcBA = job.fcBA)) )" & _
					 " AND iss.fcba='" & session("fcBA") & "'"
		  txtPeriod= " AND to_char(issue.sivdate,'mm-yyyy')='" & session("pStore") & "'"
		  txtOrder = " ORDER BY iss.sivcode"
       elseif txtTable="Attendance_GAD" then
	      txtColom = "SELECT gad.fddate,gad.employeecode, employee.fcname, job.fcname as JobName," & _
		             " gad.locationtype, gad.locationcode,gad.mandays,gad.unit,gad.Rate,gad.Output,gad.othrs" & _
					 " FROM attendance_gad gad, job job, employee employee " & _
                     " WHERE ( ((gad.jobcode = job.fccode) AND (gad.fcBA = job.fcBA)) and " & _
					 " ((gad.employeecode = employee.fccode) and (gad.fcBA = employee.fcBA)) )" & _
					 " AND gad.fcba='" & session("fcBA") & "'"
		  txtPeriod= " AND to_char(gad.fddate,'mm-yyyy')='" & session("pEmployee") & "'"
		  txtOrder = " ORDER BY EMPLOYEECODE,FDDATE"					 
	   elseif txtTable="EmpAllded" or txtTable="Empallded"  then
		  txtColom = "SELECT emp.adhocdate,emp.empcode, employee.fcname as empName, allowded.fcname," & _
					 " emp.amount FROM empallded emp, employee employee, allowded allowded" & _
					 " WHERE ( ((emp.empcode = employee.fccode) and (emp.fcBA = employee.fcBA)) " & _
					 " AND ((emp.allowdedcode = allowded.FCcode) and (emp.fcBA = allowded.fcBA)) )" & _
		             " and emp.fcba='" & session("fcBA") & "'"
		  txtPeriod= " AND to_char(emp.adhocdate,'mm-yyyy')='" & session("pEmployee") & "'"
		  txtOrder = " ORDER BY emp.empcode,emp.adhocdate"	  
	   elseif txtTable="HarvestingQuality" then
		  txtColom = "SELECT har.empcode, employee.fcname, har.fieldcode, har.under_ripe," & _
					 " har.over_ripe, har.abnormal, har.long_stalk, har.eaten_by_rat," & _
					 " har.unharvest_ffb, har.uncollect_lf_circle, har.uncollect_lf_piece," & _
					 " har.unarrange_ffb, har.unprune_frond, har.DAY, har.MONTH, har.YEAR" & _
					 " FROM harvestingquality har, employee employee" & _
					 " WHERE ((har.empcode = employee.fccode) and (har.fcBA = employee.fcBA))" & _
		             " and har.fcba='" & session("fcBA") & "'"
		  txtPeriod= " AND har.Month='" & mid(session("pEmployee"),1,2) & "' and har.Year='" & mid(session("pEmployee"),4,4) & "'"
		  txtOrder = " ORDER BY har.empcode"	  
	   elseif txtTable="WorkshopActivity" then
		  txtColom = "SELECT wor.dateactivity, wor.workshopcode, job.fcname, wor.locationtype," & _
					 " wor.LOCATION, wor.TIME, wor.hours, wor.amount" & _
					 " FROM workshopactivity wor, job job" & _
					 " WHERE ((wor.job = job.fccode) and (wor.fcBA = job.fcBA))" & _
					 " AND wor.fcba='" & session("fcBA") & "'"
		  txtPeriod= " AND to_char(wor.dateactivity,'mm-yyyy')='" & session("pWorkshop") & "'"
		  txtOrder = " ORDER BY wor.dateactivity"
	   elseif txtTable="VehicleActivity" then
		  txtColom = "SELECT veh.vehdate, veh.vehiclecode, veh.jobcode, job.fcname, veh.TIME," & _
					 " veh.locationtype, veh.locationcode, veh.hours,veh.unit, veh.output,veh.amount" & _
					 " FROM vehicleactivity veh, job job" & _
					 " WHERE ((veh.jobcode = job.fccode) and (veh.fcBA = job.fcBA))" & _
					 " AND veh.fcba='" & session("fcBA") & "'"
		  txtPeriod= " AND to_char(veh.vehdate,'mm-yyyy')='" & session("pVehicle") & "'"
		  txtOrder = " ORDER BY veh.vehiclecode"
	   elseif txtTable="AGREEMENTDETAIL" then
		  txtColom = "SELECT AGREEMENTCODE, LINENO, AGREEMENTDETAIL.JOB,JOB.FCNAME, LOCATIONTYPE, LOCATIONCODE, DUEDATE, RETENTION, QTY, RATE " & _
		             " FROM AGREEMENTDETAIL,JOB " & _
		             " Where ( ((AGREEMENTDETAIL.JOB = JOB.FCCODE) AND (AGREEMENTDETAIL.FCBA = JOB.FCBA)) )" & _
					 " AND AGREEMENTDETAIL.fcba='" & session("fcBA") & "'"
		  txtPeriod= " and to_char(DUEDATE,'mm-yyyy')='" & Session("pContractor") & "'"
		  txtOrder = " ORDER BY AGREEMENTCODE"
	   elseif txtTable="WIP" then
		  txtColom = "SELECT AGREEMENTCODE, WIPDATE, AGREEMENTLINENO, QTY FROM WORKINPROGRESS" & _
		             " Where fcba='" & session("fcBA") & "'"
		  txtPeriod= " and to_char(WIPDATE,'mm-yyyy')='" & Session("pContractor") & "'"
		  txtOrder = " ORDER BY AGREEMENTCODE"

	   elseif txtTable="Nursery" then
		  txtColom = "SELECT NURSERYCODE, DESCRIPTION, DATEPLANTED, CROP, PLANTINGMATERIAL," & _
		             " CLONE, SOURCE, TYPE, STATUS, SOURCETYPE, SOURCECODE, QUANTITY FROM Nursery" & _
		             " Where fcba='" & session("fcBA") & "'"
		  txtOrder = " ORDER BY NURSERYCODE"
	   elseif txtTable="NurseryReceived" then
		  txtColom = "SELECT NURSERYCODE, RECEIVEDDATE, QUANTITY FROM NURSERYRECEIVED" & _
		             " Where fcba='" & session("fcBA") & "'"
		  txtPeriod= " AND to_char(RECEIVEDDATE,'mm-yyyy')='" & session("pNursery") & "'"
		  txtOrder = " ORDER BY NURSERYCODE"
	   elseif txtTable="NurseryCulling" then
		  txtColom = "SELECT NURSERYCODE, CULLINGDATE, QUANTITY FROM NURSERYCULLING" & _
		             " Where fcba='" & session("fcBA") & "'"
		  txtPeriod= " AND to_char(CULLINGDATE,'mm-yyyy')='" & session("pNursery") & "'"
		  txtOrder = " ORDER BY NURSERYCODE" 
	   elseif txtTable="NurseryIssue" then
		  txtColom = "SELECT NURSERYCODE, ISSUEDATE, LOCATIONTYPE, LOCATIONCODE, JOBCODE, QUANTITY FROM NURSERYISSUE" & _
		             " Where fcba='" & session("fcBA") & "'"
		  txtPeriod= " AND to_char(ISSUEDATE,'mm-yyyy')='" & session("pNursery") & "'"
		  txtOrder = " ORDER BY NURSERYCODE"

	   elseif txtTable="Harvesting" then
		  txtColom = "SELECT CHITNO, FIELDCODE, CROPCODE, PRODUCTCODE, COLLECTIONDATE,HARVESTINGDATE,RECEPTIONDATE, " & _
                     "VEHICLE, LOADERCODE, RAMPATTENDANTCODE,WEIGHT, QUANTITY  FROM Harvesting" & _
					 " where fcba='" & session("fcBA") & "'"
		  txtPeriod= " AND to_char(RECEPTIONDATE,'mm-yyyy')='" & session("pEmployee") & "'"
		  txtOrder = " ORDER BY RECEPTIONDATE"	  
	   elseif txtTable="HarvestingRound" then
		  txtColom = "SELECT HRNDYEAR, HRNDMONTH, FIELDCODE, CROPCODE, NUMBEROFROUNDS, HECTARAGE, STANDPERHECTHARVESTED FROM HarvestingRound" & _
					 " where fcba='" & session("fcBA") & "'"
		  txtPeriod= " AND HRNDMONTH='" & mid(session("pEmployee"),1,2) & "' and HRNDYEAR='" & mid(session("pEmployee"),4,4) & "'"
		  txtOrder = " ORDER BY FIELDCODE"	  
	   elseif txtTable="Weather" then
		  txtColom = "SELECT WIDATE, WIDIVISION, TOTRNINCHES, TOTRNDAYS, AVGRAINDAY, AVGTEMP, AVGWIND FROM Weather" & _
					 " where fcba='" & session("fcBA") & "'"
		  txtPeriod= " AND to_char(WIDATE,'mm-yyyy')='" & session("pEmployee") & "'"
		  txtOrder = " ORDER BY WIDATE"	  
	   elseif txtTable="WriteOff" then
		  txtColom = "SELECT * FROM WriteOff" & _
					 " where fcba='" & session("fcBA") & "'"
		  txtPeriod= " AND to_char(WODATE,'mm-yyyy')='" & session("pPurchase") & "'"
		  txtOrder = " ORDER BY WODATE"	  
	   elseif txtTable="WEIGHBRIDGETICKET" then
		  txtColom = "SELECT * FROM WEIGHBRIDGETICKET" & _
					 " where fcba='" & session("fcBA") & "'"
		  txtPeriod= " AND to_char(datein,'mm-yyyy')='" & session("pProduction") & "'"
		  txtOrder = " ORDER BY chitno"	  
	   elseif txtTable="GRADINGRECORD" then
		  txtColom = "SELECT * FROM GRADINGRECORD" & _
					 " where fcba='" & session("fcBA") & "'"
		  txtPeriod= " AND to_char(DateTime,'mm-yyyy')='" & session("pProduction") & "'"
		  txtOrder = " ORDER BY GradingNo"
	   elseif txtTable="SHIFTREPORT" then
		  txtColom = "SELECT * FROM SHIFTREPORT" & _
					 " where fcba='" & session("fcBA") & "'"
		  txtPeriod= " AND to_char(PRODUCTIONDATE,'mm-yyyy')='" & session("pProduction") & "'"
		  txtOrder = " ORDER BY ACTUALSTART"	  
	   elseif txtTable="PRODUCTQUALITY" then
		  txtColom = "SELECT * FROM PRODUCTQUALITY" & _
					 " where fcba='" & session("fcBA") & "'"
		  txtPeriod= " AND to_char(QUALITYDATE,'mm-yyyy')='" & session("pProduction") & "'"
		  txtOrder = " ORDER BY PSTORAGE_STORAGECODE"	
	   elseif txtTable="PRODUCTSTORAGEREADING" then
		  txtColom = "SELECT * FROM PRODUCTSTORAGEREADING" & _
					 " where fcba='" & session("fcBA") & "'"
		  txtPeriod= " AND to_char(READINGDATE,'mm-yyyy')='" & session("pProduction") & "'"
		  txtOrder = " ORDER BY READINGDATE"
	   elseif txtTable="DELIVERYCONTRACT" then
		  txtColom = "SELECT * FROM DELIVERYCONTRACT" & _
					 " where fcba='" & session("fcBA") & "'"
		  txtPeriod= " AND to_char(STARTDATE,'mm-yyyy')='" & session("pProduction") & "'"
		  txtOrder = " ORDER BY CONTRACTCODE"
	   elseif txtTable="DELIVERYNOTE" then
		  txtColom = "SELECT * FROM DELIVERYNOTE" & _
					 " where fcba='" & session("fcBA") & "'"
		  txtPeriod= " AND to_char(DNOTE_DATE,'mm-yyyy')='" & session("pProduction") & "'"
		  txtOrder = " ORDER BY DELIVERYNOTECODE"      
	   elseif txtTable="Allowded" then
		  txtColom = "SELECT fccode,description,jobcode from allowded har" & _
					 " WHERE har.fcba='" & session("fcBA") & "'"
		  txtOrder = " ORDER BY har.fccode"	  
	   else
	      txtColom = "SELECT * FROM " & txtTable 
	   end if
	   'Sebelumnya
	   'if txtWhere <> "" and txtPeriod<> "" and txtOrder <> ""  then
	   '   q = txtColom & txtWhere & txtOrder
	   'elseif txtWhere = "" and txtPeriod<> "" and txtOrder <> ""  then
	   '   q = txtColom & txtPeriod & txtOrder
	   'else
	   '   q = txtColom
	   'end if

	   if txtWhere <> "" then
	      q = txtColom & txtWhere & txtOrder
	   elseif txtWhere = "" then
	      q = txtColom & txtPeriod & txtOrder
	   else
	      q = txtColom
	   end if
	end if   
	'response.Write "sql : " & q
	%>
