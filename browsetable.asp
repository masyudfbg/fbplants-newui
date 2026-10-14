	<%
	Dim txtTable,txtModule,q,nCountofField,NO,txtSearch,txtField,txtColom,txtWhere,txtPeriod,txtOrder,txtValidForm,txtFcba,txtIsActive
	Dim txtDIVISION,txtSTATUS,txtCROPTYPE,txtSOILTYPE,txtTERRAINTYPE,txtOWNERSHIP,txtHARVESTINGSECTION,txtACTIVATION
	Dim i
	txtTable =Request("txtTable")
	txtValidForm =Request("txtValidForm")
	txtModule=Request("txtModule")
	txtSearch=Request("txtSearch")
	txtField =Request("txtField")
    txtWhere =Request("txtsort")
	txtDIVISION =Request("DIVISION")
	txtSTATUS =Request("STATUS")
	txtCROPTYPE =Request("CROPTYPE")
	txtSOILTYPE =Request("SOILTYPE")
	txtTERRAINTYPE =Request("TERRAINTYPE")
	txtOWNERSHIP =Request("OWNERSHIP")
	txtHARVESTINGSECTION =Request("HARVESTINGSECTION")
	txtACTIVATION =Request("ACTIVATION")
	txtFcba = Request("txtFcba")
	txtIsActive  =  Request("txtisActive")
	
    txtColom =""

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
	   elseif txtTable="Transporter" then
		  txtColom = "SELECT fccode, fcname, controljob, datecreated, activation, address," & _
		             " state, city, remarks, bankname, bankaccount FROM transporter" & _
		             " Where fcba='" & session("fcBA") & "'"
		  txtOrder = " ORDER BY FCCODE"
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
		  txtColom = "SELECT FCCODE, FCBA, FCNAME, ALLDED_GROUP, ALLDED, CROP,CONTROL_JOB, MAN_AUTO, TYPE_1, TYPE_2, TYPE_3, TYPE_4, MIN_AMT, MAX_AMT, RATE_1, RATE_2, RATE_3, RATE_4, REMARKS from ALLOWDED har" & _
					 " WHERE har.fcba='" & session("fcba") & "'"
		  txtOrder = " ORDER BY har.fccode"	  
       elseif txtTable="Gang" then
		  txtColom = "SELECT FCCODE, FCNAME, REMARKS, ACTIVATION, GANGTYPECODE from Gang " & _
					 " WHERE fcba='" & session("fcba") & "'"
		  txtOrder = " ORDER BY fccode"	 
	   elseif txtTable="UnitOfMeasure" then
		  txtColom = "SELECT UNITOFMEASURECODE, UNITOFMEASUREDESC from unitofmeasure " & _
					 " WHERE fcba='" & session("fcba") & "'"
		  txtOrder = " ORDER BY UNITOFMEASURECODE"	 
	   elseif txtTable="Building" then
		  txtColom = "SELECT FCCODE, FCNAME, BUILDGROUPCODE_DESC, ITEM, ACTIVATION from Building " & _
					 " WHERE fcba='" & session("fcba") & "'"
		  txtOrder = " ORDER BY fccode"	
	   elseif txtTable="GCMASTER" then
		  txtColom = "SELECT ITEM, FCCODE, FCNAME, GCGROUP, ACTIVATION from GCMASTER " & _
					 " WHERE fcba='" & session("fcba") & "'"
		  txtOrder = " ORDER BY fccode"	
	   elseif txtTable="Stock" then
		  txtColom = 	" SELECT FCCODE," & _
						"        FCNAME," & _
						"        ITEMTYPE," & _
						"        MANUFACTUREPARTNUMBER," & _
						"        MANUFACTURECODE," & _
						"        SUPPLIERPARTNUMBER," & _
						"        SUPPLIERCODE," & _
						"        UNITOFMEASUREMENT," & _
						"        STOCKGROUPCODE," & _
						"        ISLESSZERO " & _
						" FROM STOCK" & _
						" WHERE FCBA = '" & session("FCBA") & "'" & _
						"       and (validuntil is null or to_number(to_char(validuntil, 'yyyymmdd')) > to_number(to_char(sysdate, 'yyyymmdd')))"
		  txtOrder = 	" ORDER BY FCCODE"
		elseif txtTable="BankStock" then
		  txtColom = 	" SELECT FCCODE," & _
						"        FCNAME," & _
						"        ITEMTYPE," & _
						"        MANUFACTUREPARTNUMBER," & _
						"        MANUFACTURECODE," & _
						"        SUPPLIERPARTNUMBER," & _
						"        SUPPLIERCODE," & _
						"        UNITOFMEASUREMENT," & _
						"        STOCKGROUPCODE||'-'||stockgroup_api.get_stockgrpname(STOCKGROUPCODE,'ARJ') STOCKGROUPCODE," & _
						"        STOCKSUBGROUPCODE||'-'||stocksubgroup_api.get_stocksubgrpname(STOCKSUBGROUPCODE) STOCKSUBGROUPCODE," & _
						"        ISLESSZERO " & _
						" FROM BANKSTOCK" & _
						" WHERE isActive='" & txtIsActive & "' AND (validuntil is null or to_number(to_char(validuntil, 'yyyymmdd')) > to_number(to_char(sysdate, 'yyyymmdd')))"   
		  txtOrder = 	" ORDER BY FCCODE"
	   elseif txtTable="StockGroup" then
		    txtColom = 	" SELECT FCCODE, FCNAME, CONTROLJOB" & _
						"   FROM STOCKGROUP" & _
						" WHERE ACTIVATION = 'Y' AND FCBA = '" & session("FCBA") & "'"
		    txtOrder = 	" ORDER BY FCCODE"
	   elseif txtTable="Supplier" then
		    txtColom = 	" SELECT FCCODE, FCNAME, CONTROLJOB, DATECREATED, ACTIVATION, ADDRESS, CITY, STATE, REMARKS, BANKNAME, BANKACCOUNT, BANK_OWNERSHIP, PHONE, FAX, FCTYPE " & _
						"   FROM SUPPLIER" & _
						" WHERE FCBA = '" & session("FCBA") & "'"
		    txtOrder = 	" ORDER BY FCCODE"
	   elseif txtTable="Store" then
		    txtColom = 	" SELECT FCCODE, FCNAME, REMARKS, ACTIVATION" & _
						"   FROM STORE" & _
						" WHERE ACTIVATION = 'Y' AND FCBA = '" & session("FCBA") & "'"
		    txtOrder = 	" ORDER BY FCCODE"
	   elseif txtTable="SectionName" then
		    txtColom = 	" SELECT FCCODE, FCNAME" & _
						"   FROM SectionName" 
		    txtOrder = 	" ORDER BY FCCODE"
	   elseif txtTable="EmpPosition" then
		    txtColom = 	" SELECT POSITION, DESCRIPTION" & _
						"   FROM EmpPosition" & _
						" WHERE  FCBA = '" & session("FCBA") & "'"
		    txtOrder = 	" ORDER BY POSITION"
	   elseif txtTable="VehicleGroup" then
		    txtColom = 	" SELECT VEHICLEGROUPCODE, DESCRIPTION, UNITS, WORKSHOPFACTOR, COSTPERVEHICLE, NOOFVEHICLE, VEHICLETYPECODE" & _
						"   FROM VehicleGroup" & _
						" WHERE  FCBA = '" & session("FCBA") & "'"
		    txtOrder = 	" ORDER BY VEHICLEGROUPCODE"
	   elseif txtTable="OpHarvRate" then
		    txtColom = 	" SELECT CROPCODE, ABW_CENSUS, TERRAIN_TYPE, ABW_MIN, ABW_MAX, BASIS_QTY, BASIS_RATE, EXTRA_BASIS_RATE, BASIS_HA" & _
						"   FROM OpHarvRate" & _
						" WHERE  FCBA = '" & session("FCBA") & "'"
		    txtOrder = 	" ORDER BY CROPCODE"
	   elseif txtTable="Vehicle" then
		    txtColom = 	" SELECT FCCODE, FCNAME, VEHICLEGROUPCODE,VEHICLESUBGROUPCODE, DATECREATED, ACTIVATION, REGISTRATIONNO, MAKE" & _
						" , MODEL, ENGINENO, CHASISNO, YEAROFMADE, YEAROFPURCHASE, OWNERSHIP, VEHICLESTATUS" & _
						" , CONTROLALLOCATION, CONTROLJOB, DATETERMINATE" & _
						"   FROM VEHICLE" & _
						" WHERE ACTIVATION = 'Y' AND FCBA = '" & session("FCBA") & "'"
		    txtOrder = 	" ORDER BY FCCODE"
		elseif txtTable="Bank" then
		    txtColom = 	" SELECT FCCODE, FCNAME, CONTROLJOB, ACCOUNTCODE, ADDRESS, CITY, STATE, REMARKS" & _
						" , DATECREATED, ACTIVATION, TRIGGER_SYSTEM, BANKTYPE" & _
						"   FROM Bank" & _
						" WHERE ACTIVATION = 'Y' AND FCBA = '" & session("FCBA") & "'"
		    txtOrder = 	" ORDER BY FCCODE"
	    elseif txtTable="Field" then
			if  len(trim(txtValidForm)) = 0 then txtValidForm = textdate(date)
		    txtColom = " SELECT fccode" & _
			           "        , field_api.getValid_Name(fccode, fcba, to_date('" & txtValidForm & "', 'mm/dd/yyyy')) fcname" & _
			           "        , field_api.getValid_Afdeling(fccode, fcba, to_date('" & txtValidForm & "', 'mm/dd/yyyy')) division" & _
			           "        , field_api.getValid_HA(fccode, fcba, to_date('" & txtValidForm & "', 'mm/dd/yyyy')) hectarageplanted" & _
			           "        , field_api.GetValid_terraintype(fccode, fcba, to_date('" & txtValidForm & "', 'mm/dd/yyyy')) terraintype" & _
			           "        , field_api.getvalid_soiltype(fccode, fcba, to_date('" & txtValidForm & "', 'mm/dd/yyyy')) soiltype" & _
			           "        , field_api.GetValid_ownership(fccode, fcba, to_date('" & txtValidForm & "', 'mm/dd/yyyy')) ownership" & _
			           "        , field_api.GetValid_field_info_01(fccode, fcba, to_date('" & txtValidForm & "', 'mm/dd/yyyy')) koperasi" & _
			           "        , field_api.GetValid_Activation(fccode, fcba, to_date('" & txtValidForm & "', 'mm/dd/yyyy')) activation" & _
			           "        , field_api.GetValid_CropType(fccode, fcba, to_date('" & txtValidForm & "', 'mm/dd/yyyy')) croptype" & _
			           "        , field_api.getValid_PlantingDate(fccode, fcba, to_date('" & txtValidForm & "', 'mm/dd/yyyy')) plantingdate" & _
			           "        , field_api.getvalid_plantingmaterial(fccode, fcba, to_date('" & txtValidForm & "', 'mm/dd/yyyy')) plantingmaterial" & _
			           "        , field_api.getValid_POKOK(fccode, fcba, to_date('" & txtValidForm & "', 'mm/dd/yyyy')) totalstand" & _
			           "        , field_api.GetValid_status(fccode, fcba, to_date('" & txtValidForm & "', 'mm/dd/yyyy')) status" & _
			           "        , field_api.getvalid_prevsemester_abw(fccode, fcba, to_date('" & txtValidForm & "', 'mm/dd/yyyy')) prevsemester_abw" & _
			           "        , field_api.GetValid_pressemester_abw(fccode, fcba, to_date('" & txtValidForm & "', 'mm/dd/yyyy')) pressemester_abw" & _
			           "        , field_api.GetValid_harvestingbased_abw(fccode, fcba, to_date('" & txtValidForm & "', 'mm/dd/yyyy')) harvestingbased_abw" & _
			           "        , field_api.getvalid_main_road(fccode, fcba, to_date('" & txtValidForm & "', 'mm/dd/yyyy')) main_road" & _
			           "        , field_api.getvalid_collectionroad(fccode, fcba, to_date('" & txtValidForm & "', 'mm/dd/yyyy')) collectionroad" & _
			           "        , field_api.getvalid_maindrain(fccode, fcba, to_date('" & txtValidForm & "', 'mm/dd/yyyy')) maindrain" & _
			           "        , field_api.getvalid_collectiondrain(fccode, fcba, to_date('" & txtValidForm & "', 'mm/dd/yyyy')) collectiondrain" & _
			           "        , field_api.getvalid_subsidiarydrain(fccode, fcba, to_date('" & txtValidForm & "', 'mm/dd/yyyy')) subsidiarydrain" & _
			           "        , field_api.getvalid_collectionpoint(fccode, fcba, to_date('" & txtValidForm & "', 'mm/dd/yyyy')) collectionpoint" & _
			           "        , field_api.getvalid_harvestingbridge(fccode, fcba, to_date('" & txtValidForm & "', 'mm/dd/yyyy')) harvestingbridge" & _
			           "        , field_api.getvalid_permanentbridge(fccode, fcba, to_date('" & txtValidForm & "', 'mm/dd/yyyy')) permanentbridge" & _
			           "        , field_api.getvalid_woodbridge(fccode, fcba, to_date('" & txtValidForm & "', 'mm/dd/yyyy')) woodbridge" & _
			           "        , field_api.getvalid_buist(fccode, fcba, to_date('" & txtValidForm & "', 'mm/dd/yyyy')) buist" & _
			           "        , field_api.getvalid_owlnest(fccode, fcba, to_date('" & txtValidForm & "', 'mm/dd/yyyy')) owlnest" & _
			           "        , field_api.GetValid_isactive_harvested(fccode, fcba, to_date('" & txtValidForm & "', 'mm/dd/yyyy')) isactive_harvested" & _
			           "        , field_api.GetValid_ValidFrom(fccode, fcba, to_date('" & txtValidForm & "', 'mm/dd/yyyy')) valid_from" & _
			           "        , field_api.getvalid_fcentry(fccode, fcba, to_date('" & txtValidForm & "', 'mm/dd/yyyy')) fcentry" & _
			           "        , field_api.getvalid_fcedit(fccode, fcba, to_date('" & txtValidForm & "', 'mm/dd/yyyy')) fcedit" & _
			           "        , field_api.getvalid_fcip(fccode, fcba, to_date('" & txtValidForm & "', 'mm/dd/yyyy')) fcip" & _
			           "        , '" & session("FCBA") & "' fcba" & _
			           "        , field_api.getvalid_lastupdate(fccode, fcba, to_date('" & txtValidForm & "', 'mm/dd/yyyy')) lastupdate" & _
			           "        , field_api.getvalid_lasttime(fccode, fcba, to_date('" & txtValidForm & "', 'mm/dd/yyyy')) lasttime" & _
			           " from field"
			if  LEN(TRIM(txtValidForm)) > 0 then	
				varVALIDFORM = " to_number(to_char(field_api.getvalid_validfrom(fccode,fcba,to_date('" & _
				txtValidForm & "','mm/dd/yyyy')),'YYYYMMDD'))=to_number(to_char(field_api.GetValid_ValidFrom(fccode, fcba, to_date('" & txtValidForm & "', 'mm/dd/yyyy')),'YYYYMMDD')) AND"
				else 
				varVALIDFORM = ""
			end if
				
			if 	LEN(TRIM(txtDIVISION)) > 0 then
				varDIVISION = " field_api.getValid_Afdeling(fccode, fcba, to_date('" & txtValidForm & "', 'mm/dd/yyyy')) = '" & txtDIVISION & "' AND"
				else 
				varDIVISION = ""
			end if
			if 	LEN(TRIM(txtSTATUS)) > 0 then
				varSTATUS = " field_api.GetValid_status(fccode, fcba, to_date('" & txtValidForm & "', 'mm/dd/yyyy')) = '" & txtSTATUS & "' AND"
				else 
				varSTATUS = ""
			end if
			if 	LEN(TRIM(txtCROPTYPE)) > 0 then
				varCROPTYPE = " field_api.GetValid_CropType(fccode, fcba, to_date('" & txtValidForm & "', 'mm/dd/yyyy')) = '" & txtCROPTYPE & "' AND"
				else 
				varCROPTYPE = ""
			end if
			if 	LEN(TRIM(txtSOILTYPE)) > 0 then
				varSOILTYPE = " field_api.getvalid_soiltype(fccode, fcba, to_date('" & txtValidForm & "', 'mm/dd/yyyy')) = '" & txtSOILTYPE & "' AND"
				else 
				varSOILTYPE = ""
			end if
			if 	LEN(TRIM(txtTERRAINTYPE)) > 0 then
				varTERRAINTYPE = " field_api.GetValid_terraintype(fccode, fcba, to_date('" & txtValidForm & "', 'mm/dd/yyyy')) = '" & txtTERRAINTYPE & "' AND"
				else 
				varTERRAINTYPE = ""
			end if
			if 	LEN(TRIM(txtOWNERSHIP)) > 0 then
				varOWNERSHIP = " field_api.GetValid_ownership(fccode, fcba, to_date('" & txtValidForm & "', 'mm/dd/yyyy')) = '" & txtOWNERSHIP & "' AND"
				else 
				varOWNERSHIP = ""
			end if
			if 	LEN(TRIM(txtHARVESTINGSECTION)) > 0 then
				varHARVESTINGSECTION = " field_api.GetValid_Harvestingsection(fccode, fcba, to_date('" & txtValidForm & "', 'mm/dd/yyyy')) = '" & txtHARVESTINGSECTION & "' AND"
				else 
				varHARVESTINGSECTION = ""
			end if
			if 	LEN(TRIM(txtACTIVATION)) > 0 then
				varACTIVATION = " field_api.GetValid_Activation(fccode, fcba, to_date('" & txtValidForm & "', 'mm/dd/yyyy')) = '" & txtACTIVATION & "' AND"
				else 
				varACTIVATION = ""
			end if

			txtColom = txtColom & " WHERE" & varVALIDFORM & varDIVISION & varSTATUS & varCROPTYPE & varSOILTYPE & _
					   varTERRAINTYPE & varOWNERSHIP & varHARVESTINGSECTION & varACTIVATION & " FCBA = '" & session("FCBA") & "'"			
		    txtOrder = 	" ORDER BY FCCODE"
			 'response.write (txtColom)
	    elseif txtTable= "RangeTaxGroup" then
		if REQUEST("FCCODE") = "" then
		    val_FCCODE = ""
		else
			val_FCCODE = " AND FCCODE='" & REQUEST("FCCODE") & "'"
		 end if
		  txtColom = 	" SELECT" & _
						" FCCODE, FCNAME, RANGEFROM, RANGEUNTIL, TARIFFPERCENTAGE," & _
						" REMARKS FROM " & txtTable & "" & _
						" WHERE FCBA = '" & Session("FCBA") & "'" & _
						val_FCCODE
		  txtOrder = 	" ORDER BY FCCODE"
	   elseif txtTable= "MaritalStatus" then
	   if REQUEST("FCCODE") = "" then
		    val_FCCODE = ""
			else
			val_FCCODE = " WHERE FCCODE='" & REQUEST("FCCODE") & "'"
		 end if
		  txtColom = 	" SELECT" & _
		                " FCCODE, FCNAME" & _
						"  FROM " & txtTable & "" & _
						val_FCCODE
	      txtOrder = 	" ORDER BY FCCODE"
	   elseif txtTable="UntaxableIncome" then
	   if REQUEST("FCCODE") = "" then
		    val_MARITALSTATUSCODE = ""
			else
			val_MARITALSTATUSCODE = " AND MARITALSTATUSCODE='" & REQUEST("FCCODE") & "'"
		 end if
		  txtColom = 	" SELECT" & _
						" MARITALSTATUSCODE, FCNAME, UTI_EMPLOYEE, UTI_SPOUSE, UTI_CHILD," & _
						" REMARKS FROM " & txtTable & "" & _
						" WHERE FCBA = '" & Session("FCBA") & "'" & _
						val_MARITALSTATUSCODE
		  txtOrder = 	" ORDER BY MARITALSTATUSCODE"

	   elseif txtTable="EmployeeAllowdedDetail" then
	   
	      if REQUEST("ALLOWDEDCODE") = "" then
		    val_ALLOWDEDCODE = ""
			else
			val_ALLOWDEDCODE = " AND ALLOWDEDCODE='" & REQUEST("ALLOWDEDCODE") & "'"
		 end if
		 if REQUEST("FCCODE") = "" then
		    val_EMPLOYEECODE = ""
			else
			val_EMPLOYEECODE = " AND EMPLOYEECODE='" & REQUEST("FCCODE") & "'"
		 end if
		     
			txtColom = " select employeecode" & _
			           "        , employee_api.get_employeename(fcba, employeecode) employeename" & _
			           "        , allowdedcode" & _
			           "        , allowded_api.get_fcname(allowdedcode, fcba) allowdedname" & _
			           "        , rate_1" & _
			           "        , rate_2" & _
			           "        , rate_3" & _
			           "        , rate_4" & _
			           "        , rate_5" & _
			           "        , remarks" & _
			           "        , fcentry" & _
			           "        , fcedit" & _
			           "        , fcip" & _
			           "        , fcba" & _
			           "        , lastupdate" & _
			           "        , lasttime" & _
			           " from " & txttable & "" & _
			           " where fcba = '" & Session("FCBA") & "'" & _
						val_ALLOWDEDCODE + val_EMPLOYEECODE
			txtOrder = " order by employeecode"				
		
	   elseif txtTable="Employee" then
			txtColom = " select e.fccode" & _
			           "        , e.fcname" & _
			           "        , e.sectionname" & _
			           "        , e.gangcode" & _
			           "        , to_char(e.dateofbirth,'dd/mm/yyyy')dateofbirth " & _
			           "        , p.fcname religion" & _
			           "        , e.race" & _
			           "        , e.identitycard" & _
			           "        , e.highesteducation" & _
			           "        , e.yearofgraduation" & _
			           "        , to_char(e.dateofjoin,'dd/mm/yyyy') dateofjoin" & _
			           "        , to_char(e.dateofdesignation,'dd/mm/yyyy')dateofdesignation" & _
			           "        , e.presentgrade" & _
			           "        , e.presentgrade_detail" & _
			           "        , e.presentrate" & _
			           "        , e.previousgrade" & _
			           "        , e.previousgrade_detail" & _
			           "        , e.previousrate" & _
			           "        , e.maritalstatus" & _
			           "        , e.spouse" & _
			           "        , e.dependentchild" & _
			           "        , e.jamsostekno" & _
			           "        , e.jamsostekcategory" & _
			           "        , e.position" & _
			           "        , e.dateterminate" & _
			           "        , e.nationality" & _
			           "        , e.npwp" & _
			           "        , e.naturetype" & _
			           "        , e.includedtaxcalculation" & _
			           "        , e.taxstatus" & _
			           "        , e.paymenttype" & _
			           "        , e.bankname" & _
			           "        , e.accountno" & _
			           "        , e.ownershipofaccount" & _
			           "        , e.gender" & _
			           "        , e.taxpaidbycompany" & _
			           "        , e.address" & _
			           "        , e.bpjskes_category" & _
					   "        , e.bpjskes_no" & _
					   "        , e.subgrade" & _
					   "        , e.company_nik" & _ 
                       " from employee e" & _
					   " INNER JOIN PARAMETER p ON p.fccode=e.religion and p.fcba=e.fcba" & _
                       " where e.fcba = '" & session("FCBA") & "' AND p.PARCODE='E0003'  AND e.dateterminate IS NULL"
			txtOrder = " order by e.fcname "
	   elseif txtTable="Job" then
			if 	LEN(TRIM(txtSearch)) > 0 then
				SELECT CASE TRIM(txtSearch) 
				CASE "BKM"  
					txtColom = 	" SELECT FCCODE, FCNAME, JOBGROUPCODE, JOB_FORCROPTYPE, JOB_FORFIELDSTATUS, JOB_CATEGORY, JOB_FORBALANCESHEETNAME, JOB_FORPROFITANDLOSSNAME,  JOB_FORCASHFLOWNAME,UNITOFMEASUREMENT, ACCOUNTCODE " & _
								" FROM JOB" & _
								" WHERE  JobGroupCode in ('1114000','1310100','1310200','1310300','1310400','1310400','1310500', " & _
								"                         '1320100','1320200','1320300','1320400','1320500','1320600','1320700','1320800','1320900','1321000','1321100','1321200','1321300','1321400', " & _
								"						  '6310000','6210000','6400000', " & _
								"						  '6500100','6500200','6500300','6500400','6500500','6500600','6500700','6500800','6500900','6501000','6502100','6502200','6610100') AND ACTIVATION = 'Y' AND FCBA = '" & session("FCBA") & "'"
					txtOrder = 	" ORDER BY FCCODE"
				CASE "SIV"  
					txtColom = 	" SELECT FCCODE, FCNAME, JOBGROUPCODE, JOB_FORCROPTYPE, JOB_FORFIELDSTATUS, JOB_CATEGORY, JOB_FORBALANCESHEETNAME, JOB_FORPROFITANDLOSSNAME,  JOB_FORCASHFLOWNAME, UNITOFMEASUREMENT, ACCOUNTCODE " & _
								" FROM JOB" & _
								" WHERE  JobGroupCode in ('1114000','1310100','1310200','1310300','1310400','1310400','1310500', " & _
								"                         '1320100','1320200','1320300','1320400','1320500','1320600','1320700','1320800','1320900','1321000','1321100','1321200','1321300','1321400', " & _
								"						  '1210000','6310000','6210000','6400000','1130000', " & _
								"						  '6500100','6500200','6500300','6500400','6500500','6500600','6500700','6500800','6500900','6501000','6502100','6502200','6610100') AND ACTIVATION = 'Y' AND FCBA = '" & session("FCBA") & "'"
					txtOrder = 	" ORDER BY FCCODE"
			    CASE ELSE			
                    txtColom = 	" SELECT FCCODE, FCNAME, JOBGROUPCODE, JOB_FORCROPTYPE, JOB_FORFIELDSTATUS, JOB_CATEGORY, JOB_FORBALANCESHEETNAME, JOB_FORPROFITANDLOSSNAME,  JOB_FORCASHFLOWNAME, UNITOFMEASUREMENT, ACCOUNTCODE " & _
								" FROM JOB" & _
								" WHERE ACTIVATION = 'Y' AND FCBA = '" & session("FCBA") & "'"
					txtOrder = 	" ORDER BY FCCODE"				

                END SELECT				 
		  Else 
				    txtColom = 	" SELECT FCCODE, FCNAME, JOBGROUPCODE, JOB_FORCROPTYPE, JOB_FORFIELDSTATUS, JOB_CATEGORY, JOB_FORBALANCESHEETNAME, JOB_FORPROFITANDLOSSNAME,  JOB_FORCASHFLOWNAME, UNITOFMEASUREMENT, ACCOUNTCODE " & _
								" FROM JOB" & _
								" WHERE ACTIVATION = 'Y' AND FCBA = '" & session("FCBA") & "'"
					txtOrder = 	" ORDER BY FCCODE"				
		  End if
	      
		elseif txtTable="Jobmaster" then
			if 	LEN(TRIM(txtSearch)) > 0 then
				SELECT CASE TRIM(txtSearch) 
				CASE "BKM"  
					txtColom = 	" SELECT FCCODE, FCNAME, JOBGROUPCODE, JOB_FORCROPTYPE, JOB_FORFIELDSTATUS, JOB_CATEGORY, JOB_FORBALANCESHEETCODE||'-'||JOB_FORBALANCESHEETNAME JOB_FORBALANCESHEETNAME, JOB_FORPROFITANDLOSSCODE||'-'||JOB_FORPROFITANDLOSSNAME JOB_FORPROFITANDLOSSNAME, JOB_FORCASHFLOWCODE||'-'||JOB_FORCASHFLOWNAME JOB_FORCASHFLOWNAME,JOB_FORBUDGETCODE||'-'||JOB_FORBUDGETNAME JOB_FORBUDGETNAME,UOM_UNIT,  UNITOFMEASUREMENT, ACCOUNTCODE " & _
								" FROM JOBMASTER" & _
								" WHERE  JobGroupCode in ('1114000','1310100','1310200','1310300','1310400','1310400','1310500', " & _
								"                         '1320100','1320200','1320300','1320400','1320500','1320600','1320700','1320800','1320900','1321000','1321100','1321200','1321300','1321400', " & _
								"						  '6310000','6210000','6400000', " & _
								"						  '6500100','6500200','6500300','6500400','6500500','6500600','6500700','6500800','6500900','6501000','6502100','6502200','6610100') AND ACTIVATION = 'Y' AND FCBA = '" & session("FCBA") & "'"
					txtOrder = 	" ORDER BY FCCODE"
				CASE "SIV"  
					txtColom = 	" SELECT FCCODE, FCNAME, JOBGROUPCODE, JOB_FORCROPTYPE, JOB_FORFIELDSTATUS, JOB_CATEGORY, JOB_FORBALANCESHEETCODE||'-'||JOB_FORBALANCESHEETNAME JOB_FORBALANCESHEETNAME, JOB_FORPROFITANDLOSSCODE||'-'||JOB_FORPROFITANDLOSSNAME JOB_FORPROFITANDLOSSNAME, JOB_FORCASHFLOWCODE||'-'||JOB_FORCASHFLOWNAME JOB_FORCASHFLOWNAME,JOB_FORBUDGETCODE||'-'||JOB_FORBUDGETNAME JOB_FORBUDGETNAME, UOM_UNIT, UNITOFMEASUREMENT, ACCOUNTCODE " & _
								" FROM JOBMASTER" & _
								" WHERE  JobGroupCode in ('1114000','1310100','1310200','1310300','1310400','1310400','1310500', " & _
								"                         '1320100','1320200','1320300','1320400','1320500','1320600','1320700','1320800','1320900','1321000','1321100','1321200','1321300','1321400', " & _
								"						  '1210000','6310000','6210000','6400000','1130000', " & _
								"						  '6500100','6500200','6500300','6500400','6500500','6500600','6500700','6500800','6500900','6501000','6502100','6502200','6610100') AND ACTIVATION = 'Y' AND FCBA = '" & session("FCBA") & "'"
					txtOrder = 	" ORDER BY FCCODE"
			    CASE ELSE			
                    txtColom = 	" SELECT FCCODE, FCNAME, JOBGROUPCODE, JOB_FORCROPTYPE, JOB_FORFIELDSTATUS, JOB_CATEGORY, JOB_FORBALANCESHEETCODE||'-'||JOB_FORBALANCESHEETNAME JOB_FORBALANCESHEETNAME, JOB_FORPROFITANDLOSSCODE||'-'||JOB_FORPROFITANDLOSSNAME JOB_FORPROFITANDLOSSNAME, JOB_FORCASHFLOWCODE||'-'||JOB_FORCASHFLOWNAME JOB_FORCASHFLOWNAME,JOB_FORBUDGETCODE||'-'||JOB_FORBUDGETNAME JOB_FORBUDGETNAME, UOM_UNIT, UNITOFMEASUREMENT, ACCOUNTCODE " & _
								" FROM JOBMASTER" & _
								" WHERE ACTIVATION = 'Y' "
					txtOrder = 	" ORDER BY FCCODE"				

                END SELECT				 
		  Else 
				    txtColom = 	" SELECT FCCODE, FCNAME, JOBGROUPCODE, JOB_FORCROPTYPE, JOB_FORFIELDSTATUS, JOB_CATEGORY, JOB_FORBALANCESHEETCODE||'-'||JOB_FORBALANCESHEETNAME JOB_FORBALANCESHEETNAME, JOB_FORPROFITANDLOSSCODE||'-'||JOB_FORPROFITANDLOSSNAME JOB_FORPROFITANDLOSSNAME, JOB_FORCASHFLOWCODE||'-'||JOB_FORCASHFLOWNAME JOB_FORCASHFLOWNAME,JOB_FORBUDGETCODE||'-'||JOB_FORBUDGETNAME JOB_FORBUDGETNAME, UOM_UNIT,  UNITOFMEASUREMENT, ACCOUNTCODE " & _
								" FROM JOBMASTER" & _
								" WHERE ACTIVATION = 'Y' "
					txtOrder = 	" ORDER BY FCCODE"				
		  End if			
	   elseif txtTable="JobGroup" then
		  txtColom = 	" SELECT FCCODE, FCNAME, PARENTCODE, PARENTDESC" & _
						"   FROM JOBGROUP" & _
						" WHERE FCBA = '" & session("FCBA") & "'"
		  txtOrder = 	" ORDER BY FCCODE"
	   elseif txtTable="JobGroupMaster" then
		  txtColom = 	" SELECT FCCODE, FCNAME, PARENTCODE, PARENTDESC" & _
						"   FROM JOBGROUPMASTER" & _
		  txtOrder = 	" ORDER BY FCCODE"
	   elseif txtTable="BusinessUnit" then
	      txtColom = 	"SELECT FCCODE, FCNAME, FCCOMPANYCODE, FCCOMPANYNAME, FCAREA, FCOWNERSHIP, FCMANAGER, FCTELEPHONE " & _
	      				" FROM BUSINESSUNIT"  
	      txtOrder = 	" ORDER BY FCCODE"
	   elseif txtTable="Trigger_Control" then
	      txtColom = 	"SELECT ACTIVITY, SOURCE, TARGET " & _
	      				" FROM TRIGGER_CONTROL"  
	      txtOrder = 	" ORDER BY ACTIVITY"		
	   elseif txtTable="LocationType" then
	      txtColom = 	"SELECT LOCATIONTYPECODE, LOCATIONTYPENAME, ITEM " & _
	      				" FROM LocationType"  
	      txtOrder = 	" ORDER BY LOCATIONTYPECODE"		
	   elseif txtTable="LocationCode" then
	      if 	LEN(TRIM(txtSearch)) > 0 then
				SELECT CASE TRIM(txtSearch) 
				CASE "FF"  
								txtColom = " SELECT FCCODE,TRIM(FCNAME)||'-'||TRIM(DIVISION)||'-'||TRIM(DX_BLOCK_STATUS) FCNAME FROM FIELD WHERE fcba = '" & txtFcba & "' AND ACTIVATION='Y'"
				                txtOrder = " ORDER BY FCCODE "					
				CASE "CB"  
								txtColom = " SELECT FCCODE,FCNAME FROM gcmaster WHERE fcba = '" & txtFcba & "' AND ACTIVATION='Y' AND Item='CB'"
				                txtOrder = " ORDER BY FCCODE "					
				CASE "GC"  
								txtColom = " SELECT FCCODE,FCNAME FROM gcmaster WHERE fcba = '" & txtFcba & "' AND ACTIVATION='Y' AND Item='GC'"
				                txtOrder = " ORDER BY FCCODE "					
				CASE "BA"  
								txtColom = " SELECT FCCODE,FCNAME FROM bank WHERE fcba = '" & txtFcba & "' AND ACTIVATION='Y'"
				                txtOrder = " ORDER BY FCCODE "		
                CASE "MA"  
								txtColom = " SELECT FCCODE,Description FCNAME FROM Machine WHERE fcba = '" & txtFcba & "' AND ACTIVATION='Y'"
				                txtOrder = " ORDER BY FCCODE "	
				CASE "VH"  
								txtColom = " SELECT FCCODE,FCNAME FROM Vehicle WHERE fcba = '" & txtFcba & "' AND ACTIVATION='Y'"
				                txtOrder = " ORDER BY FCCODE "	
				CASE "ST"  
								txtColom = " SELECT FCCODE,FCNAME FROM store WHERE fcba = '" & txtFcba & "' AND ACTIVATION='Y'"
				                txtOrder = " ORDER BY FCCODE "	
				CASE "MN"  
								txtColom = " SELECT NurseryCode FCCODE,Description FCNAME FROM nursery WHERE fcba = '" & txtFcba & "' AND Status='Y'"
				                txtOrder = " ORDER BY NurseryCode "	
				CASE "PN"  
								txtColom = " SELECT NurseryCode FCCODE,Description FCNAME FROM nursery WHERE fcba = '" & txtFcba & "' AND Status='Y'"
				                txtOrder = " ORDER BY NurseryCode "	
	            CASE "BM"  
								txtColom = " SELECT FCCODE,FCNAME FROM building WHERE fcba = '" & txtFcba & "' AND ACTIVATION='Y'"
				                txtOrder = " ORDER BY FCCODE "			
				CASE "AP"  
								txtColom = " SELECT FCCODE,FCNAME FROM Supplier WHERE fcba = '" & txtFcba & "' AND ACTIVATION='Y' UNION ALL SELECT FCCODE,FCNAME FROM Transporter WHERE fcba = '" & txtFcba & "' AND ACTIVATION='Y' "
				                txtOrder = " ORDER BY FCCODE "			
				CASE "AR"  
								txtColom = " SELECT CustomerCode FCCODE,Description FCNAME FROM customer WHERE fcba = '" & txtFcba & "' AND isactive='Y'"
				                txtOrder = " ORDER BY CustomerCode "			
				CASE "CA"  
								txtColom = " SELECT FCCODE,FCNAME FROM Contractor WHERE fcba = '" & txtFcba & "' AND ACTIVATION='Y'"
				                txtOrder = " ORDER BY FCCODE "			
				CASE "FA"  
								txtColom = " SELECT FCCODE,FCNAME FROM Fixedasset WHERE fcba = '" & txtFcba & "' AND isactive='Y'"
				                txtOrder = " ORDER BY FCCODE "	
                CASE "WS"  
								txtColom = " SELECT FCCODE,FCNAME FROM Workshop WHERE fcba = '" & txtFcba & "' AND ACTIVATION='Y'"
				                txtOrder = " ORDER BY FCCODE "									
				CASE ELSE				
	 							
				txtColom =  " SELECT '' FCCODE, '-' FCNAME " & _
	      				    " FROM DUAL"  
	             txtOrder = ""		

                END SELECT				 
		  Else 
				txtColom =  " SELECT '' FCCODE, '-' FCNAME " & _
	      				    " FROM DUAL"  
	             txtOrder = ""		
		  End if
	      
		  
	   ' Special for Administrator, Added by NAF, 2 May 2013	  
	   elseif txtTable="Roles" then
	      txtColom = " select fccode, fcname from roles_tab "
		  txtOrder = "  order by fccode"
	   elseif txtTable="RolesMenu" then
		  txtColom = " select a.fcrolescode, a.fcmenucode, b.fcname from " & _
						" menu_roles_tab a, " & _
						" menu_tab b " & _
						" where " & _
						" a.fcmenucode = b.fccode "
		  txtOrder = "  order by a.fcrolescode, a.fcmenucode"
	   elseif txtTable="Users" then
	      txtColom = " select fccode, fcpassword, fullname, department, accessrights, isactive, valid_from, valid_until from users_tab " 
	      txtOrder = " order by fccode "
	   elseif txtTable="Users" then
	      txtColom = " select fccode, fcpassword, fullname, department, accessrights, isactive, valid_from, valid_until from users_tab " 
	      txtOrder = " order by fccode "
	   elseif txtTable="UsersRoles" then
	      txtColom = " select FCUSERCODE, FCROLESCODE, ISACTIVE, VALID_FROM, VALID_UNTIL from user_roles_tab " 
	      txtOrder = " order by fcrolescode, fcusercode "
	   elseif txtTable="UsersCompany" then
	      txtColom = " select FCUSERCODE, FCCOMPANYCODE, ISDEFAULT from user_company_tab " 
	      txtOrder = " order by FCUSERCODE, FCCOMPANYCODE, ISDEFAULT "
	   elseif txtTable="UsersBusinessUnit" then
	      txtColom = " select FCUSERCODE, FCBA, ISDEFAULT from user_fcba_tab " 
	      txtOrder = " order by FCUSERCODE, FCBA, ISDEFAULT "		  
	   elseif txtTable="UsersChanged" then
	      txtColom = "  SELECT a.fcusercode, a.fcba, b.fccompanycode " & _
							" FROM (SELECT fcba, fcusercode " & _
									" FROM user_fcba_tab " & _
								   " WHERE isdefault = 'TRUE') a,  " & _
								 " (SELECT fccompanycode, fcusercode  " & _
									" FROM user_company_tab  " & _
								   " WHERE isdefault = 'TRUE') b  " & _
						   " WHERE a.fcusercode = b.fcusercode  " 
	      txtOrder = " ORDER BY a.fcusercode  " 	
   
	   ' End of Added by NAF 
	   
	   else
			txtColom = 	"SELECT * FROM " & txtTable & _
						" WHERE FCBA='" & session("fcba") & "'"
		'	txtOrder = 	" ORDER BY FCCODE"
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
	
	'response.Write "sql : " & q
	'RESPONSE.Flush()
	%>
