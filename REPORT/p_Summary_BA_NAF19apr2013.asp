<%@Language=VBScript%>
<% Option Explicit %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%
	Dim v_txtBudgetYear, v_txtBudgetVersion, v_txtModuleVersion
	Dim txtBudgetYear, txtBudgetVersion, txtModuleVersion
	Dim v_txtdatefrom, v_txtdateto,  v_txtfcba, v_txtyear, v_txtmonth,v_txtComp
	Dim txtdatefrom, txtdateto,  txtfcba, txtyear, txtmonth , txtComp
	Dim strProduction, strHandling, strHandlingAmount, strFertilizer, strEndingBalance, strMandays, strVehicleGroup
	Dim rsProduction, rsHandling, rsHandlingAmount, rsFertilizer, rsEndingBalance, rsMandays, rsVehicleGroup
	Dim strSql, no, cClass, tempData
	Dim tempRefCode, strHRef, rsHRef, tempReference, tempExtendedReference, tempRowSpan
	Dim strSQLPreParameter, rsSQLPreParameter
	Dim strCodeDesc, rsCodeDesc, tempCodeDescription
	Dim sPrevDesc
	Dim tempRounding
    Dim tempBudget_HI, tempBudget_SHI, tempBudget_SBI, tempBudgetType
	Dim strBudget, rsBudget 
					
	
	Dim tempUOM, strOutput, rsOutput
	
	Dim tempDataBudget_HI, tempDataBudget_SHI, tempDataBudget_SBI
	Dim tempCompare_HI, tempCompare_SHI, tempCompare_SBI
	Dim cClassAssigned_HI, cClassAssigned_SHI, cClassAssigned_SBI
	Dim strLandClearing, rsLandClearing

	Dim tempTBSProduksi, tempTBSProduksi_shi, tempTBSProduksi_sbi
	Dim tempBiayaAngkut, tempBiayaAngkut_shi, tempBiayaAngkut_sbi
	Dim tempData_shi, tempData_sbi 	
	Dim rscomp , strcomp,txtcompany ,txcompanyname,str1
	
    
			




	' DateTo and DateFrom will be EQUAL to remove miss-understanding
	
	txtdatefrom			= Request("txtdatefrom")
	txtdateto			= Request("txtdatefrom")
    txtfcba 			= Request("txtfcba")
	txtyear 			= Request("txtyear")
	txtmonth 			= Request("txtmonth")
	
	txtcomp          	= Request("txtComp") 
	txtcompany  		= Request("txtComp")
	txtBudgetYear 		= Request("txtBudgetYear")
	txtBudgetVersion	= Request("txtBudgetVersion")
	txtModuleVersion	= Request("txtModuleVersion")

	strSQLPreParameter = "select to_char(to_date('" & txtdatefrom & "', 'MM/DD/YYYY'),'YYYY') as CurrentYear, " & _
	       "to_char(to_date('" & txtdatefrom & "', 'MM/DD/YYYY'),'MM') as CurrentPeriod from dual "
	
	Set rsSQLPreParameter = Server.CreateObject("ADODB.Recordset")
	rsSQLPreParameter.Open strSQLPreParameter, connstr, adOpenStatic 
	
    if not rsSQLPreParameter.EOF then
		txtyear = rsSQLPreParameter("CurrentYear")
		txtmonth = rsSQLPreParameter("CurrentPeriod")
	end if
	
     
	'response.write  txtcompany
	
	 strcomp= " select a.fccompanycode,a.fccompanyname from BUSINESSUNIT a  where a.fccompanycode = '" & txtcomp & "'"
	Set rscomp = Server.CreateObject("ADODB.Recordset")
	rscomp.Open  strcomp,connstr, adOpenStatic
	if not rscomp.eof then
		txcompanyname   = 	rscomp("fccompanyname")
	End if
	rscomp.Close
	set rscomp = nothing	
	
	v_txtdatefrom		=	txtdatefrom
	v_txtdateto			=	txtdateto
	v_txtfcba     		=   txtfcba
	v_txtyear 			=   txtyear
	v_txtmonth			=   txtmonth
	v_txtComp			=   txtComp
	v_txtBudgetYear 	= 	txtBudgetYear
	v_txtBudgetVersion 	= 	txtBudgetVersion
	v_txtModuleVersion 	= 	txtModuleVersion
	
	if len(trim(v_txtFcba)) > 0 then 
		if v_txtFcba <> "--ALL--" then
	        str1 = str1 & " AND fcba = '" & v_txtFcba& "' " 
		end if
	end if
	'response.write str1
	
	
	' Zero 
	' Land Clearing
	
	strLandClearing = " select codedesc, ROWNUM thenumber, JobGenCode, description, unitofmeasurement, hi, shi, sbi, budget_hi, budget_shi, budget_sbi FROM ( " & _ 
                           " SELECT 'SR003_LC' codedesc,  " & _
                                  " JobGenCode,   " & _
                                  " JobGenName description,   " & _
                                  " orderdata,   " & _
                                  " UnitOfMeasurement,   " & _
                                  " SUM (hi) hi,   " & _
                                  " SUM (shi) shi,   " & _
                                  " SUM (sbi) sbi,   " & _
								  " SUM (budget_hi) budget_hi,   " & _
                                  " SUM (budget_shi) budget_shi,   " & _
                                  " SUM (budget_sbi) budget_sbi   " & _
                             " FROM (SELECT fcba, FCCODE JobGenCode,   " & _
                                          " FCNAME JobGenName,   " & _
                                          " param_2 orderdata,   " & _
                                          " param_3 unitofmeasurement, " & _  
                                          " 0 hi,   " & _
                                          " 0 shi,   " & _
                                          " 0 sbi,   " & _
										  " 0 budget_hi,   " & _
                                          " 0 budget_shi,   " & _
                                          " 0 budget_sbi   " & _
                                     " FROM PARAMETER   " & _
                                    " WHERE PARCODE = 'SR003_LC' AND PARAM_1 = 'DEFAULT'   " & _
									" UNION " & _
									 " SELECT fcba,jobGenCode, " & _
											" JobGenName, " & _
											" OrderData, " & _
											" UnitOfMeasureMent, " & _
											" SUM (hi) hi, " & _
											" SUM (shi) shi, " & _
											" SUM (sbi) sbi, " & _
											" SUM (budget_hi) hi, " & _
											" SUM (budget_shi) shi, " & _
											" SUM (budget_sbi) sbi " & _
										" FROM (  SELECT fcba,jobGenCode,  " & _
													   " JobGenName,  " & _
													   " OrderData,  " & _
													   " UnitOfMeasureMent,  " & _
													   " SUM (qty) hi,  " & _
													   " 0 shi,  " & _
													   " 0 sbi,  " & _
													   " SUM (budget_qty) budget_hi,  " & _
													   " 0 budget_shi,  " & _
													   " 0 budget_sbi  " & _
												  " FROM (SELECT b.fcba,c.jobGenCode, " & _
															   " c.JobGenName, " & _
															   " c.OrderData, " & _
															   " c.UnitOfMeasureMent, " & _
															   " a.wipdate, " & _
															   " a.qty, " & _
															   " 0 budget_qty " & _	   
														  " FROM workinprogress a, " & _
															   " agreementdetail b, " & _
															   " (SELECT a.fcba, " & _
																	   " a.FCCODE JobGenCode, " & _
																	   " a.FCNAME JobGenName, " & _
																	   " a.param_2 orderdata, " & _
																	   " a.param_3 unitofmeasurement, " & _
																	   " b.param_1 jobcode " & _
																  " FROM PARAMETER a, parameterdetail b " & _
																 " WHERE     a.fccode = b.parheadcode " & _
																	   " AND a.parcode = b.parcode " & _
																	   " AND a.fcba = b.fcba " & _
																	   " AND a.PARCODE = 'SR003_LC' " & _
																	   " AND a.PARAM_1 = 'DEFAULT' " & _
																	   " AND a.FCCODE IN ('STACKING')) c " & _
														 " WHERE     b.AGREEMENTCODE = a.AGREEMENTCODE " & _
															   " AND b.LINENO = a.AGREEMENTLINENO " & _
															   " AND b.FCBA = a.FCBA " & _
															   " AND b.job = c.jobcode " & _
														" UNION  " & _ 
													     " SELECT " & _
															   " a.fcba, a.FCCODE JobGenCode, " & _
															   " a.FCNAME JobGenName, " & _
															   " a.param_2 orderdata, " & _
															   " a.param_3 unitofmeasurement, " & _
															   " TO_DATE ('" & v_txtdatefrom & "', 'MM/DD/YYYY') wipdate, " & _
															   " 0 qty, " & _
															   " budget_transaction_estate_api.GetQuantityDay (a.fcba, " & _
																			 v_txtBudgetYear & " , " & _
																			 v_txtBudgetVersion & " , " & _
																			 v_txtModuleVersion & " , " & _
																			 " b.param_1 , " & _
																			 " TO_DATE ('" & v_txtdatefrom & "', 'MM/DD/YYYY')) budget_qty " & _
																	  " FROM PARAMETER a, parameterdetail b " & _
																	 " WHERE     a.fccode = b.parheadcode " & _
																		   " AND a.parcode = b.parcode " & _
																		   " AND a.fcba = b.fcba " & _
																		   " AND a.PARCODE = 'SR003_LC' " & _
																		   " AND a.PARAM_1 = 'DEFAULT' " & _
																		   " AND a.FCCODE IN ('STACKING') " & _	   															   
													      " ) " & _
												 " WHERE wipdate = TO_DATE ('" & v_txtdatefrom & "', 'MM/DD/YYYY') " & _
											  " GROUP BY fcba,jobGenCode, " & _
													   " JobGenName, " & _
													   " OrderData, " & _
													   " UnitOfMeasureMent " & _
											  " UNION " & _
												" SELECT fcba,jobGenCode, " & _
													   " JobGenName, " & _
													   " OrderData, " & _
													   " UnitOfMeasureMent, " & _
													   " 0 hi, " & _
													   " SUM (qty) shi, " & _
													   " 0 sbi, " & _
													   " 0 budget_hi, " & _
													   " SUM (budget_qty) budget_shi, " & _
													   " 0 budget_sbi " & _													   
												  " FROM (SELECT b.fcba,c.jobGenCode, " & _
															   " c.JobGenName, " & _
															   " c.OrderData, " & _
															   " c.UnitOfMeasureMent, " & _
															   " a.wipdate, " & _
															   " a.qty ," & _
															   " 0 budget_qty " & _	   															   
														  " FROM workinprogress a, " & _
															   " agreementdetail b, " & _
															   " (SELECT a.fcba, " & _
																	   " a.FCCODE JobGenCode, " & _
																	   " a.FCNAME JobGenName, " & _
																	   " a.param_2 orderdata, " & _
																	   " a.param_3 unitofmeasurement, " & _
																	   " b.param_1 jobcode " & _
																  " FROM PARAMETER a, parameterdetail b " & _
																 " WHERE     a.fccode = b.parheadcode " & _
																	   " AND a.parcode = b.parcode " & _
																	   " AND a.fcba = b.fcba " & _
																	   " AND a.PARCODE = 'SR003_LC' " & _
																	   " AND a.PARAM_1 = 'DEFAULT' " & _
																	   " AND a.FCCODE IN ('STACKING')) c " & _
														 " WHERE     b.AGREEMENTCODE = a.AGREEMENTCODE " & _
															   " AND b.LINENO = a.AGREEMENTLINENO " & _
															   " AND b.FCBA = a.FCBA " & _
															   " AND b.job = c.jobcode	" & _														   
													" UNION  " & _ 
													     " SELECT " & _
															   " a.fcba,a.FCCODE JobGenCode, " & _
															   " a.FCNAME JobGenName, " & _
															   " a.param_2 orderdata, " & _
															   " a.param_3 unitofmeasurement, " & _
															   " TO_DATE ('" & v_txtdatefrom & "', 'MM/DD/YYYY') wipdate, " & _
															   " 0 qty, " & _
															   " budget_transaction_estate_api.GetQuantityMonth (a.fcba, " & _
																			 v_txtBudgetYear & " , " & _
																			 v_txtBudgetVersion & " , " & _
																			 v_txtModuleVersion & " , " & _
																			 " b.param_1 , " & _
																			 " TO_DATE ('" & v_txtdatefrom & "', 'MM/DD/YYYY')) budget_qty " & _
																	  " FROM PARAMETER a, parameterdetail b " & _
																	 " WHERE     a.fccode = b.parheadcode " & _
																		   " AND a.parcode = b.parcode " & _
																		   " AND a.fcba = b.fcba " & _
																		   " AND a.PARCODE = 'SR003_LC' " & _
																		   " AND a.PARAM_1 = 'DEFAULT' " & _
																		   " AND a.FCCODE IN ('STACKING') " & _	   															   
													      " ) " & _
												 " WHERE     wipdate <= TO_DATE ('" & v_txtdatefrom & "', 'MM/DD/YYYY') " & _
													   " AND TO_CHAR (wipdate, 'YYYY') = '" & v_txtYear & "' " & _
													   " AND TO_CHAR (wipdate, 'MM') = '" & v_txtMonth & "' " & _
											  " GROUP BY fcba,jobGenCode, " & _
													   " JobGenName, " & _
													   " OrderData, " & _
													   " UnitOfMeasureMent " & _
											  " UNION " & _
												" SELECT fcba,jobGenCode, " & _
													   " JobGenName, " & _
													   " OrderData, " & _
													   " UnitOfMeasureMent, " & _
													   " 0 hi, " & _
													   " 0 shi, " & _
													   " SUM (qty) sbi, " & _
													   " 0 budget_hi, " & _
													   " 0 budget_shi, " & _
													   " SUM (budget_qty) budget_sbi " & _													   
												  " FROM (SELECT a.fcba,c.jobGenCode, " & _
															   " c.JobGenName, " & _
															   " c.OrderData, " & _
															   " c.UnitOfMeasureMent, " & _
															   " a.wipdate, " & _
															   " a.qty, " & _
															   " 0 budget_qty " & _	   															   															   
														  " FROM workinprogress a, " & _
															   " agreementdetail b, " & _
															   " (SELECT a.fcba, " & _
																	   " a.FCCODE JobGenCode, " & _
																	   " a.FCNAME JobGenName, " & _
																	   " a.param_2 orderdata, " & _
																	   " a.param_3 unitofmeasurement, " & _
																	   " b.param_1 jobcode " & _
																  " FROM PARAMETER a, parameterdetail b " & _
																 " WHERE     a.fccode = b.parheadcode " & _
																	   " AND a.parcode = b.parcode " & _
																	   " AND a.fcba = b.fcba " & _
																	   " AND a.PARCODE = 'SR003_LC' " & _
																	   " AND a.PARAM_1 = 'DEFAULT' " & _
																	   " AND a.FCCODE IN ('STACKING')) c " & _
														 " WHERE     b.AGREEMENTCODE = a.AGREEMENTCODE " & _
															   " AND b.LINENO = a.AGREEMENTLINENO " & _
															   " AND b.FCBA = a.FCBA " & _
															  "  AND b.job = c.jobcode " & _
													" UNION  " & _ 
													     " SELECT " & _
															   " a.fcba,a.FCCODE JobGenCode, " & _
															   " a.FCNAME JobGenName, " & _
															   " a.param_2 orderdata, " & _
															   " a.param_3 unitofmeasurement, " & _
															   " TO_DATE ('" & v_txtdatefrom & "', 'MM/DD/YYYY') wipdate, " & _
															   " 0 qty, " & _
															   " budget_transaction_estate_api.GetQuantityUntilMonth (a.fcba, " & _
																			 v_txtBudgetYear & " , " & _
																			 v_txtBudgetVersion & " , " & _
																			 v_txtModuleVersion & " , " & _
																			 " b.param_1 , " & _
																			 " TO_DATE ('" & v_txtdatefrom & "', 'MM/DD/YYYY')) budget_qty " & _
																	  " FROM PARAMETER a, parameterdetail b " & _
																	 " WHERE     a.fccode = b.parheadcode " & _
																		   " AND a.parcode = b.parcode " & _
																		   " AND a.fcba = b.fcba " & _
																		   " AND a.PARCODE = 'SR003_LC' " & _
																		   " AND a.PARAM_1 = 'DEFAULT' " & _
																		   " AND a.FCCODE IN ('STACKING') " & _	   															   
													      " ) " & _															  
												"  WHERE wipdate <= TO_DATE ('" & v_txtdatefrom & "', 'MM/DD/YYYY') " & _
													  "  AND TO_CHAR (wipdate, 'YYYY') = '" & v_txtYear & "' " & _
											 "  GROUP BY fcba,jobGenCode, " & _
													   " JobGenName, " & _
													   " OrderData, " & _
													   " UnitOfMeasureMent) " & _
									" GROUP BY fcba,jobGenCode, " & _
											 " JobGenName, " & _
											 " OrderData, " & _
											 " UnitOfMeasureMent " & _
                                   " UNION   " & _
                                     " SELECT a.fcba,c.JobGenCode,   " & _
                                            " c.JobGenName,   " & _
                                            " c.orderdata,   " & _
                                            " c.UnitOfMeasurement, " & _  
                                            " SUM (a.output) hi,   " & _
                                            " 0 shi,   " & _
                                            " 0 sbi,   " & _
											" SUM ( budget_transaction_estate_api.GetOutputDay (a.fcba, " & _
																						 v_txtBudgetYear & " , " & _
																						 v_txtBudgetVersion & " , " & _
																						 v_txtModuleVersion & " , " & _
																						 " a.jobcode , " & _
																						 " TO_DATE ('" & v_txtdatefrom & "', 'MM/DD/YYYY'))) budget_hi,  " & _
                                            " 0 budget_shi,   " & _
                                            " 0 budget_sbi   " & _
                                       " FROM Attendance_Gad a, " & _
                                            " (SELECT a.fcba,   " & _
                                                    " a.FCCODE JobGenCode,   " & _
                                                    " a.FCNAME JobGenName,   " & _
                                                    " a.param_2 orderdata,   " & _
                                                    " a.param_3 unitofmeasurement,   " & _
                                                    " b.param_1 jobcode   " & _
                                               " FROM PARAMETER a, parameterdetail b   " & _
                                              " WHERE     a.fccode = b.parheadcode   " & _
                                                    " AND a.parcode = b.parcode   " & _
                                                    " AND a.fcba = b.fcba   " & _
                                                    " AND a.PARCODE = 'SR003_LC'  " & _ 
                                                    " AND a.PARAM_1 = 'DEFAULT') c   " & _
                                      " WHERE    a.jobcode  = c.jobcode   " & _
                                            " AND a.fddate = TO_DATE ('" & v_txtdatefrom & "', 'MM/DD/YYYY')   " & _
                                            " AND a.mandays > 0 " & _
                                   " GROUP BY a.fcba,c.JobGenCode,   " & _
                                            " c.JobGenName,   " & _
                                            " c.orderdata,   " & _
                                            " c.UnitOfMeasurement  " & _ 
                                   " UNION   " & _								   
								   " SELECT a.fcba,a.FCCODE JobGenCode,   " & _
                                                    " a.FCNAME JobGenName,   " & _
                                                    " a.param_2 orderdata,   " & _
                                                    " a.param_3 unitofmeasurement, " & _  
													" 0 hi," & _  
													" 0 shi," & _  
													" 0 sbi," & _  
													" SUM (budget_transaction_estate_api.GetOutputDay (a.fcba, " & _
																						 v_txtBudgetYear & " , " & _
																						 v_txtBudgetVersion & " , " & _
																						 v_txtModuleVersion & " , " & _
																						 " b.param_1 , " & _
																						 " TO_DATE ('" & v_txtdatefrom & "', 'MM/DD/YYYY'))) budget_hi,   " & _
													" SUM (budget_transaction_estate_api.GetOutputMonth (a.fcba, " & _
																						 v_txtBudgetYear & " , " & _
																						 v_txtBudgetVersion & " , " & _
																						 v_txtModuleVersion & " , " & _
																						 " b.param_1 , " & _
																						 " TO_DATE ('" & v_txtdatefrom & "', 'MM/DD/YYYY'))) budget_shi,   " & _
													" SUM (budget_transaction_estate_api.GetOutputUntilMonth (a.fcba, " & _
																						 v_txtBudgetYear & " , " & _
																						 v_txtBudgetVersion & " , " & _
																						 v_txtModuleVersion & " , " & _
																						 " b.param_1 , " & _
																						 " TO_DATE ('" & v_txtdatefrom & "', 'MM/DD/YYYY'))) budget_sbi  " & _
                                               " FROM PARAMETER a, parameterdetail b   " & _
                                              " WHERE     a.fccode = b.parheadcode   " & _
                                                    " AND a.parcode = b.parcode   " & _
                                                    " AND a.fcba = b.fcba   " & _
                                                    " AND a.PARCODE = 'SR003_LC'   " & _
                                                    " AND a.PARAM_1 = 'DEFAULT' " & _
													" AND a.FCCODE IN ('KACANGAN') " & _  
	                                   " GROUP BY a.fcba,a.FCCODE,   " & _
                                            " a.FCNAME,   " & _
                                            " a.param_2,   " & _
                                            " a.param_3 " & _												
                                   " UNION   " & _								   									
                                   " SELECT a.fcba,c.JobGenCode,   " & _
                                            " c.JobGenName,   " & _
                                            " c.orderdata,   " & _
                                            " c.UnitOfMeasurement, " & _  
                                            " SUM (a.output) hi,   " & _
                                            " 0 shi,   " & _
                                            " 0 sbi,   " & _
                                            " 0 budget_hi,   " & _
                                            " 0 budget_shi,   " & _
                                            " 0 budget_sbi   " & _
                                       " FROM Attendance_Gad a, " & _
                                            " (SELECT a.fcba,   " & _
                                                    " a.FCCODE JobGenCode,   " & _
                                                    " a.FCNAME JobGenName,   " & _
                                                    " a.param_2 orderdata,   " & _
                                                    " a.param_3 unitofmeasurement, " & _  
                                                    " b.param_1 jobcode   " & _
                                               " FROM PARAMETER a, parameterdetail b   " & _
                                              " WHERE     a.fccode = b.parheadcode   " & _
                                                    " AND a.parcode = b.parcode   " & _
                                                    " AND a.fcba = b.fcba   " & _
                                                    " AND a.PARCODE = 'SR003_LC'   " & _
                                                    " AND a.PARAM_1 = 'DEFAULT' " & _
													" AND a.FCCODE IN ('KACANGAN')) c   " & _
                                      " WHERE    a.jobcode  = c.jobcode   " & _
                                            " AND a.fddate = TO_DATE ('" & v_txtdatefrom & "', 'MM/DD/YYYY')   " & _
                                            " AND a.mandays > 0 " & _
                                   " GROUP BY a.fcba,c.JobGenCode,   " & _
                                            " c.JobGenName,   " & _
                                            " c.orderdata,   " & _
                                            " c.UnitOfMeasurement " & _
									" UNION " & _								   
                                   " SELECT a.fcba,c.JobGenCode,   " & _
                                            " c.JobGenName,   " & _
                                            " c.orderdata,   " & _
                                            " c.UnitOfMeasurement, " & _  
                                            " 0 hi,   " & _
                                            " SUM (a.output) shi,   " & _
                                            " 0 sbi,   " & _
											" 0 budget_hi,   " & _
                                            " 0 budget_shi,   " & _
                                            " 0 budget_sbi   " & _
                                       " FROM Attendance_Gad a, " & _
                                            " (SELECT a.fcba,   " & _
                                                    " a.FCCODE JobGenCode,   " & _
                                                    " a.FCNAME JobGenName,   " & _
                                                    " a.param_2 orderdata,   " & _
                                                    " a.param_3 unitofmeasurement, " & _  
                                                    " b.param_1 jobcode   " & _
                                               " FROM PARAMETER a, parameterdetail b   " & _
                                              " WHERE     a.fccode = b.parheadcode   " & _
                                                    " AND a.parcode = b.parcode   " & _
                                                    " AND a.fcba = b.fcba   " & _
                                                    " AND a.PARCODE = 'SR003_LC'   " & _
                                                    " AND a.PARAM_1 = 'DEFAULT' " & _
													" AND a.FCCODE IN ('KACANGAN')) c   " & _
                                      " WHERE    a.jobcode  = c.jobcode   " & _
                                            " AND a.fddate <= TO_DATE ('" & v_txtdatefrom & "', 'MM/DD/YYYY')   " & _
                                            " AND to_char(a.fddate, 'YYYY') = '" & v_txtYear & "'    " & _
                                            " AND to_char(a.fddate, 'MM') = '" & v_txtMonth & "'   " & _
                                            " AND a.mandays > 0 " & _
                                   " GROUP BY a.fcba,c.JobGenCode,   " & _
                                            " c.JobGenName,   " & _
                                            " c.orderdata,   " & _
                                            " c.UnitOfMeasurement " & _
									" UNION " & _
                                   " SELECT a.fcba,c.JobGenCode,   " & _
                                            " c.JobGenName,   " & _
                                            " c.orderdata,   " & _
                                            " c.UnitOfMeasurement, " & _  
                                            " 0 hi,   " & _
											" 0 shi,   " & _
											" SUM (a.output) sbi,   " & _
                                            " 0 budget_hi,   " & _
											" 0 budget_shi,   " & _
											" 0 budget_sbi   " & _											
                                       " FROM Attendance_Gad a, " & _
                                            " (SELECT a.fcba,   " & _
                                                    " a.FCCODE JobGenCode,   " & _
                                                    " a.FCNAME JobGenName,   " & _
                                                    " a.param_2 orderdata,   " & _
                                                    " a.param_3 unitofmeasurement, " & _  
                                                    " b.param_1 jobcode   " & _
                                               " FROM PARAMETER a, parameterdetail b   " & _
                                              " WHERE     a.fccode = b.parheadcode   " & _
                                                    " AND a.parcode = b.parcode   " & _
                                                    " AND a.fcba = b.fcba   " & _
                                                    " AND a.PARCODE = 'SR003_LC'   " & _
                                                    " AND a.PARAM_1 = 'DEFAULT'   " & _
													" AND a.FCCODE IN ('KACANGAN' )) c   " & _
                                      " WHERE    a.jobcode  = c.jobcode   " & _
                                            " AND a.fddate <= TO_DATE ('" & v_txtdatefrom & "', 'MM/DD/YYYY')   " & _
                                            " AND to_char(a.fddate, 'YYYY') = '" & v_txtYear & "'    " & _
                                            " AND a.mandays > 0 " & _
                                   " GROUP BY a.fcba,c.JobGenCode,   " & _
                                            " c.JobGenName,   " & _
                                            " c.orderdata,   " & _
                                            " c.UnitOfMeasurement " & _
								   " UNION " & _								   
								   " SELECT a.fcba,a.FCCODE JobGenCode,   " & _
                                                    " a.FCNAME JobGenName,   " & _
                                                    " a.param_2 orderdata,   " & _
                                                    " a.param_3 unitofmeasurement, " & _  
													" 0 hi," & _  
													" 0 shi," & _  
													" 0 sbi," & _  
													" SUM (budget_transaction_estate_api.GetOutputDay (a.fcba, " & _
																						 v_txtBudgetYear & " , " & _
																						 v_txtBudgetVersion & " , " & _
																						 v_txtModuleVersion & " , " & _
																						 " b.param_1 , " & _
																						 " TO_DATE ('" & v_txtdatefrom & "', 'MM/DD/YYYY'))) budget_hi,   " & _
													" SUM (budget_transaction_estate_api.GetOutputMonth (a.fcba, " & _
																						 v_txtBudgetYear & " , " & _
																						 v_txtBudgetVersion & " , " & _
																						 v_txtModuleVersion & " , " & _
																						 " b.param_1 , " & _
																						 " TO_DATE ('" & v_txtdatefrom & "', 'MM/DD/YYYY'))) budget_shi,   " & _
													" SUM (budget_transaction_estate_api.GetOutputUntilMonth (a.fcba, " & _
																						 v_txtBudgetYear & " , " & _
																						 v_txtBudgetVersion & " , " & _
																						 v_txtModuleVersion & " , " & _
																						 " b.param_1 , " & _
																						 " TO_DATE ('" & v_txtdatefrom & "', 'MM/DD/YYYY'))) budget_sbi  " & _
                                               " FROM PARAMETER a, parameterdetail b   " & _
                                              " WHERE     a.fccode = b.parheadcode   " & _
                                                    " AND a.parcode = b.parcode   " & _
                                                    " AND a.fcba = b.fcba   " & _
                                                    " AND a.PARCODE = 'SR003_LC'   " & _
                                                    " AND a.PARAM_1 = 'DEFAULT' " & _
													" AND a.FCCODE IN ('TANAMSAWIT') " & _  
	                                   " GROUP BY a.fcba,a.FCCODE,   " & _
                                            " a.FCNAME,   " & _
                                            " a.param_2,   " & _
                                            " a.param_3 " & _												
								   " UNION " & _								   
								" SELECT a.fcba,c.JobGenCode,   " & _
                                            " c.JobGenName,   " & _
                                            " c.orderdata,   " & _
                                            " c.UnitOfMeasurement, " & _  
                                            " SUM (a.unit) hi,   " & _
                                            " 0 shi,   " & _
                                            " 0 sbi,   " & _
                                            " 0 budget_hi,   " & _
                                            " 0 budget_shi,   " & _
                                            " 0 budget_sbi   " & _
                                       " FROM Attendance_Gad a, " & _
                                            " (SELECT a.fcba,   " & _
                                                    " a.FCCODE JobGenCode,   " & _
                                                    " a.FCNAME JobGenName,   " & _
                                                    " a.param_2 orderdata,   " & _
                                                    " a.param_3 unitofmeasurement, " & _  
                                                    " b.param_1 jobcode   " & _
                                               " FROM PARAMETER a, parameterdetail b   " & _
                                              " WHERE     a.fccode = b.parheadcode   " & _
                                                    " AND a.parcode = b.parcode   " & _
                                                    " AND a.fcba = b.fcba   " & _
                                                    " AND a.PARCODE = 'SR003_LC'   " & _
                                                    " AND a.PARAM_1 = 'DEFAULT' " & _
													" AND a.FCCODE IN ('TANAMSAWIT')) c   " & _
                                      " WHERE    a.jobcode  = c.jobcode   " & _
                                            " AND a.fddate = TO_DATE ('" & v_txtdatefrom & "', 'MM/DD/YYYY')   " & _
                                   " GROUP BY a.fcba, c.JobGenCode,   " & _
                                            " c.JobGenName,   " & _
                                            " c.orderdata,   " & _
                                            " c.UnitOfMeasurement " & _
									" UNION " & _								   
								 " SELECT a.fcba,c.JobGenCode,   " & _
                                            " c.JobGenName,   " & _
                                            " c.orderdata,   " & _
                                            " c.UnitOfMeasurement, " & _  
                                            " 0 hi,   " & _
                                            " SUM (a.unit) shi,   " & _
                                            " 0 sbi,   " & _
											" 0 budget_hi,   " & _
                                            " 0 budget_shi,   " & _
                                            " 0 budget_sbi   " & _
                                       " FROM Attendance_Gad a, " & _
                                            " (SELECT a.fcba,   " & _
                                                    " a.FCCODE JobGenCode,   " & _
                                                    " a.FCNAME JobGenName,   " & _
                                                    " a.param_2 orderdata,   " & _
                                                    " a.param_3 unitofmeasurement, " & _  
                                                    " b.param_1 jobcode   " & _
                                               " FROM PARAMETER a, parameterdetail b   " & _
                                              " WHERE     a.fccode = b.parheadcode   " & _
                                                    " AND a.parcode = b.parcode   " & _
                                                    " AND a.fcba = b.fcba   " & _
                                                    " AND a.PARCODE = 'SR003_LC'   " & _
                                                    " AND a.PARAM_1 = 'DEFAULT' " & _
													" AND a.FCCODE IN ('TANAMSAWIT')) c   " & _
                                      " WHERE    a.jobcode  = c.jobcode   " & _
                                            " AND a.fddate <= TO_DATE ('" & v_txtdatefrom & "', 'MM/DD/YYYY')   " & _
                                            " AND to_char(a.fddate, 'YYYY') = '" & v_txtYear & "'    " & _
                                            " AND to_char(a.fddate, 'MM') = '" & v_txtMonth & "'   " & _
                                   " GROUP BY a.fcba,c.JobGenCode,   " & _
                                            " c.JobGenName,   " & _
                                            " c.orderdata,   " & _
                                            " c.UnitOfMeasurement " & _
									" UNION " & _
                                   " SELECT a.fcba,c.JobGenCode,   " & _
                                            " c.JobGenName,   " & _
                                            " c.orderdata,   " & _
                                            " c.UnitOfMeasurement, " & _  
                                            " 0 hi,   " & _
											" 0 shi,   " & _
											" SUM (a.unit) sbi,   " & _
									        " 0 budget_hi,   " & _
											" 0 budget_shi,   " & _
											" 0 budget_sbi   " & _		
                                       " FROM Attendance_Gad a, " & _
                                            " (SELECT a.fcba,   " & _
                                                    " a.FCCODE JobGenCode,   " & _
                                                    " a.FCNAME JobGenName,   " & _
                                                    " a.param_2 orderdata,   " & _
                                                    " a.param_3 unitofmeasurement, " & _  
                                                    " b.param_1 jobcode   " & _
                                               " FROM PARAMETER a, parameterdetail b   " & _
                                              " WHERE     a.fccode = b.parheadcode   " & _
                                                    " AND a.parcode = b.parcode   " & _
                                                    " AND a.fcba = b.fcba   " & _
                                                    " AND a.PARCODE = 'SR003_LC'   " & _
                                                    " AND a.PARAM_1 = 'DEFAULT'   " & _
													" AND a.FCCODE IN ('TANAMSAWIT' )) c   " & _
                                      " WHERE    a.jobcode  = c.jobcode   " & _
                                            " AND a.fddate <= TO_DATE ('" & v_txtdatefrom & "', 'MM/DD/YYYY')   " & _
                                            " AND to_char(a.fddate, 'YYYY') = '" & v_txtYear & "'    " & _
                                   " GROUP BY a.fcba,c.JobGenCode,   " & _
                                            " c.JobGenName,   " & _
                                            " c.orderdata,   " & _
                                            " c.UnitOfMeasurement " & _
									"		)   " & _
									  "			     where ACCESSRIGHTSCOMPANY_API.IsAccessRights('" & Session("CodeName") & "',FCBA,'" & txtcompany & "') = 'TRUE' " & str1 & " " & _
                         " GROUP BY JobGenCode,   " & _
                                  " JobGenName,   " & _
                                  " orderdata,   " & _
                                  " UnitOfMeasurement  " & _ 
                         " ORDER BY orderdata  ) "	 
	
 	Set rsLandClearing = Server.CreateObject("ADODB.Recordset")
	'response.write strLandClearing
	
	rsLandClearing.Open strLandClearing, connstr, 3,3

	
	
	' First
    ' To Handle the Query of Production
    ' strProduction = " select codedesc, rownum thenumber, description, hi, shi, sbi, budget_hi, budget_shi, budget_sbi from " & _
				    ' "( Select codedesc, " & _
	                    ' " (SELECT fcname " & _
									' " FROM PARAMETERDETAIL " & _
								   ' " WHERE parcode = 'SR003' " & _
									 ' " AND fccode = codedesc " & _
									 ' " AND fcba = '" & session("fcba") & "') description, " & _
	                    ' " (SELECT param_3 orderdata " & _
									' " FROM PARAMETERDETAIL " & _
								   ' " WHERE parcode = 'SR003' " & _
									 ' " AND fccode = codedesc " & _
									 ' " AND fcba = '" & session("fcba") & "') orderdata, " & _
	                ' " hi, shi, sbi, budget_hi, budget_shi, budget_sbi from ( " & _
					' " SELECT   codedesc, '1' theNumber, " & _
								 ' " sum(hi) hi, sum(shi) shi, sum(sbi) sbi, sum( budget_hi) budget_hi, sum(budget_shi) budget_shi, sum(budget_sbi) budget_sbi " & _
							' " FROM ((SELECT 'SR003_01' AS codedesc, SUM (totaldailyreceivetonase0) hi, 0 shi, 0 sbi, 0 budget_hi, 0 budget_shi, 0 budget_sbi  " & _
									' " FROM view_dailyharvesting_mill " & _
								   ' " WHERE accountingyear = '" & v_txtyear & "' " & _
									 ' " AND accountingperiod = '"& v_txtmonth & "' " & _
									 ' " AND receptiondate = TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY') " & _
								  ' " UNION " & _ 
									' " SELECT 'SR003_01' AS codedesc, 0 hi, SUM (totaldailyreceivetonase0) shi, 0 sbi, 0 budget_hi, 0 budget_shi, 0 budget_sbi   " & _
									' " FROM view_dailyharvesting_mill " & _
								   ' " WHERE accountingyear = '" & v_txtyear & "' " & _
									 ' " AND accountingperiod = '"& v_txtmonth & "' " & _
									 ' " AND receptiondate <= TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY') " & _
								  ' " UNION " & _ 
									' " SELECT 'SR003_01' AS codedesc, 0 hi, 0 shi, SUM (totaldailyreceivetonase0) sbi, 0 budget_hi, 0 budget_shi, 0 budget_sbi   " & _
									' " FROM view_dailyharvesting_mill " & _
								   ' " WHERE accountingyear = '" & v_txtyear & "' " & _
									 ' " AND receptiondate <= TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY') " & _
								  ' " UNION " & _ 
									' " SELECT 'SR003_01' AS codedesc, 0 hi, 0 shi, 0 sbi, " & _
										' " BUDGET_TRANSACTION_PRODQTY_API.GetQuantityDay (nvl('" & v_txtfcba & "' ,'" & session("fcba") & "'), " & _
																			 ' v_txtBudgetYear & " , " & _
																			 ' v_txtBudgetVersion & " , " & _
																			 ' v_txtModuleVersion & " , " & _
																			 ' " TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY')) budget_hi, " & _
										' " BUDGET_TRANSACTION_PRODQTY_API.GetQuantityMonth (nvl('" & v_txtfcba & "' ,'" & session("fcba") & "'), " & _
																			 ' v_txtBudgetYear & " , " & _
																			 ' v_txtBudgetVersion & " , " & _
																			 ' v_txtModuleVersion & " , " & _
																			 ' " TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY')) budget_shi, " & _
										' " BUDGET_TRANSACTION_PRODQTY_API.GetQuantityUntilMonth (nvl('" & v_txtfcba & "' ,'" & session("fcba") & "'), " & _
																			 ' v_txtBudgetYear & " , " & _
																			 ' v_txtBudgetVersion & " , " & _
																			 ' v_txtModuleVersion & " , " & _
																			 ' " TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY')) budget_sbi from dual " & _
								  ' " UNION " & _
								  ' " SELECT 'SR003_02' AS codedesc, " & _
										 ' " (SUM (variantemp) + SUM (restantemp)) hi, (SUM (variantemp) + SUM (restantemp)) shi, (SUM (variantemp) + SUM (restantemp)) sbi, 0 budget_hi, 0 budget_shi, 0 budget_sbi " & _
									' " FROM (SELECT varian variantemp, 0 restantemp " & _
											' " FROM view_dailyharvesting_varian " & _
										   ' " WHERE dateprod = TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY') " & _
										  ' " UNION " & _
										  ' " SELECT 0 variantemp, SUM (varian) restantemp " & _
											' " FROM view_dailyharvesting_varian " & _
										   ' " WHERE dateprod < TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY')) " & _
								  ' " UNION " & _
								  ' " SELECT 'SR003_03' AS codedesc, 0 hi, 0 shi, 0 sbi, 0 budget_hi, 0 budget_shi, 0 budget_sbi " & _
									' " FROM DUAL " & _
								  ' " UNION " & _
								  ' " SELECT 'SR003_04' AS codedesc, SUM (totrninches) hi, 0 shi, 0 sbi, 0 budget_hi, 0 budget_shi, 0 budget_sbi " & _
									' " FROM WEATHER " & _
								   ' " WHERE widate = TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY') " & _
								  ' " UNION " & _
								  ' " SELECT 'SR003_04' AS codedesc, 0 shi, SUM (totrninches) shi, 0 sbi, 0 budget_hi, 0 budget_shi, 0 budget_sbi  " & _
									' " FROM WEATHER " & _
								   ' " WHERE widate <= TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY') " & _
								   ' " and to_char(widate, 'MM' ) = '" & v_txtmonth & "' " & _
								   ' " and to_char(widate, 'YYYY') = '" & v_txtyear & "' " & _
								  ' " UNION " & _
								  ' " SELECT 'SR003_04' AS codedesc, 0 hi, 0 shi, SUM (totrninches) sbi, 0 budget_hi, 0 budget_shi, 0 budget_sbi  " & _
									' " FROM WEATHER " & _
								   ' " WHERE widate <= TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY') " & _
								   ' " and to_char(widate, 'YYYY') = '" & v_txtyear & "' " & _
								  ' " UNION " & _
									' " select codedesc, sum(hi) hi, sum(shi) shi, sum(sbi) sbi, sum(budget_hi) budget_hi, sum(budget_shi) budget_shi, sum(budget_sbi) budget_sbi " & _
									' " from ( " & _
									' " SELECT 'SR003_OP' codedesc, " & _
										   ' " SUM (mandays) hi, " & _
										   ' " 0 shi, " & _
										   ' " 0 sbi, " & _
										   ' " 0 budget_hi, " & _ 
										   ' " 0 budget_shi, " & _ 
										   ' " 0 budget_sbi " & _ 
									  ' " FROM attendance_gad " & _
									 ' " WHERE     GetCompanyCode (fcba) = '" & Session("COMPANYCODE") & "' " & _
										   ' " AND jobcode = '6502101' " & _
										   ' " AND fddate = TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY') " & _
										   ' " AND mandays > 0 " & _
									' " UNION " & _
									' " SELECT 'SR003_OP' codedesc, " & _
										   ' " 0 hi, " & _
										   ' " SUM (mandays) shi, " & _
										   ' " 0 sbi, " & _
										   ' " 0 budget_hi, " & _ 
										   ' " 0 budget_shi, " & _ 
										   ' " 0 budget_sbi " & _ 
									  ' " FROM attendance_gad " & _
									 ' " WHERE     GetCompanyCode (fcba) = '"& Session("COMPANYCODE") &"' " & _
										   ' " AND jobcode = '6502101' " & _
										   ' " AND fddate <= TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY') " & _
										   ' " AND TO_CHAR (fddate, 'YYYY') = '" & v_txtyear & "' " & _
										   ' " AND TO_CHAR (fddate, 'MM') = '" & v_txtmonth & "' " & _
										   ' " AND mandays > 0 " & _
									' " UNION " & _
									' " SELECT 'SR003_OP' codedesc, " & _
										   ' " 0 hi, " & _
										   ' " 0 shi, " & _
										   ' " SUM (mandays) sbi, " & _
										   ' " 0 budget_hi, " & _ 
										   ' " 0 budget_shi, " & _ 
										   ' " 0 budget_sbi " & _ 
									  ' " FROM attendance_gad " & _
									 ' " WHERE     GetCompanyCode (fcba) = '" & Session("COMPANYCODE") & "' " & _
										   ' " AND jobcode = '6502101' " & _
										   ' " AND fddate <= TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY') " & _
										   ' " AND TO_CHAR (fddate, 'YYYY') = '" & v_txtyear & "' " & _
										   ' " AND mandays > 0 " & _
									  ' " UNION " & _   
									' " SELECT 'SR003_OP' AS codedesc, 0 hi, 0 shi, 0 sbi, " & _
										' " SUM( BUDGET_TRANSACTION_ESTATE_API.GetQuantityDay (nvl('" & v_txtfcba & "' ,'" & session("fcba") & "'), " & _
																			 ' v_txtBudgetYear & " , " & _
																			 ' v_txtBudgetVersion & " , " & _
																			 ' v_txtModuleVersion & " , " & _
																			 ' "'6502101', " & _ 
																			 ' " TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY'))) / " & _
										' " SUM( BUDGET_TRANSACTION_ESTATE_API.GetOutputDay (nvl('" & v_txtfcba & "' ,'" & session("fcba") & "'), " & _
																			 ' v_txtBudgetYear & " , " & _
																			 ' v_txtBudgetVersion & " , " & _
																			 ' v_txtModuleVersion & " , " & _
																			 ' "'6502101', " & _ 
																			 ' " TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY')))  budget_hi, " & _
										' " SUM(BUDGET_TRANSACTION_ESTATE_API.GetQuantityMonth (nvl('" & v_txtfcba & "' ,'" & session("fcba") & "'), " & _
																			 ' v_txtBudgetYear & " , " & _
																			 ' v_txtBudgetVersion & " , " & _
																			 ' v_txtModuleVersion & " , " & _
																			 ' "'6502101', " & _ 
																			 ' " TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY'))) / " & _ 
										' " SUM(BUDGET_TRANSACTION_ESTATE_API.GetOutputMonth (nvl('" & v_txtfcba & "' ,'" & session("fcba") & "'), " & _
																			 ' v_txtBudgetYear & " , " & _
																			 ' v_txtBudgetVersion & " , " & _
																			 ' v_txtModuleVersion & " , " & _
																			 ' "'6502101', " & _ 
																			 ' " TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY'))) budget_shi, " & _
										' " SUM(BUDGET_TRANSACTION_ESTATE_API.GetQuantityUntilMonth (nvl('" & v_txtfcba & "' ,'" & session("fcba") & "'), " & _
																			 ' v_txtBudgetYear & " , " & _
																			 ' v_txtBudgetVersion & " , " & _
																			 ' v_txtModuleVersion & " , " & _
																			 ' "'6502101', " & _ 
																			 ' " TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY'))) / " & _
										' " SUM(BUDGET_TRANSACTION_ESTATE_API.GetOutputUntilMonth (nvl('" & v_txtfcba & "' ,'" & session("fcba") & "'), " & _
																			 ' v_txtBudgetYear & " , " & _
																			 ' v_txtBudgetVersion & " , " & _
																			 ' v_txtModuleVersion & " , " & _
																			 ' "'6502101', " & _ 
																			 ' " TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY'))) budget_sbi from dual " & _										   
									 ' " ) " & _
									 ' " group by codedesc	 " & _							  
								  ' " UNION " & _
								  ' "  SELECT codedesc,  " & _
										' " SUM (hi) hi, " & _
										' " SUM (shi) shi, " & _
										' " SUM (sbi) sbi, " & _
										' " SUM (budget_hi) hi, " & _
										' " SUM (budget_shi) shi, " & _
										' " SUM (budget_sbi) sbi " & _
									' " FROM (SELECT 'SR003_BA' AS codedesc, " & _
												 ' " SUM (totalAmount) hi, " & _
												 ' " 0 shi, " & _
												 ' " 0 sbi, " & _
												 ' " 0 budget_hi, " & _
												 ' " 0 budget_shi, " & _
												 ' " 0 budget_sbi " & _												 
											' " FROM (  SELECT b.vehicledate, " & _
														   ' " SUM (a.average * b.totalQuantity) TotalAmount " & _
													  ' " FROM (  SELECT vehiclecode, " & _
																	 ' " SUM (amount) / SUM (quantity) average " & _
																' " FROM vehiclecb " & _
															   ' " WHERE TO_CHAR (transdate, 'YYYYMM') = GetPreviousMonth ('" & v_txtyear & v_txtmonth & "') " & _
																	 ' " AND jobcode IN " & _
																			' " ('6502201', " & _
																			' "  '6502202', " & _
																			' "  '6502203', " & _
																			' "  '6502204') " & _
																	 ' " AND targettype = 'FF' " & _
															' " GROUP BY vehiclecode) a, " & _
														   ' " (  SELECT vehiclecode, " & _
																	 ' " vehicledate, " & _
																	 ' " SUM (unit) totalQuantity " & _
																' " FROM vehicleactivity " & _
															   ' " WHERE vehicledate = " & _
																		' " TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY') " & _
																	 ' " AND jobcode IN " & _
																			' " ('6502201', " & _
																			' "  '6502202', " & _
																			' "  '6502203', " & _
																			' "  '6502204') " & _
																	 ' " AND locationtype = 'FF' " & _
															' " GROUP BY vehiclecode, vehicledate) b " & _
													 ' " WHERE a.vehiclecode = b.vehiclecode " & _
												  ' " GROUP BY b.vehicledate) " & _
										  ' " UNION " & _
										  ' " SELECT 'SR003_BA' AS codedesc, " & _
												 ' " 0 hi, " & _
												 ' " SUM (totalAmount) shi, " & _
												 ' " 0 sbi, " & _
												 ' " 0 budget_hi, " & _
												 ' " 0 budget_shi, " & _
												 ' " 0 budget_sbi " & _													 
											' " FROM (  SELECT b.vehicledate, " & _
														   ' " SUM (a.average * b.totalQuantity) TotalAmount " & _
													  ' " FROM (  SELECT vehiclecode, " & _
																	 ' " SUM (amount) / SUM (quantity) average " & _
																' " FROM vehiclecb " & _
															   ' " WHERE TO_CHAR (transdate, 'YYYYMM') = GetPreviousMonth ('" & v_txtyear & v_txtmonth & "') " & _
																	 ' " AND jobcode IN " & _
																			' " ('6502201', " & _
																			' "  '6502202', " & _
																			' "  '6502203', " & _
																			' "  '6502204') " & _
																	 ' " AND targettype = 'FF' " & _
															' " GROUP BY vehiclecode) a, " & _
														   ' " (  SELECT vehiclecode, " & _
																	' "  vehicledate, " & _
																	' "  SUM (unit) totalQuantity " & _
																' " FROM vehicleactivity " & _
															   ' " WHERE vehicledate <= " & _
																		' " TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY') " & _
																	 ' " AND TO_CHAR (vehicledate, 'YYYY') = '" & v_txtyear & "' " & _
																	 ' " AND TO_CHAR (vehicledate, 'MM') = '" & v_txtmonth &"' " & _
																	 ' " AND jobcode IN " & _
																			' " ('6502201', " & _
																			 ' " '6502202', " & _
																			' " '6502203', " & _
																			 ' " '6502204') " & _
																	 ' " AND locationtype = 'FF' " & _
															' " GROUP BY vehiclecode, vehicledate) b " & _
													' "  WHERE a.vehiclecode = b.vehiclecode " & _
												 ' "  GROUP BY b.vehicledate) " & _
										 ' "  UNION " & _
										 ' "  SELECT 'SR003_BA' AS codedesc, " & _
												' "  0 hi, " & _
												' "  0 shi, " & _
												' "  SUM (totalAmount) sbi, " & _
												 ' " 0 budget_hi, " & _
												 ' " 0 budget_shi, " & _
												 ' " 0 budget_sbi " & _													
											' " FROM (  SELECT b.vehicledate, " & _
														  ' "  SUM (a.average * b.totalQuantity) TotalAmount " & _
													 ' "  FROM (  SELECT vehiclecode, " & _
																	' "  SUM (amount) / SUM (quantity) average " & _
																' " FROM vehiclecb " & _
															 ' "   WHERE TO_CHAR (transdate, 'YYYYMM') = GetPreviousMonth ('" & v_txtyear & v_txtmonth & "') " & _
																	' "  AND jobcode IN " & _
																			' " ('6502201', " & _
																			 ' " '6502202', " & _
																			 ' " '6502203', " & _
																			 ' " '6502204') " & _
																	 ' " AND targettype = 'FF' " & _
															' " GROUP BY vehiclecode) a, " & _
														   ' " (  SELECT vehiclecode, " & _
																	 ' " vehicledate, " & _
																	 ' " SUM (unit) totalQuantity " & _
																' " FROM vehicleactivity " & _
															   ' " WHERE vehicledate <= " & _
																		' " TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY') " & _
																	' "  AND TO_CHAR (vehicledate, 'YYYY') = '" & v_txtyear & "' " & _
																	' "  AND jobcode IN " & _
																			' " ('6502201', " & _
																			 ' " '6502202', " & _
																			' "  '6502203', " & _
																			' "  '6502204') " & _
																	' "  AND locationtype = 'FF' " & _
															' " GROUP BY vehiclecode, vehicledate) b " & _
													 ' " WHERE a.vehiclecode = b.vehiclecode " & _
												  ' " GROUP BY b.vehicledate)) " & _
									' " GROUP BY codedesc " & _
									' " UNION " & _
								  ' " SELECT 'SR003_05' AS codedesc, " & _
										 ' " DECODE (SUM (pom_tonase_hi), " & _
												 ' " 0, 0, " & _
												 ' " SUM (mandays_amount_hi) " & _
												 ' " / (SUM (pom_tonase_hi) * 1000) " & _
												' " ) hi, 0 shi, 0 sbi, 0 budget_hi, 0 budget_shi, 0 budget_sbi " & _
									' " FROM (SELECT   fcba, accountingyear, accountingperiod, fieldcode, " & _
												   ' " division, plantingyear, fddate, " & _
												   ' " SUM (mandays_amount_hi) mandays_amount_hi, " & _
												   ' " SUM (pom_tonase_hi) pom_tonase_hi " & _
											  ' " FROM (SELECT fcba, accountingyear, accountingperiod, " & _
														   ' " fieldcode, division, plantingyear, " & _
														   ' " receptiondate fddate, 0 mandays_amount_hi, " & _
														   ' " totaldailyreceivetonase1 pom_tonase_hi " & _
													  ' " FROM view_dailyharvesting_mill " & _
													 ' " WHERE accountingyear = '" & v_txtyear & "' " & _
													   ' " AND accountingperiod = '" & v_txtmonth &"' " & _
													   ' " AND receptiondate = " & _
																  ' " TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY') " & _
													' " UNION " & _
													' " SELECT fcba, accountingyear, accountingperiod, " & _
														   ' " fieldcode, division, plantingyear, fddate, " & _
														   ' " mandays_amount mandays_amount_hi, " & _
														   ' " 0 pom_tonase_hi " & _
													  ' " FROM view_dailyharvesting_amount " & _
													 ' " WHERE accountingyear = '" & v_txtyear & "' " & _
													   ' " AND accountingperiod = '" & v_txtmonth &"' " & _
													   ' " AND fddate = " & _
																  ' " TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY')) " & _
										  ' " GROUP BY fcba, " & _
												   ' " accountingyear, " & _
												   ' " accountingperiod, " & _
												   ' " fieldcode, " & _
												   ' " division, " & _
												   ' " plantingyear, " & _
												   ' " fddate) " & _
								  ' " UNION " & _
								  ' " SELECT 'SR003_05' AS codedesc, " & _
								          ' " 0 hi , " & _
										 ' " DECODE (SUM (pom_tonase_hi), " & _
												 ' " 0, 0, " & _
												 ' " SUM (mandays_amount_hi) " & _
												 ' " / (SUM (pom_tonase_hi) * 1000) " & _
												' " ) shi, " & _
												' " 0 sbi , 0 budget_hi, 0 budget_shi, 0 budget_sbi " & _
									' " FROM (SELECT   fcba, accountingyear, accountingperiod, fieldcode, " & _
												   ' " division, plantingyear, fddate, " & _
												   ' " SUM (mandays_amount_hi) mandays_amount_hi, " & _
												   ' " SUM (pom_tonase_hi) pom_tonase_hi " & _
											  ' " FROM (SELECT fcba, accountingyear, accountingperiod, " & _
														   ' " fieldcode, division, plantingyear, " & _
														   ' " receptiondate fddate, 0 mandays_amount_hi, " & _
														   ' " totaldailyreceivetonase1 pom_tonase_hi " & _
													  ' " FROM view_dailyharvesting_mill " & _
													 ' " WHERE accountingyear = '" & v_txtyear & "' " & _
													   ' " AND accountingperiod = '" & v_txtmonth &"' " & _
													   ' " AND receptiondate <= " & _
																  ' " TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY') " & _
													' " UNION " & _
													' " SELECT fcba, accountingyear, accountingperiod, " & _
														   ' " fieldcode, division, plantingyear, fddate, " & _
														   ' " mandays_amount mandays_amount_hi, " & _
														   ' " 0 pom_tonase_hi " & _
													  ' " FROM view_dailyharvesting_amount " & _
													 ' " WHERE accountingyear = '" & v_txtyear & "' " & _
													   ' " AND accountingperiod = '" & v_txtmonth &"' " & _
													   ' " AND fddate <= " & _
																  ' " TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY')) " & _
										  ' " GROUP BY fcba, " & _
												   ' " accountingyear, " & _
												   ' " accountingperiod, " & _
												   ' " fieldcode, " & _
												   ' " division, " & _
												   ' " plantingyear, " & _
												   ' " fddate) " & _
								  ' " UNION " & _
								  ' " SELECT 'SR003_05' AS codedesc, 0 hi, 0 shi, " & _
										 ' " DECODE (SUM (pom_tonase_hi), " & _
												 ' " 0, 0, " & _
												 ' " SUM (mandays_amount_hi) " & _
												 ' " / (SUM (pom_tonase_hi) * 1000) " & _
												' " ) sbi,  0 budget_hi, 0 budget_shi, 0 budget_sbi " & _
									' " FROM (SELECT   fcba, accountingyear, accountingperiod, fieldcode, " & _
												   ' " division, plantingyear, fddate, " & _
												   ' " SUM (mandays_amount_hi) mandays_amount_hi, " & _
												   ' " SUM (pom_tonase_hi) pom_tonase_hi " & _
											  ' " FROM (SELECT fcba, accountingyear, accountingperiod, " & _
														   ' " fieldcode, division, plantingyear, " & _
														   ' " receptiondate fddate, 0 mandays_amount_hi, " & _
														   ' " totaldailyreceivetonase1 pom_tonase_hi " & _
													  ' " FROM view_dailyharvesting_mill " & _
													 ' " WHERE accountingyear = '" & v_txtyear & "' " & _
													   ' " AND receptiondate <= " & _
																  ' " TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY') " & _
													' " UNION " & _
													' " SELECT fcba, accountingyear, accountingperiod, " & _
														   ' " fieldcode, division, plantingyear, fddate, " & _
														   ' " mandays_amount mandays_amount_hi, " & _
														   ' " 0 pom_tonase_hi " & _
													  ' " FROM view_dailyharvesting_amount " & _
													 ' " WHERE accountingyear = '" & v_txtyear & "' " & _
													   ' " AND fddate <= " & _
																  ' " TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY')) " & _
										  ' " GROUP BY fcba, " & _
												   ' " accountingyear, " & _
												   ' " accountingperiod, " & _
												   ' " fieldcode, " & _
												   ' " division, " & _
												   ' " plantingyear, " & _
												   ' " fddate) " & _
												   ' " ) " & _
						' " ) " & _
						
						' " GROUP BY codedesc ) " & _
						' " ORDER BY orderdata ) " 
						
	 strProduction = " Select codedesc, rownum thenumber, description, hi, shi, sbi, budget_hi, budget_shi, budget_sbi " & _
					 "  from ( select codedesc, description, sum(hi) hi, sum(shi) shi, sum(sbi) sbi, nvl(sum( budget_hi),0) budget_hi, nvl(sum(budget_shi),0) budget_shi, nvl(sum(budget_sbi),0) budget_sbi " & _
					 "		  from ( Select  fcba,codedesc, '1' thenumber, " & _
					 "						( SELECT fcname FROM PARAMETERDETAIL WHERE parcode = 'SR003' AND fccode = codedesc AND fcba = 'KSD' ) description, " & _
					 "						( SELECT param_3 orderdata FROM PARAMETERDETAIL WHERE parcode = 'SR003' AND fccode = codedesc AND fcba = 'KSD') orderdata,  " & _
					 "						hi, shi, sbi, budget_hi, budget_shi, budget_sbi " & _
					 "				 from ( " & _
					 "						 SELECT fcba,codedesc,sum(hi) hi, sum(shi) shi, sum(sbi) sbi, sum( budget_hi) budget_hi, sum(budget_shi) budget_shi, sum(budget_sbi) budget_sbi " & _
					 "						 FROM " & _
					 "						  ( " & _
					 "						   SELECT fcba,codedesc,sum(hi) hi, sum(shi) shi, sum(sbi) sbi," & _
					 "							BUDGET_TRANSACTION_PRODQTY_API.GetQuantityDay(fcba, '" & v_txtBudgetYear &"','" & v_txtBudgetVersion &"','" & v_txtModuleVersion & "',TO_DATE('" & v_txtdateto & "', 'MM/DD/YYYY')) budget_hi, " & _ 
					 "							BUDGET_TRANSACTION_PRODQTY_API.GetQuantityMonth(fcba, '" & v_txtBudgetYear &"','" & v_txtBudgetVersion &"','" & v_txtModuleVersion & "',TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY')) budget_shi," & _
					 "							BUDGET_TRANSACTION_PRODQTY_API.GetQuantityUntilMonth(fcba, '" & v_txtBudgetYear &"','" & v_txtBudgetVersion &"','" & v_txtModuleVersion & "',TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY')) budget_sbi " & _
					 "							FROM   " & _                                          
					 "								( SELECT fcba ,'SR003_01' AS codedesc, SUM (totaldailyreceivetonase0) hi, 0 shi, 0 sbi, 0 budget_hi, 0 budget_shi, 0 budget_sbi " & _ 
					 "								  FROM view_dailyharvesting_mill " & _
					 "								  WHERE accountingyear = '" & v_txtyear & "' " & _
					 "								  AND accountingperiod = '" & v_txtmonth & "' " & _
					 "								  AND receptiondate = TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY')" & _
					 "								  group by fcba " & _
					 "								  UNION      " & _    
					 "								  SELECT fcba,'SR003_01' AS codedesc, 0 hi, SUM (totaldailyreceivetonase0) shi, 0 sbi, 0 budget_hi, 0 budget_shi, 0 budget_sbi  " & _ 
					 "								  FROM view_dailyharvesting_mill " & _
					 "								  WHERE accountingyear = '" & v_txtyear & "' " & _
					 "								  AND accountingperiod = '" & v_txtmonth & "' " & _
					 "								  AND receptiondate <= TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY') " & _
					 "								  group by fcba " & _
					 "								  UNION  " & _
					 "								  SELECT fcba,'SR003_01' AS codedesc, 0 hi, 0 shi, SUM (totaldailyreceivetonase0) sbi, 0 budget_hi, 0 budget_shi, 0 budget_sbi " & _  
					 "								  FROM view_dailyharvesting_mill " & _
					 "								  WHERE accountingyear = '" & v_txtyear & "' " & _
					 "								  AND receptiondate <= TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY') " & _
					 "								  group by fcba  " & _                         
					 "								) " & _
					 "						   GROUP BY fcba,codedesc  	" & _																			   
					 "						   UNION  " & _      					   
					 "						   SELECT fcba,'SR003_02' AS codedesc, " & _
					 "						   (SUM (variantemp) + SUM (restantemp)) hi, (SUM (variantemp) + SUM (restantemp)) shi, (SUM (variantemp) + SUM (restantemp)) sbi, 0 budget_hi, 0 budget_shi, 0 budget_sbi " & _
					 "						   FROM (  SELECT fcba,varian variantemp, 0 restantemp " & _
					 "								   FROM view_dailyharvesting_varian " & _
					 "								   WHERE dateprod = TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY') " & _
					 "								   UNION " & _
					 "								   SELECT fcba,0 variantemp, SUM(varian) restantemp " & _
					 "								   FROM view_dailyharvesting_varian " & _
					 "								   WHERE dateprod < TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY') " & _
					 "								   group by fcba 	" & _							   
					 "								  ) " & _
					 "						   GROUP BY fcba  " & _                              
					 "						   UNION " & _
					 "						   SELECT fccode fcba,'SR003_03' AS codedesc, 0 hi, 0 shi, 0 sbi, 0 budget_hi, 0 budget_shi, 0 budget_sbi " & _
					 "						   FROM businessunit   " & _
					 "						   UNION " & _
					 "						   SELECT fcba, codedesc,sum(hi) hi,sum(shi) shi, sum(sbi) sbi,sum(budget_hi) budget_hi, sum(budget_shi) budget_shi, sum(budget_sbi) budget_sbi  " & _      
					 "						   FROM ( SELECT fcba,'SR003_04' AS codedesc, SUM (totrninches) hi, 0 shi, 0 sbi, 0 budget_hi, 0 budget_shi, 0 budget_sbi " & _
					 "								  FROM WEATHER " & _
					 "								  WHERE widate = TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY') " & _
					 "								  group by fcba " & _
					 "								  UNION " & _
					 "								  SELECT fcba,'SR003_04' AS codedesc, 0 shi, SUM (totrninches) shi, 0 sbi, 0 budget_hi, 0 budget_shi, 0 budget_sbi " & _ 
					 "								  FROM WEATHER " & _
					 "								  WHERE widate <= TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY') " & _
					 "								  and to_char(widate, 'MM' ) = '" & v_txtmonth & "' " & _
					 "								  and to_char(widate, 'YYYY') = '" & v_txtyear & "' " & _
					 "								  group by fcba " & _
					 "								  UNION " & _
					 "								  SELECT fcba,'SR003_04' AS codedesc, 0 hi, 0 shi, SUM (totrninches) sbi, 0 budget_hi, 0 budget_shi, 0 budget_sbi " & _ 
					 "								  FROM WEATHER " & _
					 "								  WHERE widate <= TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY') " & _
					 "								  and to_char(widate, 'YYYY') = '" & v_txtyear & "' " & _
					 "								  group by fcba " & _
					 "								)  " & _
					 "						   GROUP BY fcba, codedesc" & _
					 "						   UNION " & _
					 "						   SELECT fcba,codedesc,sum(hi) hi, sum(shi) shi, sum(sbi) sbi,SUM(budget_hi) budget_hi, SUM(budget_shi) budget_shi,  SUM(budget_sbi) budget_sbi  " & _                          
					 "						   FROM        " & _
					 "							  ( SELECT fcba,'SR003_OP' codedesc,SUM (mandays) hi, 0 shi, 0 sbi, 0 budget_hi, 0 budget_shi, 0 budget_sbi  " & _
					 "								 FROM attendance_gad " & _
					 "								 WHERE   jobcode = '6502101' " & _
					 "								 AND fddate = TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY') " & _
					 "								 AND mandays > 0 " & _
					 "								 group by fcba" & _
					 "								 UNION " & _
					 "								 SELECT fcba,'SR003_OP' codedesc,0 hi, SUM (mandays) shi,0 sbi, 0 budget_hi, 0 budget_shi, 0 budget_sbi  " & _
					 "								 FROM attendance_gad " & _
					 "								 WHERE  " & _
					 "								 jobcode = '6502101' " & _
					 "								 AND fddate <= TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY') " & _
					 "								 AND TO_CHAR (fddate, 'YYYY') = '" & v_txtyear & "' " & _
					 "								 AND TO_CHAR (fddate, 'MM') = '" & v_txtmonth & "' " & _
					 "								 AND mandays > 0 " & _
					 "								 group by fcba " & _
					 "								 UNION " & _
					 "								 SELECT fcba, 'SR003_OP' codedesc, 0 hi, 0 shi, SUM (mandays) sbi, 0 budget_hi,0 budget_shi, 0 budget_sbi  " & _
					 "								 FROM attendance_gad " & _
					 "								 WHERE   jobcode = '6502101' " & _
					 "								 AND fddate <= TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY') " & _
					 "								 AND TO_CHAR (fddate, 'YYYY') = '" & v_txtyear & "' " & _
					 "								 AND mandays > 0 " & _
					 "								 group by fcba " & _
					 "								 UNION " & _
					 "								 SELECT fccode fcba, 'SR003_OP' codedesc, 0 hi, 0 shi,0 sbi, " & _
					 "								 SUM(BUDGET_TRANSACTION_ESTATE_API.GetQuantityDay(fccode, '" & v_txtBudgetYear &"','" & v_txtBudgetVersion &"','" & v_txtModuleVersion & "','6502101',TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY'))) " & _
					 "								 / SUM( BUDGET_TRANSACTION_ESTATE_API.GetOutputDay(fccode, '" & v_txtBudgetYear &"','" & v_txtBudgetVersion &"','" & v_txtModuleVersion & "','6502101',TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY'))) budget_hi, " & _ 
					 "								 SUM(BUDGET_TRANSACTION_ESTATE_API.GetQuantityMonth(fccode, '" & v_txtBudgetYear &"','" & v_txtBudgetVersion &"','" & v_txtModuleVersion & "','6502101',TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY'))) " & _
					 "								 / SUM(BUDGET_TRANSACTION_ESTATE_API.GetOutputMonth(fccode, '" & v_txtBudgetYear &"','" & v_txtBudgetVersion &"','" & v_txtModuleVersion & "','6502101',TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY'))) budget_shi, " & _
					 "								 BUDGET_TRANSACTION_ESTATE_API.GetQuantityUntilMonth(fccode, '" & v_txtBudgetYear &"','" & v_txtBudgetVersion &"','" & v_txtModuleVersion & "','6502101',TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY')) " & _
					 "								 / SUM(BUDGET_TRANSACTION_ESTATE_API.GetOutputUntilMonth(fccode, '" & v_txtBudgetYear &"','" & v_txtBudgetVersion &"','" & v_txtModuleVersion & "','6502101',TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY'))) budget_sbi " & _
					 "								 FROM BUSINESSUNIT   " & _
					 "								 group by fccode  " & _
					 "							   ) " & _
					 "							GROUP BY fcba,codedesc  " & _                     
					 "							UNION " & _
					 "							SELECT fcba,codedesc, SUM (hi) hi,SUM (shi) shi,SUM (sbi) sbi,SUM (budget_hi) hi, SUM (budget_shi) shi, SUM (budget_sbi) sbi " & _
					 "							FROM ( SELECT fcba,'SR003_BA' AS codedesc, SUM (totalAmount) hi, 0 shi, 0 sbi, 0 budget_hi, 0 budget_shi, 0 budget_sbi  " & _
					 "								   FROM ( SELECT b.fcba,b.vehicledate,SUM (a.average * b.totalQuantity) TotalAmount " & _
					 "										  FROM ( SELECT fcba,vehiclecode, SUM (amount) / SUM (quantity) average " & _
					 "												 FROM vehiclecb " & _
					 "												 WHERE TO_CHAR (transdate, 'YYYYMM') = GetPreviousMonth ('" & v_txtyear & v_txtmonth & "') " & _
					 "												 AND jobcode IN ('6502201','6502202','6502203','6502204') " & _
					 "												 AND targettype = 'FF' " & _
					 "												 GROUP BY fcba,vehiclecode " & _
					 "												) a, " & _
					 "												( SELECT fcba,vehiclecode, vehicledate, SUM (unit) totalQuantity " & _
					 "												  FROM vehicleactivity " & _
					 "												  WHERE vehicledate = TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY') " & _
					 "												  AND jobcode IN ('6502201','6502202','6502203','6502204') " & _
					 "												  AND locationtype = 'FF' " & _
					 "												  GROUP BY fcba,vehiclecode, vehicledate " & _
					 "												) b " & _
					 "										   WHERE a.vehiclecode = b.vehiclecode " & _
					 "										   AND a.fcba=b.fcba " & _
					 "										   GROUP BY b.fcba,b.vehicledate " & _
					 "										  ) " & _
					 "								   GROUP BY fcba " & _
					 "								   UNION " & _
					 "								   SELECT fcba, 'SR003_BA' AS codedesc,  0 hi, SUM (totalAmount) shi, 0 sbi, 0 budget_hi, 0 budget_shi, 0 budget_sbi  " & _                          
					 "								   FROM ( SELECT b.fcba,b.vehicledate,  SUM (a.average * b.totalQuantity) TotalAmount " & _
					 "										  FROM (  SELECT fcba,vehiclecode, SUM (amount) / SUM (quantity) average " & _
					 "												  FROM vehiclecb " & _
					 "												  WHERE TO_CHAR (transdate, 'YYYYMM') = GetPreviousMonth ('" & v_txtyear & v_txtmonth & "') " & _
					 "												  AND jobcode IN ('6502201','6502202','6502203','6502204') " & _
					 "												  AND targettype = 'FF' " & _
					 "												  GROUP BY fcba,vehiclecode " & _
					 "												) a, " & _
					 "												( SELECT fcba,vehiclecode, vehicledate,SUM (unit) totalQuantity " & _
					 "												  FROM vehicleactivity " & _
					 "												  WHERE vehicledate <= TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY') " & _
					 "												  AND TO_CHAR (vehicledate, 'YYYY') = '" & v_txtyear & "' " & _
					 "												  AND TO_CHAR (vehicledate, 'MM') = '" & v_txtmonth & "' " & _
					 "												  AND jobcode IN ('6502201','6502202','6502203','6502204') " & _
					 "												  AND locationtype = 'FF' " & _
					 "												  GROUP BY fcba,vehiclecode, vehicledate " & _
					 "												) b " & _
					 "										   WHERE a.vehiclecode = b.vehiclecode " & _
					 "										   AND a.fcba=b.fcba " & _
					 "										   GROUP BY b.fcba,b.vehicledate " & _
					 "										  ) " & _
					 "								   GROUP BY fcba " & _
					 "								   UNION " & _
					 "								   SELECT fcba,'SR003_BA' AS codedesc, 0 hi,0 shi, SUM(totalAmount) sbi,0 budget_hi,0 budget_shi, 0 budget_sbi  " & _                         
					 "								   FROM ( SELECT b.fcba,b.vehicledate, SUM (a.average * b.totalQuantity) TotalAmount " & _
					 "										  FROM ( SELECT fcba,vehiclecode, SUM (amount) / SUM (quantity) average " & _
					 "												 FROM vehiclecb " & _
					 "												 WHERE TO_CHAR (transdate, 'YYYYMM') = GetPreviousMonth ('" & v_txtyear & v_txtmonth & "') " & _
					 "												 AND jobcode IN ('6502201','6502202','6502203','6502204') " & _
					 "	 											 AND targettype = 'FF' " & _
					 "												 GROUP BY fcba,vehiclecode " & _
					 "												) a," & _
					 "											   ( SELECT fcba,vehiclecode,vehicledate, SUM (unit) totalQuantity " & _
					 "												 FROM vehicleactivity " & _
					 "												 WHERE vehicledate <=  TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY') " & _
					 "												 AND jobcode IN ('6502201','6502202','6502203','6502204') " & _
					 "												 AND locationtype = 'FF' " & _
					 "												 GROUP BY fcba,vehiclecode, vehicledate " & _
					 "												) b " & _
					 "										   WHERE a.vehiclecode = b.vehiclecode " & _
					 "										   AND a.fcba=b.fcba " & _
					 "										   GROUP BY b.fcba,b.vehicledate " & _
					 "										) " & _
					 "								   GROUP BY fcba " & _
					 "								   UNION " & _
					 "								   SELECT fccode fcba, 'SR003_BA' AS codedesc,  0 hi, 0 shi, 0 sbi, 0 budget_hi, 0 budget_shi, 0 budget_sbi  from businessunit  " & _                        
					 "								) " & _
					 "								GROUP BY fcba,codedesc  " & _    
					 "								UNION " & _
					 "								SELECT fcba,'SR003_05' AS codedesc,DECODE(SUM(pom_tonase_hi),0, 0,SUM (mandays_amount_hi)/ (SUM (pom_tonase_hi) * 1000)) hi, 0 shi, 0 sbi, 0 budget_hi, 0 budget_shi, 0 budget_sbi " & _
					 "								FROM ( " & _
					 "									   SELECT fcba, accountingyear, accountingperiod, fieldcode,division, plantingyear, fddate, SUM(mandays_amount_hi) mandays_amount_hi,SUM (pom_tonase_hi) pom_tonase_hi " & _
					 "									   FROM ( SELECT fcba, accountingyear, accountingperiod, fieldcode, division, plantingyear,receptiondate fddate, 0 mandays_amount_hi,totaldailyreceivetonase1 pom_tonase_hi " & _
					 "											  FROM view_dailyharvesting_mill " & _
					 "											  WHERE accountingyear = '" & v_txtyear & "' AND accountingperiod = '" & v_txtmonth & "' " & _
					 "											  AND receptiondate = TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY') " & _
					 "											  UNION " & _
					 "											  SELECT fcba, accountingyear, accountingperiod,fieldcode, division, plantingyear, fddate,mandays_amount mandays_amount_hi,0 pom_tonase_hi " & _
					 "											  FROM view_dailyharvesting_amount " & _
					 "											  WHERE accountingyear = '" & v_txtyear & "' AND accountingperiod = '" & v_txtmonth & "' " & _
					 "											  AND fddate = TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY') " & _
					 "											)    " & _       
					 "									   GROUP BY fcba,accountingyear,accountingperiod,fieldcode,division, plantingyear,Fddate " & _
					 "									  ) " & _
					 "								 GROUP BY fcba " & _
					 "								 UNION " & _
					 "								 SELECT fcba,'SR003_05' AS codedesc,0 hi ,DECODE (SUM (pom_tonase_hi),0, 0,SUM (mandays_amount_hi) / (SUM (pom_tonase_hi) * 1000)) shi, 0 sbi , 0 budget_hi, 0 budget_shi, 0 budget_sbi " & _
					 "								 FROM ( SELECT fcba, accountingyear, accountingperiod, fieldcode,division, plantingyear, fddate,SUM (mandays_amount_hi) mandays_amount_hi,SUM(pom_tonase_hi) pom_tonase_hi " & _
					 "										FROM ( SELECT fcba, accountingyear, accountingperiod,fieldcode, division, plantingyear,receptiondate fddate, 0 mandays_amount_hi, totaldailyreceivetonase1 pom_tonase_hi " & _
					 "											   FROM view_dailyharvesting_mill " & _
					 "											   WHERE accountingyear = '" & v_txtyear & "' AND accountingperiod = '" & v_txtmonth & "' " & _
					 "											   AND receptiondate <= TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY') " & _
					 "											   UNION " & _
					 "											   SELECT fcba, accountingyear, accountingperiod,fieldcode, division, plantingyear, fddate, mandays_amount mandays_amount_hi, 0 pom_tonase_hi " & _
					 "											   FROM view_dailyharvesting_amount " & _
					 "											   WHERE accountingyear = '" & v_txtyear & "' AND accountingperiod = '" & v_txtmonth & "' " & _
					 "											   AND fddate <=  TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY') " & _
					 "											 ) " & _
					 "										GROUP BY fcba,accountingyear,accountingperiod,fieldcode,division, plantingyear,Fddate " & _
					 "									  ) " & _
					 "								 GROUP BY fcba " & _
					 "								 UNION " & _
					 "								 SELECT fcba,'SR003_05' AS codedesc, 0 hi, 0 shi,DECODE(SUM (pom_tonase_hi),0, 0,SUM (mandays_amount_hi)/ (SUM (pom_tonase_hi) * 1000)) sbi,  0 budget_hi, 0 budget_shi, 0 budget_sbi " & _
					 "								 FROM ( SELECT   fcba, accountingyear, accountingperiod, fieldcode,division, plantingyear, fddate,SUM (mandays_amount_hi) mandays_amount_hi, SUM (pom_tonase_hi) pom_tonase_hi " & _
					 "										FROM ( SELECT fcba, accountingyear, accountingperiod,fieldcode, division, plantingyear,receptiondate fddate, 0 mandays_amount_hi,totaldailyreceivetonase1 pom_tonase_hi " & _
					 "											   FROM view_dailyharvesting_mill " & _
					 "											   WHERE accountingyear = '" & v_txtyear & "' " & _
					 "											   AND receptiondate <= TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY')  " & _                                                            
					 "											   UNION " & _
					 "											   SELECT fcba, accountingyear, accountingperiod,fieldcode, division, plantingyear, fddate,mandays_amount mandays_amount_hi, 0 pom_tonase_hi " & _
					 "											   FROM view_dailyharvesting_amount " & _
					 "											   WHERE accountingyear = '" & v_txtyear & "'  " & _
					 "											   AND fddate <= TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY') " & _
					 "											  ) " & _
					 "										GROUP BY fcba,accountingyear,accountingperiod,fieldcode,division, plantingyear,Fddate " & _
					 "									   ) " & _
					 "								  GROUP BY fcba " & _
					 "							  ) " & _
					 "			   GROUP BY fcba,codedesc " & _
					 "				) " & _
					 "			  where ACCESSRIGHTSCOMPANY_API.IsAccessRights('" & Session("CodeName") & "',FCBA,'" & txtcompany & "') = 'TRUE' " & str1 & " " & _
					 "			   ORDER BY orderdata " & _
					 "		) group by codedesc, description, thenumber " & _
				     " ) "
  
 	Set rsProduction = Server.CreateObject("ADODB.Recordset")
	'response.write strProduction
	
	rsProduction.Open strProduction, connstr, adOpenStatic 
	
	
	' OnepointHalf
	' To Handle the Query of Hasil Kerja
	
						
	' strOutput = " SELECT codedesc, " & _
					   ' " ROWNUM theNumber, " & _
					   ' " groupCode, " & _
					   ' " description, " & _
					   ' " hi, " & _
					   ' " shi, " & _
					   ' " sbi, " & _
					   ' " budget_hi, " & _
					   ' " budget_shi, " & _
					   ' " budget_sbi " & _
				  ' " FROM (  SELECT codedesc, " & _
								 ' " groupCode, " & _
								 ' " orderdata, " & _
								 ' " param_2_name description, " & _
								 ' " SUM (hk_per_ha) hi, " & _
								 ' " SUM (hk_per_ha_shi) shi, " & _
								 ' " SUM (hk_per_ha_sbi) sbi, " & _
								 ' " SUM (budget_hk_per_ha) budget_hi, " & _
								 ' " SUM (budget_hk_per_ha_shi) budget_shi, " & _
								 ' " SUM (budget_hk_per_ha_sbi) budget_sbi " & _
							' " FROM (  SELECT 'SR003_55' codedesc, " & _
							               ' " groupCode, " & _ 
										   ' " orderdata, " & _
										   ' " param_2_name, " & _
										   ' " SUM (total) hk_per_ha, " & _
										   ' " 0 hk_per_ha_shi, " & _
										   ' " 0 hk_per_ha_sbi, " & _
										   ' " 0 budget_hk_per_ha, " & _
										   ' " 0 budget_hk_per_ha_shi, " & _
										   ' " 0 budget_hk_per_ha_sbi " & _
									  ' " FROM (SELECT b_temp.param_2, " & _
									               ' " b_temp.groupCode, " & _
												   ' " b_temp.orderdata, " & _
												   ' " b_temp.param_2_name, " & _
												   ' " 0 sku, " & _
												   ' " 0 khl, " & _
												   ' " 0 total, " & _
												   ' " 0 ha_harian, " & _
												   ' " 0 bo_harian, " & _
												   ' " 0 total_bo_harian, " & _
												   ' " 0 hk_per_ha, " & _
												   ' " 0 gli_ltr, " & _
												   ' " 0 gli_ltr_ha, " & _
												   ' " 0 par_ltr, " & _
												   ' " 0 par_ltr_ha, " & _
												   ' " 0 met_kg, " & _
												   ' " 0 met_kg_ha, " & _
												   ' " 0 flo_kg, " & _
												   ' " 0 flo_kg_ha, " & _
												   ' " 0 amount_hi " & _
											  ' " FROM JOB a_temp, " & _
												   ' " (SELECT grouphead.fcba, " & _
														   ' " grouphead.param_1, " & _
														   ' " grouphead.param_2, " & _
														   ' " groupname.fccode groupCode, " & _
														   ' " groupname.orderdata, " & _
														   ' " groupname.fcname param_2_name " & _
													  ' " FROM PARAMETERDETAIL grouphead, " & _
														   ' " (SELECT fcba, fccode, fcname, param_1 orderdata " & _
															 ' "  FROM PARAMETERDETAIL " & _
															 ' " WHERE parcode = 'SR002' " & _
																' "    AND parheadcode = 'GROUPJOB') groupname " & _
													 ' " WHERE grouphead.parcode = 'SR002' " & _
														   ' " AND grouphead.parheadcode IN " & _
																  ' " ('ACTIVITY_TBM', 'ACTIVITY_TM') " & _
														   ' " AND groupname.fcba = grouphead.fcba " & _
														   ' " AND groupname.fccode = " & _
																  ' " grouphead.param_2) b_temp " & _
											 ' " WHERE a_temp.fccode = b_temp.param_1 " & _
												   ' " AND a_temp.fcba = b_temp.fcba " & _
											' " UNION " & _
											  ' " SELECT c.param_2, " & _
											         ' " c.groupCode, " & _
													 ' " c.orderdata, " & _
													 ' " c.param_2_name, " & _
													 ' " SUM (a.sku) sku, " & _
													 ' " SUM (a.khl) khl, " & _
													 ' " SUM (a.ha_harian) total, " & _
													 ' " SUM (a.ha_harian) ha_harian, " & _
													 ' " SUM (a.bo_harian) bo_harian, " & _
													 ' " SUM (a.total_bo_harian) total_bo_harian, " & _
													 ' " DECODE (SUM (a.total), " & _
															 ' " 0, 0, " & _
															 ' " SUM (a.ha_harian) / SUM (a.total)) " & _
														' " hk_per_ha, " & _
													 ' " SUM (a.gli_ltr) gli_ltr, " & _
													 ' " DECODE ( " & _
														' " SUM (a.total_bo_harian), " & _
														' " 0, 0, " & _
														' " SUM (a.gli_ltr) / SUM (a.total_bo_harian)) " & _
														' " gli_ltr_ha, " & _
													 ' " SUM (a.par_ltr) par_ltr, " & _
													 ' " DECODE ( " & _
														' " SUM (a.total_bo_harian), " & _
														' " 0, 0, " & _
														' " SUM (a.par_ltr) / SUM (a.total_bo_harian)) " & _
														' " par_ltr_ha, " & _
													 ' " SUM (a.met_kg) met_kg, " & _
													 ' " DECODE ( " & _
														' " SUM (a.total_bo_harian), " & _
														' " 0, 0, " & _
														' " SUM (a.met_kg) / SUM (a.total_bo_harian)) " & _
														' " met_kg_ha, " & _
													 ' " SUM (flo_kg) flo_kg, " & _
													 ' " DECODE ( " & _
														' " SUM (a.total_bo_harian), " & _
														' " 0, 0, " & _
														' " SUM (flo_kg) / SUM (a.total_bo_harian)) " & _
														' " flo_kg_ha, " & _
													 ' " DECODE (SUM (a.ha_harian), " & _
															 ' " 0, 0, " & _
															 ' " SUM (a.amount_hi) / SUM (a.ha_harian)) " & _
														' " amount_hi_per_ha " & _
												' " FROM view_attendance_tm_smry a, " & _
													 ' " JOB b, " & _
													 ' " (SELECT grouphead.fcba, " & _
															 ' " grouphead.param_1, " & _
															 ' " grouphead.param_2, " & _
															 ' " groupname.fccode groupCode, " & _
															 ' " groupname.param_1 orderdata, " & _
															 ' " groupname.fcname param_2_name " & _
														' " FROM PARAMETERDETAIL grouphead, " & _
															 ' " (SELECT fcba, fccode, fcname, param_1 " & _
																' " FROM PARAMETERDETAIL " & _
															   ' " WHERE parcode = 'SR002' " & _
																	 ' " AND parheadcode = 'GROUPJOB') groupname " & _
													   ' " WHERE grouphead.parcode = 'SR002' " & _
															 ' " AND grouphead.parheadcode IN " & _
																	' " ('ACTIVITY_TBM', 'ACTIVITY_TM') " & _
															 ' " AND groupname.fcba = grouphead.fcba " & _
															 ' " AND groupname.fccode = " & _
																	' " grouphead.param_2) c " & _
											   ' " WHERE (a.fcba = b.fcba AND a.jobcode = b.fccode) " & _
													 ' " AND (a.fcba = c.fcba AND a.jobcode = c.param_1) " & _
													 ' " AND fddate BETWEEN TO_DATE ('" & v_txtdatefrom & "','MM/DD/YYYY') " & _
																	' " AND TO_DATE ('" & v_txtdatefrom & "','MM/DD/YYYY') " & _
											' " GROUP BY c.param_2, c.groupCode,  c.orderdata, c.param_2_name " & _
											' " ORDER BY param_2) " & _
								  ' " GROUP BY param_2, groupCode, orderdata, param_2_name " & _
								  ' " UNION " & _
									' " SELECT 'SR003_55' codedesc, " & _
									       ' " groupCode, " & _
										   ' " orderdata, " & _
										   ' " param_2_name, " & _
										   ' " 0 hk_per_ha, " & _
										   ' " SUM (total) hk_per_ha_shi, " & _
										   ' " 0 hk_per_ha_sbi, " & _
										   ' " 0 budget_hk_per_ha, " & _
										   ' " 0 budget_hk_per_ha_shi, " & _
										   ' " 0 budget_hk_per_ha_sbi " & _										   
									  ' " FROM (SELECT b_temp.param_2, " & _
									               ' " b_temp.groupCode, " & _
												   ' " b_temp.orderdata, " & _
												   ' " b_temp.param_2_name, " & _
												   ' " 0 sku, " & _
												   ' " 0 khl, " & _
												   ' " 0 total, " & _
												   ' " 0 ha_harian, " & _
												   ' " 0 bo_harian, " & _
												   ' " 0 total_bo_harian, " & _
												   ' " 0 hk_per_ha, " & _
												   ' " 0 gli_ltr, " & _
												   ' " 0 gli_ltr_ha, " & _
												   ' " 0 par_ltr, " & _
												   ' " 0 par_ltr_ha, " & _
												   ' " 0 met_kg, " & _
												   ' " 0 met_kg_ha, " & _
												   ' " 0 flo_kg, " & _
												   ' " 0 flo_kg_ha, " & _
												   ' " 0 amount_hi " & _
											  ' " FROM JOB a_temp, " & _
												   ' " (SELECT grouphead.fcba, " & _
														   ' " grouphead.param_1, " & _
														   ' " grouphead.param_2, " & _
														   ' " groupname.fccode groupCode, " & _
														   ' " groupname.param_1 orderdata, " & _
														   ' " groupname.fcname param_2_name " & _
													  ' " FROM PARAMETERDETAIL grouphead, " & _
														   ' " (SELECT fcba, fccode, fcname, param_1 " & _
															  ' " FROM PARAMETERDETAIL " & _
															 ' " WHERE parcode = 'SR002' " & _
																   ' " AND parheadcode = 'GROUPJOB') groupname " & _
													 ' " WHERE grouphead.parcode = 'SR002' " & _
														   ' " AND grouphead.parheadcode IN " & _
																  ' " ('ACTIVITY_TBM', 'ACTIVITY_TM') " & _
														   ' " AND groupname.fcba = grouphead.fcba " & _
														   ' " AND groupname.fccode = " & _
																  ' " grouphead.param_2) b_temp " & _
											 ' " WHERE a_temp.fccode = b_temp.param_1 " & _
												  ' " AND a_temp.fcba = b_temp.fcba " & _
											' " UNION " & _
											 ' "  SELECT c.param_2, " & _
											         ' " c.groupCode, " & _
													 ' " c.orderdata, " & _
													 ' " c.param_2_name, " & _
													 ' " SUM (a.sku) sku, " & _
													 ' " SUM (a.khl) khl, " & _
													 ' " SUM (a.ha_harian) total, " & _
													' "  SUM (a.ha_harian) ha_harian, " & _
													 ' " SUM (a.bo_harian) bo_harian, " & _
													 ' " SUM (a.total_bo_harian) total_bo_harian, " & _
													 ' " DECODE (SUM (a.total), " & _
															 ' " 0, 0, " & _
															' "  SUM (a.ha_harian) / SUM (a.total)) " & _
														' " hk_per_ha, " & _
													' "  SUM (a.gli_ltr) gli_ltr, " & _
													' "  DECODE ( " & _
													' " 	SUM (a.total_bo_harian), " & _
													' " 	0, 0, " & _
													' " 	SUM (a.gli_ltr) / SUM (a.total_bo_harian)) " & _
														' " gli_ltr_ha, " & _
													' "  SUM (a.par_ltr) par_ltr, " & _
													' "  DECODE ( " & _
													' " 	SUM (a.total_bo_harian), " & _
													' " 	0, 0, " & _
													' " 	SUM (a.par_ltr) / SUM (a.total_bo_harian)) " & _
													' " 	par_ltr_ha, " & _
													' "  SUM (a.met_kg) met_kg, " & _
													 ' " DECODE ( " & _
													' " 	SUM (a.total_bo_harian), " & _
													' " 	0, 0, " & _
														' " SUM (a.met_kg) / SUM (a.total_bo_harian)) " & _
														' " met_kg_ha, " & _
													' "  SUM (flo_kg) flo_kg, " & _
													 ' " DECODE ( " & _
														' " SUM (a.total_bo_harian), " & _
														' " 0, 0, " & _
														' " SUM (flo_kg) / SUM (a.total_bo_harian)) " & _
														' " flo_kg_ha, " & _
													' "  DECODE (SUM (a.ha_harian), " & _
															' "  0, 0, " & _
															' "  SUM (a.amount_hi) / SUM (a.ha_harian)) " & _
													' " 	amount_hi_per_ha " & _
												' " FROM view_attendance_tm_smry a, " & _
													' "  JOB b, " & _
													 ' " (SELECT grouphead.fcba, " & _
															' "  grouphead.param_1, " & _
															' "  grouphead.param_2, " & _
															' "  groupname.fccode groupCode, " & _
															' "  groupname.param_1 orderdata, " & _
															' "  groupname.fcname param_2_name " & _
														' " FROM PARAMETERDETAIL grouphead, " & _
															 ' " (SELECT fcba, fccode, fcname, param_1 " & _
																' " FROM PARAMETERDETAIL " & _
															  ' "  WHERE parcode = 'SR002' " & _
																	' "  AND parheadcode = 'GROUPJOB') groupname " & _
													  ' "  WHERE grouphead.parcode = 'SR002' " & _
															' "  AND grouphead.parheadcode IN " & _
																	' " ('ACTIVITY_TBM', 'ACTIVITY_TM') " & _
															' "  AND groupname.fcba = grouphead.fcba " & _ 
															 ' " AND groupname.fccode = " & _
																	' " grouphead.param_2) c " & _
											  ' "  WHERE (a.fcba = b.fcba AND a.jobcode = b.fccode) " & _
													' "  AND (a.fcba = c.fcba AND a.jobcode = c.param_1) " & _
													' "  AND fddate <= " & _
														' " 	TO_DATE ('" & v_txtdatefrom & "', 'MM/DD/YYYY') " & _
													' "  AND TO_CHAR (fddate, 'MM') = '" & v_txtMonth & "' " & _
													' "  AND TO_CHAR (fddate, 'YYYY') = '" & v_txtYear & "' " & _
											' " GROUP BY c.param_2,c.groupCode, c.orderdata,  c.param_2_name " & _
											' " ORDER BY param_2) " & _
								 ' "  GROUP BY param_2, groupCode, orderdata, param_2_name " & _
								' "   UNION " & _
								' " 	SELECT 'SR003_55' codedesc, " & _
									     ' "   groupCode, " & _
										 ' "   orderdata, " & _
										 ' "   param_2_name, " & _
										  ' "  0 hk_per_ha, " & _
										 ' "   0 hk_per_ha_shi, " & _
										 ' "   SUM (total) hk_per_ha_sbi, " & _
										   ' " 0 budget_hk_per_ha, " & _
										   ' " 0 budget_hk_per_ha_shi, " & _
										   ' " 0 budget_hk_per_ha_sbi " & _										 
									' "   FROM (SELECT b_temp.param_2, " & _
												 ' "   b_temp.groupCode, " & _
												 ' "   b_temp.orderdata, " & _
												 ' "   b_temp.param_2_name, " & _
												  ' "  0 sku, " & _
												  ' "  0 khl, " & _
												  ' "  0 total, " & _
												 ' "   0 ha_harian, " & _
												 ' "   0 bo_harian, " & _
												 ' "   0 total_bo_harian, " & _
												 ' "   0 hk_per_ha, " & _
												 ' "   0 gli_ltr, " & _
												 ' "  0 gli_ltr_ha, " & _
												 ' "   0 par_ltr, " & _
												 ' "   0 par_ltr_ha, " & _
												 ' "   0 met_kg, " & _
												  ' "  0 met_kg_ha, " & _
												  ' "  0 flo_kg, " & _
												  ' "  0 flo_kg_ha, " & _
												  ' "  0 amount_hi " & _
											 ' "  FROM JOB a_temp, " & _
												   ' " (SELECT grouphead.fcba, " & _
														' "    grouphead.param_1, " & _
														' "    grouphead.param_2, " & _
														' "    groupname.fccode groupCode, " & _
														' "    groupname.param_1 orderdata, " & _
														' "    groupname.fcname param_2_name " & _
													 ' "  FROM PARAMETERDETAIL grouphead, " & _
														' "    (SELECT fcba, fccode, fcname, param_1 " & _
														' " 	  FROM PARAMETERDETAIL " & _
														' " 	 WHERE parcode = 'SR002' " & _
															' " 	   AND parheadcode = 'GROUPJOB') groupname " & _
													' "  WHERE grouphead.parcode = 'SR002' " & _
												' " 		   AND grouphead.parheadcode IN " & _
												' " 				  ('ACTIVITY_TBM', 'ACTIVITY_TM') " & _
												' " 		   AND groupname.fcba = grouphead.fcba " & _
												' " 		   AND groupname.fccode = " & _
												' " 				  grouphead.param_2) b_temp " & _
											' "  WHERE a_temp.fccode = b_temp.param_1 " & _
											' " 	   AND a_temp.fcba = b_temp.fcba " & _
										' " 	UNION " & _
										' " 	  SELECT c.param_2, " & _
										' " 			 c.groupCode, " & _
										' " 			 c.orderdata, " & _
										' " 			 c.param_2_name, " & _
										' " 			 SUM (a.sku) sku, " & _
										' " 			 SUM (a.khl) khl, " & _
										' " 			 SUM (a.ha_harian) total, " & _
										' " 			 SUM (a.ha_harian) ha_harian, " & _
										' " 			 SUM (a.bo_harian) bo_harian, " & _
										' " 			 SUM (a.total_bo_harian) total_bo_harian, " & _
										' " 			 DECODE (SUM (a.total), " & _
										' " 					 0, 0, " & _
										' " 					 SUM (a.ha_harian) / SUM (a.total)) " & _
										' " 				hk_per_ha, " & _
										' " 			 SUM (a.gli_ltr) gli_ltr, " & _
										' " 			 DECODE ( " & _
										' " 				SUM (a.total_bo_harian), " & _
										' " 				0, 0, " & _
										' " 				SUM (a.gli_ltr) / SUM (a.total_bo_harian)) " & _
										' " 				gli_ltr_ha, " & _
										' " 			 SUM (a.par_ltr) par_ltr, " & _
										' " 			 DECODE ( " & _
										' " 				SUM (a.total_bo_harian), " & _
										' " 				0, 0, " & _
										' " 				SUM (a.par_ltr) / SUM (a.total_bo_harian)) " & _
										' " 				par_ltr_ha, " & _
										' " 			 SUM (a.met_kg) met_kg, " & _
										' " 			 DECODE ( " & _
										' " 				SUM (a.total_bo_harian), " & _
										' " 				0, 0, " & _
										' " 				SUM (a.met_kg) / SUM (a.total_bo_harian)) " & _
										' " 				met_kg_ha, " & _
										' " 			 SUM (flo_kg) flo_kg, " & _
											' " 		 DECODE ( " & _
											' " 			SUM (a.total_bo_harian), " & _
											' " 			0, 0, " & _
											' " 			SUM (flo_kg) / SUM (a.total_bo_harian)) " & _
											' " 			flo_kg_ha, " & _
											' " 		 DECODE (SUM (a.ha_harian), " & _
											' " 				 0, 0, " & _
											' " 				 SUM (a.amount_hi) / SUM (a.ha_harian)) " & _
											' " 			amount_hi_per_ha " & _
											' " 	FROM view_attendance_tm_smry a, " & _
											' " 		 JOB b, " & _
											' " 		 (SELECT grouphead.fcba, " & _
											' " 				 grouphead.param_1, " & _
											' " 				 grouphead.param_2, " & _
											' " 				 groupname.fccode groupCode, " & _
											' " 				 groupname.param_1 orderdata, " & _
											' " 				 groupname.fcname param_2_name " & _
											' " 			FROM PARAMETERDETAIL grouphead, " & _
											' " 				 (SELECT fcba, fccode, fcname, param_1 " & _
											' " 					FROM PARAMETERDETAIL " & _
											' " 				   WHERE parcode = 'SR002' " & _
											' " 						 AND parheadcode = 'GROUPJOB') groupname " & _
											' "		   WHERE grouphead.parcode = 'SR002' " & _
											' " 				 AND grouphead.parheadcode IN " & _
											' " 						('ACTIVITY_TBM', 'ACTIVITY_TM') " & _
											' " 				 AND groupname.fcba = grouphead.fcba " & _
											' " 				 AND groupname.fccode = " & _
											' " 						grouphead.param_2) c " & _
											' "    WHERE (a.fcba = b.fcba AND a.jobcode = b.fccode) " & _
											' " 		 AND (a.fcba = c.fcba AND a.jobcode = c.param_1) " & _
											' " 		 AND fddate <= " & _
											' " 				TO_DATE ('" & v_txtdatefrom & "', 'MM/DD/YYYY') " & _
											' " 		 AND TO_CHAR (fddate, 'YYYY') = '" & v_txtYear & "' " & _
										' " 	GROUP BY c.param_2, c.groupCode,  c.orderdata, c.param_2_name " & _
										' " 	ORDER BY param_2) " & _
								' "   GROUP BY param_2, groupCode, orderdata, param_2_name " & _								
								' " UNION " & _
								' "SELECT 'SR003_55' codedesc, " & _
							               ' " groupCode, " & _ 
										   ' " orderdata, " & _
										   ' " param_2_name, " & _
										   ' " 0 hk_per_ha, " & _
										   ' " 0 hk_per_ha_shi, " & _
										   ' " 0 hk_per_ha_sbi, " & _
										   ' " BUDGET_TRANSACTION_ESTATE_API.GetOutputDay(nvl('" & v_txtfcba & "' ,'" & session("fcba") & "'), " & _
										   									 ' v_txtBudgetYear & " , " & _
																			 ' v_txtBudgetVersion & " , " & _
																			 ' v_txtModuleVersion & " , " & _
																			' " a_temp.fccode, " & _
																			' " TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY')) budget_hk_per_ha , " & _
										   ' " BUDGET_TRANSACTION_ESTATE_API.GetOutputMonth(nvl('" & v_txtfcba & "' ,'" & session("fcba") & "'), " & _
										   									 ' v_txtBudgetYear & " , " & _
																			 ' v_txtBudgetVersion & " , " & _
																			 ' v_txtModuleVersion & " , " & _
																			' " a_temp.fccode, " & _
																			' " TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY')) budget_hk_per_ha_shi , " & _
										   ' " BUDGET_TRANSACTION_ESTATE_API.GetOutputUntilMonth(nvl('" & v_txtfcba & "' ,'" & session("fcba") & "'), " & _
										   									 ' v_txtBudgetYear & " , " & _
																			 ' v_txtBudgetVersion & " , " & _
																			 ' v_txtModuleVersion & " , " & _
																			' " a_temp.fccode, " & _
																			' " TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY')) budget_hk_per_ha_sbi " & _
										' "  FROM JOB a_temp, " & _
												   ' " (SELECT grouphead.fcba, " & _
														' "    grouphead.param_1, " & _
														' "    grouphead.param_2, " & _
														' "    groupname.fccode groupCode, " & _
														' "    groupname.param_1 orderdata, " & _
														' "    groupname.fcname param_2_name " & _
													 ' "  FROM PARAMETERDETAIL grouphead, " & _
														' "    (SELECT fcba, fccode, fcname, param_1 " & _
														' " 	  FROM PARAMETERDETAIL " & _
														' " 	 WHERE parcode = 'SR002' " & _
															' " 	   AND parheadcode = 'GROUPJOB') groupname " & _
													' "  WHERE grouphead.parcode = 'SR002' " & _
												' " 		   AND grouphead.parheadcode IN " & _
												' " 				  ('ACTIVITY_TBM', 'ACTIVITY_TM') " & _
												' " 		   AND groupname.fcba = grouphead.fcba " & _
												' " 		   AND groupname.fccode = " & _
												' " 				  grouphead.param_2) b_temp " & _
											' "  WHERE a_temp.fccode = b_temp.param_1 " & _
											' " 	   AND a_temp.fcba = b_temp.fcba  " & _
								' ") " & _
						' " GROUP BY codedesc, groupCode, orderdata, param_2_name)" 
	strOutput = "  SELECT codedesc, ROWNUM theNumber,groupCode,description,hi,shi,sbi,budget_hi,budget_shi,budget_sbi " & _
				"  FROM  ( SELECT codedesc, groupCode, orderdata,param_2_name description,SUM (hk_per_ha) hi,SUM (hk_per_ha_shi) shi,SUM (hk_per_ha_sbi) sbi, " & _
				"					SUM (budget_hk_per_ha) budget_hi,SUM (budget_hk_per_ha_shi) budget_shi,SUM (budget_hk_per_ha_sbi) budget_sbi " & _
				"			 FROM (  SELECT fcba,'SR003_55' codedesc,groupCode,orderdata,param_2_name,SUM (total) hk_per_ha,0 hk_per_ha_shi,0 hk_per_ha_sbi, " & _
				"							0 budget_hk_per_ha,0 budget_hk_per_ha_shi, 0 budget_hk_per_ha_sbi " & _
				"					 FROM ( SELECT a_temp.fcba,b_temp.param_2,b_temp.groupCode,b_temp.orderdata,b_temp.param_2_name,0 sku,0 khl,0 total,0 ha_harian,0 bo_harian, " & _
				"								   0 total_bo_harian,0 hk_per_ha,0 gli_ltr, 0 gli_ltr_ha,0 par_ltr,0 par_ltr_ha, 0 met_kg,0 met_kg_ha,0 flo_kg,0 flo_kg_ha,0 amount_hi " & _
				"							FROM JOB a_temp, " & _
				"								( SELECT grouphead.fcba, grouphead.param_1,grouphead.param_2,groupname.fccode groupCode,groupname.orderdata, groupname.fcname param_2_name " & _
				"								  FROM PARAMETERDETAIL grouphead, " & _
				"								  ( SELECT fcba, fccode, fcname, param_1 orderdata " & _
				"									FROM PARAMETERDETAIL " & _
				"									WHERE parcode = 'SR002' " & _
				"									 AND parheadcode = 'GROUPJOB' " & _
				"								  ) groupname " & _
				"								  WHERE grouphead.parcode = 'SR002' " & _
				"								  AND grouphead.parheadcode IN  ('ACTIVITY_TBM', 'ACTIVITY_TM') " & _
				"								  AND groupname.fcba = grouphead.fcba " & _
				"								  AND groupname.fccode = grouphead.param_2 " & _
				"								) b_temp " & _
				"							WHERE a_temp.fccode = b_temp.param_1 " & _
				"							AND a_temp.fcba = b_temp.fcba " & _
				"							UNION " & _
				"							SELECT c.fcba,c.param_2,c.groupCode,c.orderdata, c.param_2_name,SUM(a.sku) sku,SUM (a.khl) khl,SUM(a.ha_harian) total,SUM(a.ha_harian) ha_harian, " & _
				"								   SUM(a.bo_harian) bo_harian,SUM(a.total_bo_harian) total_bo_harian,DECODE(SUM(a.total),0, 0,SUM (a.ha_harian) / SUM (a.total))hk_per_ha, " & _
				"								   SUM(a.gli_ltr) gli_ltr,DECODE(SUM(a.total_bo_harian),0, 0,SUM (a.gli_ltr) / SUM (a.total_bo_harian)) gli_ltr_ha, " & _
				"								   SUM(a.par_ltr) par_ltr,DECODE(SUM (a.total_bo_harian),0, 0,SUM (a.par_ltr) / SUM (a.total_bo_harian)) par_ltr_ha, " & _
				"								   SUM(a.met_kg) met_kg,DECODE(SUM (a.total_bo_harian),0, 0,SUM (a.met_kg) / SUM (a.total_bo_harian)) met_kg_ha, " & _
				"								   SUM(flo_kg) flo_kg,DECODE(SUM (a.total_bo_harian),0, 0,SUM (flo_kg) / SUM (a.total_bo_harian)) flo_kg_ha, " & _
				"								   DECODE(SUM(a.ha_harian), 0, 0,SUM (a.amount_hi) / SUM (a.ha_harian)) amount_hi_per_ha " & _
				"							 FROM view_attendance_tm_smry a, JOB b, " & _
				"							 ( SELECT grouphead.fcba, grouphead.param_1,grouphead.param_2,groupname.fccode groupCode,groupname.param_1 orderdata, " & _
				"									  groupname.fcname param_2_name " & _
				"							   FROM PARAMETERDETAIL grouphead, " & _
				"							   ( SELECT fcba, fccode, fcname, param_1 " & _
				"								 FROM PARAMETERDETAIL " & _
				"								 WHERE parcode = 'SR002' " & _
				"								 AND parheadcode = 'GROUPJOB' " & _
				"							   ) groupname " & _
				"							   WHERE grouphead.parcode = 'SR002' " & _
				"							   AND grouphead.parheadcode IN ('ACTIVITY_TBM', 'ACTIVITY_TM') " & _
				"							   AND groupname.fcba = grouphead.fcba " & _
				"							   AND groupname.fccode = grouphead.param_2 " & _
				"							 ) c " & _
				"							 WHERE (a.fcba = b.fcba AND a.jobcode = b.fccode) " & _
				"							 AND (a.fcba = c.fcba AND a.jobcode = c.param_1) " & _
				"							 AND fddate BETWEEN TO_DATE ('" & v_txtdatefrom & "','MM/DD/YYYY') " & _
				"							 AND TO_DATE ('" & v_txtdatefrom  & "','MM/DD/YYYY') " & _
				"							 GROUP BY c.fcba,c.param_2, c.groupCode,  c.orderdata, c.param_2_name " & _
				"							 ORDER BY param_2 " & _
				"						  ) " & _
				"					 GROUP BY fcba,param_2, groupCode, orderdata, param_2_name " & _
				"					 UNION    " & _             
				"					 SELECT fcba,'SR003_55' codedesc,groupCode,orderdata,param_2_name,0 hk_per_ha,SUM (total) hk_per_ha_shi, " & _
				"							0 hk_per_ha_sbi,0 budget_hk_per_ha,0 budget_hk_per_ha_shi, 0 budget_hk_per_ha_sbi   " & _                     
				"					 FROM ( SELECT a_temp.fcba,b_temp.param_2,b_temp.groupCode, b_temp.orderdata,b_temp.param_2_name,0 sku,0 khl,0 total,0 ha_harian, " & _
				"								   0 bo_harian,0 total_bo_harian,0 hk_per_ha,0 gli_ltr,0 gli_ltr_ha,0 par_ltr,0 par_ltr_ha,0 met_kg, " & _
				"								   0 met_kg_ha,0 flo_kg,0 flo_kg_ha,0 amount_hi " & _
				"							FROM JOB a_temp, " & _
				"							( SELECT grouphead.fcba,grouphead.param_1,grouphead.param_2, groupname.fccode groupCode,groupname.param_1 orderdata,groupname.fcname param_2_name " & _
				"							  FROM PARAMETERDETAIL grouphead, " & _
				"							  ( SELECT fcba, fccode, fcname, param_1 " & _
				"								FROM PARAMETERDETAIL " & _
				"								WHERE parcode = 'SR002' " & _
				"								AND parheadcode = 'GROUPJOB' " & _
				"							  ) groupname " & _
				"							  WHERE grouphead.parcode = 'SR002' " & _
				"							  AND grouphead.parheadcode IN ('ACTIVITY_TBM', 'ACTIVITY_TM') " & _
				"							  AND groupname.fcba = grouphead.fcba " & _
				"							  AND groupname.fccode = grouphead.param_2 " & _
				"							) b_temp " & _
				"							WHERE a_temp.fccode = b_temp.param_1 " & _
				"							AND a_temp.fcba = b_temp.fcba " & _
				"							UNION " & _
				"							SELECT c.fcba,c.param_2,c.groupCode,c.orderdata,c.param_2_name,SUM (a.sku) sku,SUM (a.khl) khl,SUM (a.ha_harian) total,SUM (a.ha_harian) ha_harian, " & _
				"								   SUM(a.bo_harian) bo_harian,SUM (a.total_bo_harian) total_bo_harian, " & _
				"								   DECODE(SUM (a.total),0, 0,SUM (a.ha_harian) / SUM (a.total)) hk_per_ha, " & _
				"								   SUM(a.gli_ltr) gli_ltr,DECODE(SUM(a.total_bo_harian),0, 0,SUM(a.gli_ltr) / SUM(a.total_bo_harian)) gli_ltr_ha, " & _
				"								   SUM(a.par_ltr) par_ltr,DECODE( SUM(a.total_bo_harian),0, 0,SUM (a.par_ltr) / SUM (a.total_bo_harian)) par_ltr_ha, " & _
				"								   SUM(a.met_kg) met_kg,DECODE(SUM (a.total_bo_harian),0, 0,SUM (a.met_kg) / SUM (a.total_bo_harian)) met_kg_ha, " & _
				"								   SUM (flo_kg) flo_kg,DECODE(SUM (a.total_bo_harian),0, 0,SUM (flo_kg) / SUM (a.total_bo_harian)) flo_kg_ha, " & _
				"								   DECODE(SUM (a.ha_harian),0, 0,SUM (a.amount_hi) / SUM (a.ha_harian)) amount_hi_per_ha " & _
				"							 FROM view_attendance_tm_smry a, JOB b, " & _
				"							 ( SELECT grouphead.fcba, grouphead.param_1,grouphead.param_2,groupname.fccode groupCode,groupname.param_1 orderdata,groupname.fcname param_2_name " & _
				"							   FROM PARAMETERDETAIL grouphead, " & _
				"									( SELECT fcba, fccode, fcname, param_1 " & _
				"									  FROM PARAMETERDETAIL " & _
				"									  WHERE parcode = 'SR002' " & _
				"									  AND parheadcode = 'GROUPJOB' " & _
				"									) groupname " & _
				"							   WHERE grouphead.parcode = 'SR002' " & _
				"							   AND grouphead.parheadcode IN ('ACTIVITY_TBM', 'ACTIVITY_TM') " & _
				"							   AND groupname.fcba = grouphead.fcba  " & _
				"							   AND groupname.fccode = grouphead.param_2 " & _
				"							 ) c " & _
				"							 WHERE (a.fcba = b.fcba AND a.jobcode = b.fccode) " & _
				"							 AND (a.fcba = c.fcba AND a.jobcode = c.param_1) " & _
				"							 AND fddate <= TO_DATE ('" & v_txtdatefrom & "', 'MM/DD/YYYY') " & _
				"							 AND TO_CHAR (fddate, 'MM') = '" & v_txtmonth & "' " & _
				"							 AND TO_CHAR (fddate, 'YYYY') = '" & v_txtyear & "' " & _
				"							 GROUP BY c.fcba,c.param_2,c.groupCode, c.orderdata,  c.param_2_name " & _
				"							 ORDER BY param_2 " & _
				"						 )     " & _               
				"					GROUP BY fcba,param_2, groupCode, orderdata, param_2_name " & _
				"					UNION 	" & _				
				"					SELECT fcba,'SR003_55' codedesc,groupCode,orderdata,param_2_name,0 hk_per_ha,0 hk_per_ha_shi, " & _
				"							SUM (total) hk_per_ha_sbi,0 budget_hk_per_ha,0 budget_hk_per_ha_shi,0 budget_hk_per_ha_sbi  " & _                    
				"					FROM ( SELECT a_temp.fcba,b_temp.param_2,b_temp.groupCode,b_temp.orderdata,b_temp.param_2_name,0 sku,0 khl,0 total,0 ha_harian, " & _
				"								  0 bo_harian,0 total_bo_harian,0 hk_per_ha,0 gli_ltr,0 gli_ltr_ha,0 par_ltr,0 par_ltr_ha,0 met_kg, " & _
				"								  0 met_kg_ha,0 flo_kg,0 flo_kg_ha,0 amount_hi " & _
				"						   FROM JOB a_temp, " & _
				"								( SELECT grouphead.fcba,grouphead.param_1,grouphead.param_2,groupname.fccode groupCode,groupname.param_1 orderdata,groupname.fcname param_2_name " & _
				"								  FROM PARAMETERDETAIL grouphead, " & _
				"								  ( SELECT fcba, fccode, fcname, param_1 " & _
				"									FROM PARAMETERDETAIL " & _
				"									WHERE parcode = 'SR002' " & _
				"									AND parheadcode = 'GROUPJOB' " & _
				"								  ) groupname " & _
				"								  WHERE grouphead.parcode = 'SR002' " & _
				"								  AND grouphead.parheadcode IN ('ACTIVITY_TBM', 'ACTIVITY_TM') " & _
				"								  AND groupname.fcba = grouphead.fcba " & _
				"								  AND groupname.fccode = grouphead.param_2 " & _
				"								) b_temp " & _
				"						  WHERE a_temp.fccode = b_temp.param_1 " & _
				"						  AND a_temp.fcba = b_temp.fcba " & _
				"						  UNION " & _
				"						  SELECT c.fcba,c.param_2,c.groupCode,c.orderdata,c.param_2_name,SUM(a.sku) sku, SUM(a.khl) khl,SUM(a.ha_harian) total, SUM(a.ha_harian) ha_harian, " & _
				"								 SUM(a.bo_harian) bo_harian,SUM(a.total_bo_harian) total_bo_harian,DECODE (SUM (a.total),0, 0,SUM (a.ha_harian) / SUM (a.total)) hk_per_ha, " & _
				"								 SUM(a.gli_ltr) gli_ltr,DECODE(SUM(a.total_bo_harian),0, 0,SUM (a.gli_ltr) / SUM (a.total_bo_harian)) gli_ltr_ha, " & _
				"								 SUM(a.par_ltr) par_ltr,DECODE(SUM(a.total_bo_harian),0, 0,SUM (a.par_ltr) / SUM (a.total_bo_harian)) par_ltr_ha, " & _
				"								 SUM(a.met_kg) met_kg,DECODE(SUM (a.total_bo_harian),0, 0,SUM (a.met_kg) / SUM (a.total_bo_harian)) met_kg_ha, " & _
				"								 SUM(flo_kg) flo_kg,DECODE(SUM(a.total_bo_harian),0, 0,SUM (flo_kg) / SUM (a.total_bo_harian)) flo_kg_ha, " & _
				"								 DECODE(SUM (a.ha_harian),0, 0,SUM(a.amount_hi) / SUM (a.ha_harian)) amount_hi_per_ha " & _
				"						  FROM view_attendance_tm_smry a, JOB b, " & _
				"						  ( SELECT grouphead.fcba, grouphead.param_1, grouphead.param_2, groupname.fccode groupCode,groupname.param_1 orderdata, groupname.fcname param_2_name " & _
				"							FROM PARAMETERDETAIL grouphead, " & _
				"							( SELECT fcba, fccode, fcname, param_1 " & _
				"							  FROM PARAMETERDETAIL " & _
				"							  WHERE parcode = 'SR002' " & _
				"							  AND parheadcode = 'GROUPJOB' " & _
				"							) groupname " & _
				"							WHERE grouphead.parcode = 'SR002' " & _
				"							AND grouphead.parheadcode IN ('ACTIVITY_TBM', 'ACTIVITY_TM') " & _
				"							AND groupname.fcba = grouphead.fcba " & _
				"							AND groupname.fccode = grouphead.param_2 " & _
				"						   ) c " & _
				"						   WHERE (a.fcba = b.fcba AND a.jobcode = b.fccode) " & _
				"						   AND (a.fcba = c.fcba AND a.jobcode = c.param_1) " & _
				"						   AND fddate <=  TO_DATE ('" & v_txtdatefrom & "', 'MM/DD/YYYY') " & _
				"						   AND TO_CHAR (fddate, 'YYYY') = '" & v_txtyear & "' " & _
				"						   GROUP BY c.fcba,c.param_2, c.groupCode,  c.orderdata, c.param_2_name " & _
				"						   ORDER BY param_2 " & _
				"					 ) " & _
				"				   GROUP BY fcba,param_2, groupCode, orderdata, param_2_name  " & _                             
				"				   UNION " & _
				"				   SELECT a_temp.fcba,'SR003_55' codedesc,groupCode, orderdata, param_2_name,0 hk_per_ha, 0 hk_per_ha_shi, 0 hk_per_ha_sbi, " & _
				"						  BUDGET_TRANSACTION_ESTATE_API.GetOutputDay(a_temp.fcba,'" & v_txtBudgetYear &"','" & v_txtBudgetVersion &"','" & v_txtModuleVersion & "',a_temp.fccode,TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY')) budget_hk_per_ha , " & _
				"						  BUDGET_TRANSACTION_ESTATE_API.GetOutputMonth(a_temp.fcba,'" & v_txtBudgetYear &"','" & v_txtBudgetVersion &"','" & v_txtModuleVersion & "',a_temp.fccode,TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY')) budget_hk_per_ha_shi , " & _
				"						  BUDGET_TRANSACTION_ESTATE_API.GetOutputUntilMonth(a_temp.fcba,'" & v_txtBudgetYear &"','" & v_txtBudgetVersion &"','" & v_txtModuleVersion & "',a_temp.fccode,TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY')) budget_hk_per_ha_sbi " & _
				"				   FROM JOB a_temp,  " & _        
				"				   ( SELECT grouphead.fcba,grouphead.param_1,grouphead.param_2,groupname.fccode groupCode,groupname.param_1 orderdata,groupname.fcname param_2_name " & _
				"					 FROM PARAMETERDETAIL grouphead, " & _
				"					 ( SELECT fcba, fccode, fcname, param_1 " & _
				"					   FROM PARAMETERDETAIL " & _
				"					   WHERE parcode = 'SR002' " & _
				"					   AND parheadcode = 'GROUPJOB' " & _
				"					 ) groupname " & _
				"					 WHERE grouphead.parcode = 'SR002' " & _
				"					 AND grouphead.parheadcode IN ('ACTIVITY_TBM', 'ACTIVITY_TM') " & _
				"					 AND groupname.fcba = grouphead.fcba " & _
				"					 AND groupname.fccode = grouphead.param_2 " & _
				"				   ) b_temp " & _
				"					 WHERE a_temp.fccode = b_temp.param_1 " & _
				"					 AND a_temp.fcba = b_temp.fcba  " & _  
				"				  ) " & _
				"			     where ACCESSRIGHTSCOMPANY_API.IsAccessRights('" & Session("CodeName") & "',FCBA,'" & txtcompany & "') = 'TRUE' " & str1 & " " & _ 
				"				 GROUP BY codedesc, groupCode, orderdata, param_2_name ) "


    	
 	Set rsOutput = Server.CreateObject("ADODB.Recordset")
	'response.write strOutput
	
	rsOutput.Open strOutput, connstr, adOpenStatic 
	
	
 	' Second
    ' To Handle the Query of Perawatan Output Ratio
	
	' strHandling = " SELECT codedesc, ROWNUM theNumber,groupCode, description, hi, shi, sbi, budget_hi, budget_shi,  budget_sbi FROM ( " & _
	              ' " SELECT codedesc, groupCode, param_2_name description , sum(hk_per_ha) hi, sum(hk_per_ha_shi) shi, sum(hk_per_ha_sbi) sbi, " & _
                        ' " sum(budget_hk_per_ha) budget_hi, sum(budget_hk_per_ha_shi) budget_shi, sum(budget_hk_per_ha_sbi) budget_sbi FROM ( " & _
						' " SELECT   'SR003_06' codedesc , orderdata, groupCode, param_2_name,  DECODE(SUM (total), 0, 0,  SUM (total) / SUM (ha_harian)) hk_per_ha, " & _
						    ' " 0 hk_per_ha_shi, 0 hk_per_ha_sbi, " & _
							' " 0 budget_hk_per_ha, 0 budget_hk_per_ha_shi, 0 budget_hk_per_ha_sbi " & _
							' " FROM (SELECT   b_temp.param_2,b_temp.groupCode, b_temp.orderdata, b_temp.param_2_name, 0 sku, 0 khl, 0 total, " & _
										   ' " 0 ha_harian, 0 bo_harian, 0 total_bo_harian, 0 hk_per_ha, " & _
										   ' " 0 gli_ltr, 0 gli_ltr_ha, 0 par_ltr, 0 par_ltr_ha, 0 met_kg, " & _
										   ' " 0 met_kg_ha, 0 flo_kg, 0 flo_kg_ha, 0 amount_hi " & _
									  ' " FROM JOB a_temp, " & _
										   ' " (SELECT grouphead.fcba, grouphead.param_1, " & _
												   ' " grouphead.param_2,groupname.fccode groupCode, groupname.param_1 orderdata, groupname.fcname param_2_name " & _
											  ' " FROM PARAMETERDETAIL grouphead, " & _
												   ' " (SELECT fcba, fccode, fcname, param_1 " & _
													  ' " FROM PARAMETERDETAIL " & _
													 ' " WHERE parcode = 'SR002' " & _
													   ' " AND parheadcode = 'GROUPJOB') groupname " & _
											 ' " WHERE grouphead.parcode = 'SR002' " & _
											   ' " AND grouphead.parheadcode IN " & _
																	  ' " ('ACTIVITY_TBM', 'ACTIVITY_TM') " & _
											   ' " AND groupname.fcba = grouphead.fcba " & _
											   ' " AND groupname.fccode = grouphead.param_2) b_temp " & _
									 ' " WHERE a_temp.fccode = b_temp.param_1 " & _
										   ' " AND a_temp.fcba = b_temp.fcba " & _
								  ' " UNION " & _
								  ' " SELECT   c.param_2, c.orderdata, c.groupCode, c.param_2_name, SUM (a.sku) sku, SUM (a.khl) khl, " & _
										   ' " SUM (a.total) total, SUM (a.ha_harian) ha_harian, " & _
										   ' " SUM (a.bo_harian) bo_harian, " & _
										   ' " SUM (a.total_bo_harian) total_bo_harian, " & _
										   ' " DECODE (SUM (a.total), " & _
												   ' " 0, 0, " & _
												   ' " SUM (a.ha_harian) / SUM (a.total) " & _
												  ' " ) hk_per_ha, " & _
										   ' " SUM (a.gli_ltr) gli_ltr, " & _
										   ' " DECODE (SUM (a.total_bo_harian), " & _
												   ' " 0, 0, " & _
												   ' " SUM (a.gli_ltr) / SUM (a.total_bo_harian) " & _
												  ' " ) gli_ltr_ha, " & _
										   ' " SUM (a.par_ltr) par_ltr, " & _
										   ' " DECODE (SUM (a.total_bo_harian), " & _
												   ' " 0, 0, " & _
												   ' " SUM (a.par_ltr) / SUM (a.total_bo_harian) " & _
												  ' " ) par_ltr_ha, " & _
										   ' " SUM (a.met_kg) met_kg, " & _
										   ' " DECODE (SUM (a.total_bo_harian), " & _
												   ' " 0, 0, " & _
												   ' " SUM (a.met_kg) / SUM (a.total_bo_harian) " & _
												  ' " ) met_kg_ha, " & _
										   ' " SUM (flo_kg) flo_kg, " & _
										   ' " DECODE (SUM (a.total_bo_harian), " & _
												   ' " 0, 0, " & _
												   ' " SUM (flo_kg) / SUM (a.total_bo_harian) " & _
												  ' " ) flo_kg_ha, " & _
										   ' " DECODE (SUM (a.ha_harian), " & _
												   ' " 0, 0, " & _
												   ' " SUM (a.amount_hi) / SUM (a.ha_harian) " & _
												  ' " ) amount_hi_per_ha " & _
									  ' " FROM view_attendance_tm_smry a, " & _
										   ' " JOB b, " & _
										   ' " (SELECT grouphead.fcba, grouphead.param_1, " & _
												   ' " grouphead.param_2, groupname.fccode groupCode, groupname.param_1 orderdata, groupname.fcname param_2_name " & _
											  ' " FROM PARAMETERDETAIL grouphead, " & _
												   ' " (SELECT fcba, fccode, fcname, param_1 " & _
													  ' " FROM PARAMETERDETAIL " & _
													 ' " WHERE parcode = 'SR002' " & _
													   ' " AND parheadcode = 'GROUPJOB') groupname " & _
											 ' " WHERE grouphead.parcode = 'SR002' " & _
											   ' " AND grouphead.parheadcode IN " & _
																	  ' " ('ACTIVITY_TBM', 'ACTIVITY_TM') " & _
											   ' " AND groupname.fcba = grouphead.fcba " & _
											   ' " AND groupname.fccode = grouphead.param_2) c " & _
									 ' " WHERE (a.fcba = b.fcba AND a.jobcode = b.fccode) " & _
									   ' " AND (a.fcba = c.fcba AND a.jobcode = c.param_1) " & _
									   ' " AND fddate BETWEEN TO_DATE ('" & v_txtdatefrom & "', 'MM/DD/YYYY') " & _
													  ' " AND TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY') " & _
								  ' " GROUP BY c.param_2, c.groupCode, c.orderdata, c.param_2_name " & _
								  ' " ORDER BY param_2) " & _
						' " GROUP BY param_2, groupcode, orderdata, param_2_name " & _
						' " UNION " & _
						' " SELECT   'SR003_06' codedesc , orderdata,  groupCode, param_2_name, 0 hk_per_ha,  DECODE(SUM (total), 0, 0,  SUM (total) / SUM (ha_harian)) hk_per_ha_shi , " & _
						    ' " 0 hk_per_ha_sbi, " & _
							' " 0 budget_hk_per_ha, 0 budget_hk_per_ha_shi, 0 budget_hk_per_ha_sbi " & _
							' " FROM (SELECT   b_temp.groupCode, b_temp.param_2, b_temp.orderdata, b_temp.param_2_name, 0 sku, 0 khl, 0 total, " & _
										   ' " 0 ha_harian, 0 bo_harian, 0 total_bo_harian, 0 hk_per_ha, " & _
										   ' " 0 gli_ltr, 0 gli_ltr_ha, 0 par_ltr, 0 par_ltr_ha, 0 met_kg, " & _
										   ' " 0 met_kg_ha, 0 flo_kg, 0 flo_kg_ha, 0 amount_hi " & _
									  ' " FROM JOB a_temp, " & _
										   ' " (SELECT grouphead.fcba, grouphead.param_1, " & _
												   ' " grouphead.param_2, groupname.fccode groupCode, groupname.param_1 orderdata, groupname.fcname param_2_name " & _
											  ' " FROM PARAMETERDETAIL grouphead, " & _
												   ' " (SELECT fcba, fccode, fcname, param_1 " & _
													  ' " FROM PARAMETERDETAIL " & _
													 ' " WHERE parcode = 'SR002' " & _
													   ' " AND parheadcode = 'GROUPJOB') groupname " & _
											 ' " WHERE grouphead.parcode = 'SR002' " & _
											   ' " AND grouphead.parheadcode IN " & _
																	  ' " ('ACTIVITY_TBM', 'ACTIVITY_TM') " & _
											   ' " AND groupname.fcba = grouphead.fcba " & _
											   ' " AND groupname.fccode = grouphead.param_2) b_temp " & _
									 ' " WHERE a_temp.fccode = b_temp.param_1 " & _
										   ' " AND a_temp.fcba = b_temp.fcba " & _
								  ' " UNION " & _
								  ' " SELECT   c.param_2, c.orderdata, c.groupCode, c.param_2_name, SUM (a.sku) sku, SUM (a.khl) khl, " & _
										   ' " SUM (a.total) total, SUM (a.ha_harian) ha_harian, " & _
										   ' " SUM (a.bo_harian) bo_harian, " & _
										   ' " SUM (a.total_bo_harian) total_bo_harian, " & _
										   ' " DECODE (SUM (a.total), " & _
												   ' " 0, 0, " & _
												   ' " SUM (a.ha_harian) / SUM (a.total) " & _
												  ' " ) hk_per_ha, " & _
										   ' " SUM (a.gli_ltr) gli_ltr, " & _
										   ' " DECODE (SUM (a.total_bo_harian), " & _
												   ' " 0, 0, " & _
												   ' " SUM (a.gli_ltr) / SUM (a.total_bo_harian) " & _
												  ' " ) gli_ltr_ha, " & _
										   ' " SUM (a.par_ltr) par_ltr, " & _
										   ' " DECODE (SUM (a.total_bo_harian), " & _
												   ' " 0, 0, " & _
												   ' " SUM (a.par_ltr) / SUM (a.total_bo_harian) " & _
												  ' " ) par_ltr_ha, " & _
										   ' " SUM (a.met_kg) met_kg, " & _
										   ' " DECODE (SUM (a.total_bo_harian), " & _
												   ' " 0, 0, " & _
												   ' " SUM (a.met_kg) / SUM (a.total_bo_harian) " & _
												  ' " ) met_kg_ha, " & _
										   ' " SUM (flo_kg) flo_kg, " & _
										   ' " DECODE (SUM (a.total_bo_harian), " & _
												   ' " 0, 0, " & _
												   ' " SUM (flo_kg) / SUM (a.total_bo_harian) " & _
												  ' " ) flo_kg_ha, " & _
										   ' " DECODE (SUM (a.ha_harian), " & _
												   ' " 0, 0, " & _
												   ' " SUM (a.amount_hi) / SUM (a.ha_harian) " & _
												  ' " ) amount_hi_per_ha " & _
									  ' " FROM view_attendance_tm_smry a, " & _
										   ' " JOB b, " & _
										   ' " (SELECT grouphead.fcba, grouphead.param_1, " & _
												   ' " grouphead.param_2, groupname.fccode groupCode, groupname.param_1 orderdata, groupname.fcname param_2_name " & _
											  ' " FROM PARAMETERDETAIL grouphead, " & _
												   ' " (SELECT fcba, fccode, fcname, param_1 " & _
													  ' " FROM PARAMETERDETAIL " & _
													 ' " WHERE parcode = 'SR002' " & _
													   ' " AND parheadcode = 'GROUPJOB') groupname " & _
											 ' " WHERE grouphead.parcode = 'SR002' " & _
											   ' " AND grouphead.parheadcode IN " & _
																	  ' " ('ACTIVITY_TBM', 'ACTIVITY_TM') " & _
											   ' " AND groupname.fcba = grouphead.fcba " & _
											   ' " AND groupname.fccode = grouphead.param_2) c " & _
									 ' " WHERE (a.fcba = b.fcba AND a.jobcode = b.fccode) " & _
									   ' " AND (a.fcba = c.fcba AND a.jobcode = c.param_1) " & _
									 ' " AND fddate <= TO_DATE ('" & v_txtdatefrom & "', 'MM/DD/YYYY') " & _
									   ' " AND to_char(fddate, 'MM' ) = '" & v_txtmonth & "' " & _
								       ' " AND to_char(fddate, 'YYYY') = '" & v_txtyear & "' " & _		
								  ' " GROUP BY c.param_2, c.groupCode, c.orderdata, c.param_2_name " & _
								  ' " ORDER BY param_2) " & _
						' " GROUP BY param_2, orderdata, groupCode, param_2_name " & _
						' " UNION " & _
						' " SELECT   'SR003_06' codedesc , orderdata, groupCode,  param_2_name, 0 hk_per_ha, 0 hk_per_ha_shi, DECODE(SUM (total), 0, 0,  SUM (total) / SUM (ha_harian)) hk_per_ha_sbi , " & _
							' " 0 budget_hk_per_ha, 0 budget_hk_per_ha_shi, 0 budget_hk_per_ha_sbi " & _
							' " FROM (SELECT   b_temp.groupCode, b_temp.param_2, b_temp.orderdata, b_temp.param_2_name, 0 sku, 0 khl, 0 total, " & _
										   ' " 0 ha_harian, 0 bo_harian, 0 total_bo_harian, 0 hk_per_ha, " & _
										   ' " 0 gli_ltr, 0 gli_ltr_ha, 0 par_ltr, 0 par_ltr_ha, 0 met_kg, " & _
										   ' " 0 met_kg_ha, 0 flo_kg, 0 flo_kg_ha, 0 amount_hi " & _
									  ' " FROM JOB a_temp, " & _
										   ' " (SELECT grouphead.fcba, grouphead.param_1, " & _
												   ' " grouphead.param_2,groupname.fccode groupCode, groupname.param_1 orderdata, groupname.fcname param_2_name " & _
											  ' " FROM PARAMETERDETAIL grouphead, " & _
												   ' " (SELECT fcba, fccode, fcname, param_1 " & _
													  ' " FROM PARAMETERDETAIL " & _
													 ' " WHERE parcode = 'SR002' " & _
													   ' " AND parheadcode = 'GROUPJOB') groupname " & _
											 ' " WHERE grouphead.parcode = 'SR002' " & _
											   ' " AND grouphead.parheadcode IN " & _
																	  ' " ('ACTIVITY_TBM', 'ACTIVITY_TM') " & _
											   ' " AND groupname.fcba = grouphead.fcba " & _
											   ' " AND groupname.fccode = grouphead.param_2) b_temp " & _
									 ' " WHERE a_temp.fccode = b_temp.param_1 " & _
										   ' " AND a_temp.fcba = b_temp.fcba " & _
								  ' " UNION " & _
								  ' " SELECT   c.param_2, c.orderdata, c.groupCode, c.param_2_name, SUM (a.sku) sku, SUM (a.khl) khl, " & _
										   ' " SUM (a.total) total, SUM (a.ha_harian) ha_harian, " & _
										   ' " SUM (a.bo_harian) bo_harian, " & _
										   ' " SUM (a.total_bo_harian) total_bo_harian, " & _
										   ' " DECODE (SUM (a.total), " & _
												   ' " 0, 0, " & _
												   ' " SUM (a.ha_harian) / SUM (a.total) " & _
												  ' " ) hk_per_ha, " & _
										   ' " SUM (a.gli_ltr) gli_ltr, " & _
										   ' " DECODE (SUM (a.total_bo_harian), " & _
												   ' " 0, 0, " & _
												   ' " SUM (a.gli_ltr) / SUM (a.total_bo_harian) " & _
												  ' " ) gli_ltr_ha, " & _
										   ' " SUM (a.par_ltr) par_ltr, " & _
										   ' " DECODE (SUM (a.total_bo_harian), " & _
												   ' " 0, 0, " & _
												   ' " SUM (a.par_ltr) / SUM (a.total_bo_harian) " & _
												  ' " ) par_ltr_ha, " & _
										   ' " SUM (a.met_kg) met_kg, " & _
										   ' " DECODE (SUM (a.total_bo_harian), " & _
												   ' " 0, 0, " & _
												   ' " SUM (a.met_kg) / SUM (a.total_bo_harian) " & _
												  ' " ) met_kg_ha, " & _
										   ' " SUM (flo_kg) flo_kg, " & _
										   ' " DECODE (SUM (a.total_bo_harian), " & _
												   ' " 0, 0, " & _
												   ' " SUM (flo_kg) / SUM (a.total_bo_harian) " & _
												  ' " ) flo_kg_ha, " & _
										   ' " DECODE (SUM (a.ha_harian), " & _
												   ' " 0, 0, " & _
												   ' " SUM (a.amount_hi) / SUM (a.ha_harian) " & _
												  ' " ) amount_hi_per_ha " & _
									  ' " FROM view_attendance_tm_smry a, " & _
										   ' " JOB b, " & _
										   ' " (SELECT grouphead.fcba, grouphead.param_1, " & _
												   ' " grouphead.param_2, groupname.fccode groupCode, groupname.param_1 orderdata, groupname.fcname param_2_name " & _
											  ' " FROM PARAMETERDETAIL grouphead, " & _
												   ' " (SELECT fcba, fccode, fcname, param_1 " & _
													  ' " FROM PARAMETERDETAIL " & _
													 ' " WHERE parcode = 'SR002' " & _
													   ' " AND parheadcode = 'GROUPJOB') groupname " & _
											 ' " WHERE grouphead.parcode = 'SR002' " & _
											   ' " AND grouphead.parheadcode IN " & _
																	  ' " ('ACTIVITY_TBM', 'ACTIVITY_TM') " & _
											   ' " AND groupname.fcba = grouphead.fcba " & _
											   ' " AND groupname.fccode = grouphead.param_2) c " & _
									 ' " WHERE (a.fcba = b.fcba AND a.jobcode = b.fccode) " & _
									   ' " AND (a.fcba = c.fcba AND a.jobcode = c.param_1) " & _
									   ' " AND fddate <= TO_DATE ('" & v_txtdatefrom & "', 'MM/DD/YYYY') " & _
								       ' " AND to_char(fddate, 'YYYY') = '" & v_txtyear & "' " & _		
								  ' " GROUP BY c.param_2, c.orderdata, c.groupCode, c.param_2_name " & _
								  ' " ORDER BY param_2) " & _
						' " GROUP BY param_2,orderdata, groupCode, param_2_name " & _
						' " UNION " & _
								' "SELECT 'SR003_06' codedesc, " & _
							               ' " groupCode, " & _ 
										   ' " orderdata, " & _
										   ' " param_2_name, " & _
										   ' " 0 hk_per_ha, " & _
										   ' " 0 hk_per_ha_shi, " & _
										   ' " 0 hk_per_ha_sbi, " & _
											' " sum(BUDGET_TRANSACTION_ESTATE_API.GetQuantityDay(nvl('" & v_txtfcba & "' ,'" & session("fcba") & "'), " & _
										   									 ' v_txtBudgetYear & " , " & _
																			 ' v_txtBudgetVersion & " , " & _
																			 ' v_txtModuleVersion & " , " & _
																			' " a_temp.fccode, " & _
																			' " TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY'))) /  " & _										   
										   ' " sum(BUDGET_TRANSACTION_ESTATE_API.GetOutputDay(nvl('" & v_txtfcba & "' ,'" & session("fcba") & "'), " & _
										   									 ' v_txtBudgetYear & " , " & _
																			 ' v_txtBudgetVersion & " , " & _
																			 ' v_txtModuleVersion & " , " & _
																			' " a_temp.fccode, " & _
																			' " TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY'))) budget_hk_per_ha , " & _
										   ' " sum(BUDGET_TRANSACTION_ESTATE_API.GetQuantityMonth(nvl('" & v_txtfcba & "' ,'" & session("fcba") & "'), " & _
										   									 ' v_txtBudgetYear & " , " & _
																			 ' v_txtBudgetVersion & " , " & _
																			 ' v_txtModuleVersion & " , " & _
																			' " a_temp.fccode, " & _
																			' " TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY'))) / " & _
										   ' " sum(BUDGET_TRANSACTION_ESTATE_API.GetOutputMonth(nvl('" & v_txtfcba & "' ,'" & session("fcba") & "'), " & _
										   									 ' v_txtBudgetYear & " , " & _
																			 ' v_txtBudgetVersion & " , " & _
																			 ' v_txtModuleVersion & " , " & _
																			' " a_temp.fccode, " & _
																			' " TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY'))) budget_hk_per_ha_shi , " & _
										   ' " sum(BUDGET_TRANSACTION_ESTATE_API.GetQuantityUntilMonth(nvl('" & v_txtfcba & "' ,'" & session("fcba") & "'), " & _
										   									 ' v_txtBudgetYear & " , " & _
																			 ' v_txtBudgetVersion & " , " & _
																			 ' v_txtModuleVersion & " , " & _
																			' " a_temp.fccode, " & _
																			' " TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY'))) / " & _																			
										   ' " sum(BUDGET_TRANSACTION_ESTATE_API.GetOutputUntilMonth(nvl('" & v_txtfcba & "' ,'" & session("fcba") & "'), " & _
										   									 ' v_txtBudgetYear & " , " & _
																			 ' v_txtBudgetVersion & " , " & _
																			 ' v_txtModuleVersion & " , " & _
																			' " a_temp.fccode, " & _
																			' " TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY'))) budget_hk_per_ha_sbi " & _
										' "  FROM JOB a_temp, " & _
												   ' " (SELECT grouphead.fcba, " & _
														' "    grouphead.param_1, " & _
														' "    grouphead.param_2, " & _
														' "    groupname.fccode groupCode, " & _
														' "    groupname.param_1 orderdata, " & _
														' "    groupname.fcname param_2_name " & _
													 ' "  FROM PARAMETERDETAIL grouphead, " & _
														' "    (SELECT fcba, fccode, fcname, param_1 " & _
														' " 	  FROM PARAMETERDETAIL " & _
														' " 	 WHERE parcode = 'SR002' " & _
															' " 	   AND parheadcode = 'GROUPJOB') groupname " & _
													' "  WHERE grouphead.parcode = 'SR002' " & _
												' " 		   AND grouphead.parheadcode IN " & _
												' " 				  ('ACTIVITY_TBM', 'ACTIVITY_TM') " & _
												' " 		   AND groupname.fcba = grouphead.fcba " & _
												' " 		   AND groupname.fccode = " & _
												' " 				  grouphead.param_2) b_temp " & _
											' "  WHERE a_temp.fccode = b_temp.param_1 " & _
											' " 	   AND a_temp.fcba = b_temp.fcba group by groupCode, orderdata, param_2_name" & _						
						' " ) " & _
						' " GROUP BY codedesc, orderdata, groupCode,  param_2_name ) " 
	strHandling = "  SELECT codedesc, ROWNUM theNumber,groupCode, description, hi, shi, sbi, budget_hi, budget_shi,  budget_sbi " & _
				  "	 FROM ( SELECT codedesc, groupCode, param_2_name description , sum(hk_per_ha) hi, sum(hk_per_ha_shi) shi, sum(hk_per_ha_sbi) sbi, " & _
				  "				   sum(budget_hk_per_ha) budget_hi, sum(budget_hk_per_ha_shi) budget_shi, sum(budget_hk_per_ha_sbi) budget_sbi " & _
				  "				FROM ( SELECT fcba, 'SR003_06' codedesc , orderdata, groupCode, param_2_name,  DECODE(SUM (total), 0, 0,  SUM (total) / SUM (ha_harian)) hk_per_ha, " & _
				  "								0 hk_per_ha_shi, 0 hk_per_ha_sbi,0 budget_hk_per_ha, 0 budget_hk_per_ha_shi, 0 budget_hk_per_ha_sbi " & _
				  "					   FROM ( SELECT b_temp.fcba,  b_temp.param_2,b_temp.groupCode, b_temp.orderdata, b_temp.param_2_name, 0 sku, 0 khl, 0 total, " & _
				  "								0 ha_harian, 0 bo_harian, 0 total_bo_harian, 0 hk_per_ha,0 gli_ltr, 0 gli_ltr_ha, 0 par_ltr, 0 par_ltr_ha, 0 met_kg, " & _
				  "								0 met_kg_ha, 0 flo_kg, 0 flo_kg_ha, 0 amount_hi " & _
				  "							  FROM JOB a_temp, " & _
				  "								  ( SELECT grouphead.fcba, grouphead.param_1,grouphead.param_2,groupname.fccode groupCode, groupname.param_1 orderdata, " & _
				  "										   groupname.fcname param_2_name " & _
				  "	 								FROM PARAMETERDETAIL grouphead, " & _
				  "									( SELECT fcba, fccode, fcname, param_1 " & _
				  "									  FROM PARAMETERDETAIL " & _
				  "									  WHERE parcode = 'SR002' " & _
				  "									  AND parheadcode = 'GROUPJOB' " & _
				  "									) groupname " & _
				  "									WHERE grouphead.parcode = 'SR002' " & _
				  "									AND grouphead.parheadcode IN ('ACTIVITY_TBM', 'ACTIVITY_TM') " & _
				  "									AND groupname.fcba = grouphead.fcba " & _
				  "									AND groupname.fccode = grouphead.param_2 " & _
				  "								) b_temp " & _
				  "							  WHERE a_temp.fccode = b_temp.param_1 " & _
				  "							  AND a_temp.fcba = b_temp.fcba " & _
				  "							  UNION " & _
				  "							  SELECT c.fcba,c.param_2, c.orderdata, c.groupCode, c.param_2_name, SUM (a.sku) sku, SUM(a.khl) khl, " & _
				  "								SUM(a.total) total, SUM(a.ha_harian) ha_harian,SUM(a.bo_harian) bo_harian,SUM(a.total_bo_harian) total_bo_harian, " & _
				  "								DECODE(SUM(a.total),0, 0,SUM(a.ha_harian) / SUM(a.total)) hk_per_ha,SUM(a.gli_ltr) gli_ltr, " & _
				  "								DECODE (SUM(a.total_bo_harian),0, 0,SUM(a.gli_ltr) / SUM(a.total_bo_harian)) gli_ltr_ha, " & _
				  "								SUM(a.par_ltr) par_ltr,DECODE (SUM (a.total_bo_harian),0, 0,SUM(a.par_ltr) / SUM(a.total_bo_harian)) par_ltr_ha, " & _
				  "								SUM(a.met_kg) met_kg,DECODE (SUM (a.total_bo_harian),0, 0,SUM (a.met_kg) / SUM (a.total_bo_harian)) met_kg_ha," & _
				  "								SUM(flo_kg) flo_kg,DECODE(SUM(a.total_bo_harian),0, 0,SUM (flo_kg) / SUM (a.total_bo_harian)) flo_kg_ha, " & _
				  "								DECODE(SUM (a.ha_harian),0, 0,SUM (a.amount_hi) / SUM (a.ha_harian)) amount_hi_per_ha " & _
				  "							 FROM view_attendance_tm_smry a, JOB b, " & _
				  "							 ( SELECT grouphead.fcba, grouphead.param_1,grouphead.param_2, groupname.fccode groupCode, groupname.param_1 orderdata, " & _
				  "									 groupname.fcname param_2_name " & _
				  "							   FROM PARAMETERDETAIL grouphead, " & _
				  "							   ( SELECT fcba, fccode, fcname, param_1 " & _
				  "								 FROM PARAMETERDETAIL " & _
				  "								 WHERE parcode = 'SR002' " & _
				  "								 AND parheadcode = 'GROUPJOB' " & _
				  "							   ) groupname " & _
				  "							   WHERE grouphead.parcode = 'SR002' " & _
				  "							   AND grouphead.parheadcode IN ('ACTIVITY_TBM', 'ACTIVITY_TM') " & _
				  "							   AND groupname.fcba = grouphead.fcba " & _
				  "							   AND groupname.fccode = grouphead.param_2 " & _
				  "							 ) c " & _
				  "							 WHERE (a.fcba = b.fcba AND a.jobcode = b.fccode) " & _
				  "							 AND (a.fcba = c.fcba AND a.jobcode = c.param_1) " & _
				  "							 AND fddate BETWEEN TO_DATE ('" & v_txtdatefrom & "', 'MM/DD/YYYY') " & _
				  "							 AND TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY') " & _
				  "							 GROUP BY c.fcba,c.param_2, c.groupCode, c.orderdata, c.param_2_name " & _
				  "							 ORDER BY param_2 " & _
				  "						  ) " & _
				  "					   GROUP BY fcba,param_2, groupcode, orderdata, param_2_name " & _			 
				  "					   UNION " & _
				  "					   SELECT fcba, 'SR003_06' codedesc , orderdata,  groupCode, param_2_name, 0 hk_per_ha,  DECODE(SUM (total), 0, 0,  SUM (total) / SUM (ha_harian)) hk_per_ha_shi , " & _
				  "							  0 hk_per_ha_sbi,0 budget_hk_per_ha, 0 budget_hk_per_ha_shi, 0 budget_hk_per_ha_sbi " & _
				  "					   FROM ( SELECT b_temp.fcba, b_temp.groupCode, b_temp.param_2, b_temp.orderdata, b_temp.param_2_name, 0 sku, 0 khl, 0 total, " & _
				  "									 0 ha_harian, 0 bo_harian, 0 total_bo_harian, 0 hk_per_ha, 0 gli_ltr, 0 gli_ltr_ha, 0 par_ltr, 0 par_ltr_ha, 0 met_kg, " & _
				  "									 0 met_kg_ha, 0 flo_kg, 0 flo_kg_ha, 0 amount_hi " & _
				  "							  FROM JOB a_temp, " & _
				  "							  ( SELECT grouphead.fcba, grouphead.param_1,grouphead.param_2, groupname.fccode groupCode, groupname.param_1 orderdata, groupname.fcname param_2_name " & _
				  "								FROM PARAMETERDETAIL grouphead, " & _
				  "								( SELECT fcba, fccode, fcname, param_1 " & _
				  "								  FROM PARAMETERDETAIL " & _
				  "								  WHERE parcode = 'SR002' " & _
				  "								  AND parheadcode = 'GROUPJOB' " & _
				  "								) groupname " & _
				  "								WHERE grouphead.parcode = 'SR002' " & _
				  "								AND grouphead.parheadcode IN ('ACTIVITY_TBM', 'ACTIVITY_TM') " & _
				  "								AND groupname.fcba = grouphead.fcba " & _
				  "								AND groupname.fccode = grouphead.param_2 " & _
				  "							  ) b_temp " & _
				  "							  WHERE a_temp.fccode = b_temp.param_1 " & _
				  "							  AND a_temp.fcba = b_temp.fcba " & _
				  "							  UNION " & _
				  "							  SELECT c.fcba,c.param_2, c.orderdata, c.groupCode, c.param_2_name, SUM(a.sku) sku, SUM(a.khl) khl, " & _
				  "									 SUM(a.total) total, SUM(a.ha_harian) ha_harian,SUM (a.bo_harian) bo_harian,SUM (a.total_bo_harian) total_bo_harian, " & _
				  "									 DECODE(SUM(a.total),0, 0,SUM(a.ha_harian) / SUM(a.total)) hk_per_ha,SUM(a.gli_ltr) gli_ltr, " & _
				  "									 DECODE(SUM(a.total_bo_harian),0, 0,SUM(a.gli_ltr) / SUM(a.total_bo_harian)) gli_ltr_ha,SUM(a.par_ltr) par_ltr, " & _
				  "									 DECODE(SUM(a.total_bo_harian),0, 0,SUM (a.par_ltr) / SUM (a.total_bo_harian)) par_ltr_ha, " & _
				  "									 SUM(a.met_kg) met_kg,DECODE (SUM (a.total_bo_harian),0, 0,SUM(a.met_kg) / SUM(a.total_bo_harian)) met_kg_ha, " & _
				  "									 SUM(flo_kg) flo_kg,DECODE(SUM(a.total_bo_harian),0, 0,SUM (flo_kg) / SUM (a.total_bo_harian)) flo_kg_ha, " & _
				  "									 DECODE(SUM(a.ha_harian),0, 0,SUM(a.amount_hi) / SUM(a.ha_harian)) amount_hi_per_ha " & _
				  "							  FROM view_attendance_tm_smry a,JOB b, " & _
				  "							  ( SELECT grouphead.fcba, grouphead.param_1,grouphead.param_2, groupname.fccode groupCode, groupname.param_1 orderdata, " & _
				  "									  groupname.fcname param_2_name " & _
				  "								FROM PARAMETERDETAIL grouphead, " & _
				  "								( SELECT fcba, fccode, fcname, param_1 " & _
				  "								  FROM PARAMETERDETAIL " & _
				  "								  WHERE parcode = 'SR002' " & _
				  "								  AND parheadcode = 'GROUPJOB' " & _
				  "								) groupname " & _
				  "								WHERE grouphead.parcode = 'SR002' " & _
				  "								AND grouphead.parheadcode IN ('ACTIVITY_TBM', 'ACTIVITY_TM') " & _
				  "								AND groupname.fcba = grouphead.fcba " & _
				  "								AND groupname.fccode = grouphead.param_2 " & _
				  "							  ) c " & _
				  "							  WHERE (a.fcba = b.fcba AND a.jobcode = b.fccode) " & _
				  "	 						  AND (a.fcba = c.fcba AND a.jobcode = c.param_1) " & _
				  "							  AND fddate <= TO_DATE ('" & v_txtdatefrom & "', 'MM/DD/YYYY') " & _
				  "							  AND to_char(fddate, 'MM' ) = ' & v_txtmonth & ' " & _
				  "							  AND to_char(fddate, 'YYYY') = '" & v_txtyear & "'  " & _   
				  "							  GROUP BY c.fcba, c.param_2, c.groupCode, c.orderdata, c.param_2_name " & _
				  "							 ORDER BY param_2 " & _
				  "						 ) " & _
				  "	 				 GROUP BY fcba, param_2, orderdata, groupCode, param_2_name " & _
				  "					 UNION " & _
				  "					 SELECT fcba,'SR003_06' codedesc , orderdata, groupCode,  param_2_name, 0 hk_per_ha, 0 hk_per_ha_shi, DECODE(SUM (total), 0, 0,  SUM (total) / SUM (ha_harian)) hk_per_ha_sbi , " & _
				  "  							0 budget_hk_per_ha, 0 budget_hk_per_ha_shi, 0 budget_hk_per_ha_sbi " & _
				  "					 FROM ( SELECT b_temp.fcba,  b_temp.groupCode, b_temp.param_2, b_temp.orderdata, b_temp.param_2_name, 0 sku, 0 khl, 0 total, " & _
				  "								   0 ha_harian, 0 bo_harian, 0 total_bo_harian, 0 hk_per_ha,0 gli_ltr, 0 gli_ltr_ha, 0 par_ltr, 0 par_ltr_ha, 0 met_kg, " & _
				  "								   0 met_kg_ha, 0 flo_kg, 0 flo_kg_ha, 0 amount_hi " & _
				  "							FROM JOB a_temp, " & _
				  "							( SELECT grouphead.fcba, grouphead.param_1,grouphead.param_2,groupname.fccode groupCode, groupname.param_1 orderdata, groupname.fcname param_2_name " & _
				  "							  FROM PARAMETERDETAIL grouphead, " & _
				  "							  ( SELECT fcba, fccode, fcname, param_1 " & _
				  "								FROM PARAMETERDETAIL " & _
				  "								WHERE parcode = 'SR002' " & _
				  "								AND parheadcode = 'GROUPJOB' " & _
				  "							  ) groupname " & _
				  "							  WHERE grouphead.parcode = 'SR002' " & _
				  "							  AND grouphead.parheadcode IN ('ACTIVITY_TBM', 'ACTIVITY_TM') " & _
				  "							  AND groupname.fcba = grouphead.fcba " & _
				  "							  AND groupname.fccode = grouphead.param_2 " & _
				  "							) b_temp " & _
				  "							WHERE a_temp.fccode = b_temp.param_1 " & _
				  "							AND a_temp.fcba = b_temp.fcba " & _
				  "						   UNION " & _
				  "						   SELECT c.fcba,c.param_2, c.orderdata, c.groupCode, c.param_2_name, SUM (a.sku) sku, SUM (a.khl) khl, " & _
				  "								  SUM (a.total) total, SUM (a.ha_harian) ha_harian,SUM (a.bo_harian) bo_harian,SUM (a.total_bo_harian) total_bo_harian, " & _
				  "								  DECODE(SUM(a.total),0, 0,SUM(a.ha_harian) / SUM(a.total)) hk_per_ha, SUM(a.gli_ltr) gli_ltr, " & _
				  "								  DECODE(SUM(a.total_bo_harian),0, 0,SUM (a.gli_ltr) / SUM (a.total_bo_harian)) gli_ltr_ha, " & _
				  "								  SUM(a.par_ltr) par_ltr,DECODE (SUM (a.total_bo_harian),0, 0,SUM (a.par_ltr) / SUM (a.total_bo_harian)) par_ltr_ha, " & _
				  "								  SUM(a.met_kg) met_kg,DECODE (SUM (a.total_bo_harian),0, 0,SUM(a.met_kg) / SUM(a.total_bo_harian)) met_kg_ha, " & _
				  "								  SUM(flo_kg) flo_kg,DECODE(SUM(a.total_bo_harian),0, 0,SUM(flo_kg) / SUM(a.total_bo_harian)) flo_kg_ha, " & _
				  "								  DECODE(SUM (a.ha_harian),0, 0,SUM (a.amount_hi) / SUM (a.ha_harian)) amount_hi_per_ha " & _
				  "							 FROM view_attendance_tm_smry a,JOB b, " & _
				  "							 ( SELECT grouphead.fcba, grouphead.param_1,grouphead.param_2, groupname.fccode groupCode, groupname.param_1 orderdata, groupname.fcname param_2_name " & _
				  "							   FROM PARAMETERDETAIL grouphead, " & _
				  "									( SELECT fcba, fccode, fcname, param_1 " & _
				  "									  FROM PARAMETERDETAIL " & _
				  "									  WHERE parcode = 'SR002' " & _
				  "									  AND parheadcode = 'GROUPJOB' " & _
				  "									) groupname " & _
				  "							   WHERE grouphead.parcode = 'SR002' " & _
				  "							   AND grouphead.parheadcode IN ('ACTIVITY_TBM', 'ACTIVITY_TM') " & _
				  "							   AND groupname.fcba = grouphead.fcba " & _
				  "							   AND groupname.fccode = grouphead.param_2 " & _
				  "							  ) c " & _
				  "							  WHERE (a.fcba = b.fcba AND a.jobcode = b.fccode) " & _
				  "							  AND (a.fcba = c.fcba AND a.jobcode = c.param_1) " & _
				  "							  AND fddate <= TO_DATE ('" & v_txtdatefrom & "', 'MM/DD/YYYY') " & _
				  "							  AND to_char(fddate, 'YYYY') = '" & v_txtyear & "' " & _    
				  "							  GROUP BY c.fcba,c.param_2, c.orderdata, c.groupCode, c.param_2_name " & _
				  "							  ORDER BY param_2" & _
				  "						   ) " & _
				  "					 GROUP BY fcba,param_2,orderdata, groupCode, param_2_name " & _	 
				  "					 UNION " & _
				  "					 SELECT a_temp.fcba,'SR003_06' codedesc,groupCode,orderdata,param_2_name,0 hk_per_ha,0 hk_per_ha_shi,0 hk_per_ha_sbi, " & _
				  "							sum(BUDGET_TRANSACTION_ESTATE_API.GetQuantityDay(a_temp.fcba,'" & v_txtBudgetYear &"','" & v_txtBudgetVersion &"','" & v_txtModuleVersion & "',a_temp.fccode,TO_DATE('" & v_txtdateto & "', 'MM/DD/YYYY'))) " & _
				  "							/ sum(BUDGET_TRANSACTION_ESTATE_API.GetOutputDay(a_temp.fcba,'" & v_txtBudgetYear &"','" & v_txtBudgetVersion &"','" & v_txtModuleVersion & "',a_temp.fccode,TO_DATE('" & v_txtdateto & "', 'MM/DD/YYYY'))) budget_hk_per_ha , " & _
				  "							sum(BUDGET_TRANSACTION_ESTATE_API.GetQuantityMonth(a_temp.fcba,'" & v_txtBudgetYear &"','" & v_txtBudgetVersion &"','" & v_txtModuleVersion & "',a_temp.fccode,TO_DATE('" & v_txtdateto & "', 'MM/DD/YYYY'))) " & _
				  "							/ sum(BUDGET_TRANSACTION_ESTATE_API.GetOutputMonth(a_temp.fcba,'" & v_txtBudgetYear &"','" & v_txtBudgetVersion &"','" & v_txtModuleVersion & "',a_temp.fccode,TO_DATE('" & v_txtdateto & "', 'MM/DD/YYYY'))) budget_hk_per_ha_shi , " & _
				  "							sum(BUDGET_TRANSACTION_ESTATE_API.GetQuantityUntilMonth(a_temp.fcba,'" & v_txtBudgetYear &"','" & v_txtBudgetVersion &"','" & v_txtModuleVersion & "',a_temp.fccode,TO_DATE('" & v_txtdateto & "', 'MM/DD/YYYY'))) " & _
				  "							/ sum(BUDGET_TRANSACTION_ESTATE_API.GetOutputUntilMonth(a_temp.fcba,'" & v_txtBudgetYear &"','" & v_txtBudgetVersion &"','" & v_txtModuleVersion & "',a_temp.fccode,TO_DATE('" & v_txtdateto & "', 'MM/DD/YYYY'))) budget_hk_per_ha_sbi " & _
				  "					 FROM JOB a_temp, " & _
				  "					 ( SELECT grouphead.fcba,grouphead.param_1,grouphead.param_2,groupname.fccode groupCode,groupname.param_1 orderdata, " & _
				  "							  groupname.fcname param_2_name " & _
				  "					   FROM PARAMETERDETAIL grouphead, " & _
				  "					   ( SELECT fcba, fccode, fcname, param_1 " & _
				  "						 FROM PARAMETERDETAIL " & _
				  "						 WHERE parcode = 'SR002' " & _
				  "						 AND parheadcode = 'GROUPJOB' " & _
				  "					   ) groupname " & _
				  "					   WHERE grouphead.parcode = 'SR002' " & _
				  "					   AND grouphead.parheadcode IN ('ACTIVITY_TBM', 'ACTIVITY_TM') " & _
				  "					   AND groupname.fcba = grouphead.fcba " & _
				  "					   AND groupname.fccode = grouphead.param_2 " & _
				  "					  ) b_temp " & _
				  "					  WHERE a_temp.fccode = b_temp.param_1 " & _
				  "					  AND a_temp.fcba = b_temp.fcba " & _
				  "	 				  group by a_temp.fcba,groupCode, orderdata, param_2_name " & _          
				  "					 ) " & _
				   "			     where ACCESSRIGHTSCOMPANY_API.IsAccessRights('" & Session("CodeName") & "',FCBA,'" & txtcompany & "') = 'TRUE' " & str1 & " " & _ 
				  "					 GROUP BY codedesc, orderdata, groupCode,  param_2_name )  " 


	
 	Set rsHandling = Server.CreateObject("ADODB.Recordset")
	'response.write strHandling
	
	rsHandling.Open strHandling, connstr, adOpenStatic 

 	' Third
    ' To Handle the Query of Perawatan Amount

	' strHandlingAmount =	" SELECT codedesc, ROWNUM theNumber, groupCode, param_2_name description ,  amount_hi_per_ha hi, amount_hi_per_ha_shi shi, amount_hi_per_ha_sbi sbi, budget_hi, budget_shi, budget_sbi FROM ( " & _
						' " SELECT   'SR003_07' codedesc ,  groupCode, param_2_name, SUM (amount_hi_per_ha) amount_hi_per_ha, " & _
						    ' " sum(amount_hi_per_ha_shi) amount_hi_per_ha_shi, sum(amount_hi_per_ha_sbi) amount_hi_per_ha_sbi, " & _
							' " sum(budget_data_hi)/sum(budget_output_hi) budget_hi, sum(budget_data_shi)/sum(budget_output_shi) budget_shi, sum(budget_data_sbi)/sum(budget_output_sbi) budget_sbi " & _
							' " FROM (SELECT   b_temp.groupCode, b_temp.param_2, b_temp.param_2_name, 0 sku, 0 khl, 0 total, " & _
										   ' " 0 ha_harian, 0 bo_harian, 0 total_bo_harian, 0 ha_per_hk, " & _
										   ' " 0 gli_ltr, 0 gli_ltr_ha, 0 par_ltr, 0 par_ltr_ha, 0 met_kg, " & _
										   ' " 0 met_kg_ha, 0 flo_kg, 0 flo_kg_ha, 0 amount_hi_per_ha, 0 amount_hi_per_ha_shi, 0 amount_hi_per_ha_sbi, " & _
										   ' " BUDGET_TRANSACTION_ESTATE_API.GetBudgetDay(nvl('" & v_txtfcba & "' ,'" & session("fcba") & "'), " & _
										   									 ' v_txtBudgetYear & " , " & _
																			 ' v_txtBudgetVersion & " , " & _
																			 ' v_txtModuleVersion & " , " & _
																			' " a_temp.fccode, " & _
																			' " TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY')) budget_data_hi, " & _
										   ' " BUDGET_TRANSACTION_ESTATE_API.GetOutputDay(nvl('" & v_txtfcba & "' ,'" & session("fcba") & "'), " & _
										   									 ' v_txtBudgetYear & " , " & _
																			 ' v_txtBudgetVersion & " , " & _
																			 ' v_txtModuleVersion & " , " & _
																			' " a_temp.fccode, " & _
																			' " TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY')) budget_output_hi , " & _
										   ' " BUDGET_TRANSACTION_ESTATE_API.GetBudgetMonth(nvl('" & v_txtfcba & "' ,'" & session("fcba") & "'), " & _
										   									 ' v_txtBudgetYear & " , " & _
																			 ' v_txtBudgetVersion & " , " & _
																			 ' v_txtModuleVersion & " , " & _
																			' " a_temp.fccode, " & _
																			' " TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY')) budget_data_shi, " & _
										   ' " BUDGET_TRANSACTION_ESTATE_API.GetOutputMonth(nvl('" & v_txtfcba & "' ,'" & session("fcba") & "'), " & _
										   									 ' v_txtBudgetYear & " , " v_txtBudgetVersion & " , " & _
																			 ' v_txtModuleVersion & " , " & _
																			' " a_temp.fccode, " & _
																			' " TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY')) budget_output_shi , " & _
										   ' " BUDGET_TRANSACTION_ESTATE_API.GetBudgetUntilMonth(nvl('" & v_txtfcba & "' ,'" & session("fcba") & "'), " & _
										   									 ' v_txtBudgetYear & " , " & _
																			 ' v_txtBudgetVersion & " , " & _
																			 ' v_txtModuleVersion & " , " & _
																			' " a_temp.fccode, " & _
																			' " TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY')) budget_data_sbi, " & _
										   ' " BUDGET_TRANSACTION_ESTATE_API.GetOutputUntilMonth(nvl('" & v_txtfcba & "' ,'" & session("fcba") & "'), " & _
										   									 ' v_txtBudgetYear & " , " & _
																			 ' v_txtBudgetVersion & " , " & _
																			 ' v_txtModuleVersion & " , " & _
																			' " a_temp.fccode, " & _
																			' " TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY')) budget_output_sbi " & _
									  ' " FROM JOB a_temp, " & _
										   ' " (SELECT grouphead.fcba, grouphead.param_1, " & _
												   ' " grouphead.param_2, groupname.fccode groupCode, groupname.fcname param_2_name " & _
											  ' " FROM PARAMETERDETAIL grouphead, " & _
												   ' " (SELECT fcba, fccode, fcname " & _
													  ' " FROM PARAMETERDETAIL " & _
													 ' " WHERE parcode = 'SR002' " & _
													   ' " AND parheadcode = 'GROUPJOB') groupname " & _
											 ' " WHERE grouphead.parcode = 'SR002' " & _
											   ' " AND grouphead.parheadcode IN " & _
																	  ' " ('ACTIVITY_TBM', 'ACTIVITY_TM') " & _
											   ' " AND groupname.fcba = grouphead.fcba " & _
											   ' " AND groupname.fccode = grouphead.param_2) b_temp " & _
									 ' " WHERE a_temp.fccode = b_temp.param_1 " & _
										   ' " AND a_temp.fcba = b_temp.fcba " & _
								  ' " UNION " & _
								  ' " SELECT   c.groupCode, c.param_2, c.param_2_name, SUM (a.sku) sku, SUM (a.khl) khl, " & _
										   ' " SUM (a.total) total, SUM (a.ha_harian) ha_harian, " & _
										   ' " SUM (a.bo_harian) bo_harian, " & _
										   ' " SUM (a.total_bo_harian) total_bo_harian, " & _
										   ' " DECODE (SUM (a.total), " & _
												   ' " 0, 0, " & _
												   ' " SUM (a.ha_harian) / SUM (a.total) " & _
												  ' " ) ha_per_hk, " & _
										   ' " SUM (a.gli_ltr) gli_ltr, " & _
										   ' " DECODE (SUM (a.total_bo_harian), " & _
												   ' " 0, 0, " & _
												   ' " SUM (a.gli_ltr) / SUM (a.total_bo_harian) " & _
												  ' " ) gli_ltr_ha, " & _
										   ' " SUM (a.par_ltr) par_ltr, " & _
										   ' " DECODE (SUM (a.total_bo_harian), " & _
												   ' " 0, 0, " & _
												   ' " SUM (a.par_ltr) / SUM (a.total_bo_harian) " & _
												  ' " ) par_ltr_ha, " & _
										   ' " SUM (a.met_kg) met_kg, " & _
										   ' " DECODE (SUM (a.total_bo_harian), " & _
												   ' " 0, 0, " & _
												   ' " SUM (a.met_kg) / SUM (a.total_bo_harian) " & _
												  ' " ) met_kg_ha, " & _
										   ' " SUM (flo_kg) flo_kg, " & _
										   ' " DECODE (SUM (a.total_bo_harian), " & _
												   ' " 0, 0, " & _
												   ' " SUM (flo_kg) / SUM (a.total_bo_harian) " & _
												  ' " ) flo_kg_ha, " & _
										   ' " DECODE (SUM (a.ha_harian), " & _
												   ' " 0, 0, " & _
												   ' " SUM (a.amount_hi) / SUM (a.ha_harian) " & _
												  ' " ) amount_hi_per_ha, 0 amount_hi_per_ha_shi, 0 amount_hi_per_ha_sbi,  " & _
										   ' " 0 budget_data_hi , 0 budget_output_hi, 0 budget_data_shi, 0 budget_output_shi, 0 budget_data_sbi, 0 budget_output_sbi " & _		  
									  ' " FROM view_attendance_tm_smry a, " & _
										   ' " JOB b, " & _
										   ' " (SELECT grouphead.fcba, grouphead.param_1, " & _
												   ' " grouphead.param_2, groupname.fccode groupCode, groupname.fcname param_2_name " & _
											  ' " FROM PARAMETERDETAIL grouphead, " & _
												   ' " (SELECT fcba, fccode, fcname " & _
													  ' " FROM PARAMETERDETAIL " & _
													 ' " WHERE parcode = 'SR002' " & _
													   ' " AND parheadcode = 'GROUPJOB') groupname " & _
											 ' " WHERE grouphead.parcode = 'SR002' " & _
											   ' " AND grouphead.parheadcode IN " & _
																	  ' " ('ACTIVITY_TBM', 'ACTIVITY_TM') " & _
											   ' " AND groupname.fcba = grouphead.fcba " & _
											   ' " AND groupname.fccode = grouphead.param_2) c " & _
									 ' " WHERE (a.fcba = b.fcba AND a.jobcode = b.fccode) " & _
									   ' " AND (a.fcba = c.fcba AND a.jobcode = c.param_1) " & _
									   ' " AND fddate BETWEEN TO_DATE ('" & v_txtdatefrom & "', 'MM/DD/YYYY') " & _
													  ' " AND TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY') " & _
								  ' " GROUP BY c.param_2, c.groupCode, c.param_2_name " & _
								  ' " UNION " & _
								  ' " SELECT   c.groupCode, c.param_2, c.param_2_name, SUM (a.sku) sku, SUM (a.khl) khl, " & _
										   ' " SUM (a.total) total, SUM (a.ha_harian) ha_harian, " & _
										   ' " SUM (a.bo_harian) bo_harian, " & _
										   ' " SUM (a.total_bo_harian) total_bo_harian, " & _
										   ' " DECODE (SUM (a.total), " & _
												   ' " 0, 0, " & _
												   ' " SUM (a.ha_harian) / SUM (a.total) " & _
												  ' " ) ha_per_hk, " & _
										   ' " SUM (a.gli_ltr) gli_ltr, " & _
										   ' " DECODE (SUM (a.total_bo_harian), " & _
												   ' " 0, 0, " & _
												   ' " SUM (a.gli_ltr) / SUM (a.total_bo_harian) " & _
												  ' " ) gli_ltr_ha, " & _
										   ' " SUM (a.par_ltr) par_ltr, " & _
										   ' " DECODE (SUM (a.total_bo_harian), " & _
												   ' " 0, 0, " & _
												   ' " SUM (a.par_ltr) / SUM (a.total_bo_harian) " & _
												  ' " ) par_ltr_ha, " & _
										   ' " SUM (a.met_kg) met_kg, " & _
										   ' " DECODE (SUM (a.total_bo_harian), " & _
												   ' " 0, 0, " & _
												   ' " SUM (a.met_kg) / SUM (a.total_bo_harian) " & _
												  ' " ) met_kg_ha, " & _
										   ' " SUM (flo_kg) flo_kg, " & _
										   ' " DECODE (SUM (a.total_bo_harian), " & _
												   ' " 0, 0, " & _
												   ' " SUM (flo_kg) / SUM (a.total_bo_harian) " & _
												  ' " ) flo_kg_ha, " & _
										    ' " 0 amount_hi_per_ha, " & _		  
										   ' " DECODE (SUM (a.ha_harian), " & _
												   ' " 0, 0, " & _
												   ' " SUM (a.amount_hi) / SUM (a.ha_harian) " & _
												  ' " ) amount_hi_per_ha_shi, " & _
											' " 0 amount_hi_per_ha_sbi, " & _	 
										   ' " 0 budget_data_hi , 0 budget_output_hi, 0 budget_data_shi, 0 budget_output_shi, 0 budget_data_sbi, 0 budget_output_sbi " & _		  											
									  ' " FROM view_attendance_tm_smry a, " & _
										   ' " JOB b, " & _
										   ' " (SELECT grouphead.fcba, grouphead.param_1, " & _
												   ' " grouphead.param_2, groupname.fccode groupCode, groupname.fcname param_2_name " & _
											  ' " FROM PARAMETERDETAIL grouphead, " & _
												   ' " (SELECT fcba, fccode, fcname " & _
													  ' " FROM PARAMETERDETAIL " & _
													 ' " WHERE parcode = 'SR002' " & _
													   ' " AND parheadcode = 'GROUPJOB') groupname " & _
											 ' " WHERE grouphead.parcode = 'SR002' " & _
											   ' " AND grouphead.parheadcode IN " & _
																	  ' " ('ACTIVITY_TBM', 'ACTIVITY_TM') " & _
											   ' " AND groupname.fcba = grouphead.fcba " & _
											   ' " AND groupname.fccode = grouphead.param_2) c " & _
									 ' " WHERE (a.fcba = b.fcba AND a.jobcode = b.fccode) " & _
									   ' " AND (a.fcba = c.fcba AND a.jobcode = c.param_1) " & _
									   ' " AND fddate <= TO_DATE ('" & v_txtdatefrom & "', 'MM/DD/YYYY') " & _
									   ' " AND to_char(fddate, 'MM' ) = '" & v_txtmonth & "' " & _
								       ' " AND to_char(fddate, 'YYYY') = '" & v_txtyear & "' " & _
								  ' " GROUP BY c.groupCode, c.param_2, c.param_2_name " & _
								  ' " UNION " & _
								  ' " SELECT   c.groupCode, c.param_2, c.param_2_name, SUM (a.sku) sku, SUM (a.khl) khl, " & _
										   ' " SUM (a.total) total, SUM (a.ha_harian) ha_harian, " & _
										   ' " SUM (a.bo_harian) bo_harian, " & _
										   ' " SUM (a.total_bo_harian) total_bo_harian, " & _
										   ' " DECODE (SUM (a.total), " & _
												   ' " 0, 0, " & _
												   ' " SUM (a.ha_harian) / SUM (a.total) " & _
												  ' " ) ha_per_hk, " & _
										   ' " SUM (a.gli_ltr) gli_ltr, " & _
										   ' " DECODE (SUM (a.total_bo_harian), " & _
												   ' " 0, 0, " & _
												   ' " SUM (a.gli_ltr) / SUM (a.total_bo_harian) " & _
												  ' " ) gli_ltr_ha, " & _
										   ' " SUM (a.par_ltr) par_ltr, " & _
										   ' " DECODE (SUM (a.total_bo_harian), " & _
												   ' " 0, 0, " & _
												   ' " SUM (a.par_ltr) / SUM (a.total_bo_harian) " & _
												  ' " ) par_ltr_ha, " & _
										   ' " SUM (a.met_kg) met_kg, " & _
										   ' " DECODE (SUM (a.total_bo_harian), " & _
												   ' " 0, 0, " & _
												   ' " SUM (a.met_kg) / SUM (a.total_bo_harian) " & _
												  ' " ) met_kg_ha, " & _
										   ' " SUM (flo_kg) flo_kg, " & _
										   ' " DECODE (SUM (a.total_bo_harian), " & _
												   ' " 0, 0, " & _
												   ' " SUM (flo_kg) / SUM (a.total_bo_harian) " & _
												  ' " ) flo_kg_ha, " & _
										    ' " 0 amount_hi_per_ha, " & _		  
										    ' " 0 amount_hi_per_ha_shi, " & _		  
										   ' " DECODE (SUM (a.ha_harian), " & _
												   ' " 0, 0, " & _
												   ' " SUM (a.amount_hi) / SUM (a.ha_harian) " & _
												  ' " ) amount_hi_per_ha_sbi ," & _
										   ' " 0 budget_data_hi , 0 budget_output_hi, 0 budget_data_shi, 0 budget_output_shi, 0 budget_data_sbi, 0 budget_output_sbi " & _		  
									  ' " FROM view_attendance_tm_smry a, " & _
										   ' " JOB b, " & _
										   ' " (SELECT grouphead.fcba, grouphead.param_1, " & _
												   ' " grouphead.param_2, groupName.fccode groupCode, groupname.fcname param_2_name " & _
											  ' " FROM PARAMETERDETAIL grouphead, " & _
												   ' " (SELECT fcba, fccode, fcname " & _
													  ' " FROM PARAMETERDETAIL " & _
													 ' " WHERE parcode = 'SR002' " & _
													   ' " AND parheadcode = 'GROUPJOB') groupname " & _
											 ' " WHERE grouphead.parcode = 'SR002' " & _
											   ' " AND grouphead.parheadcode IN " & _
																	  ' " ('ACTIVITY_TBM', 'ACTIVITY_TM') " & _
											   ' " AND groupname.fcba = grouphead.fcba " & _
											   ' " AND groupname.fccode = grouphead.param_2) c " & _
									 ' " WHERE (a.fcba = b.fcba AND a.jobcode = b.fccode) " & _
									   ' " AND (a.fcba = c.fcba AND a.jobcode = c.param_1) " & _
									   ' " AND fddate <= TO_DATE ('" & v_txtdatefrom & "', 'MM/DD/YYYY') " & _
								       ' " AND to_char(fddate, 'YYYY') = '" & v_txtyear & "' " & _									   
								  ' " GROUP BY c.param_2,c.groupCode, c.param_2_name " & _
								  ' " ) " & _
						' " GROUP BY param_2, groupCode, param_2_name " & _
						' " ORDER BY param_2 " & _
						' " ) " 
						
     strHandlingAmount = " SELECT codedesc, ROWNUM theNumber, groupCode, param_2_name description ,  hi, shi,  sbi, budget_hi, budget_shi, budget_sbi " & _
						 " FROM ( SELECT codedesc, groupCode, param_2_name , sum(amount_hi_per_ha) hi, sum(amount_hi_per_ha_shi) shi, sum(amount_hi_per_ha_sbi) sbi, sum(budget_hi) budget_hi, sum(budget_shi) budget_shi, sum(budget_sbi) budget_sbi " & _
						 "	 FROM ( SELECT fcba,'SR003_07' codedesc,'1' nomer, groupCode, param_2_name, SUM (amount_hi_per_ha) amount_hi_per_ha, " & _
						 "				   sum(amount_hi_per_ha_shi) amount_hi_per_ha_shi, sum(amount_hi_per_ha_sbi) amount_hi_per_ha_sbi, " & _
						 "				   sum(budget_data_hi)/sum(budget_output_hi) budget_hi, sum(budget_data_shi)/sum(budget_output_shi) budget_shi, sum(budget_data_sbi)/sum(budget_output_sbi) budget_sbi " & _
						 "			FROM ( SELECT   a_temp.fcba,b_temp.groupCode, b_temp.param_2, b_temp.param_2_name, 0 sku, 0 khl, 0 total, " & _
						 "							0 ha_harian, 0 bo_harian, 0 total_bo_harian, 0 ha_per_hk, " & _
						 "							0 gli_ltr, 0 gli_ltr_ha, 0 par_ltr, 0 par_ltr_ha, 0 met_kg, " & _
						 "							0 met_kg_ha, 0 flo_kg, 0 flo_kg_ha, 0 amount_hi_per_ha, 0 amount_hi_per_ha_shi, 0 amount_hi_per_ha_sbi, " & _
						 "							BUDGET_TRANSACTION_ESTATE_API.GetBudgetDay(a_temp.fcba,'" & v_txtBudgetYear &"','" & v_txtBudgetVersion &"','" & v_txtModuleVersion & "' ,a_temp.fccode,TO_DATE('" & v_txtdateto & "', 'MM/DD/YYYY')) budget_data_hi, " & _
						 "							BUDGET_TRANSACTION_ESTATE_API.GetOutputDay(a_temp.fcba,'" & v_txtBudgetYear &"','" & v_txtBudgetVersion &"','" & v_txtModuleVersion & "' ,a_temp.fccode,TO_DATE('" & v_txtdateto & "', 'MM/DD/YYYY')) budget_output_hi , " & _
						 "							BUDGET_TRANSACTION_ESTATE_API.GetBudgetMonth(a_temp.fcba,'" & v_txtBudgetYear &"','" & v_txtBudgetVersion &"','" & v_txtModuleVersion & "' ,a_temp.fccode,TO_DATE('" & v_txtdateto & "', 'MM/DD/YYYY')) budget_data_shi, " & _
						 "							BUDGET_TRANSACTION_ESTATE_API.GetOutputMonth(a_temp.fcba,'" & v_txtBudgetYear &"','" & v_txtBudgetVersion &"','" & v_txtModuleVersion & "' ,a_temp.fccode,TO_DATE('" & v_txtdateto & "', 'MM/DD/YYYY')) budget_output_shi , " & _
						 "							BUDGET_TRANSACTION_ESTATE_API.GetBudgetUntilMonth(a_temp.fcba,'" & v_txtBudgetYear &"','" & v_txtBudgetVersion &"','" & v_txtModuleVersion & "' ,a_temp.fccode,TO_DATE('" & v_txtdateto & "', 'MM/DD/YYYY')) budget_data_sbi, " & _
						 "							BUDGET_TRANSACTION_ESTATE_API.GetOutputUntilMonth(a_temp.fcba,'" & v_txtBudgetYear &"','" & v_txtBudgetVersion &"','" & v_txtModuleVersion & "',a_temp.fccode,TO_DATE('" & v_txtdateto & "', 'MM/DD/YYYY')) budget_output_sbi " & _
						 "				   FROM JOB a_temp, " & _
						 "				   ( SELECT grouphead.fcba, grouphead.param_1,grouphead.param_2, groupname.fccode groupCode, groupname.fcname param_2_name " & _
						 "					 FROM PARAMETERDETAIL grouphead, " & _
						 "					 ( SELECT fcba, fccode, fcname " & _
						 "					   FROM PARAMETERDETAIL " & _
						 "					   WHERE parcode = 'SR002' " & _
						 "					   AND parheadcode = 'GROUPJOB' " & _
						 "					 ) groupname " & _
						 "					 WHERE grouphead.parcode = 'SR002' " & _
						 "					 AND grouphead.parheadcode IN ('ACTIVITY_TBM', 'ACTIVITY_TM') " & _
						 "					 AND groupname.fcba = grouphead.fcba " & _
						 "					 AND groupname.fccode = grouphead.param_2 " & _
						 "				   ) b_temp " & _
						 "				   WHERE a_temp.fccode = b_temp.param_1 " & _
						 "				   AND a_temp.fcba = b_temp.fcba " & _
						 "				   UNION " & _
						 "				   SELECT c.fcba,c.groupCode, c.param_2, c.param_2_name, SUM(a.sku) sku, SUM(a.khl) khl, " & _
						 "						  SUM(a.total) total, SUM(a.ha_harian) ha_harian,SUM(a.bo_harian) bo_harian, " & _
						 "						  SUM(a.total_bo_harian) total_bo_harian,DECODE(SUM(a.total),0, 0,SUM (a.ha_harian) / SUM (a.total)) ha_per_hk, " & _
						 "						  SUM(a.gli_ltr) gli_ltr,DECODE(SUM(a.total_bo_harian), 0, 0,SUM(a.gli_ltr) / SUM (a.total_bo_harian)) gli_ltr_ha, " & _
						 "						  SUM(a.par_ltr) par_ltr,DECODE(SUM(a.total_bo_harian),0, 0,SUM (a.par_ltr) / SUM (a.total_bo_harian)) par_ltr_ha, " & _
						 "						  SUM(a.met_kg) met_kg, DECODE(SUM(a.total_bo_harian),0, 0,SUM (a.met_kg) / SUM (a.total_bo_harian)) met_kg_ha, " & _
						 "						  SUM(flo_kg) flo_kg,DECODE(SUM(a.total_bo_harian), 0, 0,SUM (flo_kg) / SUM (a.total_bo_harian)) flo_kg_ha, " & _
						 "						  DECODE(SUM(a.ha_harian),0, 0,SUM (a.amount_hi) / SUM (a.ha_harian)) amount_hi_per_ha, 0 amount_hi_per_ha_shi, 0 amount_hi_per_ha_sbi,  " & _
						 "						  0 budget_data_hi , 0 budget_output_hi, 0 budget_data_shi, 0 budget_output_shi, 0 budget_data_sbi, 0 budget_output_sbi   " & _    
						 "				   FROM view_attendance_tm_smry a,JOB b, " & _
						 "				   ( SELECT grouphead.fcba, grouphead.param_1,grouphead.param_2, groupname.fccode groupCode, groupname.fcname param_2_name " & _
						 "					 FROM PARAMETERDETAIL grouphead, " & _
						 "					 ( SELECT fcba, fccode, fcname " & _
						 "					   FROM PARAMETERDETAIL " & _
						 "					   WHERE parcode = 'SR002' " & _
						 "					   AND parheadcode = 'GROUPJOB' " & _
						 "					 ) groupname " & _
						 "					 WHERE grouphead.parcode = 'SR002' " & _
						 "					 AND grouphead.parheadcode IN ('ACTIVITY_TBM', 'ACTIVITY_TM') " & _
						 "					 AND groupname.fcba = grouphead.fcba " & _
						 "					 AND groupname.fccode = grouphead.param_2 " & _
						 "				   ) c " & _
						 "				   WHERE (a.fcba = b.fcba AND a.jobcode = b.fccode) " & _
						 "				   AND (a.fcba = c.fcba AND a.jobcode = c.param_1) " & _
						 "				   AND fddate BETWEEN TO_DATE ('" & v_txtdatefrom & "', 'MM/DD/YYYY') " & _
						 "				   AND TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY') " & _
						 "				   GROUP BY c.fcba,c.param_2, c.groupCode, c.param_2_name " & _
						 "				   UNION " & _
						 "				   SELECT   c.fcba,c.groupCode, c.param_2, c.param_2_name, SUM (a.sku) sku, SUM (a.khl) khl, " & _
						 "							SUM (a.total) total, SUM (a.ha_harian) ha_harian,SUM (a.bo_harian) bo_harian, " & _
						 "							SUM (a.total_bo_harian) total_bo_harian,DECODE (SUM (a.total),0, 0,SUM (a.ha_harian) / SUM (a.total)) ha_per_hk, " & _
						 "							SUM (a.gli_ltr) gli_ltr,DECODE (SUM (a.total_bo_harian),0, 0,SUM (a.gli_ltr) / SUM (a.total_bo_harian)) gli_ltr_ha, " & _
						 "							SUM (a.par_ltr) par_ltr,DECODE (SUM (a.total_bo_harian),0, 0,SUM (a.par_ltr) / SUM (a.total_bo_harian)) par_ltr_ha, " & _
						 "							SUM (a.met_kg) met_kg,DECODE (SUM (a.total_bo_harian), 0, 0,SUM (a.met_kg) / SUM (a.total_bo_harian)) met_kg_ha, " & _
						 "							SUM (flo_kg) flo_kg,DECODE (SUM (a.total_bo_harian),0, 0,SUM (flo_kg) / SUM (a.total_bo_harian)) flo_kg_ha, " & _
						 "							0 amount_hi_per_ha,DECODE (SUM (a.ha_harian),0, 0,SUM (a.amount_hi) / SUM (a.ha_harian)) amount_hi_per_ha_shi, " & _
						 "							0 amount_hi_per_ha_sbi,0 budget_data_hi , 0 budget_output_hi, 0 budget_data_shi, 0 budget_output_shi, 0 budget_data_sbi, 0 budget_output_sbi  " & _                           
						 "				   FROM view_attendance_tm_smry a, JOB b, " & _
						 "				   ( SELECT grouphead.fcba, grouphead.param_1,grouphead.param_2, groupname.fccode groupCode, groupname.fcname param_2_name " & _
						 "					 FROM PARAMETERDETAIL grouphead, " & _
						 "					 ( SELECT fcba, fccode, fcname " & _
						 "					   FROM PARAMETERDETAIL " & _
						 "					   WHERE parcode = 'SR002' " & _
						 "					   AND parheadcode = 'GROUPJOB' " & _
						 "					 ) groupname " & _
						 "					 WHERE grouphead.parcode = 'SR002' " & _
						 "					 AND grouphead.parheadcode IN ('ACTIVITY_TBM', 'ACTIVITY_TM') " & _
						 "					 AND groupname.fcba = grouphead.fcba " & _
						 "					 AND groupname.fccode = grouphead.param_2 " & _
						 "				   ) c  " & _
						 "				   WHERE (a.fcba = b.fcba AND a.jobcode = b.fccode) " & _
						 "				   AND (a.fcba = c.fcba AND a.jobcode = c.param_1) " & _
						 "				   AND fddate <= TO_DATE ('" & v_txtdatefrom & "', 'MM/DD/YYYY') " & _
						 "				   AND to_char(fddate, 'MM' ) = '" & v_txtmonth & "' " & _
						 "				   AND to_char(fddate, 'YYYY') = '" & v_txtyear & "' " & _
						 "				   GROUP BY c.fcba,c.groupCode, c.param_2, c.param_2_name " & _
						 "				   UNION " & _
						 "				   SELECT   c.fcba,c.groupCode, c.param_2, c.param_2_name, SUM (a.sku) sku, SUM (a.khl) khl, " & _
						 "							SUM (a.total) total, SUM (a.ha_harian) ha_harian, " & _
						 "							SUM (a.bo_harian) bo_harian, " & _
						 "							SUM (a.total_bo_harian) total_bo_harian, " & _
						 "							DECODE (SUM (a.total),0, 0,SUM (a.ha_harian) / SUM (a.total)) ha_per_hk, " & _
						 "							SUM (a.gli_ltr) gli_ltr,DECODE(SUM (a.total_bo_harian),0, 0,SUM (a.gli_ltr) / SUM (a.total_bo_harian)) gli_ltr_ha, " & _
						 "							SUM (a.par_ltr) par_ltr,DECODE(SUM (a.total_bo_harian),0, 0,SUM (a.par_ltr) / SUM (a.total_bo_harian)) par_ltr_ha, " & _
						 "							SUM (a.met_kg) met_kg,DECODE(SUM (a.total_bo_harian),0, 0,SUM (a.met_kg) / SUM(a.total_bo_harian)) met_kg_ha, " & _
						 "							SUM (flo_kg) flo_kg,DECODE(SUM(a.total_bo_harian),0, 0,SUM (flo_kg) / SUM (a.total_bo_harian)) flo_kg_ha, " & _
						 "							0 amount_hi_per_ha,0 amount_hi_per_ha_shi,DECODE (SUM (a.ha_harian),0, 0,SUM (a.amount_hi) / SUM (a.ha_harian)) amount_hi_per_ha_sbi , " & _
						 "							0 budget_data_hi , 0 budget_output_hi, 0 budget_data_shi, 0 budget_output_shi, 0 budget_data_sbi, 0 budget_output_sbi  " & _     
						 "				   FROM view_attendance_tm_smry a,JOB b, " & _
						 "				   ( SELECT grouphead.fcba, grouphead.param_1,grouphead.param_2, groupName.fccode groupCode, groupname.fcname param_2_name " & _
						 "					 FROM PARAMETERDETAIL grouphead, " & _
						 "					 ( SELECT fcba, fccode, fcname " & _
						 "					   FROM PARAMETERDETAIL " & _
						 "					   WHERE parcode = 'SR002' " & _
						 "					   AND parheadcode = 'GROUPJOB' " & _
						 "					 ) groupname " & _
						 "					 WHERE grouphead.parcode = 'SR002' " & _
						 "					 AND grouphead.parheadcode IN ('ACTIVITY_TBM', 'ACTIVITY_TM') " & _
						 "					 AND groupname.fcba = grouphead.fcba " & _
						 "					 AND groupname.fccode = grouphead.param_2 " & _
						 "				   ) c " & _
						 "				   WHERE (a.fcba = b.fcba AND a.jobcode = b.fccode) " & _
						 "				   AND (a.fcba = c.fcba AND a.jobcode = c.param_1) " & _
						 "				   AND fddate <= TO_DATE ('" & v_txtdatefrom & "', 'MM/DD/YYYY') " & _
						 "				   AND to_char(fddate, 'YYYY') = '" & v_txtyear & "'  " & _                    
						 "				   GROUP BY c.fcba,c.param_2,c.groupCode, c.param_2_name " & _
						 "				) " & _
						  "			     where ACCESSRIGHTSCOMPANY_API.IsAccessRights('" & Session("CodeName") & "',FCBA,'" & txtcompany & "') = 'TRUE' " & str1 & " " & _ 
						 "				 GROUP BY fcba,param_2, groupCode, param_2_name " & _
						 "				 ORDER BY param_2 " & _
						 "		  )  " & _                  
						 "		  group by nomer,codedesc, groupCode, param_2_name " & _
						 "	 ) "

 	
 	Set rsHandlingAmount = Server.CreateObject("ADODB.Recordset")
	'response.write strHandlingAmount
	
	rsHandlingAmount.Open strHandlingAmount, connstr, adOpenStatic 
						

	
	' Fourth
    ' To Handle the Query of PUPUK/Fertilizer
	

	' strFertilizer = " select codedesc, ROWNUM thenumber, stockGenCode, description, unitofmeasurement, hi, shi, sbi, budget_hi, budget_shi, budget_sbi FROM ( " & _
						  ' " SELECT 'SR003_08' codedesc, " & _
								 ' " stockGenCode, " & _
								 ' " stockGenName description, " & _
								 ' " orderdata, " & _
								 ' " UnitOfMeasurement, " & _
								 ' " SUM (hi) hi, " & _
								 ' " SUM (shi) shi, " & _
								 ' " SUM (sbi) sbi, " & _
								 ' " SUM (budget_hi) budget_hi, " & _
								 ' " SUM (budget_shi) budget_shi, " & _
								 ' " SUM (budget_sbi) budget_sbi " & _
							' " FROM (SELECT FCCODE stockGenCode, " & _
										 ' " FCNAME stockGenName, " & _
										 ' " param_2 orderdata, " & _
										 ' " param_3 unitofmeasurement, " & _
										 ' " 0 hi, " & _
										 ' " 0 shi, " & _
										 ' " 0 sbi, " & _
										 ' " nvl(BUDGET_TRANSACTION_ESTATE_API.GetQuantityDaySource(nvl('" & v_txtfcba & "' ,'" & session("fcba") & "'), " & _
										   									 ' v_txtBudgetYear & " , " & _
																			 ' v_txtBudgetVersion & " , " & _
																			 ' v_txtModuleVersion & " , " & _
																			' " 'ST',fccode, " & _
																			' " TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY')),0) budget_hi , " & _
										 ' " nvl(BUDGET_TRANSACTION_ESTATE_API.GetQuantityMonthSource(nvl('" & v_txtfcba & "' ,'" & session("fcba") & "'), " & _
										   									 ' v_txtBudgetYear & " , " & _
																			 ' v_txtBudgetVersion & " , " & _
																			 ' v_txtModuleVersion & " , " & _
																			' " 'ST',fccode, " & _
																			' " TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY')),0) budget_shi , " & _
										 ' " nvl(BUDGET_TRANSACTION_ESTATE_API.GetQuantityUntilMonthSource(nvl('" & v_txtfcba & "' ,'" & session("fcba") & "'), " & _
										   									 ' v_txtBudgetYear & " , " & _
																			 ' v_txtBudgetVersion & " , " & _
																			 ' v_txtModuleVersion & " , " & _
																			' " 'ST',fccode, " & _
																			' " TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY')),0) budget_sbi " & _
									' " FROM PARAMETER " & _
								   ' " WHERE PARCODE = 'SR003_PUPDET' AND PARAM_1 = 'DEFAULT' " & _
								  ' " UNION " & _
									' " SELECT c.stockGenCode, " & _
										   ' " c.stockGenName, " & _
										   ' " c.orderdata, " & _
										   ' " c.UnitOfMeasurement, " & _
										   ' " SUM (a.quantity) hi, " & _
										   ' " 0 shi, " & _
										   ' " 0 sbi, " & _
										   ' " 0 budget_hi, " & _
										   ' " 0 budget_shi, " & _
										   ' " 0 budget_sbi " & _
									  ' " FROM ISSUEDETAILS a, " & _
										   ' " ISSUE b, " & _
										   ' " (SELECT a.fcba, " & _
												   ' " a.FCCODE stockGenCode, " & _
												   ' " a.FCNAME StockGenName, " & _
												   ' " a.param_2 orderdata, " & _
												   ' " a.param_3 unitofmeasurement, " & _
												   ' " b.param_1 stockcode " & _
											  ' " FROM PARAMETER a, parameterdetail b " & _
											 ' " WHERE     a.fccode = b.parheadcode " & _
												   ' " AND a.parcode = b.parcode " & _
												   ' " AND a.fcba = b.fcba " & _
												   ' " AND a.PARCODE = 'SR003_PUPDET' " & _
												   ' " AND a.PARAM_1 = 'DEFAULT') c " & _
									 ' " WHERE     a.sivcode = b.sivcode " & _
										   ' " AND a.stockcode = c.stockcode " & _
										   ' " AND b.sivdate = TO_DATE ('" & v_txtdatefrom & "', 'MM/DD/YYYY') " & _
										   ' " AND a.fcba = b.fcba " & _
										   ' " AND a.fcba = c.fcba " & _
								  ' " GROUP BY c.stockGenCode, " & _
										   ' " c.stockGenName, " & _
										   ' " c.orderdata, " & _
										   ' " c.UnitOfMeasurement " & _
								  ' " UNION " & _
									' " SELECT c.stockGenCode, " & _
										   ' " c.stockGenName, " & _
										   ' " c.orderdata, " & _
										   ' " c.UnitOfMeasurement, " & _
										   ' " 0 hi, " & _
										   ' " SUM (a.quantity) shi, " & _
										   ' " 0 sbi, " & _
										   ' " 0 budget_hi, " & _
										   ' " 0 budget_shi, " & _
										   ' " 0 budget_sbi " & _
									  ' " FROM ISSUEDETAILS a, " & _
										   ' " ISSUE b, " & _
										   ' " (SELECT a.fcba, " & _
												   ' " a.FCCODE stockGenCode, " & _
												   ' " a.FCNAME StockGenName, " & _
												   ' " a.param_2 orderdata, " & _
												   ' " a.param_3 unitofmeasurement, " & _
												   ' " b.param_1 stockcode " & _
											  ' " FROM PARAMETER a, parameterdetail b " & _
											 ' " WHERE     a.fccode = b.parheadcode " & _
												   ' " AND a.parcode = b.parcode " & _
												   ' " AND a.fcba = b.fcba " & _
												   ' " AND a.PARCODE = 'SR003_PUPDET' " & _
												   ' " AND a.PARAM_1 = 'DEFAULT') c " & _
									 ' " WHERE     a.sivcode = b.sivcode " & _
										   ' " AND a.stockcode = c.stockcode " & _
										   ' " AND b.sivdate <= TO_DATE ('" & v_txtdatefrom & "', 'MM/DD/YYYY') " & _
										   ' " AND to_char(b.sivdate, 'YYYY') = '" & v_txtYear & "'  " & _
										   ' " AND to_char(b.sivdate, 'MM') = '" & v_txtMonth & "' " & _
										   ' " AND a.fcba = b.fcba " & _
										   ' " AND a.fcba = c.fcba " & _
								  ' " GROUP BY c.stockGenCode, " & _
										   ' " c.stockGenName, " & _
										   ' " c.orderdata, " & _
										   ' " c.UnitOfMeasurement " & _
											' " UNION " & _
									' " SELECT c.stockGenCode, " & _
										   ' " c.stockGenName, " & _
										   ' " c.orderdata, " & _
										   ' " c.UnitOfMeasurement, " & _
										   ' " 0 hi, " & _
										   ' " 0 shi, " & _
										   ' " SUM (a.quantity) sbi, " & _
										   ' " 0 budget_hi, " & _
										   ' " 0 budget_shi, " & _
										   ' " 0 budget_sbi " & _
									  ' " FROM ISSUEDETAILS a, " & _
										   ' " ISSUE b, " & _
										   ' " (SELECT a.fcba, " & _
												   ' " a.FCCODE stockGenCode, " & _
												   ' " a.FCNAME StockGenName, " & _
												   ' " a.param_2 orderdata, " & _
												   ' " a.param_3 unitofmeasurement, " & _
												   ' " b.param_1 stockcode " & _
											  ' " FROM PARAMETER a, parameterdetail b " & _
											 ' " WHERE     a.fccode = b.parheadcode " & _
												   ' " AND a.parcode = b.parcode " & _
												   ' " AND a.fcba = b.fcba " & _
												   ' " AND a.PARCODE = 'SR003_PUPDET' " & _
												   ' " AND a.PARAM_1 = 'DEFAULT') c " & _
									 ' " WHERE     a.sivcode = b.sivcode " & _
										   ' " AND a.stockcode = c.stockcode " & _
										   ' " AND b.sivdate <= TO_DATE ('" & v_txtdatefrom & "', 'MM/DD/YYYY') " & _
										   ' " AND to_char(b.sivdate, 'YYYY') = '" & v_txtYear & "'  " & _
										   ' " AND a.fcba = b.fcba " & _
										   ' " AND a.fcba = c.fcba " & _
								  ' " GROUP BY c.stockGenCode, " & _
										   ' " c.stockGenName, " & _
										   ' " c.orderdata, " & _
										   ' " c.UnitOfMeasurement    " & _
										' " ) " & _
						' " GROUP BY stockGenCode, " & _
								 ' " stockGenName, " & _
								 ' " orderdata, " & _
								 ' " UnitOfMeasurement " & _
						' " ORDER BY orderdata " & _
						' " ) "	
	
	strFertilizer =  "  select codedesc, ROWNUM thenumber, stockGenCode, description, unitofmeasurement, hi, shi, sbi, budget_hi, budget_shi, budget_sbi " & _
					 "	FROM ( SELECT codedesc,stockGenCode,description,orderdata,UnitOfMeasurement,SUM (hi) hi,SUM (shi) shi, " & _
					 "				   SUM(sbi) sbi, SUM(budget_hi) budget_hi,SUM(budget_shi) budget_shi,SUM(budget_sbi) budget_sbi " & _
					 "	 FROM ( SELECT fcba,'1' nomer, 'SR003_08' codedesc,stockGenCode,stockGenName description,orderdata,UnitOfMeasurement,SUM (hi) hi,SUM (shi) shi, " & _
					 "				   SUM(sbi) sbi, SUM(budget_hi) budget_hi,SUM(budget_shi) budget_shi,SUM(budget_sbi) budget_sbi " & _
					 "			FROM ( SELECT FCBA,FCCODE stockGenCode,FCNAME stockGenName,param_2 orderdata,param_3 unitofmeasurement,0 hi,0 shi,0 sbi, " & _
					 "						  nvl(BUDGET_TRANSACTION_ESTATE_API.GetQuantityDaySource(fcba,'" & v_txtBudgetYear &"' ,'1' ,'1' ,'ST',fccode,TO_DATE('" & v_txtdateto & "', 'MM/DD/YYYY')),0) budget_hi , " & _
					 "						  nvl(BUDGET_TRANSACTION_ESTATE_API.GetQuantityMonthSource(fcba,'" & v_txtBudgetYear &"' ,'1' ,'1' ,'ST',fccode,TO_DATE('" & v_txtdateto & "', 'MM/DD/YYYY')),0) budget_shi , " & _
					 "						  nvl(BUDGET_TRANSACTION_ESTATE_API.GetQuantityUntilMonthSource(fcba,'" & v_txtBudgetYear &"' ,'1' ,'1' ,'ST',fccode,TO_DATE('" & v_txtdateto & "', 'MM/DD/YYYY')),0) budget_sbi " & _
					 "				   FROM PARAMETER " & _
					 "				   WHERE PARCODE = 'SR003_PUPDET' AND PARAM_1 = 'DEFAULT' " & _
					 "				   UNION " & _
					 "				   SELECT c.fcba ,c.stockGenCode,c.stockGenName,c.orderdata,c.UnitOfMeasurement,SUM(a.quantity) hi,0 shi,0 sbi, " & _
					 "						  0 budget_hi,0 budget_shi,0 budget_sbi " & _
					 "				   FROM ISSUEDETAILS a,ISSUE b, " & _
					 "				   ( SELECT a.fcba,a.FCCODE stockGenCode,a.FCNAME StockGenName,a.param_2 orderdata,a.param_3 unitofmeasurement,b.param_1 stockcode " & _
					 "					 FROM PARAMETER a, parameterdetail b " & _
					 "					 WHERE a.fccode = b.parheadcode " & _
					 "					 AND a.parcode = b.parcode " & _
					 "					 AND a.fcba = b.fcba " & _
					 "					 AND a.PARCODE = 'SR003_PUPDET' " & _
					 "					 AND a.PARAM_1 = 'DEFAULT' " & _
					 "				   ) c  " & _
					 "				   WHERE a.sivcode = b.sivcode " & _
					 "				   AND a.stockcode = c.stockcode " & _
					 "				   AND b.sivdate = TO_DATE ('" & v_txtdatefrom & "', 'MM/DD/YYYY') " & _
					 "				   AND a.fcba = b.fcba " & _
					 "				   AND a.fcba = c.fcba " & _
					 "				   GROUP BY c.fcba ,c.stockGenCode,c.stockGenName,c.orderdata,c.UnitOfMeasurement " & _
					 "				   UNION " & _
					 "				   SELECT c.fcba ,c.stockGenCode,c.stockGenName,c.orderdata,c.UnitOfMeasurement,0 hi,SUM (a.quantity) shi, " & _
					 "						  0 sbi,0 budget_hi,0 budget_shi,0 budget_sbi " & _
					 "				   FROM ISSUEDETAILS a,ISSUE b, " & _
					 "				   ( SELECT a.fcba,a.FCCODE stockGenCode, a.FCNAME StockGenName,a.param_2 orderdata,a.param_3 unitofmeasurement,b.param_1 stockcode " & _
					 "					 FROM PARAMETER a, parameterdetail b " & _
					 "					 WHERE a.fccode = b.parheadcode " & _
					 "					 AND a.parcode = b.parcode " & _
					 "					 AND a.fcba = b.fcba " & _
					 "					 AND a.PARCODE = 'SR003_PUPDET' " & _
					 "					 AND a.PARAM_1 = 'DEFAULT' " & _
					 "					) c " & _
					 "					WHERE a.sivcode = b.sivcode " & _
					 "					AND a.stockcode = c.stockcode " & _
					 "					AND b.sivdate <= TO_DATE ('" & v_txtdatefrom & "', 'MM/DD/YYYY') " & _
					 "					AND to_char(b.sivdate, 'YYYY') = '" & v_txtyear & "'  " & _
					 "					AND to_char(b.sivdate, 'MM') = '" & v_txtmonth & "' " & _
					 "					AND a.fcba = b.fcba " & _
					 "					AND a.fcba = c.fcba " & _
					 "					GROUP BY c.fcba, c.stockGenCode,c.stockGenName,c.orderdata,c.UnitOfMeasurement " & _
					 "					UNION " & _
					 "					SELECT c.fcba, c.stockGenCode,c.stockGenName,c.orderdata,c.UnitOfMeasurement, 0 hi,0 shi,SUM (a.quantity) sbi, " & _
					 "						   0 budget_hi,0 budget_shi,0 budget_sbi " & _
					 "					FROM ISSUEDETAILS a,ISSUE b, " & _
					 "					( SELECT a.fcba,a.FCCODE stockGenCode,a.FCNAME StockGenName,a.param_2 orderdata,a.param_3 unitofmeasurement,b.param_1 stockcode " & _
					 "					  FROM PARAMETER a, parameterdetail b " & _
					 "					  WHERE a.fccode = b.parheadcode " & _
					 "					  AND a.parcode = b.parcode " & _
					 "					  AND a.fcba = b.fcba " & _
					 "					  AND a.PARCODE = 'SR003_PUPDET' " & _
					 "					  AND a.PARAM_1 = 'DEFAULT' " & _
					 "					) c " & _
					 "					WHERE a.sivcode = b.sivcode " & _
					 "					AND a.stockcode = c.stockcode " & _
					 "					AND b.sivdate <= TO_DATE ('" & v_txtdatefrom & "', 'MM/DD/YYYY') " & _
					 "					AND to_char(b.sivdate, 'YYYY') = '" & v_txtyear & "'  " & _
					 "					AND a.fcba = b.fcba " & _
					 "					AND a.fcba = c.fcba " & _
					 "					GROUP BY c.fcba, c.stockGenCode,c.stockGenName,c.orderdata,c.UnitOfMeasurement   " & _ 
					 "				  ) " & _
					 "			     where ACCESSRIGHTSCOMPANY_API.IsAccessRights('" & Session("CodeName") & "',FCBA,'" & txtcompany & "') = 'TRUE' " & str1 & " " & _ 
					 "				 GROUP BY fcba,stockGenCode,stockGenName,orderdata,UnitOfMeasurement " & _
					 "				 ORDER BY orderdata " & _
					 "		)  " & _
					 "	   group by nomer,codedesc,stockGenCode,description,orderdata,UnitOfMeasurement " & _
					 "	 ) "
	
 	Set rsFertilizer = Server.CreateObject("ADODB.Recordset")
	'response.write strFertilizer
	
	rsFertilizer.Open strFertilizer, connstr, adOpenStatic 
	

    ' Fifth
    ' To Handle the Query of Stock Ending Balance

								
	' strEndingBalance = " select codedesc, ROWNUM thenumber, description, unitofmeasurement, hi, '' shi, '' sbi FROM ( " & _
						  ' " SELECT 'SR003_09' codedesc, " & _
								 ' " stockGenCode, " & _
								 ' " stockGenName description, " & _
								 ' " orderdata, " & _
								 ' " UnitOfMeasurement, " & _
								 ' " SUM (hi) hi, " & _
								 ' " SUM (shi) shi, " & _
								 ' " SUM (sbi) sbi " & _
							' " FROM (SELECT FCCODE stockGenCode, " & _
										 ' " FCNAME stockGenName, " & _
										 ' " param_2 orderdata, " & _
										 ' " param_3 unitofmeasurement, " & _
										 ' " 0 hi, " & _
										 ' " 0 shi, " & _
										 ' " 0 sbi " & _
									' " FROM PARAMETER " & _
								   ' " WHERE PARCODE = 'SR003_STOCKSALDO' AND PARAM_1 = 'DEFAULT' " & _
								  ' " UNION " & _
									' " SELECT c.stockGenCode, " & _
										   ' " c.stockGenName, " & _
										   ' " c.orderdata, " & _
										   ' " c.UnitOfMeasurement, " & _
										   ' " SUM (a.quantity) hi, " & _
										   ' " 0 shi, " & _
										   ' " 0 sbi " & _
									  ' " FROM STOCKTRANSHIST a, " & _
										   ' " (SELECT a.fcba, " & _
												   ' " a.FCCODE stockGenCode, " & _
												   ' " a.FCNAME StockGenName, " & _
												   ' " a.param_2 orderdata, " & _
												   ' " a.param_3 unitofmeasurement, " & _
												   ' " b.param_1 stockcode " & _
											  ' " FROM PARAMETER a, parameterdetail b " & _
											 ' " WHERE     a.fccode = b.parheadcode " & _
												   ' " AND a.parcode = b.parcode " & _
												   ' " AND a.fcba = b.fcba " & _
												   ' " AND a.PARCODE = 'SR003_STOCKSALDO' " & _
												   ' " AND a.PARAM_1 = 'DEFAULT') c " & _
									 ' " WHERE     a.stockcode = c.stockcode " & _
										   ' " AND a.transdate <= TO_DATE ('" & v_txtdatefrom & "', 'MM/DD/YYYY') " & _
										   ' " AND a.fcba = c.fcba " & _
								  ' " GROUP BY c.stockGenCode, " & _
										   ' " c.stockGenName, " & _
										   ' " c.orderdata, " & _
										   ' " c.UnitOfMeasurement " & _
								  ' " UNION " & _
									' " SELECT c.stockGenCode, " & _
										   ' " c.stockGenName, " & _
										   ' " c.orderdata, " & _
										   ' " c.UnitOfMeasurement, " & _
										   ' " 0 hi, " & _
										   ' " SUM (a.quantity) shi, " & _
										   ' " 0 sbi " & _
									  ' " FROM STOCKTRANSHIST a, " & _
										   ' " (SELECT a.fcba, " & _
												   ' " a.FCCODE stockGenCode, " & _
												   ' " a.FCNAME StockGenName, " & _
												   ' " a.param_2 orderdata, " & _
												   ' " a.param_3 unitofmeasurement, " & _
												   ' " b.param_1 stockcode " & _
											  ' " FROM PARAMETER a, parameterdetail b " & _
											 ' " WHERE     a.fccode = b.parheadcode " & _
												   ' " AND a.parcode = b.parcode " & _
												   ' " AND a.fcba = b.fcba " & _
												   ' " AND a.PARCODE = 'SR003_STOCKSALDO' " & _
												   ' " AND a.PARAM_1 = 'DEFAULT') c " & _
									 ' " WHERE    a.stockcode = c.stockcode " & _
										   ' " AND a.transdate <= TO_DATE ('" & v_txtdatefrom & "', 'MM/DD/YYYY') " & _
										  ' " AND a.fcba = c.fcba " & _
								  ' " GROUP BY c.stockGenCode, " & _
										   ' " c.stockGenName, " & _
										   ' " c.orderdata, " & _
										   ' " c.UnitOfMeasurement " & _
											' " UNION " & _
									' " SELECT c.stockGenCode, " & _
										   ' " c.stockGenName, " & _
										   ' " c.orderdata, " & _
										   ' " c.UnitOfMeasurement, " & _
										   ' " 0 hi, " & _
										   ' " 0 shi, " & _
										   ' " SUM (a.quantity) sbi " & _
									  ' " FROM STOCKTRANSHIST a, " & _
										   ' " (SELECT a.fcba, " & _
												   ' " a.FCCODE stockGenCode, " & _
												   ' " a.FCNAME StockGenName, " & _
												   ' " a.param_2 orderdata, " & _
												   ' " a.param_3 unitofmeasurement, " & _
												   ' " b.param_1 stockcode " & _
											  ' " FROM PARAMETER a, parameterdetail b " & _
											 ' " WHERE     a.fccode = b.parheadcode " & _
												   ' " AND a.parcode = b.parcode " & _
												   ' " AND a.fcba = b.fcba " & _
												   ' " AND a.PARCODE = 'SR003_STOCKSALDO' " & _
												   ' " AND a.PARAM_1 = 'DEFAULT') c " & _
									 ' " WHERE     a.stockcode = c.stockcode " & _
										   ' " AND a.transdate <= TO_DATE ('" & v_txtdatefrom & "', 'MM/DD/YYYY') " & _
										   ' " AND a.fcba = c.fcba " & _
								  ' " GROUP BY c.stockGenCode, " & _
										   ' " c.stockGenName, " & _
										   ' " c.orderdata, " & _
										   ' " c.UnitOfMeasurement    " & _
										' " ) " & _
						' " GROUP BY stockGenCode, " & _
								 ' " stockGenName, " & _
								 ' " orderdata, " & _
								 ' " UnitOfMeasurement " & _
						' " ORDER BY to_number(orderdata) " & _
						' " )	 " 
     strEndingBalance = " select codedesc, ROWNUM thenumber, description, unitofmeasurement, hi, shi, sbi " & _
						"  From ( SELECT codedesc,stockGenCode,description,orderdata,UnitOfMeasurement,SUM(hi) hi,SUM (shi) shi,SUM (sbi) sbi " & _
						"	 FROM ( SELECT fcba,'1' nomer,'SR003_09' codedesc,stockGenCode,stockGenName description,orderdata,UnitOfMeasurement,SUM(hi) hi,SUM (shi) shi,SUM (sbi) sbi " & _
						"			FROM ( SELECT fcba,FCCODE stockGenCode,FCNAME stockGenName,param_2 orderdata,param_3 unitofmeasurement,0 hi,0 shi,0 sbi " & _
						"				   FROM PARAMETER " & _
						"				   WHERE PARCODE = 'SR003_STOCKSALDO' AND PARAM_1 = 'DEFAULT' " & _
						"				   UNION " & _
						"				   SELECT c.fcba,c.stockGenCode,c.stockGenName,c.orderdata,c.UnitOfMeasurement,SUM (a.quantity) hi,0 shi,0 sbi " & _
						"				   FROM STOCKTRANSHIST a, " & _
						"				   ( SELECT a.fcba,a.FCCODE stockGenCode,a.FCNAME StockGenName,a.param_2 orderdata,a.param_3 unitofmeasurement,b.param_1 stockcode " & _
						"					 FROM PARAMETER a, parameterdetail b " & _
						"					 WHERE a.fccode = b.parheadcode " & _
						"					 AND a.parcode = b.parcode " & _
						"					 AND a.fcba = b.fcba " & _
						"					 AND a.PARCODE = 'SR003_STOCKSALDO' " & _
						"					 AND a.PARAM_1 = 'DEFAULT') c " & _
						"					 WHERE  a.stockcode = c.stockcode " & _
						"					 AND a.transdate <= TO_DATE ('" & v_txtdatefrom & "', 'MM/DD/YYYY') " & _
						"					 AND a.fcba = c.fcba " & _
						"					 GROUP BY c.fcba,c.stockGenCode,c.stockGenName,c.orderdata,c.UnitOfMeasurement " & _
						"					 UNION " & _
						"					 SELECT c.fcba,c.stockGenCode,c.stockGenName,c.orderdata,c.UnitOfMeasurement,0 hi,SUM (a.quantity) shi,0 sbi " & _
						"					 FROM STOCKTRANSHIST a, " & _
						"					 ( SELECT a.fcba,a.FCCODE stockGenCode, a.FCNAME StockGenName, a.param_2 orderdata,a.param_3 unitofmeasurement, b.param_1 stockcode " & _
						"					   FROM PARAMETER a, parameterdetail b " & _
						"					   WHERE a.fccode = b.parheadcode " & _
						"					   AND a.parcode = b.parcode " & _
						"					   AND a.fcba = b.fcba " & _
						"					   AND a.PARCODE = 'SR003_STOCKSALDO' " & _
						"					   AND a.PARAM_1 = 'DEFAULT' " & _
						"					 ) c " & _
						"					 WHERE    a.stockcode = c.stockcode " & _
						"					 AND a.transdate <= TO_DATE ('" & v_txtdatefrom & "', 'MM/DD/YYYY') " & _
						"					 AND a.fcba = c.fcba " & _
						"					 GROUP BY c.fcba,c.stockGenCode,c.stockGenName,c.orderdata,c.UnitOfMeasurement " & _
						"					 UNION " & _
						"					 SELECT c.fcba,c.stockGenCode,c.stockGenName,c.orderdata, c.UnitOfMeasurement,0 hi,0 shi,SUM (a.quantity) sbi " & _
						"					 FROM STOCKTRANSHIST a, " & _
						"					 ( SELECT a.fcba,a.FCCODE stockGenCode,a.FCNAME StockGenName,a.param_2 orderdata, a.param_3 unitofmeasurement,b.param_1 stockcode " & _
						"					   FROM PARAMETER a, parameterdetail b " & _
						"					   WHERE     a.fccode = b.parheadcode " & _
						"					   AND a.parcode = b.parcode " & _
						"					   AND a.fcba = b.fcba " & _
						"					   AND a.PARCODE = 'SR003_STOCKSALDO' " & _
						"					   AND a.PARAM_1 = 'DEFAULT' " & _
						"					  ) c " & _
						"					  WHERE     a.stockcode = c.stockcode " & _
						"					  AND a.transdate <= TO_DATE ('" & v_txtdatefrom & "', 'MM/DD/YYYY') " & _
						"					  AND a.fcba = c.fcba " & _
						"					  GROUP BY c.fcba,c.stockGenCode,c.stockGenName,c.orderdata,c.UnitOfMeasurement  " & _ 
						"					 ) "& _
						"			       where ACCESSRIGHTSCOMPANY_API.IsAccessRights('" & Session("CodeName") & "',FCBA,'" & txtcompany & "') = 'TRUE' " & str1 & " " & _ 
						"				 GROUP BY fcba,stockGenCode,stockGenName,orderdata,UnitOfMeasurement " & _
						"				 )    " & _
						"				 group by nomer,codedesc,stockGenCode,description,orderdata,UnitOfMeasurement " & _
						"				 ) "
								 
	 
 	Set rsEndingBalance = Server.CreateObject("ADODB.Recordset")
	'response.write strEndingBalance
	
	rsEndingBalance.Open strEndingBalance, connstr, adOpenStatic 
	
	
    ' strMandays = " SELECT codedesc, ROWNUM thenumber, description, hi, shi, sbi, budget_hi, budget_shi, budget_sbi FROM ( " & _
				' " SELECT codedesc, description, sum(hi) hi, sum(shi) shi, sum(sbi) sbi, sum(budget_hi) budget_hi, sum(budget_shi) budget_shi, sum(budget_sbi) budget_sbi " & _
					  ' " FROM (SELECT 'SR003_12' codedesc, 'Total Tenaga Kerja' description,  " & _
								   ' " SUM (kbt + kht + khl) hi, 0 shi, 0 sbi, 0 budget_hi, 0 budget_shi, 0 budget_sbi  " & _
							  ' " FROM view_attendance_daily_summmary  " & _ 
							 ' " WHERE fddate = TO_DATE ('" & v_txtdatefrom & "', 'MM/DD/YYYY') " & _
							 ' " UNION " & _
							 ' " SELECT 'SR003_12' codedesc, 'Total Tenaga Kerja' description,  " & _
								   ' " 0 hi, SUM (kbt + kht + khl) shi, 0 sbi , 0 budget_hi, 0 budget_shi, 0 budget_sbi   " & _
							  ' " FROM view_attendance_daily_summmary  " & _ 
							 ' " WHERE fddate <= TO_DATE ('" & v_txtdatefrom & "', 'MM/DD/YYYY') " & _
									   ' " AND to_char(fddate, 'MM' ) = '" & v_txtmonth & "' " & _
								       ' " AND to_char(fddate, 'YYYY') = '" & v_txtyear & "' " & _
							 ' " UNION " & _
							 ' " SELECT 'SR003_12' codedesc, 'Total Tenaga Kerja' description,  " & _
								   ' " 0 hi, 0 shi, SUM (kbt + kht + khl) sbi, 0 budget_hi, 0 budget_shi, 0 budget_sbi    " & _
							  ' " FROM view_attendance_daily_summmary  " & _ 
							 ' " WHERE fddate <= TO_DATE ('" & v_txtdatefrom & "', 'MM/DD/YYYY') " & _
									   ' " AND to_char(fddate, 'YYYY') = '" & v_txtyear & "' " & _
							 ' " ) GROUP BY codedesc, description	" & _
							 ' " ORDER BY codedesc ) " 
	strMandays = "   SELECT codedesc, ROWNUM thenumber, description, hi, shi, sbi, budget_hi, budget_shi, budget_sbi " & _
				 "	 FROM ( Select codedesc, description, sum(hi) hi, sum(shi) shi, sum(sbi) sbi, sum(budget_hi) budget_hi, sum(budget_shi) budget_shi, sum(budget_sbi) budget_sbi  " & _
				 "			From ( SELECT fcba,'1' nomer,codedesc, description, sum(hi) hi, sum(shi) shi, sum(sbi) sbi, sum(budget_hi) budget_hi, sum(budget_shi) budget_shi, sum(budget_sbi) budget_sbi " & _
				 "			 FROM (SELECT fcba,'SR003_12' codedesc, 'Total Tenaga Kerja' description,  " & _
				 "						SUM (kbt + kht + khl) hi, 0 shi, 0 sbi, 0 budget_hi, 0 budget_shi, 0 budget_sbi  " & _
				 "					 FROM view_attendance_daily_summmary   " & _
				 "					WHERE fddate = TO_DATE ('" & v_txtdatefrom & "', 'MM/DD/YYYY') " & _
				 "					group by fcba " & _
				 "					UNION " & _
				 "					SELECT fcba,'SR003_12' codedesc, 'Total Tenaga Kerja' description,  " & _
				 "						0 hi, SUM (kbt + kht + khl) shi, 0 sbi , 0 budget_hi, 0 budget_shi, 0 budget_sbi   " & _
				 "					 FROM view_attendance_daily_summmary  " & _ 
				 "					WHERE fddate <= TO_DATE ('" & v_txtdatefrom & "', 'MM/DD/YYYY') " & _
				 "						  AND to_char(fddate, 'MM' ) = '" & v_txtmonth & "' " & _
				 "	 						AND to_char(fddate, 'YYYY') = '" & v_txtyear & "' " & _
				 "					group by fcba " & _
				 "					UNION " & _
				 "					SELECT fcba,'SR003_12' codedesc, 'Total Tenaga Kerja' description, " & _ 
				 "						0 hi, 0 shi, SUM (kbt + kht + khl) sbi, 0 budget_hi, 0 budget_shi, 0 budget_sbi    " & _
				 "					 FROM view_attendance_daily_summmary   " & _
				 "					WHERE fddate <= TO_DATE ('" & v_txtdatefrom & "', 'MM/DD/YYYY') " & _
				 "					AND to_char(fddate, 'YYYY') = '" & v_txtyear & "' " & _
				 "					group by fcba " & _
				 "					) " & _
				 "			       where ACCESSRIGHTSCOMPANY_API.IsAccessRights('" & Session("CodeName") & "',FCBA,'" & txtcompany & "') = 'TRUE' " & str1 & " " & _
				 "					GROUP BY fcba,codedesc, description  " & _
				 "				   )  " & _
				 "				   group by nomer,codedesc, description " & _
				 "	 			 ) "

	Set rsMandays = Server.CreateObject("ADODB.Recordset")
	'response.write strMandays
	
	rsMandays.Open strMandays, connstr, adOpenStatic 	


    ' strVehicleGroup = " SELECT codedesc, ROWNUM theNumber, description, hi, shi, sbi, budget_hi, budget_shi, budget_sbi FROM (" & _
	                  ' " SELECT codedesc, description, sum(hi) hi, sum(shi) shi, sum(sbi) sbi, sum(budget_hi) budget_hi, sum(budget_shi) budget_shi, sum(budget_sbi) budget_sbi " & _
						  ' " FROM (SELECT  orderreport,  'SR003_13' codedesc, " & _
										 ' " GetVehicleReportDesc(vehiclegroupcode) description, " & _
												 ' " SUM (NVL (solar, 0)) hi,  " & _
												' " 0 shi, 0 sbi, 0 budget_hi, 0 budget_shi, 0 budget_sbi " & _
									' " FROM (SELECT c.orderreport, a.fcba, c.vehiclegroupcode vehiclegroupcode, " & _
												 ' " c.description description, a.vehiclecode vehiclecode, " & _
												 ' " b.fcname vehiclename, a.unit unit, a.hours hours, " & _
												 ' " a.solar solar, " & _
													   ' " a.solar rasio " & _
											' " FROM view_vehicleproductivity a, VEHICLE b, " & _
												 ' " VEHICLEGROUP c " & _
										   ' " WHERE a.actiondate BETWEEN TO_DATE ('" & v_txtdatefrom & "','MM/DD/YYYY' ) " & _
																  ' " AND TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY') " & _
											 ' " AND a.vehiclecode = b.fccode " & _
											 ' " AND b.vehiclegroupcode = c.vehiclegroupcode " & _
											 ' " AND a.fcba = b.fcba " & _
											 ' " AND b.fcba = c.fcba) " & _
								   ' " WHERE vehiclegroupcode LIKE '%' " & _
								' " GROUP BY orderreport, vehiclegroupcode, description " & _
								' " UNION " & _
								' " SELECT  orderreport,  'SR003_13' codedesc, " & _
										 ' " GetVehicleReportDesc(vehiclegroupcode) description, 0 hi, " & _
										 ' " SUM (NVL (solar, 0)) shi,  " & _
												' " 0 sbi , 0 budget_hi, 0 budget_shi, 0 budget_sbi" & _
									' " FROM (SELECT c.orderreport, a.fcba, c.vehiclegroupcode vehiclegroupcode, " & _
												 ' " c.description description, a.vehiclecode vehiclecode, " & _
												 ' " b.fcname vehiclename, a.unit unit, a.hours hours, " & _
												 ' " a.solar solar, " & _
													   ' " a.solar rasio " & _
											' " FROM view_vehicleproductivity a, VEHICLE b, " & _
												 ' " VEHICLEGROUP c " & _
										   ' " WHERE a.actiondate <= TO_DATE ('" & v_txtdatefrom & "', 'MM/DD/YYYY') " & _
												     ' " AND to_char(a.actiondate, 'MM' ) = '" & v_txtmonth & "' " & _
												     ' " AND to_char(a.actiondate, 'YYYY') = '" & v_txtyear & "' " & _																  
													 ' " AND a.vehiclecode = b.fccode " & _
													 ' " AND b.vehiclegroupcode = c.vehiclegroupcode " & _
													 ' " AND a.fcba = b.fcba " & _
													 ' " AND b.fcba = c.fcba) " & _
								   ' " WHERE vehiclegroupcode LIKE '%' " & _
								' " GROUP BY orderreport, vehiclegroupcode, description " & _
								' " UNION " & _
								' " SELECT  orderreport,  'SR003_13' codedesc, " & _
										 ' " GetVehicleReportDesc(vehiclegroupcode) description, 0 hi, 0 shi, " & _
										 ' " SUM (NVL (solar, 0)) sbi, 0 budget_hi, 0 budget_shi, 0 budget_sbi " & _
									' " FROM (SELECT c.orderreport, a.fcba, c.vehiclegroupcode vehiclegroupcode, " & _
												 ' " c.description description, a.vehiclecode vehiclecode, " & _
												 ' " b.fcname vehiclename, a.unit unit, a.hours hours, " & _
												 ' " a.solar solar, " & _
												 ' " a.solar rasio " & _
											' " FROM view_vehicleproductivity a, VEHICLE b, " & _
												 ' " VEHICLEGROUP c " & _
										   ' " WHERE a.actiondate <= TO_DATE ('" & v_txtdatefrom & "', 'MM/DD/YYYY') " & _
												     ' " AND to_char(a.actiondate, 'YYYY') = '" & v_txtyear & "' " & _	
											 ' " AND a.vehiclecode = b.fccode " & _
											 ' " AND b.vehiclegroupcode = c.vehiclegroupcode " & _
											 ' " AND a.fcba = b.fcba " & _
											 ' " AND b.fcba = c.fcba) " & _
								   ' " WHERE vehiclegroupcode LIKE '%' " & _
								' " GROUP BY orderreport, vehiclegroupcode, description " & _
								' " ) GROUP BY codedesc, description ORDER BY codedesc ) " 
                               		
	strVehicleGroup = "  SELECT codedesc, ROWNUM theNumber, description, hi, shi, sbi, budget_hi, budget_shi, budget_sbi " & _
					  "  FROM ( Select codedesc, description, sum(hi) hi, sum(shi) shi, sum(sbi) sbi, sum(budget_hi) budget_hi, sum(budget_shi) budget_shi, sum(budget_sbi) budget_sbi " & _
					  "	 FROM ( SELECT fcba, '1' nomer,codedesc, description, sum(hi) hi, sum(shi) shi, sum(sbi) sbi, sum(budget_hi) budget_hi, sum(budget_shi) budget_shi, sum(budget_sbi) budget_sbi " & _
					  "			   FROM (SELECT fcba, orderreport,  'SR003_13' codedesc, description,SUM (NVL (solar, 0)) hi,  " & _
					  "						  0 shi, 0 sbi, 0 budget_hi, 0 budget_shi, 0 budget_sbi " & _
					  "			  FROM ( SELECT c.fcba,c.orderreport,c.vehiclegroupcode vehiclegroupcode,c.description description, a.vehiclecode vehiclecode, " & _
					  "								b.fcname vehiclename, a.unit unit, a.hours hours, a.solar solar,a.solar rasio " & _
					  "						 FROM view_vehicleproductivity a, VEHICLE b, VEHICLEGROUP c " & _
					  "					 WHERE a.actiondate BETWEEN TO_DATE ('" & v_txtdatefrom & "','MM/DD/YYYY' ) " & _
					  "						 AND TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY') " & _
					  "						 AND a.vehiclecode = b.fccode " & _
					  "						 AND b.vehiclegroupcode = c.vehiclegroupcode " & _
					  "						 AND a.fcba = b.fcba " & _
					  "						 AND b.fcba = c.fcba " & _
					  "					   ) " & _
					  "				  GROUP BY fcba,orderreport, vehiclegroupcode, description " & _
					  "				  UNION " & _
					  "				  SELECT fcba, orderreport,  'SR003_13' codedesc, description, 0 hi, " & _
					  "						  SUM (NVL (solar, 0)) shi,0 sbi , 0 budget_hi, 0 budget_shi, 0 budget_sbi " & _
					  "				  FROM ( SELECT c.fcba,c.orderreport,c.vehiclegroupcode vehiclegroupcode,c.description description, a.vehiclecode vehiclecode, " & _
					  "							   b.fcname vehiclename, a.unit unit, a.hours hours,a.solar solar,a.solar rasio " & _
					  "						 FROM view_vehicleproductivity a, VEHICLE b,VEHICLEGROUP c " & _
					  "						 WHERE a.actiondate <= TO_DATE ('" & v_txtdatefrom & "', 'MM/DD/YYYY') " & _
					  "						 AND to_char(a.actiondate, 'MM' ) = '" & v_txtmonth & "' " & _
					  "						 AND to_char(a.actiondate, 'YYYY') = '" & v_txtyear & "'  " & _                                 
					  "						 AND a.vehiclecode = b.fccode " & _
					  "						 AND b.vehiclegroupcode = c.vehiclegroupcode " & _
					  "						 AND a.fcba = b.fcba " & _
					  "						 AND b.fcba = c.fcba " & _
					  "					   ) " & _
					  "				  GROUP BY fcba,orderreport, vehiclegroupcode, description " & _
					  "				  UNION " & _
					  "			  SELECT  fcba,orderreport,  'SR003_13' codedesc, description, 0 hi, 0 shi, " & _
					  "					  SUM(NVL(solar, 0)) sbi, 0 budget_hi, 0 budget_shi, 0 budget_sbi " & _
					  "				  FROM ( SELECT c.fcba,c.orderreport,c.vehiclegroupcode vehiclegroupcode,c.description description, a.vehiclecode vehiclecode, " & _
					  "								b.fcname vehiclename, a.unit unit, a.hours hours,a.solar solar, a.solar rasio " & _
					  "						 FROM view_vehicleproductivity a, VEHICLE b,VEHICLEGROUP c " & _
					  "						 WHERE a.actiondate <= TO_DATE ('" & v_txtdatefrom & "', 'MM/DD/YYYY') " & _
					  "						 AND to_char(a.actiondate, 'YYYY') = '" & v_txtyear & "'  " & _ 
					  "						 AND a.vehiclecode = b.fccode " & _
					  "						 AND b.vehiclegroupcode = c.vehiclegroupcode " & _
					  "						 AND a.fcba = b.fcba " & _
					  "						 AND b.fcba = c.fcba " & _
					  "				   ) " & _
					  "				   GROUP BY fcba,orderreport, vehiclegroupcode,description " & _
					  "				  ) " & _
					  "			 where ACCESSRIGHTSCOMPANY_API.IsAccessRights('" & Session("CodeName") & "',FCBA,'" & txtcompany & "') = 'TRUE' " & str1 & " " & _  
					  "				 GROUP BY fcba,codedesc, description " & _
					  "			 )  " & _
					  "				 GROUP BY nomer,codedesc, description " & _
					  "			 ) "			
								
	Set rsVehicleGroup = Server.CreateObject("ADODB.Recordset")
	'response.write strVehicleGroup
	
	rsVehicleGroup.Open strVehicleGroup, connstr, adOpenStatic 	
	



	
	
%>


<html>

<head>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>IPla'S Reporting - Maintenance Report</title>

<script type='text/javascript' src='/js/x.js'></script>
<script type='text/javascript' src='/js/xtableheaderfixed.js'></script>
<script type='text/javascript'>
xAddEventListener(window, 'load',
  function() {
    new xTableHeaderFixed('xthf-site','tbl-container',0);
    new xTableHeaderFixed('xthf-green', window);
    new xTableHeaderFixed('xthf-blue', window);
  }, false
);
</script>

<script language="JavaScript" src="/js/date-picker.js"></script> 
<script language="JavaScript">
var newwin;
function launchwin(winurl,winname,winfeatures)
{
	//This launches a new window and then
	//focuses it if window.focus() is supported.
	newwin = window.open(winurl,winname,winfeatures);

}
</script>
</head>
<form name="frmEntry" method="post" action="p_Summary_BA_NAF.asp">

<body>
	<div id="wrapper">
	<div id="header-wrapper">
	<table id="menu-table" class="centered-margin">
		<tr>
			<td align="center"><a href="/login.asp">Back to Home</a></td>
			<td align="center"><a href="/report/DashboardGroupManager.asp">Back to Report</a></td>
		</tr>
	</table>
	<div id="header">
		<h4>PT. <%=txcompanyname%></h4>
	</div>
	<div id="subheader" class="centered">
		<h2>LAPORAN HARIAN KEBUN</h3> 
		    <% if (v_txtdatefrom <> "" or v_txtdateto <> "") and v_txtdatefrom = v_txtdateto then%>
			<h5>TANGGAL: <%=v_txtdateto%>, KEBUN : <%=v_txtFcba%></h5>
		<% else %>
			<h5>TANGGAL: <%=v_txtdatefrom%> - <%=v_txtdateto%>, KEBUN : <%=v_txtFcba%></h5>
		<% end if %>
	</div>
	</div>
<table id="form-table" class="centered-margin">
	<tr>
		<th>BUDGET YEAR</th>
		<th>BUDGET VERSION</th>
		<th>MODULE VERSION</th>
		<th>TAHUN</th>
		<th>BULAN</th>	
		<th>TANGGAL</th>
		<th>KEBUN</th>
		<th>&nbsp;</th>
	</tr>
	<tr>
		<td align="center"><input type="text" name="txtBudgetYear" value="<%=v_txtBudgetYear%>"></td>
		<td align="center"><input type="text" name="txtBudgetVersion" value="<%=v_txtBudgetVersion%>"></td>
		<td align="center"><input type="text" name="txtModuleVersion" value="<%=v_txtModuleVersion%>"></td>
		<input type="hidden" name="txtComp" value="<%=v_txtComp%>">
		<td align="center"><input type="text" name="txtyear" value="<%=v_txtyear%>"></td>
		<td align="center"><input type="text" name="txtmonth" value="<%=v_txtmonth%>"></td>

		<td align="center">
		<% if v_txtdatefrom <> "" and v_txtdatefrom = v_txtdatefrom then %>
			<input type="text" maxlength="10" name="txtdatefrom" size="10" value="<%=v_txtdatefrom%>" title="Type Date mm/dd/yyyy" >
		<% else %>
			<input type="text" maxlength="10" name="txtdatefrom" size="10" value="<%=textDate(Date)%>" title="Type Date mm/dd/yyyy" >
		<% end if%>
			<a href="javascript:show_calendar('frmEntry.txtdatefrom');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
				<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top>
			</a>
		</td>

		
		<td align="center">
		<select name="txtFcba">	
		<option value="">--ALL--</option>		
		<%
			Dim rsFcba
			Set rsFcba = Server.CreateObject("ADODB.Recordset")
			rsFcba.CursorLocation = adUseClient
			strsql =  "SELECT fccode FROM businessunit where fccompanycode = '" & txtcompany & "'" 
			response.write strsql
			rsFcba.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
			
			if not rsFcba.eof then
				DO while not rsFcba.eof
				if v_txtFcba = rsFcba("fccode") then
		%>			<option value="<%=rsFcba("fccode")%>" Selected><%=rsFcba("fccode")%></option>
				<%else%>
					<option value="<%=rsFcba("fccode")%>"><%=rsFcba("fccode")%></option>
				<%end if%>
		<%			rsFcba.movenext
				loop
			End if
			rsFcba.Close
			set rsFcba = nothing
		%>
		</select>		
		</td>

		
		
		<td align="center"><input type="submit" value="View" /></td>
	</tr>
</table>

<table border="0" width="100%" id="product-table" class='xthf-site'>
<thead>
	<tr>
		<th class="information" colspan="3"> INFORMASI </th>
	    <th class="uom" rowspan="2" width="20px"> UOM </th>
		<th class="actual" colspan="3"> ACTUAL </th>
		<th class="budget" colspan="3"> TARGET-BUDGET-STANDARD </th>
		<th class="indicater" colspan="3"> INDICATOR (%) </th>
	</tr>
	<tr>
	    <th class="information" width="20px"> GROUP </th>
	    <th class="information" width="20px" > NO. </th>
	    <th class="information" width="200px"> DESKRIPSI </th>
	    
		<th class="actual" width="20px">  HI  </th>
	    <th class="actual" width="20px"> SHI  </th>
	    <th class="actual" width="20px"> SBI  </th>
		
	    <th class="budget" width="20px">  HI  </th>
	    <th class="budget" width="20px"> SHI  </th>
	    <th class="budget" width="20px"> SBI  </th>
	    
		<th class="indicater" width="20px">  HI  </th>
	    <th class="indicater" width="20px"> SHI  </th>
	    <th class="indicater" width="20px"> SBI  </th>
	</tr>
</thead>
	
<!-- Start Land Clearing -->	
<tbody>
<tr>
<th class="separator" colspan="13" align="left"></th>				
</tr>	



<!-- Start HandlingAmount -->	
<%
	if not rsLandClearing.EOF then
	    no = 0

		rsLandClearing.MoveLast
		tempRowSpan = rsLandClearing("theNumber")
		rsLandClearing.MoveFirst

		while not rsLandClearing.EOF
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
			no = no + 1
%>


			<tr class="<%=cClass%>">
			
			    <%
					strCodeDesc = "SELECT param_1 CodeDescription, rate_1 theRowSpan FROM PARAMETERDETAIL " & _
									" WHERE " & _
									" parcode = 'SR003' AND " & _
									" parheadcode = 'REPORTCODE' AND " & _
									" param_2 = 'DEFAULT' AND " & _
									" fccode = '" & rsLandClearing("codedesc") & "'" 
					
				 	Set rsCodeDesc = Server.CreateObject("ADODB.Recordset")	
					rsCodeDesc.Open strCodeDesc, connstr, adOpenStatic 		

					if not rsCodeDesc.EOF then
						tempCodeDescription = rsCodeDesc ("CodeDescription")
						'tempRowSpan = cdbl(rsCodeDesc ("theRowSpan"))
                    else
                        tempCodeDescription = ""					
						'tempRowSpan = cdbl(rsCodeDesc ("theRowSpan"))
					end if					
									
									
                %>
				
				<% if no = 1 then %>
					<% if sPrevDesc <> tempCodeDescription then %>				
						<td class="boldfont" rowspan="<%= tempRowSpan %>" align="center" ><%=tempCodeDescription%></td>
						<% sPrevDesc = tempCodeDescription %>
					<% else %>
						<td align="left" ></td>
					<% end if %>
				<% end if %>
								
			
				<td align="center"  ><%=rsLandClearing("theNumber")%></td>
				
<% 
				    tempRefCode = rsLandClearing("codedesc")
				    strHRef = " SELECT param_5 reference FROM PARAMETERDETAIL " & _
								" WHERE " & _
								" parcode = 'SR003' AND " & _
								" fccode = '" & tempRefCode & "' AND " & _
								" param_2 = 'DEFAULT' " 
				
				 	Set rsHRef = Server.CreateObject("ADODB.Recordset")	
					rsHRef.Open strHRef, connstr, adOpenStatic 
					
					if not rsHRef.EOF then
						tempReference = rsHRef ("reference")
                    else
                        tempReference = ""					
					end if
					tempUOM = rsLandClearing("UnitOfMeasurement")
					
					rsHRef.Close
					set rsHRef = Nothing
				%>
				<td align="left" ><a href="<%=tempReference%>"><%=rsLandClearing("description")%></a></td>
				<td align="center" ><%=tempUOM%></td>		

				<% 
				    strHRef = " SELECT rate_3 rounding, param_4 tempBudgetType  FROM PARAMETER " & _
								" WHERE " & _
								" parcode = 'SR003_LC' AND " & _
								" fccode = '" & rsLandClearing("JobGenCode") & "'"
				
				 	Set rsHRef = Server.CreateObject("ADODB.Recordset")	
					rsHRef.Open strHRef, connstr, adOpenStatic 
					
					if not rsHRef.EOF then
						if rsHRef ("rounding") <> "" then 
							tempRounding = cdbl(rsHRef ("rounding"))
                            tempBudget_HI = cdbl(rsLandClearing ("budget_hi"))
                            tempBudget_SHI = cdbl(rsLandClearing ("budget_shi"))
                            tempBudget_SBI = cdbl(rsLandClearing ("budget_sbi"))
                            tempBudgetType = rsHRef ("tempBudgetType")
						else
						    tempRounding = 0
                            tempBudget_HI = 0
                            tempBudget_SHI = 0
                            tempBudget_SBI = 0
                            tempBudgetType = "QUANTITY"
						end if
                    else
                        tempRounding = 0				
					end if

					rsHRef.Close
					set rsHRef = Nothing
				%>
				
				
				
				
				<% ' Start Compare to Budget 

					tempDataBudget_HI = tempBudget_HI
					tempDataBudget_SHI = tempBudget_SHI
					tempDataBudget_SBI = tempBudget_SBI	
                %>

				
				<%  if rsLandClearing("hi") <>  "" then 
						tempData = cdbl(rsLandClearing("hi"))
				    else
					    tempData = 0
					end if
				%>
				<td align="right" ><%=FormatNumber(tempData,tempRounding)%></td>
				<%  if rsLandClearing("shi") <>  "" then 
						tempData_SHI = cdbl(rsLandClearing("shi"))
				    else
					    tempData_SHI = 0
					end if
				%>
				<td align="right" ><%=FormatNumber(tempData_SHI,tempRounding)%></td>
				<%  if rsLandClearing("sbi") <>  "" then 
						tempData_SBI = cdbl(rsLandClearing("sbi"))
				    else
					    tempData_SBI = 0
					end if
				%>
				<td align="right" ><%=FormatNumber(tempData_SBI,tempRounding)%></td>
				
				<td align="right" ><%=FormatNumber(tempDataBudget_HI,tempRounding)%></td>
				<td align="right" ><%=FormatNumber(tempDataBudget_SHI,tempRounding)%></td>
				<td align="right" ><%=FormatNumber(tempDataBudget_SBI,tempRounding)%></td>


                <%

					if tempDataBudget_HI <> 0 then 
					    tempCompare_HI = tempData/tempDataBudget_HI * 100
					else
					    tempCompare_HI = 0
					end if
					
					if tempDataBudget_SHI <> 0 then 
					    tempCompare_SHI = tempData_SHI/tempDataBudget_SHI * 100
					else
					    tempCompare_SHI = 0
					end if

					if tempDataBudget_SBI <> 0 then 
					    tempCompare_SBI = tempData_SBI/tempDataBudget_SBI * 100
					else
					    tempCompare_SBI = 0
					end if
					
				%>					

				<%
                    if tempBudgetType = "QUANTITY" then  
					    if tempCompare_HI > 95 then
					        cClassAssigned_HI = "green"
					    else
					        if tempCompare_HI > 85 then 
					             cClassAssigned_HI = "yellow"
						    else
						         cClassAssigned_HI = "red"
						    end if
					    end if
					
					    if tempCompare_SHI > 95 then
					        cClassAssigned_SHI = "green"
					    else
					        if tempCompare_SHI > 85 then 
					             cClassAssigned_SHI = "yellow"
						    else
						         cClassAssigned_SHI = "red"
						    end if
					    end if
					
					    if tempCompare_SBI > 95 then
					        cClassAssigned_SBI = "green"
					    else
					        if tempCompare_SBI > 85 then 
					             cClassAssigned_SBI = "yellow"
						    else
						         cClassAssigned_SBI = "red"
						    end if
					    end if
                    else
					    if tempCompare_HI > 105 then
					        cClassAssigned_HI = "red"
					    else
					        if tempCompare_HI > 95 then 
					             cClassAssigned_HI = "yellow"
						    else
						         cClassAssigned_HI = "green"
						    end if
					    end if
					
					    if tempCompare_SHI > 105 then
					        cClassAssigned_SHI = "red"
					    else
					        if tempCompare_SHI > 95 then 
					             cClassAssigned_SHI = "yellow"
						    else
						         cClassAssigned_SHI = "green"
						    end if
					    end if
					
					    if tempCompare_SBI > 105 then
					        cClassAssigned_SBI = "red"
					    else
					        if tempCompare_SBI > 95 then 
					             cClassAssigned_SBI = "yellow"
						    else
						         cClassAssigned_SBI = "green"
						    end if
					    end if
                    end if					
				%>


				
				<td class="<%=cClassAssigned_HI%>" align="right" ><%=FormatNumber(tempCompare_HI,2)%></td>
				<td class="<%=cClassAssigned_SHI%>" align="right" ><%=FormatNumber(tempCompare_SHI,2)%></td>
				<td class="<%=cClassAssigned_SBI%>" align="right" ><%=FormatNumber(tempCompare_SBI,2)%></td>
				
			</tr>
<%

			rsLandClearing.MoveNext
		Wend
	end if
	rsLandClearing.Close
	set rsLandClearing = Nothing
%>

	
	
<!-- Start Production -->	

<tr>
<th class="separator" colspan="13" align="left"></th>				
</tr>



<%
	if not rsProduction.EOF then
        no = 0
		rsProduction.MoveLast
		tempRowSpan = rsProduction("theNumber")
		rsProduction.MoveFirst
	    sPrevDesc = ""
		while not rsProduction.EOF
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
			no = no + 1
%>


			<tr class="<%=cClass%>">
				
			    <%
					strCodeDesc = "SELECT param_1 CodeDescription, rate_1 theRowSpan FROM PARAMETERDETAIL " & _
									" WHERE " & _
									" parcode = 'SR003' AND " & _
									" parheadcode = 'REPORTCODE' AND " & _
									" param_2 = 'DEFAULT' AND " & _
									" fccode = '" & rsProduction("codedesc") & "'" 
					
				 	Set rsCodeDesc = Server.CreateObject("ADODB.Recordset")	
					rsCodeDesc.Open strCodeDesc, connstr, adOpenStatic 		

					if not rsCodeDesc.EOF then
						tempCodeDescription = rsCodeDesc ("CodeDescription")
						'tempRowSpan =  cdbl(rsCodeDesc ("theRowSpan"))
                    else
                        tempCodeDescription = ""	
						'tempRowSpan	= 1					
					end if					
									
									
                %>
				
				
				<% if no = 1 then %>
					<% if sPrevDesc <> tempCodeDescription then %>				
						<td class="boldfont" rowspan="<%= tempRowSpan %>" align="center" ><%=tempCodeDescription%></td>
						<% sPrevDesc = tempCodeDescription %>
					<% else %>
						<td align="left" ></td>
					<% end if %>
				<% end if %>

				
				
				
				<td align="center" ><%=rsProduction("theNumber")%></td>
				
				
				<% 
				    tempRefCode = rsProduction("codedesc")
				    strHRef = " SELECT param_5 reference,  rate_3 rounding, param_6 UOM,  param_7 budgetType FROM PARAMETERDETAIL " & _
								" WHERE " & _
								" parcode = 'SR003' AND " & _
								" fccode = '" & tempRefCode & "' AND " & _
								" param_2 = 'DEFAULT' " 
				
				 	Set rsHRef = Server.CreateObject("ADODB.Recordset")	
					rsHRef.Open strHRef, connstr, adOpenStatic 

                    tempBudget_HI = 0 
                    tempBudget_SHI = 0
                    tempBudget_SBI = 0

					
					if not rsHRef.EOF then
						tempReference = rsHRef ("reference")
						tempUOM = rsHRef("UOM")
						tempRounding = cdbl(rsHRef("rounding"))
                        tempBudget_HI = cdbl(rsProduction("budget_hi"))
                        tempBudget_SHI = cdbl(rsProduction("budget_shi"))
                        tempBudget_SBI = cdbl(rsProduction("budget_sbi"))
                        tempBudgetType = rsHRef("budgetType")
                    else
                        tempReference = ""	
						tempUOM = ""
						tempRounding = 0
                        tempBudget_HI = 0
                        tempBudget_SHI = 0
                        tempBudget_SBI = 0
                        tempBudgetType = "QUANTITY"
					end if
					
					rsHRef.Close
					set rsHRef = Nothing

                    tempDataBudget_HI = tempBudget_HI
					tempDataBudget_SHI = tempBudget_SHI
					tempDataBudget_SBI = tempBudget_SBI
					
					tempExtendedReference = "txtyear="& v_txtyear & "&txtMonth=" & v_txtMonth & "&txttodaydate=" & _ 
					v_txtdatefrom & "&txtDivision=--SUMMARY--&txtBlok=--SUMMARY--&txtPlantingDate=--SUMMARY--&txtfcba=" & v_txtfcba & "&txtComp=" & v_txtComp & ""
					
					tempReference = tempReference & tempExtendedReference
					
				%>
				
				<td align="left" ><a href="<%=tempReference%>"><%=rsProduction("description")%></a></td>
				<td align="center" ><%=tempUOM%></td>
				

				
				

					
				
				
				
				
				<%  if rsProduction("hi") <>  "" then 
						tempData = cdbl(rsProduction("hi"))
				    else
					    tempData = 0
					end if
				%>
				
				<%  if rsProduction("shi") <>  "" then 
						tempData_shi = cdbl(rsProduction("shi"))
				    else
					    tempData_shi = 0
					end if
				%>
				
				<%  if rsProduction("sbi") <>  "" then 
						tempData_sbi = cdbl(rsProduction("sbi"))
				    else
					    tempData_sbi = 0
					end if
				%>
				

				
				<% if tempRefCode = "SR003_01" then 
				    tempTBSProduksi = cdbl(rsProduction("hi"))
					tempTBSProduksi_shi = cdbl(rsProduction("shi"))
					tempTBSProduksi_sbi = cdbl(rsProduction("sbi"))
				   end if
				   
				   if tempRefCode = "SR003_BA" then 
				      if tempTBSProduksi <> 0 then 
				          tempBiayaAngkut = cdbl(rsProduction("hi")) / tempTBSProduksi
				      else
					      tempBiayaAngkut = cdbl(rsProduction("hi")) 
				      end if
				      
					  if tempTBSProduksi_shi <> 0 then 
					      tempBiayaAngkut_shi = cdbl(rsProduction("shi")) / tempTBSProduksi_shi
					  else
					      tempBiayaAngkut_shi = cdbl(rsProduction("shi")) 
					  end if
					  
					  if tempTBSProduksi_sbi <> 0 then 
						  tempBiayaAngkut_sbi = cdbl(rsProduction("sbi")) / tempTBSProduksi_sbi
					  else
					      tempBiayaAngkut_sbi = cdbl(rsProduction("sbi")) 
					  end if
					  
					  tempData = tempBiayaAngkut
					  tempData_shi = tempBiayaAngkut_shi
					  tempData_sbi = tempBiayaAngkut_sbi
					  
				   end if

				   if tempRefCode = "SR003_OP" then 
				      if cdbl(rsProduction("hi")) <> 0 then 
				          tempBiayaAngkut = tempTBSProduksi / cdbl(rsProduction("hi")) 
				      else
					      tempBiayaAngkut = 0
				      end if
				      
					  if cdbl(rsProduction("shi")) <> 0 then 
					      tempBiayaAngkut_shi =  tempTBSProduksi_shi / cdbl(rsProduction("shi"))
					  else
					      tempBiayaAngkut_shi = 0
					  end if
					  
					  if cdbl(rsProduction("sbi")) <> 0 then 
						  tempBiayaAngkut_sbi =  tempTBSProduksi_sbi / cdbl(rsProduction("sbi"))
					  else
					      tempBiayaAngkut_sbi = 0
					  end if
					  
					  tempData = tempBiayaAngkut
					  tempData_shi = tempBiayaAngkut_shi
					  tempData_sbi = tempBiayaAngkut_sbi
					  
				   end if
				   
				   
				   
				%>  				
				
				
				
				<td align="right" ><%=FormatNumber(tempData,tempRounding)%></td>
				
				
				<% if rsProduction("codedesc") = "SR003_02" or rsProduction("codedesc") = "SR003_03" then %>
					<td class="grey" align="right" ></td>
					<td class="grey" align="right" ></td>
				<% else %>
					<td align="right" ><%=FormatNumber(tempData_shi,tempRounding)%></td>
					<td align="right" ><%=FormatNumber(tempData_sbi,tempRounding)%></td>
				<% end if %>
				
				<% if rsProduction("codedesc") = "SR003_04" then %>
					 <td class="grey" align="right" ></td>
				<% else %>
				     <td align="right" ><%=FormatNumber(tempDataBudget_HI,tempRounding)%></td>
				<% end if %>
				
				<% if rsProduction("codedesc") = "SR003_04" or rsProduction("codedesc") = "SR003_02" or rsProduction("codedesc") = "SR003_03"  then %>
					<td class="grey" align="right" ></td>
					<td class="grey" align="right" ></td>
				<% else %>
					<td align="right" ><%=FormatNumber(tempDataBudget_SHI,tempRounding)%></td>
					<td align="right" ><%=FormatNumber(tempDataBudget_SBI,tempRounding)%></td>
				<% end if %>
					

				<% ' Start Compare to Budget 
				   ' Later, This Temporary Variable must be picked from Master Budget Table.

					
				   ' End of Modification

				    if tempDataBudget_HI <> 0 then 
					    tempCompare_HI = tempData /tempDataBudget_HI * 100
				    else
					    tempCompare_HI = 0
				    end if
					
				    if tempDataBudget_SHI <> 0 then 
					    tempCompare_SHI = tempData_shi/tempDataBudget_SHI * 100
				    else
					    tempCompare_SHI = 0
				    end if

				    if tempDataBudget_SBI <> 0 then 
					    tempCompare_SBI = tempData_sbi/tempDataBudget_SBI * 100
				    else
					    tempCompare_SBI = 0
				    end if
					
				%>


				<%
                    if tempBudgetType = "QUANTITY" then  
					    if tempCompare_HI > 95 then
					        cClassAssigned_HI = "green"
					    else
					        if tempCompare_HI > 85 then 
					             cClassAssigned_HI = "yellow"
						    else
						         cClassAssigned_HI = "red"
						    end if
					    end if
					
					    if tempCompare_SHI > 95 then
					        cClassAssigned_SHI = "green"
					    else
					        if tempCompare_SHI > 85 then 
					             cClassAssigned_SHI = "yellow"
						    else
						         cClassAssigned_SHI = "red"
						    end if
					    end if
					
					    if tempCompare_SBI > 95 then
					        cClassAssigned_SBI = "green"
					    else
					        if tempCompare_SBI > 85 then 
					             cClassAssigned_SBI = "yellow"
						    else
						         cClassAssigned_SBI = "red"
						    end if
					    end if
                    else
					    if tempCompare_HI > 105 then
					        cClassAssigned_HI = "red"
					    else
					        if tempCompare_HI > 95 then 
					             cClassAssigned_HI = "yellow"
						    else
						         cClassAssigned_HI = "green"
						    end if
					    end if
					
					    if tempCompare_SHI > 105 then
					        cClassAssigned_SHI = "red"
					    else
					        if tempCompare_SHI > 95 then 
					             cClassAssigned_SHI = "yellow"
						    else
						         cClassAssigned_SHI = "green"
						    end if
					    end if
					
					    if tempCompare_SBI > 105 then
					        cClassAssigned_SBI = "red"
					    else
					        if tempCompare_SBI > 95 then 
					             cClassAssigned_SBI = "yellow"
						    else
						         cClassAssigned_SBI = "green"
						    end if
					    end if
                    end if					
				%>
				
				<% if rsProduction("codedesc") = "SR003_04"  then %>
					<td class="grey" align="right" ></td>
				<% else %>				
					<td class="<%=cClassAssigned_HI%>" align="right" ><%=FormatNumber(tempCompare_HI,2)%></td>
				<% end if %>
				
				<% if rsProduction("codedesc") = "SR003_04" or rsProduction("codedesc") = "SR003_02" or rsProduction("codedesc") = "SR003_03"  then %>
					<td class="grey" align="right" ></td>
					<td class="grey" align="right" ></td>
				<% else %>
					<td class="<%=cClassAssigned_SHI%>" align="right" ><%=FormatNumber(tempCompare_SHI,2)%></td>
					<td class="<%=cClassAssigned_SBI%>" align="right" ><%=FormatNumber(tempCompare_SBI,2)%></td>
				<% end if %>
				
				
	
				
				
				
				
			</tr>
<%

			rsProduction.MoveNext
		Wend
	end if
	rsProduction.Close
	set rsProduction = Nothing
%>




	


	
	

<tr>
<th class="separator" colspan="13" align="left"  ></th>				
</tr>

<tr>
    <th class="separator_title" align="center"  >PERAWATAN</th>				
    <th class="separator" align="left" colspan="12" ></th>				
</tr>

<!-- Start Output -->	
<%
	if not rsOutput.EOF then
	
        rsOutput.MoveLast
		tempRowSpan = rsOutput("theNumber")
		rsOutput.MoveFirst
		no = 0
		while not rsOutput.EOF
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
			no = no + 1
%>


			<tr class="<%=cClass%>">
			
			    <%
					strCodeDesc = "SELECT param_1 CodeDescription, rate_1 theRowSpan FROM PARAMETERDETAIL " & _
									" WHERE " & _
									" parcode = 'SR003' AND " & _
									" parheadcode = 'REPORTCODE' AND " & _
									" param_2 = 'DEFAULT' AND " & _
									" fccode = '" & rsOutput("codedesc") & "'" 
					
				 	Set rsCodeDesc = Server.CreateObject("ADODB.Recordset")	
					rsCodeDesc.Open strCodeDesc, connstr, adOpenStatic 		

					if not rsCodeDesc.EOF then
						tempCodeDescription = rsCodeDesc ("CodeDescription")
						'tempRowSpan = cdbl(rsCodeDesc ("theRowSpan"))
                    else
						'tempRowSpan = 1
                        tempCodeDescription = ""					
					end if					
									
									
                %>
				
				
				<% if no = 1 then %>
					<% if sPrevDesc <> tempCodeDescription then %>				
						<td class="boldfont" rowspan="<%= tempRowSpan %>" align="center" ><%=tempCodeDescription%></td>
						<% sPrevDesc = tempCodeDescription %>
					<% else %>
						<td align="left" ></td>
					<% end if %>
				<% end if %>


			
				<td align="center" ><%=rsOutput("theNumber")%></td>
				
				<% 
				    tempRefCode = rsOutput("codedesc")
				    strHRef = " SELECT param_5 reference, rate_3 rounding, param_6 UOM FROM PARAMETERDETAIL " & _
								" WHERE " & _
								" parcode = 'SR003' AND " & _
								" fccode = '" & tempRefCode & "' AND " & _
								" param_2 = 'DEFAULT' " 
				
				 	Set rsHRef = Server.CreateObject("ADODB.Recordset")	
					rsHRef.Open strHRef, connstr, adOpenStatic 
					
					if not rsHRef.EOF then
						tempReference = rsHRef ("reference")
						tempUOM = rsHRef ("UOM")
						tempRounding = cdbl(rsHRef ("rounding"))
                    else
                        tempReference = ""					
                        tempUOM = ""					
						tempRounding = 0
					end if
					
					rsHRef.Close
					set rsHRef = Nothing
					
					tempExtendedReference = "txtfcba=" & v_txtfcba & "&txtdatefrom=" & v_txtdatefrom & "&txtdateto=" & v_txtdateto & _
								            "&txtDivision=--SUMMARY--&txtBlok=--SUMMARY--&txtPlantingDate=--SUMMARY--" & _
											"&txtGroup=&txtStatusTmTbm=--SUMMARY--&txtJobCode=--SUMMARY--&txtComp=" & v_txtComp & ""
					
					tempReference = tempReference & tempExtendedReference
					
				%>
				<td align="left" ><a href="<%=tempReference%>"><%=rsOutput("description")%></a></td>
				<td align="center" ><%=tempUOM%></td>

				<% ' Start Compare to Budget 

				    strBudget = " select param_1 tempBudgetType " & _
                                  " from parameterdetail " & _
                                  " where " & _
                                  " parcode ='BUDGET' and  " & _
                                  " parheadcode = 'BUDGET_HASILKERJA' and  " & _
                                  " param_2 = 'DEFAULT' and  " & _
                                  " fccode = '" & rsOutput("groupCode") & "'"
					Set rsBudget = Server.CreateObject("ADODB.Recordset")	
					rsBudget.Open strBudget, connstr, adOpenStatic 
					
					if not rsBudget.EOF then
					    tempBudgetType = rsBudget ("tempBudgetType")
						tempDataBudget_HI = cdbl(rsOutput ("budget_hi"))
						tempDataBudget_SHI = cdbl( rsOutput ("budget_shi"))
						tempDataBudget_SBI = cdbl( rsOutput ("budget_sbi"))
                    else
						tempDataBudget_HI = 0
						tempDataBudget_SHI = 0
						tempDataBudget_SBI = 0
						tempBudgetType = "QUANTITY"
					end if
					
					rsBudget.Close
					set rsBudget = Nothing			  
				


					if tempDataBudget_HI <> 0 then 
					    tempCompare_HI = cdbl(rsOutput("hi"))/tempDataBudget_HI * 100
					else
					    tempCompare_HI = 0
					end if
					
					if tempDataBudget_SHI <> 0 then 
					    tempCompare_SHI = cdbl(rsOutput("shi"))/tempDataBudget_SHI * 100
					else
					    tempCompare_SHI = 0
					end if

					if tempDataBudget_SBI <> 0 then 
					    tempCompare_SBI = cdbl(rsOutput("sbi"))/tempDataBudget_SBI * 100
					else
					    tempCompare_SBI = 0
					end if
					
				%>
				
				
				<%  if rsOutput("hi") <>  "" then 
						tempData = cdbl(rsOutput("hi"))
				    else
					    tempData = 0
					end if
				%>
				<td align="right" ><%=FormatNumber(tempData,tempRounding)%></td>
				<%  if rsOutput("shi") <>  "" then 
						tempData = cdbl(rsOutput("shi"))
				    else
					    tempData = 0
					end if
				%>
				<td align="right" ><%=FormatNumber(tempData,tempRounding)%></td>
				<%  if rsOutput("sbi") <>  "" then 
						tempData = cdbl(rsOutput("sbi"))
				    else
					    tempData = 0
					end if
				%>
				<td align="right" ><%=FormatNumber(tempData,tempRounding)%></td>
				
				<td align="right" ><%=FormatNumber(tempDataBudget_HI,tempRounding)%></td>
				<td align="right" ><%=FormatNumber(tempDataBudget_SHI,tempRounding)%></td>
				<td align="right" ><%=FormatNumber(tempDataBudget_SBI,tempRounding)%></td>
						
				
				<%
                    if tempBudgetType = "QUANTITY" then  
					    if tempCompare_HI > 95 then
					        cClassAssigned_HI = "green"
					    else
					        if tempCompare_HI > 85 then 
					             cClassAssigned_HI = "yellow"
						    else
						         cClassAssigned_HI = "red"
						    end if
					    end if
					
					    if tempCompare_SHI > 95 then
					        cClassAssigned_SHI = "green"
					    else
					        if tempCompare_SHI > 85 then 
					             cClassAssigned_SHI = "yellow"
						    else
						         cClassAssigned_SHI = "red"
						    end if
					    end if
					
					    if tempCompare_SBI > 95 then
					        cClassAssigned_SBI = "green"
					    else
					        if tempCompare_SBI > 85 then 
					             cClassAssigned_SBI = "yellow"
						    else
						         cClassAssigned_SBI = "red"
						    end if
					    end if
                    else
					    if tempCompare_HI > 105 then
					        cClassAssigned_HI = "red"
					    else
					        if tempCompare_HI > 95 then 
					             cClassAssigned_HI = "yellow"
						    else
						         cClassAssigned_HI = "green"
						    end if
					    end if
					
					    if tempCompare_SHI > 105 then
					        cClassAssigned_SHI = "red"
					    else
					        if tempCompare_SHI > 95 then 
					             cClassAssigned_SHI = "yellow"
						    else
						         cClassAssigned_SHI = "green"
						    end if
					    end if
					
					    if tempCompare_SBI > 105 then
					        cClassAssigned_SBI = "red"
					    else
					        if tempCompare_SBI > 95 then 
					             cClassAssigned_SBI = "yellow"
						    else
						         cClassAssigned_SBI = "green"
						    end if
					    end if
                    end if					
				%>				

				
				<td class="<%=cClassAssigned_HI%>" align="right" ><%=FormatNumber(tempCompare_HI,2)%></td>
				<td class="<%=cClassAssigned_SHI%>" align="right" ><%=FormatNumber(tempCompare_SHI,2)%></td>
				<td class="<%=cClassAssigned_SBI%>" align="right" ><%=FormatNumber(tempCompare_SBI,2)%></td>
			
				
			</tr>
<%

			rsOutput.MoveNext
		Wend
	end if
	rsOutput.Close
	set rsOutput = Nothing
%>



	



<!-- Start Handling -->	
<%
	if not rsHandling.EOF then
        rsHandling.MoveLast
		tempRowSpan = rsHandling("theNumber")
		rsHandling.MoveFirst
		no = 0
		while not rsHandling.EOF
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
			no = no + 1
%>


			<tr class="<%=cClass%>">
			
			    <%
					strCodeDesc = "SELECT param_1 CodeDescription, rate_1 theRowSpan FROM PARAMETERDETAIL " & _
									" WHERE " & _
									" parcode = 'SR003' AND " & _
									" parheadcode = 'REPORTCODE' AND " & _
									" param_2 = 'DEFAULT' AND " & _
									" fccode = '" & rsHandling("codedesc") & "'" 
					
				 	Set rsCodeDesc = Server.CreateObject("ADODB.Recordset")	
					rsCodeDesc.Open strCodeDesc, connstr, adOpenStatic 		

					if not rsCodeDesc.EOF then
						tempCodeDescription = rsCodeDesc ("CodeDescription")
						'tempRowSpan = cdbl(rsCodeDesc ("theRowSpan"))
                    else
                        tempCodeDescription = ""					
						'tempRowSpan = 1
					end if					
									
									
                %>
				
				<% if no = 1 then  %>
					<% if sPrevDesc <> tempCodeDescription then %>				
						<td class="boldfont" rowspan="<%=tempRowSpan%>" align="center" ><%=tempCodeDescription%></td>
						<% sPrevDesc = tempCodeDescription %>
					<% else %>
						<td align="left" ></td>
					<% end if %>
				<% end if %>
			
				<td align="center" ><%=rsHandling("theNumber")%></td>
				
				<% 
				    tempRefCode = rsHandling("codedesc")
				    strHRef = " SELECT param_5 reference, rate_3 rounding, param_6 UOM FROM PARAMETERDETAIL " & _
								" WHERE " & _
								" parcode = 'SR003' AND " & _
								" fccode = '" & tempRefCode & "' AND " & _
								" param_2 = 'DEFAULT' " 
				
				 	Set rsHRef = Server.CreateObject("ADODB.Recordset")	
					rsHRef.Open strHRef, connstr, adOpenStatic 
					
					if not rsHRef.EOF then
						tempReference = rsHRef ("reference")
						tempUOM = rsHRef ("UOM")
						tempRounding = cdbl(rsHRef ("rounding"))
                    else
                        tempReference = ""					
                        tempUOM = ""					
						tempRounding = 0
					end if
					
					rsHRef.Close
					set rsHRef = Nothing
					
					tempExtendedReference = "txtfcba=" & v_txtfcba & "&txtdatefrom=" & v_txtdatefrom & "&txtdateto=" & v_txtdateto & _
								            "&txtDivision=--SUMMARY--&txtBlok=--SUMMARY--&txtPlantingDate=--SUMMARY--" & _
											"&txtGroup=&txtStatusTmTbm=--SUMMARY--&txtJobCode=--SUMMARY--&txtComp=" & v_txtComp & ""
					
					tempReference = tempReference & tempExtendedReference
					
				%>
				<td align="left" ><a href="<%=tempReference%>"><%=rsHandling("description")%></a></td>
				<td align="center" ><%=tempUOM%></td>

				<% 
				    ' Start Compare to Budget 


				    strBudget = " select param_1 tempBudgetType " & _
                                  " from parameterdetail " & _
                                  " where " & _
                                  " parcode ='BUDGET' and  " & _
                                  " parheadcode = 'BUDGET_RASIO' and  " & _
                                  " param_2 = 'DEFAULT' and  " & _
                                  " fccode = '" & rsHandling("groupCode") & "'"
								  
					Set rsBudget = Server.CreateObject("ADODB.Recordset")	
					rsBudget.Open strBudget, connstr, adOpenStatic 
					
					if not rsBudget.EOF then
					    tempBudgetType = rsBudget ("tempBudgetType")
						tempDataBudget_HI = cdbl(rsHandling ("budget_hi"))
						tempDataBudget_SHI = cdbl( rsHandling ("budget_shi"))
						tempDataBudget_SBI = cdbl( rsHandling ("budget_sbi"))
                    else
						tempDataBudget_HI = 0
						tempDataBudget_SHI = 0
						tempDataBudget_SBI = 0
						tempBudgetType = "QUANTITY"
					end if
					
					rsBudget.Close
					set rsBudget = Nothing			  
				

					if tempDataBudget_HI <> 0 then 
					    tempCompare_HI = cdbl(rsHandling("hi"))/tempDataBudget_HI * 100
					else
					    tempCompare_HI = 0
					end if
					
					if tempDataBudget_SHI <> 0 then 
					    tempCompare_SHI = cdbl(rsHandling("shi"))/tempDataBudget_SHI * 100
					else
					    tempCompare_SHI = 0
					end if

					if tempDataBudget_SBI <> 0 then 
					    tempCompare_SBI = cdbl(rsHandling("sbi"))/tempDataBudget_SBI * 100
					else
					    tempCompare_SBI = 0
					end if
					
				%>				
				
				<%  if rsHandling("hi") <>  "" then 
						tempData = cdbl(rsHandling("hi"))
				    else
					    tempData = 0
					end if
				%>
				<td align="right"><%=FormatNumber(tempData,tempRounding)%></td>
				<%  if rsHandling("shi") <>  "" then 
						tempData = cdbl(rsHandling("shi"))
				    else
					    tempData = 0
					end if
				%>
				<td align="right" ><%=FormatNumber(tempData,tempRounding)%></td>
				<%  if rsHandling("sbi") <>  "" then 
						tempData = cdbl(rsHandling("sbi"))
				    else
					    tempData = 0
					end if
				%>
				<td align="right" ><%=FormatNumber(tempData,tempRounding)%></td>
				
				<td align="right" ><%=FormatNumber(tempDataBudget_HI,tempRounding)%></td>
				<td align="right" ><%=FormatNumber(tempDataBudget_SHI,tempRounding)%></td>
				<td align="right" ><%=FormatNumber(tempDataBudget_SBI,tempRounding)%></td>
				
<%
                    if tempBudgetType = "QUANTITY" then  
					    if tempCompare_HI > 95 then
					        cClassAssigned_HI = "green"
					    else
					        if tempCompare_HI > 85 then 
					             cClassAssigned_HI = "yellow"
						    else
						         cClassAssigned_HI = "red"
						    end if
					    end if
					
					    if tempCompare_SHI > 95 then
					        cClassAssigned_SHI = "green"
					    else
					        if tempCompare_SHI > 85 then 
					             cClassAssigned_SHI = "yellow"
						    else
						         cClassAssigned_SHI = "red"
						    end if
					    end if
					
					    if tempCompare_SBI > 95 then
					        cClassAssigned_SBI = "green"
					    else
					        if tempCompare_SBI > 85 then 
					             cClassAssigned_SBI = "yellow"
						    else
						         cClassAssigned_SBI = "red"
						    end if
					    end if
                    else
					    if tempCompare_HI > 105 then
					        cClassAssigned_HI = "red"
					    else
					        if tempCompare_HI > 95 then 
					             cClassAssigned_HI = "yellow"
						    else
						         cClassAssigned_HI = "green"
						    end if
					    end if
					
					    if tempCompare_SHI > 105 then
					        cClassAssigned_SHI = "red"
					    else
					        if tempCompare_SHI > 95 then 
					             cClassAssigned_SHI = "yellow"
						    else
						         cClassAssigned_SHI = "green"
						    end if
					    end if
					
					    if tempCompare_SBI > 105 then
					        cClassAssigned_SBI = "red"
					    else
					        if tempCompare_SBI > 95 then 
					             cClassAssigned_SBI = "yellow"
						    else
						         cClassAssigned_SBI = "green"
						    end if
					    end if
                    end if					
				%>					

				
				
				<td class="<%=cClassAssigned_HI%>" align="right" ><%=FormatNumber(tempCompare_HI,2)%></td>
				<td class="<%=cClassAssigned_SHI%>" align="right" ><%=FormatNumber(tempCompare_SHI,2)%></td>
				<td class="<%=cClassAssigned_SBI%>" align="right" ><%=FormatNumber(tempCompare_SBI,2)%></td>
			
				
			</tr>
<%

			rsHandling.MoveNext
		Wend
	end if
	rsHandling.Close
	set rsHandling = Nothing
%>








<!-- Start HandlingAmount -->	
<%
	if not rsHandlingAmount.EOF then
        rsHandlingAmount.MoveLast
		tempRowSpan = rsHandlingAmount("theNumber")
		rsHandlingAmount.MoveFirst
		no = 0
		while not rsHandlingAmount.EOF
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
			no = no + 1
%>


			<tr class="<%=cClass%>">
	
			    <%
					strCodeDesc = "SELECT param_1 CodeDescription, rate_1 theRowSpan FROM PARAMETERDETAIL " & _
									" WHERE " & _
									" parcode = 'SR003' AND " & _
									" parheadcode = 'REPORTCODE' AND " & _
									" param_2 = 'DEFAULT' AND " & _
									" fccode = '" & rsHandlingAmount("codedesc") & "'" 
					
				 	Set rsCodeDesc = Server.CreateObject("ADODB.Recordset")	
					rsCodeDesc.Open strCodeDesc, connstr, adOpenStatic 		

					if not rsCodeDesc.EOF then
						tempCodeDescription = rsCodeDesc ("CodeDescription")
						'tempRowSpan = cdbl(rsCodeDesc ("theRowSpan"))
                    else
                        tempCodeDescription = ""					
						'tempRowSpan = 1
					end if					
									
									
                %>
				
				<% if no = 1 then %>
					<% if sPrevDesc <> tempCodeDescription then %>				
						<td class="boldfont" rowspan="<%=tempRowSpan%>" align="center" ><%=tempCodeDescription%></td>
						<% sPrevDesc = tempCodeDescription %>
					<% else %>
						<td align="left" ></td>
					<% end if %>
				<% end if %>
	
				<td align="center" ><%=rsHandlingAmount("theNumber")%></td>
				
				<% 
				    tempRefCode = rsHandlingAmount("codedesc")
				    strHRef = " SELECT param_5 reference, rate_3 rounding, param_6 UOM  FROM PARAMETERDETAIL " & _
								" WHERE " & _
								" parcode = 'SR003' AND " & _
								" fccode = '" & tempRefCode & "' AND " & _
								" param_2 = 'DEFAULT' " 
				
				 	Set rsHRef = Server.CreateObject("ADODB.Recordset")	
					rsHRef.Open strHRef, connstr, adOpenStatic 
					
					if not rsHRef.EOF then
						tempReference = rsHRef ("reference")
						tempUOM = rsHRef("UOM")						
						tempRounding = cdbl(rsHRef("rounding"))						
                    else
                        tempReference = ""					
						tempUOM = ""
						tempRounding = 0
					end if
					
					rsHRef.Close
					set rsHRef = Nothing
					
					tempExtendedReference = "txtfcba=" & v_txtfcba & "&txtdatefrom=" & v_txtdatefrom & "&txtdateto=" & v_txtdateto & _
								            "&txtDivision=--SUMMARY--&txtBlok=--SUMMARY--&txtPlantingDate=--SUMMARY--" & _
											"&txtGroup=&txtStatusTmTbm=--SUMMARY--&txtJobCode=--SUMMARY--&txtComp=" & v_txtComp & ""
					
					tempReference = tempReference & tempExtendedReference					
					
				%>
				<td align="left" ><a href="<%=tempReference%>"><%=rsHandlingAmount("description")%></a></td>
				<td align="center" ><%=tempUOM%></td>
				
				<% ' Start Compare to Budget 

				    strBudget = " select param_1 tempBudgetType " & _
                                  " from parameterdetail " & _
                                  " where " & _
                                  " parcode ='BUDGET' and  " & _
                                  " parheadcode = 'BUDGET_BIAYA' and  " & _
                                  " param_2 = 'DEFAULT' and  " & _
                                  " fccode = '" & rsHandlingAmount("groupCode") & "'"
								  
					Set rsBudget = Server.CreateObject("ADODB.Recordset")	
					rsBudget.Open strBudget, connstr, adOpenStatic 
					
					if not rsBudget.EOF then
					    tempBudgetType = rsBudget ("tempBudgetType")
						if rsHandlingAmount ("budget_hi") <> "" then 
							tempDataBudget_HI = cdbl(rsHandlingAmount ("budget_hi"))
						else
							tempDataBudget_HI = 0
						end if
						
						if rsHandlingAmount ("budget_shi") <> "" then 
							tempDataBudget_SHI = cdbl( rsHandlingAmount ("budget_shi"))
						else
							tempDataBudget_SHI = 0
						end if
						
						if rsHandlingAmount ("budget_sbi") <> "" then 
							tempDataBudget_SBI = cdbl( rsHandlingAmount ("budget_sbi"))
						else
							tempDataBudget_SBI = 0
						end if
                    else
						tempDataBudget_HI = 0
						tempDataBudget_SHI = 0
						tempDataBudget_SBI = 0
						tempBudgetType = "QUANTITY"
					end if
					
					rsBudget.Close
					set rsBudget = Nothing			  
	

					if tempDataBudget_HI <> 0 then 
					    tempCompare_HI = cdbl(rsHandlingAmount("hi"))/tempDataBudget_HI * 100
					else
					    tempCompare_HI = 0
					end if
					
					if tempDataBudget_SHI <> 0 then 
					    tempCompare_SHI = cdbl(rsHandlingAmount("shi"))/tempDataBudget_SHI * 100
					else
					    tempCompare_SHI = 0
					end if

					if tempDataBudget_SBI <> 0 then 
					    tempCompare_SBI = cdbl(rsHandlingAmount("sbi"))/tempDataBudget_SBI * 100
					else
					    tempCompare_SBI = 0
					end if
					
				%>				
				
				<%  if rsHandlingAmount("hi") <>  "" then 
						tempData = cdbl(rsHandlingAmount("hi"))
				    else
					    tempData = 0
					end if
				%>
				<td align="right" ><%=FormatNumber(tempData,tempRounding)%></td>

				<%  if rsHandlingAmount("shi") <>  "" then 
						tempData = cdbl(rsHandlingAmount("shi"))
				    else
					    tempData = 0
					end if
				%>
				<td align="right" ><%=FormatNumber(tempData,tempRounding)%></td>

				<%  if rsHandlingAmount("sbi") <>  "" then 
						tempData = cdbl(rsHandlingAmount("sbi"))
				    else
					    tempData = 0
					end if
				%>
				<td align="right" ><%=FormatNumber(tempData,tempRounding)%></td>
				
				<td align="right" ><%=FormatNumber(tempDataBudget_HI,tempRounding)%></td>
				<td align="right" ><%=FormatNumber(tempDataBudget_SHI,tempRounding)%></td>
				<td align="right" ><%=FormatNumber(tempDataBudget_SBI,tempRounding)%></td>
				
				
				<%
                    if tempBudgetType = "QUANTITY" then  
					    if tempCompare_HI > 95 then
					        cClassAssigned_HI = "green"
					    else
					        if tempCompare_HI > 85 then 
					             cClassAssigned_HI = "yellow"
						    else
						         cClassAssigned_HI = "red"
						    end if
					    end if
					
					    if tempCompare_SHI > 95 then
					        cClassAssigned_SHI = "green"
					    else
					        if tempCompare_SHI > 85 then 
					             cClassAssigned_SHI = "yellow"
						    else
						         cClassAssigned_SHI = "red"
						    end if
					    end if
					
					    if tempCompare_SBI > 95 then
					        cClassAssigned_SBI = "green"
					    else
					        if tempCompare_SBI > 85 then 
					             cClassAssigned_SBI = "yellow"
						    else
						         cClassAssigned_SBI = "red"
						    end if
					    end if
                    else
					    if tempCompare_HI > 105 then
					        cClassAssigned_HI = "red"
					    else
					        if tempCompare_HI > 95 then 
					             cClassAssigned_HI = "yellow"
						    else
						         cClassAssigned_HI = "green"
						    end if
					    end if
					
					    if tempCompare_SHI > 105 then
					        cClassAssigned_SHI = "red"
					    else
					        if tempCompare_SHI > 95 then 
					             cClassAssigned_SHI = "yellow"
						    else
						         cClassAssigned_SHI = "green"
						    end if
					    end if
					
					    if tempCompare_SBI > 105 then
					        cClassAssigned_SBI = "red"
					    else
					        if tempCompare_SBI > 95 then 
					             cClassAssigned_SBI = "yellow"
						    else
						         cClassAssigned_SBI = "green"
						    end if
					    end if
                    end if					
				%>					

		
				
				<td class="<%=cClassAssigned_HI%>" align="right" ><%=FormatNumber(tempCompare_HI,2)%></td>
				<td class="<%=cClassAssigned_SHI%>" align="right" ><%=FormatNumber(tempCompare_SHI,2)%></td>
				<td class="<%=cClassAssigned_SBI%>" align="right" ><%=FormatNumber(tempCompare_SBI,2)%></td>
				
				
			</tr>
<%

			rsHandlingAmount.MoveNext
		Wend
	end if
	rsHandlingAmount.Close
	set rsHandlingAmount = Nothing
%>





	

<tr>
<th class="separator" colspan="13" align="left"  ></th>				
</tr>


<!-- Start Fertilizer -->	
<%
	if not rsFertilizer.EOF then
		no = 0
		rsFertilizer.MoveLast
		tempRowSpan = rsFertilizer("theNumber")
		rsFertilizer.MoveFirst
		
		while not rsFertilizer.EOF
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
			no = no + 1
%>


			<tr class="<%=cClass%>">
			
			     
			    <%
					strCodeDesc = "SELECT param_1 CodeDescription, rate_1 theRowSpan FROM PARAMETERDETAIL " & _
									" WHERE " & _
									" parcode = 'SR003' AND " & _
									" parheadcode = 'REPORTCODE' AND " & _
									" param_2 = 'DEFAULT' AND " & _
									" fccode = '" & rsFertilizer("codedesc") & "'" 
					
				 	Set rsCodeDesc = Server.CreateObject("ADODB.Recordset")	
					rsCodeDesc.Open strCodeDesc, connstr, adOpenStatic 		

					if not rsCodeDesc.EOF then
					    'tempRowSpan =  cdbl(rsCodeDesc ("theRowSpan"))
						tempCodeDescription = rsCodeDesc ("CodeDescription")
                    else
                        tempCodeDescription = ""	
					    'tempRowSpan =  1						
					end if					
									
									
                %>
				
				<% if no = 1 then %>
					<% if sPrevDesc <> tempCodeDescription then %>				
						<td class="boldfont" rowspan="<%=tempRowSpan%>" align="center" ><%=tempCodeDescription%></td>
						<% sPrevDesc = tempCodeDescription %>
					<% else %>
						<td align="left" ></td>
					<% end if %>
				<% end if %>
			
				<td align="center"  ><%=rsFertilizer("theNumber")%></td>
				
				<% 
				    tempRefCode = rsFertilizer("codedesc")
				    strHRef = " SELECT param_5 reference FROM PARAMETERDETAIL " & _
								" WHERE " & _
								" parcode = 'SR003' AND " & _
								" fccode = '" & tempRefCode & "' AND " & _
								" param_2 = 'DEFAULT' " 
				
				 	Set rsHRef = Server.CreateObject("ADODB.Recordset")	
					rsHRef.Open strHRef, connstr, adOpenStatic 
					
					if not rsHRef.EOF then
						tempReference = rsHRef ("reference")
                    else
                        tempReference = ""					
					end if
					tempUOM = rsFertilizer("UnitOfMeasurement")
					
					rsHRef.Close
					set rsHRef = Nothing
					tempExtendedReference = "txtfcba=" & v_txtfcba & "&txttodaydate=" & v_txtdatefrom & "&txtStockGroup=&txtStock=--SUMMARY--&txtComp=" & v_txtComp & ""
					tempReference = "p_StockGudangMonthly_BA_NAF.asp?" & tempExtendedReference
				%>
				<td align="left" ><a href="<%=tempReference%>"><%=rsFertilizer("description")%></a></td>
				<td align="center" ><%=tempUOM%></td>		


				<% 
				    strHRef = " SELECT nvl(rate_3,0) rounding  FROM PARAMETER " & _
								" WHERE " & _
								" parcode = 'SR003_PUPDET' AND " & _
								" fccode = '" & rsFertilizer("StockGenCode") & "'"
				
				 	Set rsHRef = Server.CreateObject("ADODB.Recordset")	
					rsHRef.Open strHRef, connstr, adOpenStatic 
					
					if not rsHRef.EOF then
						if rsHRef ("rounding") <> "" then 
							tempRounding = cdbl(rsHRef ("rounding"))
						else
						    tempRounding = 0
						end if
                    else
                        tempRounding = 0				
					end if

					rsHRef.Close
					set rsHRef = Nothing
				%>				
				
				<% ' Start Compare to Budget 

	                strBudget = " select param_1 tempBudgetType " & _
                                  " from parameterdetail " & _
                                  " where " & _
                                  " parcode ='BUDGET' and  " & _
                                  " parheadcode = 'BUDGET_PUPUK' and  " & _
                                  " param_2 = 'DEFAULT' and  " & _
                                  " fccode = '" & rsFertilizer("stockGenCode") & "'"
					Set rsBudget = Server.CreateObject("ADODB.Recordset")	
					rsBudget.Open strBudget, connstr, adOpenStatic 
					
					if not rsBudget.EOF then
					    tempBudgetType = rsBudget ("tempBudgetType")
						tempDataBudget_HI = cdbl(rsFertilizer ("budget_hi"))
						tempDataBudget_SHI = cdbl( rsFertilizer ("budget_shi"))
						tempDataBudget_SBI = cdbl( rsFertilizer ("budget_sbi"))
                    else
						tempDataBudget_HI = 0
						tempDataBudget_SHI = 0
						tempDataBudget_SBI = 0
						tempBudgetType = "QUANTITY"
					end if
					
					rsBudget.Close
					set rsBudget = Nothing			  	

					if tempDataBudget_HI <> 0 then 
					    tempCompare_HI = cdbl(rsFertilizer("hi"))/tempDataBudget_HI * 100
					else
					    tempCompare_HI = 0
					end if
					
					if tempDataBudget_SHI <> 0 then 
					    tempCompare_SHI = cdbl(rsFertilizer("shi"))/tempDataBudget_SHI * 100
					else
					    tempCompare_SHI = 0
					end if

					if tempDataBudget_SBI <> 0 then 
					    tempCompare_SBI = cdbl(rsFertilizer("sbi"))/tempDataBudget_SBI * 100
					else
					    tempCompare_SBI = 0
					end if
					
				%>					
				
				<%  if rsFertilizer("hi") <>  "" then 
						tempData = cdbl(rsFertilizer("hi"))
				    else
					    tempData = 0
					end if
				%>
				<td align="right" ><%=FormatNumber(tempData,tempRounding)%></td>
				<%  if rsFertilizer("shi") <>  "" then 
						tempData = cdbl(rsFertilizer("shi"))
				    else
					    tempData = 0
					end if
				%>
				<td align="right" ><%=FormatNumber(tempData,tempRounding)%></td>
				<%  if rsFertilizer("sbi") <>  "" then 
						tempData = cdbl(rsFertilizer("sbi"))
				    else
					    tempData = 0
					end if
				%>
				<td align="right" ><%=FormatNumber(tempData,tempRounding)%></td>
				
				<td align="right" ><%=FormatNumber(tempDataBudget_HI,tempRounding)%></td>
				<td align="right" ><%=FormatNumber(tempDataBudget_SHI,tempRounding)%></td>
				<td align="right" ><%=FormatNumber(tempDataBudget_SBI,tempRounding)%></td>

				
				<%
                    if tempBudgetType = "QUANTITY" then  
					    if tempCompare_HI > 95 then
					        cClassAssigned_HI = "green"
					    else
					        if tempCompare_HI > 85 then 
					             cClassAssigned_HI = "yellow"
						    else
						         cClassAssigned_HI = "red"
						    end if
					    end if
					
					    if tempCompare_SHI > 95 then
					        cClassAssigned_SHI = "green"
					    else
					        if tempCompare_SHI > 85 then 
					             cClassAssigned_SHI = "yellow"
						    else
						         cClassAssigned_SHI = "red"
						    end if
					    end if
					
					    if tempCompare_SBI > 95 then
					        cClassAssigned_SBI = "green"
					    else
					        if tempCompare_SBI > 85 then 
					             cClassAssigned_SBI = "yellow"
						    else
						         cClassAssigned_SBI = "red"
						    end if
					    end if
                    else
					    if tempCompare_HI > 105 then
					        cClassAssigned_HI = "red"
					    else
					        if tempCompare_HI > 95 then 
					             cClassAssigned_HI = "yellow"
						    else
						         cClassAssigned_HI = "green"
						    end if
					    end if
					
					    if tempCompare_SHI > 105 then
					        cClassAssigned_SHI = "red"
					    else
					        if tempCompare_SHI > 95 then 
					             cClassAssigned_SHI = "yellow"
						    else
						         cClassAssigned_SHI = "green"
						    end if
					    end if
					
					    if tempCompare_SBI > 105 then
					        cClassAssigned_SBI = "red"
					    else
					        if tempCompare_SBI > 95 then 
					             cClassAssigned_SBI = "yellow"
						    else
						         cClassAssigned_SBI = "green"
						    end if
					    end if
                    end if					
				%>					
				
				

				
				<td class="<%=cClassAssigned_HI%>" align="right" ><%=FormatNumber(tempCompare_HI,2)%></td>
				<td class="<%=cClassAssigned_SHI%>" align="right" ><%=FormatNumber(tempCompare_SHI,2)%></td>
				<td class="<%=cClassAssigned_SBI%>" align="right" ><%=FormatNumber(tempCompare_SBI,2)%></td>
				
			</tr>
<%

			rsFertilizer.MoveNext
		Wend
	end if
	rsFertilizer.Close
	set rsFertilizer = Nothing
%>












<tr>
<th class="separator" colspan="13" align="left"  ></th>				
</tr>


<!-- Start Tenaga Kerja -->	
<%
	if not rsMandays.EOF then
	    rsMandays.MoveLast
		tempRowSpan = rsMandays("theNumber")
		rsMandays.MoveFirst
		no = 0
		while not rsMandays.EOF
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
			no = no + 1
%>


			<tr class="<%=cClass%>">
				<%
				strCodeDesc = "SELECT param_1 CodeDescription, rate_1 theRowSpan, nvl(rate_10,0) budget_hi, nvl(rate_11,0) budget_shi, nvl(rate_12,0) budget_sbi, param_7 tempBudgetType FROM PARAMETERDETAIL " & _
									" WHERE " & _
									" parcode = 'SR003' AND " & _
									" parheadcode = 'REPORTCODE' AND " & _
									" param_2 = 'DEFAULT' AND " & _
									" fccode = '" & rsMandays("codedesc") & "'" 
					
				 	Set rsCodeDesc = Server.CreateObject("ADODB.Recordset")	
					rsCodeDesc.Open strCodeDesc, connstr, adOpenStatic 		

					if not rsCodeDesc.EOF then
						tempCodeDescription = rsCodeDesc ("CodeDescription")
						'tempRowSpan = cdbl(rsCodeDesc ("theRowSpan"))
						tempDataBudget_HI = cdbl(rsMandays ("budget_hi"))
						tempDataBudget_SHI = cdbl(rsMandays ("budget_shi"))
						tempDataBudget_SBI = cdbl(rsMandays ("budget_sbi"))	
						tempBudgetType = rsCodeDesc ("tempBudgetType")
                    else
						'tempRowSpan = 1
                        tempCodeDescription = ""					
						tempDataBudget_HI = 0
						tempDataBudget_SHI = 0
						tempDataBudget_SBI = 0	
						tempBudgetType = "QUANTITY"
					end if		
					
									
									
                %>
				
				<% if no = 1 then %>
					<% if sPrevDesc <> tempCodeDescription then %>				
						<td class="boldfont" rowspan="<%= tempRowSpan %>" align="center" ><%=tempCodeDescription%></td>
						<% sPrevDesc = tempCodeDescription %>
					<% else %>
						<td align="left"></td>
					<% end if %>
				<% end if %>
			
				<td align="center" ><%=rsMandays("theNumber")%></td>
				
				<% 
				    tempRefCode = rsMandays("codedesc")
				    strHRef = " SELECT param_5 reference, nvl(rate_3,0) rounding, param_6 UOM FROM PARAMETERDETAIL " & _
								" WHERE " & _
								" parcode = 'SR003' AND " & _
								" fccode = '" & tempRefCode & "' AND " & _
								" param_2 = 'DEFAULT' " 
				
				 	Set rsHRef = Server.CreateObject("ADODB.Recordset")	
					rsHRef.Open strHRef, connstr, adOpenStatic 
					
					if not rsHRef.EOF then
						tempReference = rsHRef ("reference")
						tempUOM = rsHRef ("UOM")
						tempRounding = cdbl(rsHRef ("rounding"))
                    else
                        tempReference = ""					
						tempUOM = ""
						tempRounding = 0
					end if
					
					rsHRef.Close
					set rsHRef = Nothing
					
					tempExtendedReference = "txtDate=" & v_txtdatefrom & "&txtfcba=" & v_txtfcba & "&txtComp=" & v_txtComp & ""
					
					tempReference = tempReference & tempExtendedReference	
					
				%>
				<td align="left" ><a href="<%=tempReference%>"><%=rsMandays("description")%></a></td>
				<td align="center" ><%=tempUOM%></td>		

				<% ' Start Compare to Budget 



					if tempDataBudget_HI <> 0 then 
					    tempCompare_HI = cdbl(rsMandays("hi"))/tempDataBudget_HI * 100
					else
					    tempCompare_HI = 0
					end if
					
					if tempDataBudget_SHI <> 0 then 
					    tempCompare_SHI = cdbl(rsMandays("shi"))/tempDataBudget_SHI * 100
					else
					    tempCompare_SHI = 0
					end if

					if tempDataBudget_SBI <> 0 then 
					    tempCompare_SBI = cdbl(rsMandays("sbi"))/tempDataBudget_SBI * 100
					else
					    tempCompare_SBI = 0
					end if
					
				%>					
				
				<%  if rsMandays("hi") <>  "" then 
						tempData = cdbl(rsMandays("hi"))
				    else
					    tempData = 0
					end if
				%>
				<td align="right" ><%=FormatNumber(tempData,tempRounding)%></td>
				<%  if rsMandays("shi") <>  "" then 
						tempData = cdbl(rsMandays("shi"))
				    else
					    tempData = 0
					end if
				%>
				<td align="right" ><%=FormatNumber(tempData,tempRounding)%></td>
				<%  if rsMandays("sbi") <>  "" then 
						tempData = cdbl(rsMandays("sbi"))
				    else
					    tempData = 0
					end if
				%>
				<td align="right" ><%=FormatNumber(tempData,tempRounding)%></td>
				
				<td align="right" ><%=FormatNumber(tempDataBudget_HI,tempRounding)%></td>
				<td align="right" ><%=FormatNumber(tempDataBudget_SHI,tempRounding)%></td>
				<td align="right" ><%=FormatNumber(tempDataBudget_SBI,tempRounding)%></td>
				
				<%
                    if tempBudgetType = "QUANTITY" then  
					    if tempCompare_HI > 95 then
					        cClassAssigned_HI = "green"
					    else
					        if tempCompare_HI > 85 then 
					             cClassAssigned_HI = "yellow"
						    else
						         cClassAssigned_HI = "red"
						    end if
					    end if
					
					    if tempCompare_SHI > 95 then
					        cClassAssigned_SHI = "green"
					    else
					        if tempCompare_SHI > 85 then 
					             cClassAssigned_SHI = "yellow"
						    else
						         cClassAssigned_SHI = "red"
						    end if
					    end if
					
					    if tempCompare_SBI > 95 then
					        cClassAssigned_SBI = "green"
					    else
					        if tempCompare_SBI > 85 then 
					             cClassAssigned_SBI = "yellow"
						    else
						         cClassAssigned_SBI = "red"
						    end if
					    end if
                    else
					    if tempCompare_HI > 105 then
					        cClassAssigned_HI = "red"
					    else
					        if tempCompare_HI > 95 then 
					             cClassAssigned_HI = "yellow"
						    else
						         cClassAssigned_HI = "green"
						    end if
					    end if
					
					    if tempCompare_SHI > 105 then
					        cClassAssigned_SHI = "red"
					    else
					        if tempCompare_SHI > 95 then 
					             cClassAssigned_SHI = "yellow"
						    else
						         cClassAssigned_SHI = "green"
						    end if
					    end if
					
					    if tempCompare_SBI > 105 then
					        cClassAssigned_SBI = "red"
					    else
					        if tempCompare_SBI > 95 then 
					             cClassAssigned_SBI = "yellow"
						    else
						         cClassAssigned_SBI = "green"
						    end if
					    end if
                    end if					
				%>					
				
				
				<td class="<%=cClassAssigned_HI%>" align="right" ><%=FormatNumber(tempCompare_HI,2)%></td>
				<td class="<%=cClassAssigned_SHI%>" align="right" ><%=FormatNumber(tempCompare_SHI,2)%></td>
				<td class="<%=cClassAssigned_SBI%>" align="right" ><%=FormatNumber(tempCompare_SBI,2)%></td>

				
			</tr>
<%

			rsMandays.MoveNext
		Wend
	end if
	rsMandays.Close
	set rsMandays = Nothing
%>



	




<tr>
<th class="separator" colspan="13" align="left"  ></th>				
</tr>


<!-- Start Vehicle Group -->	
<%
	if not rsVehicleGroup.EOF then
        rsVehicleGroup.MoveLast
		tempRowSpan = rsVehicleGroup("theNumber")
		rsVehicleGroup.MoveFirst
		no = 0
		while not rsVehicleGroup.EOF
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
			no = no + 1
%>


			<tr class="<%=cClass%>">
				<%
				strCodeDesc = "SELECT param_1 CodeDescription, rate_1 theRowSpan FROM PARAMETERDETAIL " & _
									" WHERE " & _
									" parcode = 'SR003' AND " & _
									" parheadcode = 'REPORTCODE' AND " & _
									" param_2 = 'DEFAULT' AND " & _
									" fccode = '" & rsVehicleGroup("codedesc") & "'" 
					
				 	Set rsCodeDesc = Server.CreateObject("ADODB.Recordset")	
					rsCodeDesc.Open strCodeDesc, connstr, adOpenStatic 		

					if not rsCodeDesc.EOF then
						tempCodeDescription = rsCodeDesc ("CodeDescription")
						'tempRowSpan = cdbl(rsCodeDesc ("theRowSpan"))
                    else
                        tempCodeDescription = ""					
						'tempRowSpan = 1
					end if					
					

									
                %>
				
				<% if no = 1 then %>
					<% if sPrevDesc <> tempCodeDescription then %>				
						<td class="boldfont" rowspan="<%= tempRowSpan %>" align="center" ><%=tempCodeDescription%></td>
						<% sPrevDesc = tempCodeDescription %>
					<% else %>
						<td align="left" ></td>
					<% end if %>
				<% end if %>
			
				<td align="center" ><%=rsVehicleGroup("theNumber")%></td>
				
				<% 
				    tempRefCode = rsVehicleGroup("codedesc")
				    strHRef = " SELECT param_5 reference, nvl(rate_3,0) rounding , param_6 UOM FROM PARAMETERDETAIL " & _
								" WHERE " & _
								" parcode = 'SR003' AND " & _
								" fccode = '" & tempRefCode & "' AND " & _
								" param_2 = 'DEFAULT' " 
				
				 	Set rsHRef = Server.CreateObject("ADODB.Recordset")	
					rsHRef.Open strHRef, connstr, adOpenStatic 
					
					if not rsHRef.EOF then
						tempReference = rsHRef ("reference")
						tempUOM = rsHRef ("UOM")
						tempRounding = cdbl(rsHRef ("rounding"))
                    else
                        tempReference = ""					
						tempUOM = ""
						tempRounding = 0
					end if
					
					rsHRef.Close
					set rsHRef = Nothing
					

					
					
					tempExtendedReference = "txtdatefrom=" & v_txtdatefrom & "&txtdateto=" & v_txtdateto & "&txtVehicleGroup=&txtVehicleCode=--SUMMARY--&txtfcba=" & v_txtfcba & "&txtComp=" & v_txtComp & ""
					
					if tempReference <> "" then 
						tempReference = tempReference & tempExtendedReference	
					else
					    tempReference = "p_DailyVehicleSolar_BA_NAF.asp?" & tempExtendedReference	
					end if
					
				%>
				<td align="left" ><a href="<%=tempReference%>"><%=rsVehicleGroup("description")%></a></td>
				<td align="center" ><%=tempUOM%></td>				
				
				<% ' Start Compare to Budget 

					tempDataBudget_HI = cdbl(rsVehicleGroup("budget_hi"))
					tempDataBudget_SHI = cdbl(rsVehicleGroup("budget_shi"))
					tempDataBudget_SBI = cdbl(rsVehicleGroup("budget_sbi"))	

					if tempDataBudget_HI <> 0 then 
					    tempCompare_HI = cdbl(rsVehicleGroup("hi"))/tempDataBudget_HI * 100
					else
					    tempCompare_HI = 0
					end if
					
					if tempDataBudget_SHI <> 0 then 
					    tempCompare_SHI = cdbl(rsVehicleGroup("shi"))/tempDataBudget_SHI * 100
					else
					    tempCompare_SHI = 0
					end if

					if tempDataBudget_SBI <> 0 then 
					    tempCompare_SBI = cdbl(rsVehicleGroup("sbi"))/tempDataBudget_SBI * 100
					else
					    tempCompare_SBI = 0
					end if
					
				%>					
			
				<%  if rsVehicleGroup("hi") <>  "" then 
						tempData = cdbl(rsVehicleGroup("hi"))
				    else
					    tempData = 0
					end if
				%>
				<td align="right" ><%=FormatNumber(tempData,tempRounding)%></td>
				<%  if rsVehicleGroup("shi") <>  "" then 
						tempData = cdbl(rsVehicleGroup("shi"))
				    else
					    tempData = 0
					end if
				%>
				<td align="right"><%=FormatNumber(tempData,tempRounding)%></td>
				<%  if rsVehicleGroup("sbi") <>  "" then 
						tempData = cdbl(rsVehicleGroup("sbi"))
				    else
					    tempData = 0
					end if
				%>
				<td align="right" ><%=FormatNumber(tempData,tempRounding)%></td>
				
				<td align="right" ><%=FormatNumber(tempDataBudget_HI,tempRounding)%></td>
				<td align="right" ><%=FormatNumber(tempDataBudget_SHI,tempRounding)%></td>
				<td align="right" ><%=FormatNumber(tempDataBudget_SBI,tempRounding)%></td>

				<%
					if tempCompare_HI > 95 then
					    cClassAssigned_HI = "green"
					else
					    if tempCompare_HI > 85 then 
					         cClassAssigned_HI = "yellow"
						else
						     cClassAssigned_HI = "red"
						end if
					end if
					
					if tempCompare_SHI > 95 then
					    cClassAssigned_SHI = "green"
					else
					    if tempCompare_SHI > 85 then 
					         cClassAssigned_SHI = "yellow"
						else
						     cClassAssigned_SHI = "red"
						end if
					end if
					
					if tempCompare_SBI > 95 then
					    cClassAssigned_SBI = "green"
					else
					    if tempCompare_SBI > 85 then 
					         cClassAssigned_SBI = "yellow"
						else
						     cClassAssigned_SBI = "red"
						end if
					end if
					
				%>
				
				<td class="<%=cClassAssigned_HI%>" align="right" ><%=FormatNumber(tempCompare_HI,2)%></td>
				<td class="<%=cClassAssigned_SHI%>" align="right" ><%=FormatNumber(tempCompare_SHI,2)%></td>
				<td class="<%=cClassAssigned_SBI%>" align="right" ><%=FormatNumber(tempCompare_SBI,2)%></td>
				
				
			</tr>
<%

			rsVehicleGroup.MoveNext
		Wend
	end if
	rsVehicleGroup.Close
	set rsVehicleGroup = Nothing
%>




<tr>
<th class="separator" colspan="13" align="left"  ></th>				
</tr>



<!-- Start Ending Balance -->	
<%
	if not rsEndingBalance.EOF then
        rsEndingBalance.MoveLast
		tempRowSpan = rsEndingBalance("theNumber")
		rsEndingBalance.MoveFirst
		no = 0
		while not rsEndingBalance.EOF
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
			no = no + 1
%>


			<tr class="<%=cClass%>">
			
				<%
				strCodeDesc = "SELECT param_1 CodeDescription, rate_1 theRowSpan FROM PARAMETERDETAIL " & _
									" WHERE " & _
									" parcode = 'SR003' AND " & _
									" parheadcode = 'REPORTCODE' AND " & _
									" param_2 = 'DEFAULT' AND " & _
									" fccode = '" & rsEndingBalance("codedesc") & "'" 
					
				 	Set rsCodeDesc = Server.CreateObject("ADODB.Recordset")	
					rsCodeDesc.Open strCodeDesc, connstr, adOpenStatic 		

					if not rsCodeDesc.EOF then
						tempCodeDescription = rsCodeDesc ("CodeDescription")
						'tempRowSpan = cdbl(rsCodeDesc ("theRowSpan"))
                    else
						'tempRowSpan = 1
                        tempCodeDescription = ""				
						
					end if		
					tempUOM = rsEndingBalance("UnitOfMeasurement")
									
									
                %>
				
				<% if no = 1 then %>
					<% if sPrevDesc <> tempCodeDescription then %>				
						<td class="boldfont" rowspan="<%=tempRowSpan%>" align="center" ><%=tempCodeDescription%></td>
						<% sPrevDesc = tempCodeDescription %>
					<% else %>
						<td align="left"></td>
					<% end if %>
				<% end if %>
				
				<td align="center" ><%=rsEndingBalance("theNumber")%></td>
				
				<% 
				    tempRefCode = rsEndingBalance("codedesc")
				    strHRef = " SELECT param_5 reference FROM PARAMETERDETAIL " & _
								" WHERE " & _
								" parcode = 'SR003' AND " & _
								" fccode = '" & tempRefCode & "' AND " & _
								" param_2 = 'DEFAULT' " 
				
				 	Set rsHRef = Server.CreateObject("ADODB.Recordset")	
					rsHRef.Open strHRef, connstr, adOpenStatic 
					
					if not rsHRef.EOF then
						tempReference = rsHRef ("reference")
                    else
                        tempReference = ""					
					end if
					tempUOM = rsEndingBalance("unitofmeasurement")
					
					rsHRef.Close
					set rsHRef = Nothing
					
					tempExtendedReference = "txtfcba=" & v_txtfcba & "&txttodaydate=" & v_txtdatefrom & "&txtStockGroup=&txtStock=--SUMMARY--&txtComp=" & v_txtComp & ""
					tempReference = tempReference & tempExtendedReference
				%>
				<td align="left" ><a href="<%=tempReference%>"><%=rsEndingBalance("description")%></a></td>
				<td align="center" ><%=tempUOM%></td>				
				

				<% ' Start Compare to Budget 

					tempDataBudget_HI = 0
					tempDataBudget_SHI = 0
					tempDataBudget_SBI = 0	

					if tempDataBudget_HI <> 0 then 
					    tempCompare_HI = cdbl(rsEndingBalance("hi"))/tempDataBudget_HI * 100
					else
					    tempCompare_HI = 0
					end if
					
					if tempDataBudget_SHI <> 0 then 
					    tempCompare_SHI = cdbl(rsEndingBalance("shi"))/tempDataBudget_SHI * 100
					else
					    tempCompare_SHI = 0
					end if

					if tempDataBudget_SBI <> 0 then 
					    tempCompare_SBI = cdbl(rsEndingBalance("sbi"))/tempDataBudget_SBI * 100
					else
					    tempCompare_SBI = 0
					end if
					
				%>					
				
				<%  if rsEndingBalance("hi") <>  "" then 
						tempData = cdbl(rsEndingBalance("hi"))
				    else
					    tempData = 0
					end if
				%>
				<td align="right" ><%=FormatNumber(tempData,2)%></td>
				<%  if rsEndingBalance("shi") <>  "" then 
						tempData = cdbl(rsEndingBalance("shi"))
				    else
					    tempData = 0
					end if
				%>
				
				<td class="grey" align="right" ></td>
				<%  if rsEndingBalance("sbi") <>  "" then 
						tempData = cdbl(rsEndingBalance("sbi"))
				    else
					    tempData = 0
					end if
				%>
				<td class="grey" align="right" ></td>
				
				<td class="grey" align="right" ></td>
				<td class="grey" align="right" ></td>
				<td class="grey" align="right" ></td>

				<%
					if tempCompare_HI > 95 then
					    cClassAssigned_HI = "green"
					else
					    if tempCompare_HI > 85 then 
					         cClassAssigned_HI = "yellow"
						else
						     cClassAssigned_HI = "red"
						end if
					end if
					
					if tempCompare_SHI > 95 then
					    cClassAssigned_SHI = "green"
					else
					    if tempCompare_SHI > 85 then 
					         cClassAssigned_SHI = "yellow"
						else
						     cClassAssigned_SHI = "red"
						end if
					end if
					
					if tempCompare_SBI > 95 then
					    cClassAssigned_SBI = "green"
					else
					    if tempCompare_SBI > 85 then 
					         cClassAssigned_SBI = "yellow"
						else
						     cClassAssigned_SBI = "red"
						end if
					end if
					
				%>
				
				<!--td class="<%=cClassAssigned_HI%>" align="right" ><%=FormatNumber(tempCompare_HI,2)%></td-->
				<!--td class="<%=cClassAssigned_SHI%>" align="right" ><%=FormatNumber(tempCompare_SHI,2)%></td-->
				<!--td class="<%=cClassAssigned_SBI%>" align="right" ><%=FormatNumber(tempCompare_SBI,2)%></td-->

				<td class="grey" align="right" ></td-->
				<td class="grey" align="right" ></td-->
				<td class="grey" align="right" ></td-->
				
			</tr>
<%

			rsEndingBalance.MoveNext
		Wend
	end if
	rsEndingBalance.Close
	set rsEndingBalance = Nothing
%>
</tbody>

</table>

</form>

</div>
</body>
</html>




