<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%	  
	v_txtGroupPlantation  = Request("txtGroupPlantation")
	v_txtPlantation  	= Request("txtPlantation")
	v_txtGroupRegion  	= Request("txtGroupRegion")
	v_txtRegion  		= Request("txtRegion")
	v_txtcompany  		= Request("txtCompany")
	v_txtfcba 			= Request("txtfcba")
	v_txtDivision		= Request("txtDivision")
	v_txtBlok			= Request("txtBlok")
	v_txtPlantingDate	= Request("txtPlantingDate")
	v_txttodaydate		= Request("txttodaydate")
	v_txtYear 			= Request("txtYear")
	v_txtMonth 			= Request("txtmonth")
	v_txtAssetGroup 	= Request("txtAssetGroup")
	v_txtAsset 			= Request("txtAsset")
	
	txtyear 		= right(Request("txttodaydate"),4)
	txtmonth 		= left(Request("txttodaydate"),2)
	txttodaydate 	= Request("txttodaydate")
	
	
	txtComp         = Request("txtCompany")
	nSpanTotal = 8
	v_txtComp           =   txtComp
	
	if request("txtYear")="" then
	   v_txtYear	=	Year(Session("v_Costbook"))
	end if
	if request("txtMonth")="" then
	   v_txtMonth	=	Month(Session("v_Costbook"))
	end if
		 
		'response.write  "GP " & v_txtGroupPlantation & "  P " & v_txtPlantation & "  GR " & v_txtGroupRegion & " R " & v_txtRegion & " C " & v_txtcompany & " S " & v_txtfcba  &  " D " & v_txtDivision & " B "  & v_txtBlok & " TT" & v_txtPlantingDate
		
		Dim selectFields,strFieldCons,selectCons,whereCons,GroupCons,orderCons 
		strFieldCons = 	"   USER_COMPANY_API.GetGroupPlantationCode('" & Session("CodeName") & "', fcba) GroupPlantation" & _
						" , USER_COMPANY_API.GetPlantationCode('" & Session("CodeName") & "', fcba) Plantation" & _
						" , USER_COMPANY_API.GetGroupRegionCode('" & Session("CodeName") & "', fcba) groupregion" & _
						" , USER_COMPANY_API.GetRegionCode('" & Session("CodeName") & "', fcba) region" & _
						" , USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "', fcba) company" & _
						" , BUSINESSUNIT_API.get_OrderData(fcba) NoUrut, fcba,"

					   
		wherecons = " where company is not null "
		if v_txtGroupPlantation <> "--SUMMARY--"  and v_txtGroupPlantation <> "--ALL--" and v_txtGroupPlantation <> "" then
		   whereCons = wherecons & " And GroupPlantation = '" & v_txtGroupPlantation & "' "
		 
		end if
		if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then
		    if v_txtPlantation <> "--ALL--" then
			   whereCons = whereCons & " And  Plantation = '" & v_txtPlantation & "' "
			end if   
			selectCons = selectCons & ",Plantation"
			orderCons  = orderCons & ", Plantation "
		else
		    nSpanTotal = nSpanTotal - 1 
		end if
		if v_txtGroupRegion <> "--SUMMARY--"   and v_txtGroupRegion <> "" then
		    if v_txtGroupRegion <> "--ALL--" then
			   whereCons = whereCons & " And  GroupRegion = '" & v_txtGroupRegion & "' "
			end if
			selectCons = selectCons & " ,GroupRegion"
			orderCons  = orderCons & ", GroupRegion "
		else
		    nSpanTotal = nSpanTotal - 1 
		end if
		if  v_txtRegion <> "--SUMMARY--"  and v_txtRegion <> "" then
		    if v_txtRegion <> "--ALL--" then
			   whereCons = whereCons & " And  Region = '" & v_txtRegion & "' "
			end if
			selectCons = selectCons & " ,Region"
			orderCons  = orderCons & ", Region "
		else
		    nSpanTotal = nSpanTotal - 1 
		end if
		if  v_txtCompany <> "--SUMMARY--" and v_txtCompany <> "" then
		    if v_txtCompany  <> "--ALL--" then
			   whereCons = whereCons & " And  company = '" & v_txtCompany & "' "
			end if
			selectCons = selectCons & " ,company"
		    orderCons  = orderCons & ", COMPANY_API.get_OrderData(company) "
			q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where FCCOMPANYCODE = '" & v_txtCompany & "'"	
			set rsCat = connStr.Execute(q)
			if not rsCat.eof then
			   CompanyName=rsCat("FCCOMPANYNAME")			   
		    end if
		else
		    nSpanTotal = nSpanTotal - 1 	
		end if
        if  v_txtfcba <> "--SUMMARY--"  and v_txtfcba <> "" then
		    if v_txtfcba  <> "--ALL--" then
		       whereCons = whereCons & " And fcba = '" & v_txtfcba & "' "
			end if
		    selectCons = selectCons & ",fcba"
		    orderCons  = orderCons & ",fcba"
			GroupCons =  GroupCons & " , fcba"
			q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & v_txtfcba & "' "
			set rsCat = connStr.Execute(q)
			if not rsCat.eof then
			   fcbaname = rsCat("fcName")
			   CompanyName=rsCat("FCCOMPANYNAME")
			   MgrName = rsCat("FCMANAGER")
			   KTUName = rsCat("FCKTU")
			   CropName=rsCat("FCCROP")
			   FcbaType=rsCat("fcType")
		    end if
		else
		    nSpanTotal = nSpanTotal - 1
		end if
		
		
		
		   if v_txtDivision <> "--SUMMARY--"  And v_txtDivision <> "" then
		      if v_txtDivision <>  "--ALL--"  then
				whereCons = whereCons & " AND DIVISION = '" & v_txtDivision & "'"
		      end if
              selectCons = selectCons &  " , DIVISION "
			  orderCons  = orderCons & ",DIVISION "
		   else
		    nSpanTotal = nSpanTotal - 1 	  
		   end if   

        
		
			if v_txtPlantingDate <> "--SUMMARY--" and v_txtPlantingDate <> ""   then
			   if v_txtPlantingDate <>  "--ALL--" then
				  whereCons = whereCons & " AND PLANTINGYEAR = '" & v_txtplantingdate & "'"
			   end if
			   selectCons = selectCons &  " , PLANTINGYEAR "
			   orderCons  = orderCons & ",PLANTINGYEAR "
			else
		    nSpanTotal = nSpanTotal - 1    
			end if
			
			if v_txtBlok <> "--SUMMARY--" and v_txtBlok <> "" then
			   if v_txtBlok <>  "--ALL--" then
				  whereCons = whereCons & " AND FIELDCODE = '" & v_txtBlok & "'"
			   end if
			   selectCons = selectCons & " , FIELDCODE "
			   orderCons  = orderCons & ",FIELDCODE "
			else
		    nSpanTotal = nSpanTotal - 1    
			end if
			
			if v_txtAssetGroup <>  "--ALL--" and v_txtAssetGroup <>  "" then
				whereCons = whereCons & " AND assetgroupcode = '" & v_txtAssetGroup & "'"
				selectCons = selectCons & " , assetgroupcode, assetgroupname "
				GroupCons =  GroupCons & " , assetgroupcode, assetgroupname"
				orderCons  = orderCons & ",assetgroupcode, assetgroupname"
			else
				selectCons = selectCons & " , assetgroupcode, assetgroupname "
				GroupCons =  GroupCons & " , assetgroupcode, assetgroupname"
				orderCons  = orderCons & ", assetgroupcode, assetgroupname"
			end if
			
			
			if v_txtAsset <> "--SUMMARY--" and v_txtAsset <> "" then
			   if v_txtAsset <>  "--ALL--" then
				  whereCons = whereCons & " AND fccode = '" & v_txtAsset & "'"
			   end if
			   'selectCons = selectCons & " , assetcode, assetname, fccode, fcname, transdate,depreciationmonth_ledger,monthtobe"
			   selectCons = selectCons & " , assetcode, assetname, fccode, fcname, transdate"
			   orderCons  = orderCons & " , assetcode, assetname, fccode, fcname, transdate"
			   GroupCons =  GroupCons & " , assetcode, assetname, fccode, fcname, transdate"
			else
		    nSpanTotal = nSpanTotal - 1    
			end if
			
' Dim wherestatus				
' if "" & txtstatus <> "" and "" & txtstatus <> "--ALL--" then
' 	wherestatus = " and field_api.GetValid_status(fieldcode,fcba,TO_DATE('" & v_txttodaydate & "', 'mm/dd/yyyy') ) = '" & txtstatus & "'"
' end if
	ScriptName	=	Request.ServerVariables("SCRIPT_NAME")
	
	' strMain = 	" 		  SELECT " & strFieldCons & _
				' "                assetgroupcode," & _
				' "                assetgroupname," & _
				' "                assetcode," & _
				' "                assetname," & _
				' "                fccode," & _
				' "                fcname," & _
				' "                transdate," & _
				' "          		 depreciationmonth_ledger," &_
				' "          		 monthtobe," &_				
				' "                acquisitionvalue_prev," & _
				' "                adjustment_amount_prev," & _
				' "                fixassetamount_prev," & _
				' "                accumulated_depreciation_prev," & _
				' "                yearlydepreciation_prev," & _
				' "				 bookvalue," & _ 
				' "				 addinValue," & _
				' "                0 AS acquisitionvalue_curr," & _
				' "                adjustment_amount_curr," & _
				' "                fixassetamount_curr," & _
				' "                accumulated_depreciation_curr," & _
				' "                yearlydepreciation_curr," & _
				' " 				 (bookvalue + addinValue + adjustment_amount_curr - fixassetamount_curr - yearlydepreciation_curr + accumulated_depreciation_curr) netbookvalue," &_
				' "                monthlydepreciation_curr" & _				
				' "           FROM (SELECT fcba," &_
				' "          assetgroupcode,                                                                                                              " &_
				' "          assetgroupname,                                                                                                              " &_
				' "          assetcode,                                                                                                                   " &_
				' "          assetname,                                                                                                                   " &_
				' "          fccode,                                                                                                                      " &_
				' "          fcname,                                                                                                                      " &_
				' "          transdate,                                                                                                                   " &_
				' "          MAX(depreciationmonth_ledger)depreciationmonth_ledger,																		" &_
				' "          MAX(monthtobe)monthtobe,																										" &_
				' "          SUM (acquisitionvalue) acquisitionvalue_prev,                                                                                " &_
				' "          SUM (adjustment_amount) adjustment_amount_prev,                                                                              " &_
				' "          SUM (fixassetamount) fixassetamount_prev,                                                                                    " &_
				' "          SUM (accumulated_depreciation) accumulated_depreciation_prev,                                                                " &_
				' "          SUM (yearlydepreciation) yearlydepreciation_prev,                                                                            " &_
				' "          case TO_CHAR(transdate, 'yyyy') when '"&v_txtYear&"' then																	" &_
				' "               0 																														" &_
				' "           else																														" &_
				' "              (SUM (acquisitionvalue) + SUM (adjustment_amount)) - SUM (fixassetamount) - SUM (yearlydepreciation)						" &_
				' "          end bookvalue,																												" &_
				' "          case TO_CHAR(transdate, 'yyyy') when '"&v_txtYear&"' then																	" &_
				' "               SUM (addinValue)  																										" &_
				' "          else																															" &_
				' "              0																														" &_
				' "          end addinValue,																												" &_
				' "          SUM (adjustment_amount_cur) adjustment_amount_curr,                                                                          " &_
				' "          SUM (fixassetamount_cur) fixassetamount_curr,                                                                                " &_
				' "          SUM (accumulated_depreciation_cur) accumulated_depreciation_curr,                                                            " &_				
				' "          CASE 																														" &_
				' "          When SUM (yearlydepreciation_cur) > 0 Then																					" &_
				' "          SUM (yearlydepreciation_cur) - SUM (yearlydepreciation)																		" &_
				' "          Else																															" &_
				' "          0																															" &_
				' "          End yearlydepreciation_curr,																									" &_
				' "          SUM (monthlydepreciation) monthlydepreciation_curr                                                                           " &_
				' "     FROM (SELECT fa.fcba,                                                                                                             " &_
				' "                  fa.assetgroupcode,                                                                                                   " &_
				' "                  fa.assetgroupname,                                                                                                   " &_
				' "                  fa.assetcode,                                                                                                        " &_
				' "                  fa.assetname,                                                                                                        " &_
				' "                  fa.fccode,                                                                                                           " &_
				' "                  fa.fcname,                                                                                                           " &_
				' "                  fa.transdate,                                                                                                        " &_
				' "				   0 depreciationmonth_ledger,																							" &_
				' "                  0 monthtobe,																											" &_				
				' "                  0 acquisitionvalue,                                                                                                  " &_
				' "                  0 adjustment_amount,                                                                                                 " &_
				' "                  0 fixassetamount,                                                                                                    " &_
				' "                  0 accumulated_depreciation,                                                                                          " &_
				' "                  0 yearlydepreciation,                                                                                                " &_
				' "                  0 addinValue,                                                                                                        " &_
				' "                  0 adjustment_amount_cur,                                                                                             " &_
				' "                  0 fixassetamount_cur,                                                                                                " &_
				' "                  0 accumulated_depreciation_cur,                                                                                      " &_
				' "                  0 yearlydepreciation_cur,                                                                                            " &_
				' "                  0 monthlydepreciation                                                                                                " &_
				' "             FROM fixedasset fa                                                                                                        " &_
				' "             INNER JOIN assetgroup ag ON fa.assetgroupcode = ag.fccode                                       							" &_
				' "            WHERE                                                                                                                      " &_
				' "            TO_DATE(TO_CHAR(transdate, 'yyyymm'),'yyyyfmMM') <= TO_DATE('"&v_txtYear&v_txtMonth&"','yyyyfmMM') 						" &_				
				' "           UNION ALL                                                                                                                   " &_
				' "             SELECT fa.fcba,                                                                                                           " &_
				' "                    fa.assetgroupcode,                                                                                                 " &_
				' "                    fa.assetgroupname,                                                                                                 " &_
				' "                    fa.assetcode,                                                                                                      " &_
				' "                    fa.assetname,                                                                                                      " &_
				' "                    fa.fccode,                                                                                                         " &_
				' "                    fa.fcname,                                                                                                         " &_
				' "                    fa.transdate,                                                                                                      " &_
				' "				   0 depreciationmonth_ledger,																							" &_
				' "                  0 monthtobe,																											" &_				
				' "                    SUM (fa.acquisitionvalue) acquisitionvalue,                                                                        " &_
				' "                    0 adjustment_amount,                                                                                               " &_
				' "                    0 fixassetamount,                                                                                                  " &_
				' "                    0 accumulated_depreciation,                                                                                        " &_
				' "                    0 yearlydepreciation,                                                                                              " &_
				' "                    0 addinValue,                                                                                                      " &_
				' "                    0 adjustment_amount_cur,                                                                                           " &_
				' "                    0 fixassetamount_cur,                                                                                              " &_
				' "                    0 accumulated_depreciation_cur,                                                                                    " &_
				' "                    0 yearlydepreciation_cur,                                                                                          " &_
				' "                    0 monthlydepreciation                                                                                              " &_
				' "               FROM fixedasset fa                                                                                                      " &_
				' "               INNER JOIN assetgroup ag ON fa.assetgroupcode = ag.fccode                                								" &_
				' "              WHERE                                                                                                                    " &_
				' "              TO_DATE(TO_CHAR(transdate, 'yyyymm'),'yyyyfmMM') < TO_DATE('"&v_txtYear&"1','yyyyfmMM')                                  " &_				
				' "           GROUP BY fa.fcba, fa.assetgroupcode, fa.assetgroupname, fa.assetcode, fa.assetname, fa.fccode, fa.fcname, fa.transdate      " &_
				' "           UNION ALL                                                                                                                   " &_
				' "             SELECT fa.fcba,                                                                                                           " &_
				' "                    fa.assetgroupcode,                                                                                                 " &_
				' "                    fa.assetgroupname,                                                                                                 " &_
				' "                    fa.assetcode,                                                                                                      " &_
				' "                    fa.assetname,                                                                                                      " &_
				' "                    fa.fccode,                                                                                                         " &_
				' "                    fa.fcname,                                                                                                         " &_
				' "                    fa.transdate,                                                                                                      " &_
				' "				   0 depreciationmonth_ledger,																							" &_
				' "                  0 monthtobe,																											" &_				
				' "                    0 acquisitionvalue,                                                                                                " &_
				' "                    SUM (fad.adjustment_amount) adjustment_amount,                                                                     " &_
				' "                    0 fixassetamount,                                                                                                  " &_
				' "                    0 accumulated_depreciation,                                                                                        " &_
				' "                    0 yearlydepreciation,                                                                                              " &_
				' "                    0 addinValue,                                                                                                      " &_
				' "                    0 adjustment_amount_cur,                                                                                           " &_
				' "                    0 fixassetamount_cur,                                                                                              " &_
				' "                    0 accumulated_depreciation_cur,                                                                                    " &_
				' "                    0 yearlydepreciation_cur,                                                                                          " &_
				' "                    0 monthlydepreciation                                                                                              " &_
				' "               FROM fixedasset_adjustment fad                                                                                          " &_
				' "                    INNER JOIN fixedasset fa ON fa.fcba = fad.fcba AND fa.fccode = fad.fixedassetcode                                  " &_
				' "                    INNER JOIN assetgroup ag ON fa.assetgroupcode = ag.fccode                               							" &_
				' "              WHERE                                                                                                                    " &_
				' "              TO_DATE(TO_CHAR(fad.transdate, 'yyyymm'),'yyyyfmMM') < TO_DATE('"&v_txtYear&"1','yyyyfmMM')                              " &_
				' "              AND fad.fctype = 'ADDITION'                                                                                              " &_
				' "           GROUP BY fa.fcba, fa.assetgroupcode, fa.assetgroupname, fa.assetcode, fa.assetname, fa.fccode, fa.fcname, fa.transdate      " &_
				' "           UNION ALL                                                                                                                   " &_
				' "             SELECT fa.fcba,                                                                                                           " &_
				' "                    fa.assetgroupcode,                                                                                                 " &_
				' "                    fa.assetgroupname,                                                                                                 " &_
				' "                    fa.assetcode,                                                                                                      " &_
				' "                    fa.assetname,                                                                                                      " &_
				' "                    fa.fccode,                                                                                                         " &_
				' "                    fa.fcname,                                                                                                         " &_
				' "                    fa.transdate,                                                                                                      " &_
				' "				   0 depreciationmonth_ledger,																							" &_
				' "                  0 monthtobe,																											" &_				
				' "                    0 acquisitionvalue,                                                                                                " &_
				' "                    0 adjustment_amount,                                                                                               " &_
				' "                    SUM (fwd.fixassetamount) fixassetamount,                                                                           " &_
				' "                    SUM (fwd.accumulated_depreciation) accumulated_depreciation,                                                       " &_
				' "                    0 yearlydepreciation,                                                                                              " &_
				' "                    0 addinValue,                                                                                                      " &_
				' "                    0 adjustment_amount_cur,                                                                                           " &_
				' "                    0 fixassetamount_cur,                                                                                              " &_
				' "                    0 accumulated_depreciation_cur,                                                                                    " &_
				' "                    0 yearlydepreciation_cur,                                                                                          " &_
				' "                    0 monthlydepreciation                                                                                              " &_
				' "               FROM fixedasset_writeoff_det fwd                                                                                        " &_
				' "                    INNER JOIN fixedasset_writeoff fw ON (fw.fccode = fwd.fccode AND fw.fcba = fwd.fcba)                               " &_
				' "                    INNER JOIN fixedasset fa ON fa.fcba = fwd.fcba AND fa.fccode = fwd.fixassetcode                                    " &_
				' "                    INNER JOIN assetgroup ag ON fa.assetgroupcode = ag.fccode                           								" &_
				' "              WHERE                                                                                                                    " &_
				' "              TO_DATE(TO_CHAR(fw.transdate, 'yyyymm'),'yyyyfmMM') < TO_DATE('"&v_txtYear&"1','yyyyfmMM')                               " &_
				' "           GROUP BY fa.fcba, fa.assetgroupcode, fa.assetgroupname, fa.assetcode, fa.assetname, fa.fccode, fa.fcname, fa.transdate      " &_
				' "           UNION ALL                                                                                                                   " &_
				' "             SELECT fa.fcba,                                                                                                           " &_
				' "                    fa.assetgroupcode,                                                                                                 " &_
				' "                    fa.assetgroupname,                                                                                                 " &_
				' "                    fa.assetcode,                                                                                                      " &_
				' "                    fa.assetname,                                                                                                      " &_
				' "                    fa.fccode,                                                                                                         " &_
				' "                    fa.fcname,                                                                                                         " &_
				' "                    fa.transdate,                                                                                                      " &_
				' "          			 MAX(depreciationmonth_ledger)depreciationmonth_ledger,																" &_
				' "          			 MAX(monthtobe)monthtobe,																							" &_
				' "                    0 acquisitionvalue,                                                                                                " &_
				' "                    0 adjustment_amount,                                                                                               " &_
				' "                    0 fixassetamount,                                                                                                  " &_
				' "                    0 accumulated_depreciation,                                                                                        " &_
				' "                    SUM (fdl.yearlydepreciation) AS yearlydepreciation,                                                                " &_
				' "                    0 addinValue,                                                                                                      " &_
				' "                    0 adjustment_amount_cur,                                                                                           " &_
				' "                    0 fixassetamount_cur,                                                                                              " &_
				' "                    0 accumulated_depreciation_cur,                                                                                    " &_
				' "                    0 yearlydepreciation_cur,                                                                                          " &_
				' "                    0 monthlydepreciation                                                                                              " &_
				' "               FROM fixedasset_depreciation_ledger fdl                                                                                 " &_
				' "                    INNER JOIN fixedasset fa ON fa.fcba = fdl.fcba AND fa.fccode = fdl.fccode                                          " &_
				' "                    INNER JOIN assetgroup ag ON fa.assetgroupcode = ag.fccode                             								" &_
				' "              WHERE                                                                                                                    " &_
				' "              TO_DATE(TO_CHAR(date_transaction, 'yyyymm'),'yyyyfmMM') < TO_DATE('"&v_txtYear&"1','yyyyfmMM')                           " &_				 
				' "           GROUP BY fa.fcba, fa.assetgroupcode, fa.assetgroupname, fa.assetcode, fa.assetname, fa.fccode, fa.fcname, fa.transdate      " &_
				' "           UNION ALL                                                                                                                   " &_
				' "             SELECT fa.fcba,                                                                                                           " &_
				' "                    fa.assetgroupcode,                                                                                                 " &_
				' "                    fa.assetgroupname,                                                                                                 " &_
				' "                    fa.assetcode,                                                                                                      " &_
				' "                    fa.assetname,                                                                                                      " &_
				' "                    fa.fccode,                                                                                                         " &_
				' "                    fa.fcname,                                                                                                         " &_
				' "                    fa.transdate,                                                                                                      " &_
				' "				   0 depreciationmonth_ledger,																							" &_
				' "                  0 monthtobe,																											" &_				
				' "                    0 acquisitionvalue,                                                                                                " &_
				' "                    0 adjustment_amount,                                                                                               " &_
				' "                    0 fixassetamount,                                                                                                  " &_
				' "                    0 accumulated_depreciation,                                                                                        " &_
				' "                    0 yearlydepreciation,                                                                                              " &_
				' "                    SUM (acquisitionvalue) addinValue,                                                                                 " &_
				' "                    0 adjustment_amount_cur,                                                                                           " &_
				' "                    0 fixassetamount_cur,                                                                                              " &_
				' "                    0 accumulated_depreciation_cur,                                                                                    " &_
				' "                    0 yearlydepreciation_cur,                                                                                          " &_
				' "                    0 monthlydepreciation                                                                                              " &_
				' "               FROM fixedasset fa                                                                                                      " &_
				' "                    INNER JOIN assetgroup ag ON fa.assetgroupcode = ag.fccode                            								" &_
				' "              WHERE                                                                                                                    " &_
				' "              TO_DATE(TO_CHAR(transdate, 'yyyymm'),'yyyyfmMM') BETWEEN TO_DATE('"&v_txtYear&"1','yyyyfmMM') AND TO_DATE('"&v_txtYear&v_txtMonth&"','yyyyfmMM') " &_				
				' "           GROUP BY fa.fcba, fa.assetgroupcode, fa.assetgroupname, fa.assetcode, fa.assetname, fa.fccode, fa.fcname, fa.transdate      " &_
				' "           UNION ALL                                                                                                                   " &_
				' "             SELECT fa.fcba,                                                                                                           " &_
				' "                    fa.assetgroupcode,                                                                                                 " &_
				' "                    fa.assetgroupname,                                                                                                 " &_
				' "                    fa.assetcode,                                                                                                      " &_
				' "                    fa.assetname,                                                                                                      " &_
				' "                    fa.fccode,                                                                                                         " &_
				' "                    fa.fcname,                                                                                                         " &_
				' "                    fa.transdate,                                                                                                      " &_
				' "				   0 depreciationmonth_ledger,																							" &_
				' "                  0 monthtobe,																											" &_				
				' "                    0 acquisitionvalue,                                                                                                " &_
				' "                    0 adjustment_amount,                                                                                               " &_
				' "                    0 fixassetamount,                                                                                                  " &_
				' "                    0 accumulated_depreciation,                                                                                        " &_
				' "                    0 yearlydepreciation,                                                                                              " &_
				' "                    0 addinValue,                                                                                                      " &_
				' "                    SUM (adjustment_amount) adjustment_amount_cur,                                                                     " &_
				' "                    0 fixassetamount_cur,                                                                                              " &_
				' "                    0 accumulated_depreciation_cur,                                                                                    " &_
				' "                    0 yearlydepreciation_cur,                                                                                          " &_
				' "                    0 monthlydepreciation                                                                                              " &_
				' "               FROM fixedasset_adjustment fad                                                                                          " &_
				' "                    INNER JOIN fixedasset fa ON fa.fcba = fad.fcba AND fa.fccode = fad.fixedassetcode                                  " &_
				' "                    INNER JOIN assetgroup ag ON fa.assetgroupcode = ag.fccode                               							" &_
				' "              WHERE                                                                                                                    " &_
				' "              TO_DATE(TO_CHAR(fad.transdate, 'yyyymm'),'yyyyfmMM') BETWEEN TO_DATE('"&v_txtYear&"1','yyyyfmMM') AND TO_DATE('"&v_txtYear&v_txtMonth&"','yyyyfmMM') " &_				
				' "              AND fad.fctype = 'ADDITION'                                                                                              " &_
				' "           GROUP BY fa.fcba, fa.assetgroupcode, fa.assetgroupname, fa.assetcode, fa.assetname, fa.fccode, fa.fcname, fa.transdate      " &_
				' "           UNION ALL                                                                                                                   " &_
				' "             SELECT fa.fcba,                                                                                                           " &_
				' "                    fa.assetgroupcode,                                                                                                 " &_
				' "                    fa.assetgroupname,                                                                                                 " &_
				' "                    fa.assetcode,                                                                                                      " &_
				' "                    fa.assetname,                                                                                                      " &_
				' "                    fa.fccode,                                                                                                         " &_
				' "                    fa.fcname,                                                                                                         " &_
				' "                    fa.transdate,                                                                                                      " &_
				' "				   0 depreciationmonth_ledger,																							" &_
				' "                  0 monthtobe,																											" &_				
				' "                    0 acquisitionvalue,                                                                                                " &_
				' "                    0 adjustment_amount,                                                                                               " &_
				' "                    0 fixassetamount,                                                                                                  " &_
				' "                    0 accumulated_depreciation,                                                                                        " &_
				' "                    0 yearlydepreciation,                                                                                              " &_
				' "                    0 addinValue,                                                                                                      " &_
				' "                    0 adjustment_amount_cur,                                                                                           " &_
				' "                    SUM (fwd.fixassetamount) fixassetamount_cur,                                                                       " &_
				' "                    SUM (fwd.accumulated_depreciation)                                                                                 " &_
				' "                       accumulated_depreciation_cur,                                                                                   " &_
				' "                    0 yearlydepreciation_cur,                                                                                          " &_
				' "                    0 monthlydepreciation                                                                                              " &_
				' "               FROM fixedasset_writeoff_det fwd                                                                                        " &_
				' "                    INNER JOIN fixedasset_writeoff fw ON (fw.fccode = fwd.fccode AND fw.fcba = fwd.fcba)                               " &_
				' "                    INNER JOIN fixedasset fa ON fa.fcba = fwd.fcba AND fa.fccode = fwd.fixassetcode                                    " &_
				' "                    INNER JOIN assetgroup ag ON     fa.assetgroupcode = ag.fccode                          							" &_
				' "              WHERE                                                                                                                    " &_
				' "              TO_DATE(TO_CHAR(fw.transdate, 'yyyymm'),'yyyyfmMM') BETWEEN TO_DATE('"&v_txtYear&"1','yyyyfmMM') AND TO_DATE('"&v_txtYear&v_txtMonth&"','yyyyfmMM') " &_				
				' "           GROUP BY fa.fcba, fa.assetgroupcode, fa.assetgroupname, fa.assetcode, fa.assetname, fa.fccode, fa.fcname, fa.transdate      " &_
				' "           UNION ALL                                                                                                                   " &_
				' "             SELECT fa.fcba,                                                                                                           " &_
				' "                    fa.assetgroupcode,                                                                                                 " &_
				' "                    fa.assetgroupname,                                                                                                 " &_
				' "                    fa.assetcode,                                                                                                      " &_
				' "                    fa.assetname,                                                                                                      " &_
				' "                    fa.fccode,                                                                                                         " &_
				' "                    fa.fcname,                                                                                                         " &_
				' "                    fa.transdate,                                                                                                      " &_
				' "          			 MAX(depreciationmonth_ledger)depreciationmonth_ledger,																" &_
				' "          			 MAX(monthtobe)monthtobe,																							" &_				
				' "                    0 acquisitionvalue,                                                                                                " &_
				' "                    0 adjustment_amount,                                                                                               " &_
				' "                    0 fixassetamount,                                                                                                  " &_
				' "                    0 accumulated_depreciation,                                                                                        " &_
				' "                    0 yearlydepreciation,                                                                                              " &_
				' "                    0 addinValue,                                                                                                      " &_
				' "                    0 adjustment_amount_cur,                                                                                           " &_
				' "                    0 fixassetamount_cur,                                                                                              " &_
				' "                    0 accumulated_depreciation_cur,                                                                                    " &_
				' "                    SUM (fdl.yearlydepreciation) yearlydepreciation_cur,                                                               " &_
				' "                    0 monthlydepreciation                                                                                              " &_
				' "               FROM fixedasset_depreciation_ledger fdl                                                                                 " &_
				' "                    INNER JOIN fixedasset fa ON fa.fcba = fdl.fcba AND fa.fccode = fdl.fccode                                          " &_
				' "                    INNER JOIN assetgroup ag ON fa.assetgroupcode = ag.fccode                           								" &_
				' "              WHERE                                                                                                                    " &_
				' "              TO_DATE(TO_CHAR(date_transaction, 'yyyymm'),'yyyyfmMM') = TO_DATE('"&v_txtYear&v_txtMonth&"','yyyyfmMM')                 " &_				
				' "           GROUP BY fa.fcba, fa.assetgroupcode, fa.assetgroupname, fa.assetcode, fa.assetname, fa.fccode, fa.fcname, fa.transdate      " &_
				' "           UNION ALL                                                                                                                   " &_
				' "             SELECT fa.fcba,                                                                                                           " &_
				' "                    fa.assetgroupcode,                                                                                                 " &_
				' "                    fa.assetgroupname,                                                                                                 " &_
				' "                    fa.assetcode,                                                                                                      " &_
				' "                    fa.assetname,                                                                                                      " &_
				' "                    fa.fccode,                                                                                                         " &_
				' "                    fa.fcname,                                                                                                         " &_
				' "                    fa.transdate,                                                                                                      " &_
				' "          			 MAX(depreciationmonth_ledger)depreciationmonth_ledger,																" &_
				' "          			 MAX(monthtobe)monthtobe,																							" &_				
				' "                    0 acquisitionvalue,                                                                                                " &_
				' "                    0 adjustment_amount,                                                                                               " &_
				' "                    0 fixassetamount,                                                                                                  " &_
				' "                    0 accumulated_depreciation,                                                                                        " &_
				' "                    0 yearlydepreciation,                                                                                              " &_
				' "                    0 addinValue,                                                                                                      " &_
				' "                    0 adjustment_amount_cur,                                                                                           " &_
				' "                    0 fixassetamount_cur,                                                                                              " &_
				' "                    0 accumulated_depreciation_cur,                                                                                    " &_
				' "                    0 yearlydepreciation_cur,                                                                                          " &_
				' "                    SUM (monthlydepreciation) monthlydepreciation                                                                      " &_
				' "               FROM fixedasset_depreciation_ledger fdl                                                                                 " &_
				' "                    INNER JOIN fixedasset fa ON fa.fcba = fdl.fcba AND fa.fccode = fdl.fccode                                          " &_
				' "                    INNER JOIN assetgroup ag ON     fa.assetgroupcode = ag.fccode                        								 " &_
				' "              WHERE                                                                                                                    " &_
				' "              TO_DATE(TO_CHAR(date_transaction, 'yyyymm'),'yyyyfmMM') = TO_DATE('"&v_txtYear&v_txtMonth&"','yyyyfmMM')                 " &_
				' "              AND fdl.lastdepreciationmonth > 0                                                                                        " &_
				' "           GROUP BY fa.fcba, fa.assetgroupcode, fa.assetgroupname, fa.assetcode, fa.assetname, fa.fccode, fa.fcname, fa.transdate)     " &_
				' " GROUP BY fcba, assetgroupcode, assetgroupname, assetcode, assetname, fccode, fcname, transdate)" 
	
		strMain = 	" SELECT " & strFieldCons & _
					" assetgroupcode,                                                                                                                                        	"&_
					" assetgroupname,                                                                                                                                            "&_
					" assetcode,                                                                                                                                                 "&_
					" assetname,                                                                                                                                                 "&_
					" fccode,                                                                                                                                                    "&_
					" fcname,                                                                                                                                                    "&_
					" transdate,                                                                                                                                                 "&_
					" acquisitionvalue_prev,                                                                                                                                     "&_
					" adjustment_amount_prev,                                                                                                                                    "&_
					" fixassetamount_prev,                                                                                                                                       "&_
					" accumulated_depreciation_prev,                                                                                                                             "&_
					" yearlydepreciation_prev,                                                                                                                                   "&_
					" bookvalue,                                                                                                                                                 "&_
					" addinvalue,                                                                                                                                                "&_
					" 0 as acquisitionvalue_curr,                                                                                                                                "&_
					" adjustment_amount_curr,                                                                                                                                    "&_
					" fixassetamount_curr,                                                                                                                                       "&_
					" accumulated_depreciation_curr,                                                                                                                             "&_
					" yearlydepreciation_curr,                                                                                                                                   "&_
					" netbookvalue,                                                                                                                                              "&_
					" monthlydepreciation_curr 				                                                                                                                    "&_
					" from (                                                                                                                                                     "&_
					"    select                                                                                                                                                 "&_
					"    fcba,                                                                                                                                                  "&_
					"    assetgroupcode,                                                                                                                                        "&_
					"    assetgroupname,                                                                                                                                        "&_
					"    assetcode,                                                                                                                                             "&_
					"    assetname,                                                                                                                                             "&_
					"    fccode,                                                                                                                                                "&_
					"    fcname,                                                                                                                                                "&_
					"    transdate,                                                                                                                                             "&_
					"    sum (acquisitionvalue) acquisitionvalue_prev,                                                                                                          "&_
					"    sum (adjustment_amount) adjustment_amount_prev,                                                                                                        "&_
					"    sum (fixassetamount) fixassetamount_prev,                                                                                                              "&_
					"    sum (accumulated_depreciation) accumulated_depreciation_prev,                                                                                          "&_
					"    sum (yearlydepreciation) yearlydepreciation_prev,                                                                                                      "&_
					"    (sum (acquisitionvalue) + sum (adjustment_amount)) - sum (fixassetamount) - sum (yearlydepreciation) bookvalue,                                        "&_
					"    sum (addinvalue) addinvalue,                                                                                                                           "&_
					"    sum (adjustment_amount_cur) adjustment_amount_curr,                                                                                                    "&_
					"    sum (fixassetamount_cur) fixassetamount_curr,                                                                                                          "&_
					"    sum (accumulated_depreciation_cur) accumulated_depreciation_curr,                                                                                      "&_
					"    sum (yearlydepreciation_cur) - sum (yearlydepreciation) yearlydepreciation_curr,                                                                       "&_
					"    ((sum (acquisitionvalue) + sum (adjustment_amount)) - sum (fixassetamount) - sum (yearlydepreciation))                                                 "&_
					"    + sum (addinvalue) + sum (adjustment_amount_cur) - sum (fixassetamount_cur) - (sum (yearlydepreciation_cur) - sum (yearlydepreciation)) netbookvalue,  "&_    
					"    sum (monthlydepreciation) monthlydepreciation_curr                                                                                                     "&_
					"    from (                                                                                                                                                 "&_
					"        select                                                                                                                                             "&_
					"        fa.fcba,                                                                                                                                           "&_
					"        fa.assetgroupcode,                                                                                                                                 "&_
					"        fa.assetgroupname,                                                                                                                                 "&_
					"        fa.assetcode,                                                                                                                                      "&_
					"        fa.assetname,                                                                                                                                      "&_
					"        fa.fccode,                                                                                                                                         "&_
					"        fa.fcname,                                                                                                                                         "&_
					"        fa.transdate,                                                                                                                                      "&_
					"        0 acquisitionvalue,                                                                                                                                "&_
					"        0 adjustment_amount,                                                                                                                               "&_
					"        0 fixassetamount,                                                                                                                                  "&_
					"        0 accumulated_depreciation,                                                                                                                        "&_
					"        0 yearlydepreciation,                                                                                                                              "&_
					"        0 addinvalue,                                                                                                                                      "&_
					"        0 adjustment_amount_cur,                                                                                                                           "&_
					"        0 fixassetamount_cur,                                                                                                                              "&_
					"        0 accumulated_depreciation_cur,                                                                                                                    "&_
					"        0 yearlydepreciation_cur,                                                                                                                          "&_
					"        0 monthlydepreciation                                                                                                                              "&_
					"        from fixedasset fa                                                                                                                                 "&_
					"        inner join assetgroup ag on fa.assetgroupcode = ag.fccode                                                                                          "&_
					"        where to_date(to_char(transdate, 'yyyymm'),'yyyyMM') <= to_date('"&v_txtYear&v_txtMonth&"','yyyyMM') 						 				                        "&_
					"        union all                                                                                                                                          "&_
					"        select                                                                                                                                             "&_
					"        fa.fcba,                                                                                                                                           "&_
					"        fa.assetgroupcode,                                                                                                                                 "&_
					"        fa.assetgroupname,                                                                                                                                 "&_
					"        fa.assetcode,                                                                                                                                      "&_
					"        fa.assetname,                                                                                                                                      "&_
					"        fa.fccode,                                                                                                                                         "&_
					"        fa.fcname,                                                                                                                                         "&_
					"        fa.transdate,                                                                                                                                      "&_
					"        sum (fa.acquisitionvalue) acquisitionvalue,                                                                                                        "&_
					"        0 adjustment_amount,                                                                                                                               "&_
					"        0 fixassetamount,                                                                                                                                  "&_
					"        0 accumulated_depreciation,                                                                                                                        "&_
					"        0 yearlydepreciation,                                                                                                                              "&_
					"        0 addinvalue,                                                                                                                                      "&_
					"        0 adjustment_amount_cur,                                                                                                                           "&_
					"        0 fixassetamount_cur,                                                                                                                              "&_
					"        0 accumulated_depreciation_cur,                                                                                                                    "&_
					"        0 yearlydepreciation_cur,                                                                                                                          "&_
					"        0 monthlydepreciation                                                                                                                              "&_
					"        from fixedasset fa                                                                                                                                 "&_
					"        inner join assetgroup ag on  fa.assetgroupcode = ag.fccode                                                                                         "&_
					"        where                                                                                                                                              "&_
					"        to_date(to_char(transdate, 'yyyymm'),'yyyyMM') <= to_date('"&v_txtYear&v_txtMonth&"','yyyyMM')                                   				                    "&_
					"        group by                                                                                                                                           "&_
					"        fa.fcba,                                                                                                                                           "&_
					"        fa.assetgroupcode,                                                                                                                                 "&_
					"        fa.assetgroupname,                                                                                                                                 "&_
					"        fa.assetcode,                                                                                                                                      "&_
					"        fa.assetname,                                                                                                                                      "&_
					"        fa.fccode,                                                                                                                                         "&_
					"        fa.fcname,                                                                                                                                         "&_
					"        fa.transdate                                                                                                                                       "&_
					"        union all                                                                                                                                          "&_
					"        select                                                                                                                                             "&_
					"        fa.fcba,                                                                                                                                           "&_
					"        fa.assetgroupcode,                                                                                                                                 "&_
					"        fa.assetgroupname,                                                                                                                                 "&_
					"        fa.assetcode,                                                                                                                                      "&_
					"        fa.assetname,                                                                                                                                      "&_
					"        fa.fccode,                                                                                                                                         "&_
					"        fa.fcname,                                                                                                                                         "&_
					"        fa.transdate,                                                                                                                                      "&_
					"        0 acquisitionvalue,                                                                                                                                "&_
					"        sum (fad.adjustment_amount) adjustment_amount,                                                                                                     "&_
					"        0 fixassetamount,                                                                                                                                  "&_
					"        0 accumulated_depreciation,                                                                                                                        "&_
					"        0 yearlydepreciation,                                                                                                                              "&_
					"        0 addinvalue,                                                                                                                                      "&_
					"        0 adjustment_amount_cur,                                                                                                                           "&_
					"        0 fixassetamount_cur,                                                                                                                              "&_
					"        0 accumulated_depreciation_cur,                                                                                                                    "&_
					"        0 yearlydepreciation_cur,                                                                                                                          "&_
					"        0 monthlydepreciation                                                                                                                              "&_
					"        from fixedasset_adjustment fad                                                                                                                     "&_
					"        inner join fixedasset fa on (                                                                                                                      "&_
					"        fa.fcba = fad.fcba and                                                                                                                             "&_
					"        fa.fccode = fad.fixedassetcode)                                                                                                                    "&_
					"        inner join assetgroup ag on  fa.assetgroupcode = ag.fccode                                                                                         "&_
					"        where to_date(to_char(fad.transdate, 'yyyymm'),'yyyyMM') < to_date('201901','yyyyMM')                                                              "&_
					"        and fad.fctype = 'ADDITION'                                                                                                                        "&_
					"        group by                                                                                                                                           "&_
					"        fa.fcba,                                                                                                                                           "&_
					"        fa.assetgroupcode,                                                                                                                                 "&_
					"        fa.assetgroupname,                                                                                                                                 "&_
					"        fa.assetcode,                                                                                                                                      "&_
					"        fa.assetname,                                                                                                                                      "&_
					"        fa.fccode,                                                                                                                                         "&_
					"        fa.fcname,                                                                                                                                         "&_
					"        fa.transdate                                                                                                                                       "&_
					"        union all                                                                                                                                          "&_
					"        select                                                                                                                                             "&_
					"        fa.fcba,                                                                                                                                           "&_
					"        fa.assetgroupcode,                                                                                                                                 "&_
					"        fa.assetgroupname,                                                                                                                                 "&_
					"        fa.assetcode,                                                                                                                                      "&_
					"        fa.assetname,                                                                                                                                      "&_
					"        fa.fccode,                                                                                                                                         "&_
					"        fa.fcname,                                                                                                                                         "&_
					"        fa.transdate,                                                                                                                                      "&_
					"        0 acquisitionvalue,                                                                                                                                "&_
					"        0 adjustment_amount,                                                                                                                               "&_
					"        sum (fwd.fixassetamount) fixassetamount,                                                                                                           "&_
					"        sum (fwd.accumulated_depreciation) accumulated_depreciation,                                                                                       "&_
					"        0 yearlydepreciation,                                                                                                                              "&_
					"        0 addinvalue,                                                                                                                                      "&_
					"        0 adjustment_amount_cur,                                                                                                                           "&_
					"        0 fixassetamount_cur,                                                                                                                              "&_
					"        0 accumulated_depreciation_cur,                                                                                                                    "&_
					"        0 yearlydepreciation_cur,                                                                                                                          "&_
					"        0 monthlydepreciation                                                                                                                              "&_
					"        from fixedasset_writeoff_det fwd                                                                                                                   "&_
					"        inner join fixedasset_writeoff fw on (                                                                                                             "&_
					"        fw.fccode = fwd.fccode and                                                                                                                         "&_
					"        fw.fcba = fwd.fcba)                                                                                                                                "&_
					"        inner join fixedasset fa on (                                                                                                                      "&_
					"        fa.fcba = fw.fcba and                                                                                                                              "&_
					"        fa.fccode = fw.fccode)                                                                                                                             "&_
					"        inner join assetgroup ag on fa.assetgroupcode = ag.fccode                                                                                          "&_
					"        where to_date(to_char(fw.transdate, 'yyyymm'),'yyyyMM') < to_date('201901','yyyyMM')                                                               "&_
					"        group by                                                                                                                                           "&_
					"        fa.fcba,                                                                                                                                           "&_
					"        fa.assetgroupcode,                                                                                                                                 "&_
					"        fa.assetgroupname,                                                                                                                                 "&_
					"        fa.assetcode,                                                                                                                                      "&_
					"        fa.assetname,                                                                                                                                      "&_
					"        fa.fccode,                                                                                                                                         "&_
					"        fa.fcname,                                                                                                                                         "&_
					"        fa.transdate                                                                                                                                       "&_
					"        union all                                                                                                                                          "&_
					"        select                                                                                                                                             "&_
					"        fa.fcba,                                                                                                                                           "&_
					"        fa.assetgroupcode,                                                                                                                                 "&_
					"        fa.assetgroupname,                                                                                                                                 "&_
					"        fa.assetcode,                                                                                                                                      "&_
					"        fa.assetname,                                                                                                                                      "&_
					"        fa.fccode,                                                                                                                                         "&_
					"        fa.fcname,                                                                                                                                         "&_
					"        fa.transdate,                                                                                                                                      "&_
					"        0 acquisitionvalue,                                                                                                                                "&_
					"        0 adjustment_amount,                                                                                                                               "&_
					"        0 fixassetamount,                                                                                                                                  "&_
					"        0 accumulated_depreciation,                                                                                                                        "&_
					"        sum (fdl.yearlydepreciation) as yearlydepreciation,                                                                                                "&_
					"        0 addinvalue,                                                                                                                                      "&_
					"        0 adjustment_amount_cur,                                                                                                                           "&_
					"        0 fixassetamount_cur,                                                                                                                              "&_
					"        0 accumulated_depreciation_cur,                                                                                                                    "&_
					"        0 yearlydepreciation_cur,                                                                                                                          "&_
					"        0 monthlydepreciation                                                                                                                              "&_
					"        from fixedasset_depreciation_ledger fdl                                                                                                            "&_
					"        inner join fixedasset fa on (                                                                                                                      "&_
					"        fa.fcba = fdl.fcba and                                                                                                                             "&_
					"        fa.fccode = fdl.fccode)                                                                                                                            "&_
					"        inner join assetgroup ag on                                                                                                                       "&_
					"        fa.assetgroupcode = ag.fccode                                                                                                                      "&_
					"        where to_date(to_char(date_transaction, 'yyyymm'),'yyyyMM') < to_date('201901','yyyyMM')                           				                "&_
					"        group by                                                                                                                                           "&_
					"        fa.fcba,                                                                                                                                           "&_
					"        fa.assetgroupcode,                                                                                                                                 "&_
					"        fa.assetgroupname,                                                                                                                                 "&_
					"        fa.assetcode,                                                                                                                                      "&_
					"        fa.assetname,                                                                                                                                      "&_
					"        fa.fccode,                                                                                                                                         "&_
					"        fa.fcname,                                                                                                                                         "&_
					"        fa.transdate                                                                                                                                       "&_
					"        union all                                                                                                                                          "&_
					"        select                                                                                                                                             "&_
					"        fa.fcba,                                                                                                                                           "&_
					"        fa.assetgroupcode,                                                                                                                                 "&_
					"        fa.assetgroupname,                                                                                                                                 "&_
					"        fa.assetcode,                                                                                                                                      "&_
					"        fa.assetname,                                                                                                                                      "&_
					"        fa.fccode,                                                                                                                                         "&_
					"        fa.fcname,                                                                                                                                         "&_
					"        fa.transdate,                                                                                                                                      "&_
					"        0 acquisitionvalue,                                                                                                                                "&_
					"        0 adjustment_amount,                                                                                                                               "&_
					"        0 fixassetamount,                                                                                                                                  "&_
					"        0 accumulated_depreciation,                                                                                                                        "&_
					"        0 yearlydepreciation,                                                                                                                              "&_
					"        sum (acquisitionvalue) addinvalue,                                                                                                                 "&_
					"        0 adjustment_amount_cur,                                                                                                                           "&_
					"        0 fixassetamount_cur,                                                                                                                              "&_
					"        0 accumulated_depreciation_cur,                                                                                                                    "&_
					"        0 yearlydepreciation_cur,                                                                                                                          "&_
					"        0 monthlydepreciation                                                                                                                              "&_
					"        from fixedasset fa                                                                                                                                 "&_
					"        inner join assetgroup ag on                                                                                                                        "&_
					"        fa.assetgroupcode = ag.fccode                                                                                                                      "&_
					"        where to_date(to_char(transdate, 'yyyymm'),'yyyyMM') between to_date('201901','yyyyMM') and to_date('"&v_txtYear&v_txtMonth&"','yyyyMM')  				            "&_
					"        group by                                                                                                                                           "&_
					"        fa.fcba,                                                                                                                                           "&_
					"        fa.assetgroupcode,                                                                                                                                 "&_
					"        fa.assetgroupname,                                                                                                                                 "&_
					"        fa.assetcode,                                                                                                                                      "&_
					"        fa.assetname,                                                                                                                                      "&_
					"        fa.fccode,                                                                                                                                         "&_
					"        fa.fcname,                                                                                                                                         "&_
					"        fa.transdate                                                                                                                                       "&_
					"        union all                                                                                                                                          "&_
					"        select                                                                                                                                             "&_
					"        fa.fcba,                                                                                                                                           "&_
					"        fa.assetgroupcode,                                                                                                                                 "&_
					"        fa.assetgroupname,                                                                                                                                 "&_
					"        fa.assetcode,                                                                                                                                      "&_
					"        fa.assetname,                                                                                                                                      "&_
					"        fa.fccode,                                                                                                                                         "&_
					"        fa.fcname,                                                                                                                                         "&_
					"        fa.transdate,                                                                                                                                      "&_
					"        0 acquisitionvalue,                                                                                                                                "&_
					"        0 adjustment_amount,                                                                                                                               "&_
					"        0 fixassetamount,                                                                                                                                  "&_
					"        0 accumulated_depreciation,                                                                                                                        "&_
					"        0 yearlydepreciation,                                                                                                                              "&_
					"        0 addinvalue,                                                                                                                                      "&_
					"        sum (adjustment_amount) adjustment_amount_cur,                                                                                                     "&_
					"        0 fixassetamount_cur,                                                                                                                              "&_
					"        0 accumulated_depreciation_cur,                                                                                                                    "&_
					"        0 yearlydepreciation_cur,                                                                                                                          "&_
					"        0 monthlydepreciation                                                                                                                              "&_
					"        from fixedasset_adjustment fad                                                                                                                     "&_
					"        inner join fixedasset fa on (                                                                                                                      "&_
					"        fa.fcba = fad.fcba and                                                                                                                             "&_
					"        fa.fccode = fad.fixedassetcode)                                                                                                                    "&_
					"        inner join assetgroup ag on                                                                                                                        "&_
					"        fa.assetgroupcode = ag.fccode                                                                                                                      "&_
					"        where to_date(to_char(fad.transdate, 'yyyymm'),'yyyyMM') between to_date('201901','yyyyMM') and to_date('"&v_txtYear&v_txtMonth&"','yyyyMM')                         "&_
					"        and fad.fctype = 'ADDITION'                                                                                                                        "&_
					"        group by                                                                                                                                           "&_
					"        fa.fcba,                                                                                                                                           "&_
					"        fa.assetgroupcode,                                                                                                                                 "&_
					"        fa.assetgroupname,                                                                                                                                 "&_
					"        fa.assetcode,                                                                                                                                      "&_
					"        fa.assetname,                                                                                                                                      "&_
					"        fa.fccode,                                                                                                                                         "&_
					"        fa.fcname,                                                                                                                                         "&_
					"        fa.transdate                                                                                                                                       "&_
					"        union all                                                                                                                                          "&_
					"        select                                                                                                                                             "&_
					"        fa.fcba,                                                                                                                                           "&_
					"        fa.assetgroupcode,                                                                                                                                 "&_
					"        fa.assetgroupname,                                                                                                                                 "&_
					"        fa.assetcode,                                                                                                                                      "&_
					"        fa.assetname,                                                                                                                                      "&_
					"        fa.fccode,                                                                                                                                         "&_
					"        fa.fcname,                                                                                                                                         "&_
					"        fa.transdate,                                                                                                                                      "&_
					"        0 acquisitionvalue,                                                                                                                                "&_
					"        0 adjustment_amount,                                                                                                                               "&_
					"        0 fixassetamount,                                                                                                                                  "&_
					"        0 accumulated_depreciation,                                                                                                                        "&_
					"        0 yearlydepreciation,                                                                                                                              "&_
					"        0 addinvalue,                                                                                                                                      "&_
					"        0 adjustment_amount_cur,                                                                                                                           "&_
					"        sum (fwd.fixassetamount) fixassetamount_cur,                                                                                                       "&_
					"        sum (fwd.accumulated_depreciation)                                                                                                                 "&_
					"        accumulated_depreciation_cur,                                                                                                                      "&_
					"        0 yearlydepreciation_cur,                                                                                                                          "&_
					"        0 monthlydepreciation                                                                                                                              "&_
					"        from fixedasset_writeoff_det fwd                                                                                                                   "&_
					"        inner join fixedasset_writeoff fw on (                                                                                                             "&_
					"        fw.fccode = fwd.fccode and                                                                                                                         "&_
					"        fw.fcba = fwd.fcba)                                                                                                                                "&_
					"        inner join fixedasset fa on (                                                                                                                      "&_
					"        fa.fcba = fw.fcba and                                                                                                                              "&_
					"        fa.fccode = fw.fccode)                                                                                                                             "&_
					"        inner join assetgroup ag on                                                                                                                        "&_
					"        fa.assetgroupcode = ag.fccode                                                                                                                      "&_
					"        where to_date(to_char(fw.transdate, 'yyyymm'),'yyyyMM') between to_date('201901','yyyyMM') and to_date('"&v_txtYear&v_txtMonth&"','yyyyMM')                          "&_
					"        group by                                                                                                                                           "&_
					"        fa.fcba,                                                                                                                                           "&_
					"        fa.assetgroupcode,                                                                                                                                 "&_
					"        fa.assetgroupname,                                                                                                                                 "&_
					"        fa.assetcode,                                                                                                                                      "&_
					"        fa.assetname,                                                                                                                                      "&_
					"        fa.fccode,                                                                                                                                         "&_
					"        fa.fcname,                                                                                                                                         "&_
					"        fa.transdate                                                                                                                                       "&_
					"        union all                                                                                                                                          "&_
					"        select                                                                                                                                             "&_
					"        fa.fcba,                                                                                                                                           "&_
					"        fa.assetgroupcode,                                                                                                                                 "&_
					"        fa.assetgroupname,                                                                                                                                 "&_
					"        fa.assetcode,                                                                                                                                      "&_
					"        fa.assetname,                                                                                                                                      "&_
					"        fa.fccode,                                                                                                                                         "&_
					"        fa.fcname,                                                                                                                                         "&_
					"        fa.transdate,                                                                                                                                      "&_
					"        0 acquisitionvalue,                                                                                                                                "&_
					"        0 adjustment_amount,                                                                                                                               "&_
					"        0 fixassetamount,                                                                                                                                  "&_
					"        0 accumulated_depreciation,                                                                                                                        "&_
					"        0 yearlydepreciation,                                                                                                                              "&_
					"        0 addinvalue,                                                                                                                                      "&_
					"        0 adjustment_amount_cur,                                                                                                                           "&_
					"        0 fixassetamount_cur,                                                                                                                              "&_
					"        0 accumulated_depreciation_cur,                                                                                                                    "&_
					"        sum (fdl.yearlydepreciation_cur) yearlydepreciation_cur,                                                                                           "&_
					"        0 monthlydepreciation                                                                                                                              "&_
					"        from (                                                                                                                                             "&_
					"            select                                                                                                                                         "&_
					"            fdl.fcba,                                                                                                                                      "&_
					"            fdl.fccode,                                                                                                                                    "&_                                                                     
					"            fdl.date_transaction,                                                                                                                          "&_
					"            fdl.lineno,                                                                                                                                    "&_                                                  
					"            sum (fdl.yearlydepreciation) yearlydepreciation_cur                                                                                            "&_                                                             
					"            from fixedasset_depreciation_ledger fdl                                                                                                        "&_
					"            inner join (                                                                                                                                   "&_
					"                select                                                                                                                                     "&_
					"                fdl.fcba,                                                                                                                                  "&_
					"                fdl.fccode,                                                                                                                                "&_                                                                         
					"                max(fdl.date_transaction) as date_transaction                                                                                              "&_
					"                from fixedasset_depreciation_ledger fdl                                                                                                    "&_
					"                group by                                                                                                                                   "&_
					"                fdl.fcba,                                                                                                                                  "&_
					"                fdl.fccode                                                                                                                                 "&_
					"            ) fdl_a                                                                                                                                        "&_
					"            on (                                                                                                                                           "&_
					"            fdl.fcba = fdl_a.fcba and                                                                                                                      "&_
					"            fdl.fccode = fdl_a.fccode and                                                                                                                  "&_
					"            fdl.date_transaction = fdl_a.date_transaction)                                                                                                 "&_                                         
					"            where  to_date(to_char(fdl.date_transaction, 'yyyymm'),'yyyyMM') between to_date('201901','yyyyMM') and to_date('201906','yyyyMM')             "&_
					"            group by                                                                                                                                       "&_
					"            fdl.fcba,                                                                                                                                      "&_
					"            fdl.fccode,                                                                                                                                    "&_                                                                     
					"            fdl.date_transaction,                                                                                                                          "&_
					"            fdl.lineno                                                                                                                                     "&_
					"            union all                                                                                                                                      "&_
					"            select                                                                                                                                         "&_
					"            fdl.fcba,                                                                                                                                      "&_
					"            fdl.fccode,                                                                                                                                    "&_                                                                     
					"            fdl.date_transaction,                                                                                                                          "&_
					"            fdl.lineno,                                                                                                                                    "&_                                                  
					"            sum (fdl.yearlydepreciation) yearlydepreciation_cur                                                                                            "&_                                                             
					"            from fixedasset_depreciation_ledger fdl                                                                                                        "&_                                                                                                                
					"            where to_date(to_char(fdl.date_transaction, 'yyyymm'),'yyyyMM') = to_date('"&v_txtYear&v_txtMonth&"','yyyyMM')                                                   "&_
					"            group by                                                                                                                                       "&_
					"            fdl.fcba,                                                                                                                                      "&_
					"            fdl.fccode,                                                                                                                                    "&_                                                                     
					"            fdl.date_transaction,                                                                                                                          "&_
					"            fdl.lineno                                                                                                                                     "&_
					"        ) fdl                                                                                                                                              "&_
					"        inner join fixedasset fa on (                                                                                                                      "&_
					"        fa.fcba = fdl.fcba and                                                                                                                             "&_
					"        fa.fccode = fdl.fccode)                                                                                                                            "&_
					"        inner join assetgroup ag on  fa.assetgroupcode = ag.fccode                                                                                         "&_
					"        group by                                                                                                                                           "&_
					"        fa.fcba,                                                                                                                                           "&_
					"        fa.assetgroupcode,                                                                                                                                 "&_
					"        fa.assetgroupname,                                                                                                                                 "&_
					"        fa.assetcode,                                                                                                                                      "&_
					"        fa.assetname,                                                                                                                                      "&_
					"        fa.fccode,                                                                                                                                         "&_
					"        fa.fcname,                                                                                                                                         "&_
					"        fa.transdate                                                                                                                                       "&_
					"        union all                                                                                                                                          "&_
					"        select                                                                                                                                             "&_
					"        fa.fcba,                                                                                                                                           "&_
					"        fa.assetgroupcode,                                                                                                                                 "&_
					"        fa.assetgroupname,                                                                                                                                 "&_
					"        fa.assetcode,                                                                                                                                      "&_
					"        fa.assetname,                                                                                                                                      "&_
					"        fa.fccode,                                                                                                                                         "&_
					"        fa.fcname,                                                                                                                                         "&_
					"        fa.transdate,                                                                                                                                      "&_
					"        0 acquisitionvalue,                                                                                                                                "&_
					"        0 adjustment_amount,                                                                                                                               "&_
					"        0 fixassetamount,                                                                                                                                  "&_
					"        0 accumulated_depreciation,                                                                                                                        "&_
					"        0 yearlydepreciation,                                                                                                                              "&_
					"        0 addinvalue,                                                                                                                                      "&_
					"        0 adjustment_amount_cur,                                                                                                                           "&_
					"        0 fixassetamount_cur,                                                                                                                              "&_
					"        0 accumulated_depreciation_cur,                                                                                                                    "&_
					"        0 yearlydepreciation_cur,                                                                                                                          "&_
					"        sum (monthlydepreciation) monthlydepreciation                                                                                                      "&_
					"        from (                                                                                                                                             "&_
					"            select                                                                                                                                         "&_
					"            fdl.fcba,                                                                                                                                      "&_
					"            fdl.fccode,                                                                                                                                    "&_                                                                     
					"            fdl.date_transaction,                                                                                                                          "&_
					"            fdl.lineno,                                                                                                                                    "&_                                                  
					"            sum (fdl.monthlydepreciation) monthlydepreciation                                                                                              "&_                                                           
					"            from fixedasset_depreciation_ledger fdl                                                                                                        "&_
					"            inner join (                                                                                                                                   "&_
					"                select                                                                                                                                     "&_
					"                fdl.fcba,                                                                                                                                  "&_
					"                fdl.fccode,                                                                                                                                "&_                                                                         
					"                max(fdl.date_transaction) as date_transaction                                                                                              "&_
					"                from fixedasset_depreciation_ledger fdl                                                                                                    "&_
					"                group by                                                                                                                                   "&_
					"                fdl.fcba,                                                                                                                                  "&_
					"                fdl.fccode                                                                                                                                 "&_
					"            ) fdl_a                                                                                                                                        "&_
					"            on (                                                                                                                                           "&_
					"            fdl.fcba = fdl_a.fcba and                                                                                                                      "&_
					"            fdl.fccode = fdl_a.fccode and                                                                                                                  "&_
					"            fdl.date_transaction = fdl_a.date_transaction)                                                                                                 "&_                                         
					"            where  to_date(to_char(fdl.date_transaction, 'yyyymm'),'yyyyMM') between to_date('201901','yyyyMM') and to_date('201906','yyyyMM')             "&_
					"            group by                                                                                                                                       "&_
					"            fdl.fcba,                                                                                                                                      "&_
					"            fdl.fccode,                                                                                                                                    "&_                                                                     
					"            fdl.date_transaction,                                                                                                                          "&_
					"            fdl.lineno                                                                                                                                     "&_
					"            union all                                                                                                                                      "&_
					"            select                                                                                                                                         "&_
					"            fdl.fcba,                                                                                                                                      "&_
					"            fdl.fccode,                                                                                                                                    "&_                                                                     
					"            fdl.date_transaction,                                                                                                                          "&_
					"            fdl.lineno,                                                                                                                                    "&_                                                  
					"            sum (fdl.monthlydepreciation) monthlydepreciation                                                                                              "&_                                                           
					"            from fixedasset_depreciation_ledger fdl                                                                                                        "&_                                                                                                                
					"            where to_date(to_char(fdl.date_transaction, 'yyyymm'),'yyyyMM') = to_date('"&v_txtYear&v_txtMonth&"','yyyyMM')                                                   "&_
					"            group by                                                                                                                                       "&_
					"            fdl.fcba,                                                                                                                                      "&_
					"            fdl.fccode,                                                                                                                                    "&_                                                                     
					"            fdl.date_transaction,                                                                                                                          "&_
					"            fdl.lineno                                                                                                                                     "&_
					"        ) fdl                                                                                                                                              "&_
					"        inner join fixedasset fa on (                                                                                                                      "&_
					"        fa.fcba = fdl.fcba and                                                                                                                             "&_
					"        fa.fccode = fdl.fccode)                                                                                                                            "&_
					"        inner join assetgroup ag on  fa.assetgroupcode = ag.fccode                                                                                         "&_
					"        group by fa.fcba, fa.assetgroupcode, fa.assetgroupname, fa.assetcode, fa.assetname, fa.fccode, fa.fcname, fa.transdate                             "&_
					"    )                                                                                                                                                      "&_
					"    group by  fcba,  assetgroupcode,  assetgroupname,  assetcode,  assetname,  fccode,  fcname,  transdate													"&_
					")																															  								"

	strEPS2 =   "        ,SUM(acquisitionvalue_prev) acquisitionvalue_prev," & _
				"        SUM(adjustment_amount_prev) adjustment_amount_prev," & _
				"        SUM(fixassetamount_prev) fixassetamount_prev," & _
				"        SUM(accumulated_depreciation_prev) accumulated_depreciation_prev," & _
				"        SUM(yearlydepreciation_prev) yearlydepreciation_prev," & _
				"        SUM(acquisitionvalue_curr) acquisitionvalue_curr," & _
				"        SUM(adjustment_amount_curr) adjustment_amount_curr ," & _
				"        SUM(fixassetamount_curr) fixassetamount_curr," & _
				"        SUM(accumulated_depreciation_curr) accumulated_depreciation_curr," & _
				"        SUM(yearlydepreciation_curr) yearlydepreciation_curr," & _
				"        SUM(monthlydepreciation_curr) monthlydepreciation_curr," & _
				" 		 SUM(bookvalue) bookvalue," & _
				"		 SUM(addinValue) addinValue," & _
         		"		 SUM(netbookvalue) netbookvalue"

	strEPS =  	" SELECT GroupPlantation " & selectCons & strEPS2 &"  From (" & strMain & ")" & whereCons & "" & _
				" Group by GroupPlantation" & selectCons & "" &_
				" Order by GroupPlantation" & orderCons &""
 
Set rsEPS = Server.CreateObject("ADODB.Recordset")
'response.write strEPS

rsEPS.Open strEPS , connstr, adOpenStatic  

%>

<html>

<head>
<title>ePlaS Reporting - Asset</title>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<link rel="stylesheet" href="/css/jquery-ui.css" />
<script src="/js/jquery-1.9.1.js"></script>
<script src="/js/jquery-ui.js"></script>
<script language="JavaScript" src='/js/x.js'></script>
<script language="JavaScript" src='/js/xtableheaderfixed.js'></script>
<script language="JavaScript">
xAddEventListener(window, 'load',
  function() {
    new xTableHeaderFixed('xthf-site', window);
    new xTableHeaderFixed('xthf-green', window);
    new xTableHeaderFixed('xthf-blue', window);
  }, false
);
</script>
 <script>
$(function() {
$( ".datepicker" ).datepicker();
});
</script>
<!--<script language="JavaScript" src="/js/date-picker.js"></script> 
<script language="JavaScript">

var newwin;
function launchwin(winurl,winname,winfeatures)
{
	//This launches a new window and then
	//focuses it if window.focus() is supported.
	newwin = window.open(winurl,winname,winfeatures);

}
</script>-->
</head>

<body>
	<div id="wrapper">
	<!--#include file="../print_excel.asp"-->
<form name="frmEntry" method="post" action="<%=ScriptName%>">


	<table id="menu-table" class="centered-margin">
		<tr>
			<td align="center"><a href="/login.asp">Back to Home</a></td>
			<td align="center"><a href="/report/Dashboard.asp">Back to Report</a></td>
		</tr>
	</table>
	

<!--#include file="FormSelectNewSum.html"-->
<table border="0" id="form-table" class="centered-margin">
    
	<input type="hidden" name="txtComp" value="<%=v_txtComp%>">
	<tr>
	    <!--%response.write v_txtfcba %-->
	    <!--%response.write v_txtyear %-->
	    <!--%response.write v_txtmonth %-->
		<!--%response.write len(trim(request("txtYear"))) %--> 
			<% 
				Dim mon, Yr, MinYr, MaxYr, i, MM
				mon = Month(Session("v_Costbook"))
				If len(v_txtMonth)=0 then v_txtMonth=mon

				Yr = Year(Session("v_Costbook"))
				MinYr = 2015
				MaxYr = Yr  + 2
				If len(v_txtYear)= 0 then v_txtYear = Yr
			%>
		<td><font face="verdana" size="1">Year</font></td>
		<td width="4">
			<select name="txtYear">
				<% For i = MinYr to MaxYr %>
					<Option Value="<%=i%>" <%If cint(i) = cint(v_txtYear) then%> Selected <%End If%>><%=i%></option>
				<% Next %>
			</select>
		</td>
		<td><font face="verdana" size="1">Month</font></td>
		<td width="3">
					<select name="txtMonth">
				<%
					For i = 1 to 15
				%>
						<Option Value="<%=i%>" <%If cint(i) = cint(v_txtMonth) then%> Selected <%End If%>><%=BulanAP(i)%></option>
				<%
					Next
				%>
					</select>
		</td>
		<td><font face="verdana" size="1">Asset Group Code</font></td>
		<td width="3">
				<select name="txtAssetGroup">	
					<option value="--ALL--">--ALL--</option>
					<%
						Dim rsVH1
						Set rsVH1 = Server.CreateObject("ADODB.Recordset")
						rsVH1.CursorLocation = adUseClient
						strsql =  "SELECT ag.FCCODE, ag.FCNAME from assetgroup ag   order by ag.FCCODE asc"

						rsVH1.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
						if not rsVH1.eof then
							DO while not rsVH1.eof
							if v_txtAssetGroup = rsVH1("FCCODE") then
								%>			
									<option value="<%=rsVH1("FCCODE")%>" Selected><%=rsVH1("FCCODE")%> - <%=rsVH1("FCNAME")%></option>
								<%else%>
									<option value="<%=rsVH1("FCCODE")%>"><%=rsVH1("FCCODE")%> - <%=rsVH1("FCNAME")%></option>
								<%end if%>
				<%				rsVH1.movenext
							loop
						End if
						rsVH1.Close
						set rsVH1 = nothing
				%>
				</select>
		</td>
		<td><font face="verdana" size="1">Asset Code</font></td>
		<td width="3">
				<select name="txtAsset">	
					<option value="--SUMMARY--">--SUMMARY--</option>
					<option value="--ALL--">--ALL--</option>
					<%
						Dim rsFA
						Set rsFA = Server.CreateObject("ADODB.Recordset")
						rsFA.CursorLocation = adUseClient
						strsql =  "SELECT ag.FCCODE, ag.FCNAME from assetgroup ag  order by ag.FCCODE asc"

						rsFA.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
						if not rsFA.eof then
							DO while not rsFA.eof
							if v_txtAssetGroup = rsFA("FCCODE") then
								%>			
									<option value="<%=rsFA("FCCODE")%>" Selected><%=rsFA("FCCODE")%> - <%=rsFA("FCNAME")%></option>
								<%else%>
									<option value="<%=rsFA("FCCODE")%>"><%=rsFA("FCCODE")%> - <%=rsFA("FCNAME")%></option>
								<%end if%>
					<%		rsFA.movenext
							loop
						End if
						rsFA.Close
						set rsFA = nothing
				%>
				</select>
		</td>
		<td align="center">
			<input type="submit" value="View" />
		</td> 
	</tr>
</table>



<table border="0" width="100%" id="product-table" class="xthf-site">
	<thead>
		<tr>
			<th colspan="22" style="background:#ffffff; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
			<div id="header-wrapper">
				<div id="header">
					<% if  CompanyName <>  "" then %>
						<h4>PT.<%=CompanyName%></h4> 
					<% End if%>
					<% If  FcbaName <> ""  then %>
						   <h5><%=FcbaName%></h5>   
					<% End if%>
				</div>
				<div id="subheader"class="centered">
					<h2>ASSET</h2>
					<% if v_txttodaydate <> "" and v_txtFcba <> "" then %>
						<h5>PERIOD : <%=v_txttodaydate%>, BUSINESSUNIT : <%=v_txtFcba%></h5>
					<% end if %>
					<h3></br></h3>
					<h3></br></h3>
				</div>
			</div> 
				<%
					if len(trim(request("txtYear"))) = 0 then
						v_txtyear = year(Now)	 
					end if

					if len(trim(request("txtmonth"))) = 0 then
						v_txtmonth = Month(Now)	 
					end if
				%> 
			</th>
		</tr>
		<tr>
			<% if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then %>
			      <th rowspan="2" style="background-color:#00A300;color:#fff;"><%=titlePlantation%></th>
			<% 			
		    end if %>
			
			<% if v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  "" then %>
			      <th rowspan="2" style="background-color:#00A300;color:#fff;"><%=titleGroupRegion%></th>
			<% 			
		    end if %>
			
			<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
			      <th rowspan="2" style="background-color:#00A300;color:#fff;"><%=titleRegion%></th>
			<%
		    end if %>
			
			<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then %>
			       <th rowspan="2" style="background-color:#00A300;color:#fff;"><%=titleCompany%></th>
			<% 		
		    end if %>
			
			<% if v_txtfcba <>  "--SUMMARY--" and v_txtfcba <>  "" then %>
			       <th rowspan="2" style="background-color:#00A300;color:#fff;"><%=titleSite%></th>
			<% 		
		    end if %>

			
			<% if v_txtDivision <> "--SUMMARY--"   and v_txtDivision <> "" then %>
			    <th rowspan="2" style="background-color:#00A300;color:#fff;">AFDELING</th>
			<% 		
		    end if %>
			
			
			<% if v_txtPlantingDate <> "--SUMMARY--" And v_txtPlantingDate <> "" then %>
			    <th rowspan="2" style="background-color:#00A300;color:#fff;">TT</th>
			<% 		
		    end if %>


			<% if v_txtBlok <> "--SUMMARY--" And v_txtBlok <> "" then %>
			    <th rowspan="2" style="background-color:#00A300;color:#fff;">BLOK</th>
			<% 		
		    end if %>
			<th rowspan="2" style="background-color:#00A300;color:#fff;">ASSET GROUP CODE</th>
			<% if v_txtAsset <> "--SUMMARY--" And v_txtAsset <> "" then %>
			   <th rowspan="2" style="background-color:#00A300;color:#fff;">ASSET CODE</th>
			   <th rowspan="2" style="background-color:#00A300;color:#fff;">DESCRIPTION</th>
			   <th rowspan="2" style="background-color:#00A300;color:#fff;">DATE</th>
			<% 		
		    end if %>
			<th style="background-color:#00A300;color:#fff;">ORIGINAL COST</th>
			<th colspan="4" style="background-color:#00A300;color:#fff;">PREVIOUS FINANCIAL YEAR</th>
			<th colspan="8" style="background-color:#00A300;color:#fff;">CURRENT FINANCIAL YEAR</th>		
		</tr>  
		<tr>
			<th style="background-color:#00A300;color:#fff;">CHANGE COST</th>
			<th style="background-color:#00A300;color:#fff;">ADDITIONAL AMOUNT</th>
			<th style="background-color:#00A300;color:#fff;">SOLD WRITE OFF AMOUNT</th>
			<th style="background-color:#00A300;color:#fff;">SOLD AMOUNT DEPRE.</th>
			<th style="background-color:#00A300;color:#fff;">TOTAL ACC. DEPRE.</th>
			<th style="background-color:#00A300;color:#fff;">BOOK VALUE</th>
			<th style="background-color:#00A300;color:#fff;">ADD IN</th>
			<th style="background-color:#00A300;color:#fff;">ADDITIONAL AMOUNT</th>
			<th style="background-color:#00A300;color:#fff;">SOLD WRITE OFF AMOUNT</th>
			<th style="background-color:#00A300;color:#fff;">SOLD AMOUNT DEPRE.</th>
			<th style="background-color:#00A300;color:#fff;">ACC. DEPRE.</th>
			<th style="background-color:#00A300;color:#fff;">NET BOOK VALUE</th>
			<th style="background-color:#00A300;color:#fff;">MONTHLY DEPRE.</th>
		</tr>
	</thead>
<tbody>
<%  
	if not rsEPS.EOF then
		NO = 1
		nacquisitionvalue_prev = 0
		nadjustment_amount_prev = 0
		nfixassetamount_prev = 0
		naccumulated_depreciation_prev = 0
		nyearlydepreciation_prev = 0

		nbookvalue = 0
		nadjustment_amount = 0
		nfixassetamount = 0
		naccumulated_depreciation = 0
		nyearlydepreciation = 0
		nnetbookvalue = 0
		nmonthlydepreciation = 0
		
		SprevPlantation= ""
		SprevGroupRegion= ""
		SprevRegion= ""
		SprevCompany= ""
		sprevFcba = ""
		sprevDivision = ""
		sprevPlantingYear = ""
		sprevFieldCode = ""
		rows = 1
		
		while not rsEPS.EOF
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
%>
		<tr class="<%=cClass%>">
			<%  
					if rows = 1 then
						sqlrow = "  select count(assetgroupname) rowsd  from ( " & strEPS & " ) where assetgroupcode = '" & rsEPS("assetgroupcode") & "'"
						' response.write sqlrow
						Set rsrow = Server.CreateObject("ADODB.Recordset")
						rsrow.Open sqlrow, connstr, adOpenStatic
						if not rsrow.EOF then
							temprows = CDbl(rsrow("rowsd"))
						else
							temprows = 0
						end if
					end if
			%>
				<% if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  "" then %>
					<% tempPlantation = rsEPS("Plantation")
					   tempGroupRegion = "--ALL--"
					   tempRegion = "--SUMMARY--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					   tempDivision = "--SUMMARY--"
					   tempPlantingDate = "--SUMMARY--"
					   tempBlok = "--SUMMARY--"
					%>
					<% if SprevPlantation <> rsEPS("Plantation") then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
							<%=rsEPS("Plantation")%>
							</a>
						</td>
						<% 
							SprevPlantation = rsEPS("plantation")
						%>	
						
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  "" then %>
				    <%  if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then
							tempPlantation = rsEPS("Plantation")
						else
							tempPlantation = ""
						end if
					   tempGroupRegion = rsEPS("GroupRegion")
					   tempRegion = "--ALL--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					   tempDivision = "--SUMMARY--"
					   tempPlantingDate = "--SUMMARY--"
					   tempBlok = "--SUMMARY--"
					%>
					
					<% if SprevGroupRegion <> rsEPS("GroupRegion") then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
							<%=rsEPS("groupregion")%>
							</a>
						</td>
						<% SprevGroupRegion = rsEPS("GroupRegion") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
				    <% if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
					      tempPlantation = rsEPS("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
					       tempGroupRegion = rsEPS("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   tempRegion = rsEPS("Region") 
					   tempCompany ="--ALL--"
					   tempfcba = "--SUMMARY--"
					   tempDivision = "--SUMMARY--"
					   tempPlantingDate = "--SUMMARY--"
					   tempBlok = "--SUMMARY--"
					%>
					
					<% if SprevRegion <> rsEPS("Region") then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
							<%=rsEPS("region")%>
							</a>
						</td>
						<% SprevRegion = rsEPS("region") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then %>
				    <% 
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
					      tempPlantation = rsEPS("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
					       tempGroupRegion = rsEPS("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
					       tempRegion = rsEPS("Region") 
					   else
					       tempRegion = ""
					   end if
					   tempCompany =rsEPS("Company")
					   tempfcba = "--ALL--"
					   tempDivision = "--SUMMARY--"
					   tempPlantingDate = "--SUMMARY--"
					   tempBlok = "--SUMMARY--"					   
					 %>
					
					<% if SprevCompany <> rsEPS("Company") then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
							<%=rsEPS("company")%>
							</a>
						</td>
						<% SprevCompany = rsEPS("Company") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	

				<% end if %>
				
				<% if v_txtFcba <>  "--SUMMARY--" and v_txtFcba <>  "" then %>
				    <% if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
					      tempPlantation = rsEPS("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
					       tempGroupRegion = rsEPS("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
					       tempRegion = rsEPS("Region") 
					   else
					       tempRegion = ""
					   end if
					   if  v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then
					       tempCompany =rsEPS("Company")
					   else
					       tempCompany =""
					   end if
					   tempfcba = rsEPS("fcba")
					   tempDivision = "--ALL--"
					   tempPlantingDate = "--SUMMARY--"
					   tempBlok = "--SUMMARY--"					   
					 %>
					
					<% if sprevFcba <> rsEPS("fcba") then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
						    <%=rsEPS("fcba")%>
							</a>
						</td>
						<% sprevFcba = rsEPS("fcba") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	 
				<% end if %>
	 
					<td align="left"><%=rsEPS("assetgroupname")%></td>
			
			
				<%
				nSpanTotalCol = 1
				
				if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then
					nSpanTotalCol = nSpanTotalCol + 1
				end if
				
				if v_txtRegion <> "--SUMMARY--" and v_txtRegion <> "" then
					nSpanTotalCol = nSpanTotalCol + 1
				end if
				
				if v_txtGroupRegion <> "--SUMMARY--" and v_txtGroupRegion <> "" then
					nSpanTotalCol = nSpanTotalCol + 1
				end if
				
				if v_txtCompany <> "--SUMMARY--" and v_txtCompany <> "" then
					nSpanTotalCol = nSpanTotalCol + 1
				end if
				
				if v_txtFcba <> "--SUMMARY--" and v_txtFcba <> "" then
					nSpanTotalCol = nSpanTotalCol + 1
				end if
				
				
				if v_txtAsset <> "--SUMMARY--" and v_txtAsset <> "" then
			%>
				<td align="center"><%=rsEPS("fcCode")%></td>
				<td align="left"><%=rsEPS("fcname")%></td>
				<td align="center"><%=rsEPS("transdate")%></td>
			<% 
					nSpanTotalCol = nSpanTotalCol +3
					nSpanTotal = nSpanTotal + 3
				end if 
			%>					
				<td align="right"><%=FormatNumber(cdbl(to_nol(rsEPS("acquisitionvalue_prev"))),2)%></td>
				<td align="right"><%=FormatNumber(cdbl(to_nol(rsEPS("adjustment_amount_prev"))),2)%></td>
				<td align="right"><%=FormatNumber(cdbl(to_nol(rsEPS("fixassetamount_prev"))),2)%></td> 
				<td align="right"><%=FormatNumber(cdbl(to_nol(rsEPS("accumulated_depreciation_prev"))),2)%></td>
				<td align="right"><%=FormatNumber(cdbl(to_nol(rsEPS("yearlydepreciation_prev"))),2)%></td> 
				<td align="right"><%=FormatNumber(cdbl(to_nol(rsEPS("bookvalue"))),2)%></td> 
				<td align="right"><%=FormatNumber(cdbl(to_nol(rsEPS("addinValue"))),2)%></td>
				<td align="right"><%=FormatNumber(cdbl(to_nol(rsEPS("adjustment_amount_curr"))),2)%></td>
				<td align="right"><%=FormatNumber(cdbl(to_nol(rsEPS("fixassetamount_curr"))),2)%></td> 
				<td align="right"><%=FormatNumber(cdbl(to_nol(rsEPS("accumulated_depreciation_curr"))),2)%></td>
				<td align="right"><%=FormatNumber(cdbl(to_nol(rsEPS("yearlydepreciation_curr"))),2)%></td> 
				<td align="right"><%=FormatNumber(cdbl(to_nol(rsEPS("netbookvalue"))),2)%></td>
				<td align="right"><%=FormatNumber(cdbl(to_nol(rsEPS("monthlydepreciation_curr"))),2)%></td> 
		</tr>

	<%	if rows = temprows then
			nacquisitionvalue_prev = nacquisitionvalue_prev + cdbl(to_nol(rsEPS("acquisitionvalue_prev")))
			nadjustment_amount_prev = nadjustment_amount_prev + cdbl(to_nol(rsEPS("adjustment_amount_prev")))	
			nfixassetamount_prev = nfixassetamount_prev + cdbl(to_nol(rsEPS("fixassetamount_prev")))
			naccumulated_depreciation_prev = naccumulated_depreciation_prev + cdbl(to_nol(rsEPS("accumulated_depreciation_prev")))
			nyearlydepreciation_prev = nyearlydepreciation_prev + cdbl(to_nol(rsEPS("yearlydepreciation_prev")))	

			nbookvalue = nbookvalue + cdbl(to_nol(rsEPS("bookvalue")))
			naddinValue = naddinValue + cdbl(to_nol(rsEPS("addinValue")))
			nadjustment_amount = nadjustment_amount + cdbl(to_nol(rsEPS("adjustment_amount_curr")))
			nfixassetamount = nfixassetamount + cdbl(to_nol(rsEPS("fixassetamount_curr")))
			naccumulated_depreciation = naccumulated_depreciation + cdbl(to_nol(rsEPS("accumulated_depreciation_curr")))
			nyearlydepreciation = nyearlydepreciation + cdbl(to_nol(rsEPS("yearlydepreciation_curr")))
			nnetbookvalue = nnetbookvalue + cdbl(to_nol(rsEPS("netbookvalue")))
			nmonthlydepreciation = nmonthlydepreciation + cdbl(to_nol(rsEPS("monthlydepreciation_curr")))
			rows = 1
			
			if v_txtAsset <> "--SUMMARY--" and v_txtAsset <> "" then
	%>
			<tr bgcolor="#F4F4A0">
				<td align="center" colspan="<%=nSpanTotalCol%>"><font style="font-weight: bold;">TOTAL</font></td>				
				<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(nacquisitionvalue_prev)),2)%></font></td> 
				<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(nadjustment_amount_prev)),2)%></font></td>  
				<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(nfixassetamount_prev)),2)%></font></td> 
				<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(naccumulated_depreciation_prev)),2)%></font></td> 
				<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(nyearlydepreciation_prev)),2)%></font></td> 
				<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(nbookvalue)),2)%></font></td>
				<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(naddinValue)),2)%></font></td>						
				<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(nadjustment_amount)),2)%></font></td> 
				<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(nfixassetamount)),2)%></font></td> 
				<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(naccumulated_depreciation)),2)%></font></td> 
				<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(nyearlydepreciation)),2)%></font></td> 
				<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(nnetbookvalue)),2)%></font></td>  
				<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(nmonthlydepreciation)),2)%></font></td>  
			</tr>
		<% end if %>
			
		<%
			nacquisitionvalue_prev = cdbl("0")
			nadjustment_amount_prev = cdbl("0")
			nfixassetamount_prev = cdbl("0")
			naccumulated_depreciation_prev = cdbl("0")
			nyearlydepreciation_prev = cdbl("0") 
			nbookvalue = cdbl("0")
			naddinValue = cdbl("0")
			nadjustment_amount = cdbl("0") 
			nfixassetamount = cdbl("0") 
			naccumulated_depreciation = cdbl("0") 
			nyearlydepreciation = cdbl("0") 
			nnetbookvalue = cdbl("0") 
			nmonthlydepreciation = cdbl("0")
		else
		' cari sub total
			nacquisitionvalue_prev = nacquisitionvalue_prev + cdbl(to_nol(rsEPS("acquisitionvalue_prev")))
			nadjustment_amount_prev = nadjustment_amount_prev + cdbl(to_nol(rsEPS("adjustment_amount_prev")))	
			nfixassetamount_prev = nfixassetamount_prev + cdbl(to_nol(rsEPS("fixassetamount_prev")))
			naccumulated_depreciation_prev = naccumulated_depreciation_prev + cdbl(to_nol(rsEPS("accumulated_depreciation_prev")))
			nyearlydepreciation_prev = nyearlydepreciation_prev + cdbl(to_nol(rsEPS("yearlydepreciation_prev")))	

			nbookvalue = nbookvalue + cdbl(to_nol(rsEPS("bookvalue")))
			naddinValue = naddinValue + cdbl(to_nol(rsEPS("addinValue")))
			nadjustment_amount = nadjustment_amount + cdbl(to_nol(rsEPS("adjustment_amount_curr")))
			nfixassetamount = nfixassetamount + cdbl(to_nol(rsEPS("fixassetamount_curr")))
			naccumulated_depreciation = naccumulated_depreciation + cdbl(to_nol(rsEPS("accumulated_depreciation_curr")))
			nyearlydepreciation = nyearlydepreciation + cdbl(to_nol(rsEPS("yearlydepreciation_curr")))
			nnetbookvalue = nnetbookvalue + cdbl(to_nol(rsEPS("netbookvalue")))
			nmonthlydepreciation = nmonthlydepreciation + cdbl(to_nol(rsEPS("monthlydepreciation_curr")))			
			rows = rows + 1
		End if
							  
			' cari grand total 
				gt_acquisitionvalue_prev = cdbl(gt_acquisitionvalue_prev) + cdbl(rsEPS("acquisitionvalue_prev"))
				gt_adjustment_amount_prev = cdbl(gt_adjustment_amount_prev) + cdbl(rsEPS("adjustment_amount_prev"))
				gt_fixassetamount_prev = cdbl(gt_fixassetamount_prev) + cdbl(rsEPS("fixassetamount_prev"))
				gt_accumulated_depreciation_prev = cdbl(gt_accumulated_depreciation_prev) + cdbl(rsEPS("accumulated_depreciation_prev"))
				gt_yearlydepreciation_prev = cdbl(gt_yearlydepreciation_prev) + cdbl(rsEPS("yearlydepreciation_prev")) 
				gt_bookvalue = cdbl(gt_bookvalue) + cdbl(rsEPS("bookvalue"))
				gt_addinValue = cdbl(gt_addinValue) + cdbl(rsEPS("addinValue"))
				gt_adjustment_amount_curr = cdbl(gt_adjustment_amount_curr) + cdbl(rsEPS("adjustment_amount_curr"))
				gt_fixassetamount_curr = cdbl(gt_fixassetamount_curr) + cdbl(rsEPS("fixassetamount_curr")) 
				gt_accumulated_depreciation_curr = cdbl(gt_accumulated_depreciation_curr) + cdbl(rsEPS("accumulated_depreciation_curr")) 
				gt_yearlydepreciation_curr = cdbl(gt_yearlydepreciation_curr) + cdbl(rsEPS("yearlydepreciation_curr")) 
				gt_netbookvalue = cdbl(gt_netbookvalue) + cdbl(rsEPS("netbookvalue")) 
				gt_monthlydepreciation_curr = cdbl(gt_monthlydepreciation_curr) + cdbl(rsEPS("monthlydepreciation_curr"))
	
			rsEPS.MoveNext			
			Wend
		end if
		rsEPS.Close
		set rsEPS = Nothing
	%>

	<tr bgcolor="#F4F4A0">
			<td align="center" colspan="<%=nSpanTotalCol%>"><font style="font-weight: bold;">Grand Total</font></td>
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(gt_acquisitionvalue_prev)),2)%></font></td> 
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(gt_adjustment_amount_prev)),2)%></font></td>  
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(gt_fixassetamount_prev)),2)%></font></td> 
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(gt_accumulated_depreciation_prev)),2)%></font></td> 
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(gt_yearlydepreciation_prev)),2)%></font></td> 
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(gt_bookvalue)),2)%></font></td>
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(gt_addinValue)),2)%></font></td>				
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(gt_adjustment_amount_curr)),2)%></font></td> 
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(gt_fixassetamount_curr)),2)%></font></td> 
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(gt_accumulated_depreciation_curr)),2)%></font></td> 
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(gt_yearlydepreciation_curr)),2)%></font></td> 
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(gt_netbookvalue)),2)%></font></td>  
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(gt_monthlydepreciation_curr)),2)%></font></td>
	</tr>
	</tbody>
</table>
</form>
</div>
</body>
</html>

