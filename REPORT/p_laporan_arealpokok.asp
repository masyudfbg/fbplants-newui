<!--  -->
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%	
	Dim rsEPS, strEPS,rscomp , strcomp,	strEPS2
	Dim v_txtyear, v_txtmonth, v_txttodaydate, v_txtDivision, v_txtBlok, v_txtPlantingDate, NO, cClass
	Dim strMain, nSpanTotal
	Dim txtyear, txtmonth,txtday, txttodaydate, txtDivision, txtBlok, txtPlantingDate,txtComp,v_txtComp
    Dim tempDivision, tempBlok, tempPlantingDate ,tempownership
	Dim tempfcba,tempCompany,tempRegion,tempGroupRegion,tempPlantation,tempGroupPlantation
	Dim nSubMandays_Amount_hi, nSubMandays_Amount_shi
	Dim nOutput_Mandays_Amount_hi, nOutput_Mandays_Amount_shi
	DIM q,FcbaName,CompanyName,MgrName,KTUName,CropName,FcbaType
	Dim SprevFcba, sprevDivision, sprevPlantingYear, sprevFieldCode,SprevPlantingDate
	Dim SprevCompany,SprevRegion,SprevGroupRegion,SprevPlantation,Sprevownership,SprevBlok


		SprevPlantation= ""
		SprevGroupRegion= ""
		SprevRegion= ""
		SprevCompany= ""
		sprevFcba = ""
		sprevDivision = ""
		sprevPlantingYear = ""
		sprevFieldCode = ""
		Sprevownership = ""
		SprevPlantingDate = ""
		SprevBlok = ""
		
			v_txtfcba 			= Request("txtfcba")
			v_txtDivision		= Request("txtDivision")
			v_txtBlok			= Request("txtBlok")
			v_txtPlantingDate	= Request("txtPlantingDate")
			txtYear = Request("txtYear")
			txtMonth = Request("txtMonth")

			txttdate = Request("txttdate")
			' v_txttodaydate		= Request("txttodaydate")
			
			' txtyear 		= YEAR(Request("txttodaydate"))
			' txtmonth 		= MONTH(Request("txttodaydate"))
			' txtday			= DAY(Request("txttodaydate")) 
			' txttodaydate 	= Request("txttodaydate")
			
			' txtComp         = Request("txtCompany")
			nSpanTotal = 8
			' v_txtComp           =   txtComp
			' v_txtyear			=	txtyear
			' v_txtmonth			=	txtmonth
			If "" & txtYear = "" Then txtYear = Year(Session("v_Costbook"))
			If "" & txtMonth = "" Then txtMonth = Month(Session("v_Costbook"))

			txttdate = textDate(txtMonth&"/01/"&txtYear)

			
			
			'response.write  "GP " & v_txtGroupPlantation & "  P " & v_txtPlantation & "  GR " & v_txtGroupRegion & " R " & v_txtRegion & " C " & v_txtcompany & " S " & v_txtfcba  &  " D " & v_txtDivision & " B "  & v_txtBlok & " TT" & v_txtPlantingDate
		strDefaultBa = " SELECT" & _
				"     FCCODE as FCBA" & _
				" FROM BUSINESSUNIT BU" & _
				" INNER JOIN USER_FCBA_TAB UFT ON (BU.FCCODE = UFT.FCBA) " & _
				" WHERE" & _
				"     ISDEFAULT = 'TRUE'" & _
				"     AND FCTYPE = 'E'" & _
				"     AND FCUSERCODE = '" & Session("CodeName") & "'"

		Set rsDefaultBa = Server.CreateObject("ADODB.Recordset")
		rsDefaultBa.Open strDefaultBa, connstr, adOpenStatic

		If Not rsDefaultBa.EOF Then
			varDefaultBa = rsDefaultBa("FCBA")
		Else
			varDefaultBa = ""
		End If
		rsDefaultBa.Close
		set rsDefaultBa = Nothing

		If (v_txtFcba = "" Or Len(Trim(v_txtFcba)) = 0) Then
			v_txtFcba = varDefaultBa
		End If	
			
	       if  v_txtfcba <> "--ALL--"  and v_txtfcba <> "" then
		       whereCons = whereCons & " And fcba = '" & v_txtfcba & "'"
			end if
		    ' selectCons = selectCons & ",fcba"
		    ' orderCons  = orderCons & ",BUSINESSUNIT_API.get_OrderData(fcba)"
			q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & v_txtfcba & "' "
			set rsCat = connStr.Execute(q)
			if not rsCat.eof then
			   fcbaname = rsCat("fcName")
			   CompanyName=rsCat("FCCOMPANYNAME")
			   MgrName = rsCat("FCMANAGER")
			   KTUName = rsCat("FCKTU")
			   CropName=rsCat("FCCROP")
			   FcbaType=rsCat("fcType")
		else
		    nSpanTotal = nSpanTotal - 1 
		end if
			

			
			   if v_txtDivision <> "--SUMMARY--"  And v_txtDivision <> "" then
			      if v_txtDivision <>  "--ALL--"  then
					whereCons = whereCons & " AND DIVISION = '" & v_txtDivision & "'"
			      end if
			   else
			   nSpanTotal = nSpanTotal - 1  
			   end if   


				if v_txtPlantingDate <> "--SUMMARY--" and v_txtPlantingDate <> ""   then
				   if v_txtPlantingDate <>  "--ALL--" then
					  whereCons = whereCons & " AND to_char(field_api.getvalid_plantingyear (f.fccode, f.fcba, LAST_DAY (TO_DATE ('"& txtyear &"', 'YYYY')))) ='" & v_txtplantingdate & "'"
				   end if
				else
				   nSpanTotal = nSpanTotal - 1 
				end if
		
			
				if v_txtBlok <> "--SUMMARY--" and v_txtBlok <> "" then
				   if v_txtBlok <>  "--ALL--" then
					  whereCons = whereCons & " AND f.fccode ='" & v_txtBlok & "'"
				   end if
				else
					nSpanTotal = nSpanTotal - 1 
				end if
				
				
   		 strMain = 	"select " & strFieldCons & "f.fccode FIELDCODE,f.fcname, f.fcba, " & _
					"field_api.getValid_Afdeling(" & _
						"f.fccode,f.fcba,last_day(to_date('" & txttdate &"', 'MM/DD/YYYY'))) DIVISION," & _
					"field_api.GetValid_ownership(" & _
						"f.fccode,f.fcba,last_day(to_date('" & txttdate &"', 'MM/DD/YYYY'))) ownership," & _
					"field_api.getvalid_plantingyear(" & _
						"f.fccode,f.fcba,last_day(to_date('" & txttdate &"', 'MM/DD/YYYY'))) plantingyear," & _
					"field_api.getvalid_plantingpoint(" & _
						"f.fccode,f.fcba,last_day(to_date('" & txttdate &"', 'MM/DD/YYYY'))) plantingpoint," & _
					"field_api.getvalid_Ha(" & _
						"f.fccode,f.fcba,last_day(to_date('" & txttdate &"', 'MM/DD/YYYY'))) hectarageplanted," & _
					"field_api.GetValid_status(" & _
						"f.fccode,f.fcba,last_day(to_date('" & txttdate &"', 'MM/DD/YYYY'))) status," & _
					"field_api.getvalid_pokok(" & _
						"f.fccode,f.fcba,last_day(to_date('" & txttdate &"', 'MM/DD/YYYY'))) totalstand," & _
					"field_api.GetValid_harvestingbased_abw(" & _
						"f.fccode,f.fcba,last_day(to_date('" & txttdate &"', 'MM/DD/YYYY'))) harvestingbased_abw," & _
					"field_api.GetValid_CropType(" & _
						"f.fccode,f.fcba,last_day(to_date('" & txttdate &"', 'MM/DD/YYYY'))) croptype," & _
					"field_api.GetValid_terraintype(" & _
						"f.fccode,f.fcba,last_day(to_date('" & txttdate &"', 'MM/DD/YYYY'))) terraintype," & _
					"opharvrate_api.Get_Basis_Qty(field_api.GetValid_harvestingbased_abw(" & _
						"f.fccode,f.fcba,last_day(to_date('" & txttdate &"', 'MM/DD/YYYY')))," & _
					"field_api.GetValid_CropType(" & _
						"f.fccode,f.fcba,last_day(to_date('" & txttdate &"', 'MM/DD/YYYY')))," & _
					"field_api.GetValid_terraintype(" & _
						"f.fccode,f.fcba,last_day(to_date('" & txttdate &"', 'MM/DD/YYYY'))),f.fcba) basis_qty," & _
					"opharvrate_api.Get_Basis_Rate(" & _
						"field_api.GetValid_harvestingbased_abw(" & _
							"f.fccode,f.fcba,last_day(to_date('" & txttdate &"', 'MM/DD/YYYY')))," & _
						"field_api.GetValid_CropType(f.fccode,f.fcba,last_day(to_date('" & txttdate &"', 'MM/DD/YYYY')))," & _
						"field_api.GetValid_terraintype(f.fccode,f.fcba,last_day(to_date('" & txttdate &"', 'MM/DD/YYYY'))),f.fcba) Basis_Rate," & _
					"opharvrate_api.Get_Extra_Basis_Rate(" & _
						"field_api.GetValid_harvestingbased_abw(" & _
							"f.fccode,f.fcba,to_date('" & txttdate & "','MM/DD/YYYY'))," & _
						"field_api.GetValid_CropType(f.fccode,f.fcba,last_day(to_date('" & txttdate &"', 'MM/DD/YYYY')))," & _
						"field_api.GetValid_terraintype(f.fccode,f.fcba,last_day(to_date('" & txttdate &"', 'MM/DD/YYYY'))),f.fcba) Extra_Basis_Rate," & _
					"opharvrate_api.Get_Basis_Ha(" & _
						"field_api.GetValid_harvestingbased_abw(" & _
							"f.fccode,f.fcba,last_day(to_date('" & txttdate &"', 'MM/DD/YYYY')))," & _
						"field_api.GetValid_CropType(f.fccode,f.fcba,last_day(to_date('" & txttdate &"', 'MM/DD/YYYY')))," & _
						"field_api.GetValid_terraintype(f.fccode,f.fcba,last_day(to_date('" & txttdate &"', 'MM/DD/YYYY'))),f.fcba) Basis_Ha " & _
					" ,field_api.getvalid_Ha (f.fccode,f.fcba,last_day(TO_DATE ('01/01/" & txtyear &"', 'MM/DD/YYYY'))) ha_akhir_jan " & _
                    " ,field_api.getvalid_Ha (f.fccode,f.fcba,last_day(TO_DATE ('02/01/" & txtyear &"', 'MM/DD/YYYY'))) ha_akhir_feb " & _
                    " ,field_api.getvalid_Ha (f.fccode,f.fcba,last_day(TO_DATE ('03/01/" & txtyear &"', 'MM/DD/YYYY'))) ha_akhir_mar " & _
                    " ,field_api.getvalid_Ha (f.fccode,f.fcba,last_day(TO_DATE ('04/01/" & txtyear &"', 'MM/DD/YYYY'))) ha_akhir_apr " & _
                    " ,field_api.getvalid_Ha (f.fccode,f.fcba,last_day(TO_DATE ('05/01/" & txtyear &"', 'MM/DD/YYYY'))) ha_akhir_mei " & _
                    " ,field_api.getvalid_Ha (f.fccode,f.fcba,last_day(TO_DATE ('06/01/" & txtyear &"', 'MM/DD/YYYY'))) ha_akhir_jun " & _
                    " ,field_api.getvalid_Ha (f.fccode,f.fcba,last_day(TO_DATE ('07/01/" & txtyear &"', 'MM/DD/YYYY'))) ha_akhir_jul " & _
                    " ,field_api.getvalid_Ha (f.fccode,f.fcba,last_day(TO_DATE ('08/01/" & txtyear &"', 'MM/DD/YYYY'))) ha_akhir_aug " & _
                    " ,field_api.getvalid_Ha (f.fccode,f.fcba,last_day(TO_DATE ('09/01/" & txtyear &"', 'MM/DD/YYYY'))) ha_akhir_sep " & _
                    " ,field_api.getvalid_Ha (f.fccode,f.fcba,last_day(TO_DATE ('10/01/" & txtyear &"', 'MM/DD/YYYY'))) ha_akhir_okt " & _
                    " ,field_api.getvalid_Ha (f.fccode,f.fcba,last_day(TO_DATE ('11/01/" & txtyear &"', 'MM/DD/YYYY'))) ha_akhir_nov " & _
                    " ,field_api.getvalid_Ha (f.fccode,f.fcba,last_day(TO_DATE ('12/01/" & txtyear &"', 'MM/DD/YYYY'))) ha_akhir_des " & _
                    " ,field_api.getvalid_pokok (f.fccode,f.fcba,last_day(TO_DATE ('01/01/" & txtyear &"', 'MM/DD/YYYY'))) pkk_akhir_jan " & _
                    " ,field_api.getvalid_pokok (f.fccode,f.fcba,last_day(TO_DATE ('02/01/" & txtyear &"', 'MM/DD/YYYY'))) pkk_akhir_feb " & _
                    " ,field_api.getvalid_pokok (f.fccode,f.fcba,last_day(TO_DATE ('03/01/" & txtyear &"', 'MM/DD/YYYY'))) pkk_akhir_mar " & _
                    " ,field_api.getvalid_pokok (f.fccode,f.fcba,last_day(TO_DATE ('04/01/" & txtyear &"', 'MM/DD/YYYY'))) pkk_akhir_apr " & _
                    " ,field_api.getvalid_pokok (f.fccode,f.fcba,last_day(TO_DATE ('05/01/" & txtyear &"', 'MM/DD/YYYY'))) pkk_akhir_mei " & _
                    " ,field_api.getvalid_pokok (f.fccode,f.fcba,last_day(TO_DATE ('06/01/" & txtyear &"', 'MM/DD/YYYY'))) pkk_akhir_jun " & _
                    " ,field_api.getvalid_pokok (f.fccode,f.fcba,last_day(TO_DATE ('07/01/" & txtyear &"', 'MM/DD/YYYY'))) pkk_akhir_jul " & _
                    " ,field_api.getvalid_pokok (f.fccode,f.fcba,last_day(TO_DATE ('08/01/" & txtyear &"', 'MM/DD/YYYY'))) pkk_akhir_aug " & _
                    " ,field_api.getvalid_pokok (f.fccode,f.fcba,last_day(TO_DATE ('09/01/" & txtyear &"', 'MM/DD/YYYY'))) pkk_akhir_sep " & _
                    " ,field_api.getvalid_pokok (f.fccode,f.fcba,last_day(TO_DATE ('10/01/" & txtyear &"', 'MM/DD/YYYY'))) pkk_akhir_okt " & _
                    " ,field_api.getvalid_pokok (f.fccode,f.fcba,last_day(TO_DATE ('11/01/" & txtyear &"', 'MM/DD/YYYY'))) pkk_akhir_nov " & _
                    " ,field_api.getvalid_pokok (f.fccode,f.fcba,last_day(TO_DATE ('12/01/" & txtyear &"', 'MM/DD/YYYY'))) pkk_akhir_des " & _
				" from field f " & _
				" where field_api.GetValid_Activation(f.fccode,f.fcba,last_day(to_date('" & txttdate &"', 'MM/DD/YYYY')))='Y' "& whereCons &" AND to_char(f.valid_from,'yyyy')='"& txtyear &"')ff" & _
				" LEFT JOIN ( select fcba, division, fccode, lineno, hectarageplanted, totalstand FROM ( " & _
                	" (select fd.fcba, fd.fccode, fd.lineno, fd.hectarageplanted, fd.totalstand, fd.division FROM fielddetail fd" & _
	                " INNER JOIN   (select fcba, fccode, min(lineno) as linemin from fielddetail" & _
	                "  WHERE TO_CHAR (valid_from, 'yyyy') = '"& txtyear &"'"& whereCons &" " & _
	                " GROUP BY fcba, fccode, division)fdmin" & _
	                " ON (fdmin.fcba = fd.fcba AND fdmin.fccode = fd.fccode and fdmin.linemin = fd.lineno))))line " & _
                " ON (line.fcba = ff.fcba and line.fccode = ff.fieldcode)" & _
                "  LEFT JOIN(select fcba, fieldcode " & _
             	" ,SUM(decode(to_number(to_char(transdate,'mm')),'1',SUPPLYING,0)) as SUPPLYING_JAN" & _
                " ,SUM(decode(to_number(to_char(transdate,'mm')),'2',SUPPLYING,0)) as SUPPLYING_FEB" & _
                " ,SUM(decode(to_number(to_char(transdate,'mm')),'3',SUPPLYING,0)) as SUPPLYING_MAR" & _
                " ,SUM(decode(to_number(to_char(transdate,'mm')),'4',SUPPLYING,0)) as SUPPLYING_APR" & _
                " ,SUM(decode(to_number(to_char(transdate,'mm')),'5',SUPPLYING,0)) as SUPPLYING_MEI" & _
                " ,SUM(decode(to_number(to_char(transdate,'mm')),'6',SUPPLYING,0)) as SUPPLYING_JUN" & _
                " ,SUM(decode(to_number(to_char(transdate,'mm')),'7',SUPPLYING,0)) as SUPPLYING_JUL" & _
                " ,SUM(decode(to_number(to_char(transdate,'mm')),'8',SUPPLYING,0)) as SUPPLYING_AUG" & _
                " ,SUM(decode(to_number(to_char(transdate,'mm')),'9',SUPPLYING,0)) as SUPPLYING_SEP" & _
                " ,SUM(decode(to_number(to_char(transdate,'mm')),'10',SUPPLYING,0)) as SUPPLYING_OKT" & _
                " ,SUM(decode(to_number(to_char(transdate,'mm')),'11',SUPPLYING,0)) as SUPPLYING_NOV" & _
                " ,SUM(decode(to_number(to_char(transdate,'mm')),'12',SUPPLYING,0)) as SUPPLYING_DES" & _
                " ,SUM(decode(to_number(to_char(transdate,'mm')),'1',COMPLETING,0)) as COMPLETING_JAN" & _
                " ,SUM(decode(to_number(to_char(transdate,'mm')),'2',COMPLETING,0)) as COMPLETING_FEB" & _
                " ,SUM(decode(to_number(to_char(transdate,'mm')),'3',COMPLETING,0)) as COMPLETING_MAR" & _
                " ,SUM(decode(to_number(to_char(transdate,'mm')),'4',COMPLETING,0)) as COMPLETING_APR" & _
                " ,SUM(decode(to_number(to_char(transdate,'mm')),'5',COMPLETING,0)) as COMPLETING_MEI" & _
                " ,SUM(decode(to_number(to_char(transdate,'mm')),'6',COMPLETING,0)) as COMPLETING_JUN" & _
                " ,SUM(decode(to_number(to_char(transdate,'mm')),'7',COMPLETING,0)) as COMPLETING_JUL" & _
                " ,SUM(decode(to_number(to_char(transdate,'mm')),'8',COMPLETING,0)) as COMPLETING_AUG" & _
                " ,SUM(decode(to_number(to_char(transdate,'mm')),'9',COMPLETING,0)) as COMPLETING_SEP" & _
                " ,SUM(decode(to_number(to_char(transdate,'mm')),'10',COMPLETING,0)) as COMPLETING_OKT" & _
                " ,SUM(decode(to_number(to_char(transdate,'mm')),'11',COMPLETING,0)) as COMPLETING_NOV" & _
                " ,SUM(decode(to_number(to_char(transdate,'mm')),'12',COMPLETING,0)) as COMPLETING_DES" & _
                " ,SUM(decode(to_number(to_char(transdate,'mm')),'1',DEADTREE,0)) as DEADTREE_JAN" & _
                " ,SUM(decode(to_number(to_char(transdate,'mm')),'2',DEADTREE,0)) as DEADTREE_FEB" & _
                " ,SUM(decode(to_number(to_char(transdate,'mm')),'3',DEADTREE,0)) as DEADTREE_MAR" & _
                " ,SUM(decode(to_number(to_char(transdate,'mm')),'4',DEADTREE,0)) as DEADTREE_APR" & _
                " ,SUM(decode(to_number(to_char(transdate,'mm')),'5',DEADTREE,0)) as DEADTREE_MEI" & _
                " ,SUM(decode(to_number(to_char(transdate,'mm')),'6',DEADTREE,0)) as DEADTREE_JUN" & _
                " ,SUM(decode(to_number(to_char(transdate,'mm')),'7',DEADTREE,0)) as DEADTREE_JUL" & _
                " ,SUM(decode(to_number(to_char(transdate,'mm')),'8',DEADTREE,0)) as DEADTREE_AUG" & _
                " ,SUM(decode(to_number(to_char(transdate,'mm')),'9',DEADTREE,0)) as DEADTREE_SEP" & _
                " ,SUM(decode(to_number(to_char(transdate,'mm')),'10',DEADTREE,0)) as DEADTREE_OKT" & _
                " ,SUM(decode(to_number(to_char(transdate,'mm')),'11',DEADTREE,0)) as DEADTREE_NOV" & _
                " ,SUM(decode(to_number(to_char(transdate,'mm')),'12',DEADTREE,0)) as DEADTREE_DES" & _
                " FROM (select sdt.fcba, last_day(sdh.transdate) as Transdate, sdt.fieldcode" & _
                " ,SUM(decode(sdt.declarationtype,'1',quantity,0)) as SUPPLYING" & _
                " ,SUM(decode(sdt.declarationtype,'2',quantity,0)) as COMPLETING" & _
                " ,SUM(decode(sdt.declarationtype,'3',quantity,0)) as DEADTREE" & _
                " FROM supplyingdeclarationdtl sdt" & _
                " inner join supplyingdeclarationhdr sdh ON (sdh.documentno = sdt.documentno and sdh.fcba = sdt.fcba)" & _
                " where to_number(to_char(last_day(to_date('" & txttdate &"', 'MM/DD/YYYY')),'yyyymm')) <= to_number(to_char(last_day(to_date('" & txttdate &"', 'MM/DD/YYYY')),'yyyymm'))" & _
                " group by sdt.fcba, last_day(sdh.transdate), sdt.fieldcode)" & _
                " group by fcba, fieldcode " & _
       			" )sd ON (sd.fcba = ff.fcba and sd.fieldcode = ff.fieldcode)" & _
       		" LEFT JOIN (" & _
            "  select fcba, fieldcode " & _
            " ,SUM(decode(to_number(to_char(transdate,'mm')),'1',quantity,0)) as PLANTING_JAN" & _
            " ,SUM(decode(to_number(to_char(transdate,'mm')),'2',quantity,0)) as PLANTING_FEB" & _
            " ,SUM(decode(to_number(to_char(transdate,'mm')),'3',quantity,0)) as PLANTING_MAR" & _
            " ,SUM(decode(to_number(to_char(transdate,'mm')),'4',quantity,0)) as PLANTING_APR" & _
            " ,SUM(decode(to_number(to_char(transdate,'mm')),'5',quantity,0)) as PLANTING_MEI" & _
            " ,SUM(decode(to_number(to_char(transdate,'mm')),'6',quantity,0)) as PLANTING_JUN" & _
            " ,SUM(decode(to_number(to_char(transdate,'mm')),'7',quantity,0)) as PLANTING_JUL" & _
            " ,SUM(decode(to_number(to_char(transdate,'mm')),'8',quantity,0)) as PLANTING_AUG" & _
            " ,SUM(decode(to_number(to_char(transdate,'mm')),'9',quantity,0)) as PLANTING_SEP" & _
            " ,SUM(decode(to_number(to_char(transdate,'mm')),'10',quantity,0)) as PLANTING_OKT" & _
            " ,SUM(decode(to_number(to_char(transdate,'mm')),'11',quantity,0)) as PLANTING_NOV" & _
            " ,SUM(decode(to_number(to_char(transdate,'mm')),'12',quantity,0)) as PLANTING_DES" & _
            " ,SUM(decode(to_number(to_char(transdate,'mm')),'1',hectare,0)) as HECTARE_JAN" & _
            " ,SUM(decode(to_number(to_char(transdate,'mm')),'2',hectare,0)) as HECTARE_FEB" & _
            " ,SUM(decode(to_number(to_char(transdate,'mm')),'3',hectare,0)) as HECTARE_MAR" & _
            " ,SUM(decode(to_number(to_char(transdate,'mm')),'4',hectare,0)) as HECTARE_APR" & _
            " ,SUM(decode(to_number(to_char(transdate,'mm')),'5',hectare,0)) as HECTARE_MEI" & _
            " ,SUM(decode(to_number(to_char(transdate,'mm')),'6',hectare,0)) as HECTARE_JUN" & _
            " ,SUM(decode(to_number(to_char(transdate,'mm')),'7',hectare,0)) as HECTARE_JUL" & _
            " ,SUM(decode(to_number(to_char(transdate,'mm')),'8',hectare,0)) as HECTARE_AUG" & _
            " ,SUM(decode(to_number(to_char(transdate,'mm')),'9',hectare,0)) as HECTARE_SEP" & _
            " ,SUM(decode(to_number(to_char(transdate,'mm')),'10',hectare,0)) as HECTARE_OKT" & _
            " ,SUM(decode(to_number(to_char(transdate,'mm')),'11',hectare,0)) as HECTARE_NOV" & _
            " ,SUM(decode(to_number(to_char(transdate,'mm')),'12',hectare,0)) as HECTARE_DES" & _
            " from (select pdt.fcba, last_day(pdh.transdate) as Transdate, pdt.fieldcode, pdt.quantity, pdt.hectare" & _
            " from plantingdeclarationdtl pdt" & _
            " inner join plantingdeclarationhdr pdh ON (pdh.documentno = pdt.documentno and pdh.fcba = pdt.fcba)" & _
            " where  to_char(last_day(to_date('" & txttdate &"', 'MM/DD/YYYY')),'yyyy') = to_char(last_day(to_date('" & txttdate &"', 'MM/DD/YYYY')),'yyyy')" & _
            " group by pdt.fcba, last_day(pdh.transdate), pdt.fieldcode,  pdt.quantity, pdt.hectare)" & _
            " group by fcba, fieldcode " & _
       		" )pd ON (pd.fcba = ff.fcBA and pd.fieldcode = ff.fieldcode"
				' field_api.GetValid_status(f.fccode,f.fcba,to_date('" & txttodaydate & "', 'MM/DD/YYYY')) = 'TM' " & _
				' " and field_api.GetValid_harvestingbased_abw(f.fccode,f.fcba,to_date('" & txttodaydate & "', 'MM/DD/YYYY')) is not null " & _
				' " And 
				' & _
				' " and field_api.GetValid_ownership(f.fccode,f.fcba,to_date('" & txttodaydate & "', 'MM/DD/YYYY')) is not null"
				
	strEPS2  =  " line.lineno as Urut," & _
				" line.hectarageplanted as lineha," & _
				" line.totalstand as linepkk," & _
				" ff.fcba," & _
				" ff.fieldcode," & _
 				" ff.division," & _
				" ff.plantingyear," & _
				" ff.ownership," & _
				" ff.status," & _
				" ff.hectarageplanted," & _
		        " ff.totalstand," & _
		        " ff.plantingpoint" & _
		        " ,ff.ha_akhir_jan as hajan" & _
		        " ,ff.ha_akhir_feb as hafeb" & _
		        " ,ff.ha_akhir_mar as hamar" & _
		        " ,ff.ha_akhir_apr as haapr" & _
		        " ,ff.ha_akhir_mei as hamei" & _
		        " ,ff.ha_akhir_jun as hajun" & _
		        " ,ff.ha_akhir_jul as hajul" & _
		        " ,ff.ha_akhir_aug as haaug" & _
		        " ,ff.ha_akhir_sep as hasep" & _
		        " ,ff.ha_akhir_okt as haokt" & _
		        " ,ff.ha_akhir_nov as hanov" & _
		        " ,ff.ha_akhir_des as hades" & _
		        " ,ff.pkk_akhir_jan as pkkjan" & _
		        " ,ff.pkk_akhir_feb as pkkfeb" & _
		        " ,ff.pkk_akhir_mar as pkkmar" & _
		        " ,ff.pkk_akhir_apr as pkkapr" & _
		        " ,ff.pkk_akhir_mei as pkkmei" & _
		        " ,ff.pkk_akhir_jun as pkkjun" & _
		        " ,ff.pkk_akhir_jul as pkkjul" & _
		        " ,ff.pkk_akhir_aug as pkkaug" & _
		        " ,ff.pkk_akhir_sep as pkksep" & _
		        " ,ff.pkk_akhir_okt as pkkokt" & _
		        " ,ff.pkk_akhir_nov as pkknov" & _
		        " ,ff.pkk_akhir_des as pkkdes" & _
			    " ,sd.SUPPLYING_JAN as SUPPLYING_JAN " & _
			    " ,sd.SUPPLYING_FEB as SUPPLYING_FEB" & _
			    " ,sd.SUPPLYING_MAR as SUPPLYING_MAR" & _
			    " ,sd.SUPPLYING_APR as SUPPLYING_APR" & _
		        " ,sd.SUPPLYING_MEI as SUPPLYING_MEI" & _
		        " ,sd.SUPPLYING_JUN as SUPPLYING_JUN" & _
		        " ,sd.SUPPLYING_JUL as SUPPLYING_JUL" & _
		        " ,sd.SUPPLYING_AUG as SUPPLYING_AUG" & _
		        " ,sd.SUPPLYING_SEP as SUPPLYING_SEP" & _
		        " ,sd.SUPPLYING_OKT as SUPPLYING_OKT" & _
		        " ,sd.SUPPLYING_NOV as SUPPLYING_NOV" & _
		        " ,sd.SUPPLYING_DES as SUPPLYING_DES" & _
		        " ,sd.COMPLETING_JAN as COMPLETING_JAN" & _
		        " ,sd.COMPLETING_FEB as COMPLETING_FEB" & _
		        " ,sd.COMPLETING_MAR as COMPLETING_MAR" & _
		        " ,sd.COMPLETING_APR as COMPLETING_APR" & _
		        " ,sd.COMPLETING_MEI as COMPLETING_MEI" & _
		        " ,sd.COMPLETING_JUN as COMPLETING_JUN" & _
		        " ,sd.COMPLETING_JUL as COMPLETING_JUL" & _
		        " ,sd.COMPLETING_AUG as COMPLETING_AUG" & _
		        " ,sd.COMPLETING_SEP as COMPLETING_SEP" & _
		        " ,sd.COMPLETING_OKT as COMPLETING_OKT" & _
		        " ,sd.COMPLETING_NOV as COMPLETING_NOV" & _
		        " ,sd.COMPLETING_DES as COMPLETING_DES" & _
		        " ,sd.DEADTREE_JAN as DEADTREE_JAN" & _
		        " ,sd.DEADTREE_FEB as DEADTREE_FEB" & _
		        " ,sd.DEADTREE_MAR as DEADTREE_MAR" & _
		        " ,sd.DEADTREE_APR as DEADTREE_APR" & _
		        " ,sd.DEADTREE_MEI as DEADTREE_MEI" & _
		        " ,sd.DEADTREE_JUN as DEADTREE_JUN" & _
		        " ,sd.DEADTREE_JUL as DEADTREE_JUL" & _
		        " ,sd.DEADTREE_AUG as DEADTREE_AUG" & _
		        " ,sd.DEADTREE_SEP as DEADTREE_SEP" & _
		        " ,sd.DEADTREE_OKT as DEADTREE_OKT" & _
		        " ,sd.DEADTREE_NOV as DEADTREE_NOV" & _
		        " ,sd.DEADTREE_DES as DEADTREE_DES" & _
		        " ,pd.PLANTING_JAN as PLANTING_JAN" & _
		        " ,pd.PLANTING_FEB as PLANTING_FEB" & _
		        " ,pd.PLANTING_MAR as PLANTING_MAR" & _
		        " ,pd.PLANTING_APR as PLANTING_APR" & _
		        " ,pd.PLANTING_MEI as PLANTING_MEI" & _
		        " ,pd.PLANTING_JUN as PLANTING_JUN" & _
		        " ,pd.PLANTING_JUL as PLANTING_JUL" & _
		        " ,pd.PLANTING_AUG as PLANTING_AUG" & _
		        " ,pd.PLANTING_SEP as PLANTING_SEP" & _
		        " ,pd.PLANTING_OKT as PLANTING_OKT" & _
		        " ,pd.PLANTING_NOV as PLANTING_NOV" & _
		        " ,pd.PLANTING_DES as PLANTING_DES" & _
		        " ,pd.HECTARE_JAN as HECTARE_JAN" & _
		        " ,pd.HECTARE_FEB as HECTARE_FEB" & _
		        " ,pd.HECTARE_MAR as HECTARE_MAR" & _
		        " ,pd.HECTARE_APR as HECTARE_APR" & _
		        " ,pd.HECTARE_MEI as HECTARE_MEI" & _
		        " ,pd.HECTARE_JUN as HECTARE_JUN" & _
		        " ,pd.HECTARE_JUL as HECTARE_JUL" & _
		        " ,pd.HECTARE_AUG as HECTARE_AUG" & _
		        " ,pd.HECTARE_SEP as HECTARE_SEP" & _
		        " ,pd.HECTARE_OKT as HECTARE_OKT" & _
		        " ,pd.HECTARE_NOV as HECTARE_NOV" & _
		        " ,pd.HECTARE_DES as HECTARE_DES"
				' , " & _
				' "sum(basis_qty)  basis_qty," & _
				' "sum(Basis_Rate)  Basis_Rate," & _
				' "sum(Extra_Basis_Rate)  Extra_Basis_Rate," & _
				' "sum(Basis_Ha)  Basis_Ha"
				
	strEPS =  "select " & strEPS2 &"  From (" & strMain & " )" & _ 
			  " order by fcba, division, plantingyear, fcCode"
	' & _
			  ' "and USER_COMPANY_API.GetfcbaType('" & Session("CodeName") & "',ff.fcba) = 'E' "
			  ' "group by GroupPlantation, status, hectarageplanted, totalstand" & selectCons & _
			  ' ",ownership " & _
			  ' "Order by GroupPlantation" & orderCons & ",ownership, status, hectarageplanted, totalstand" & orderPlantingDate & orderBlok 


	' response.write strEPS				
Set rsEPS = Server.CreateObject("ADODB.Recordset")


rsEPS.Open strEPS , connstr, adOpenStatic 


%>

<html>

<head>
<title>ePlaS Reporting - Laporan Areal dan Pokok</title>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<link rel="stylesheet" href="/css/jquery-ui.css" />
<style>
.janjang{font-size:10px;background-color:#00A300;color:#fff; }
</style>
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
<form name="frmEntry" method="post" action="p_laporan_arealpokok.asp">


	<table id="menu-table" class="centered-margin">
		<tr>
			<td align="center"><a href="/login.asp">Back to Home</a></td>
			<td align="center"><a href="/report/Dashboard.asp">Back to Report</a></td>
		</tr>
	</table>
	
<table border="0" id="form-table" class="centered-margin">
    
	<tr>
		<th>TAHUN</th>
		<th>BULAN</th>
		<th>KEBUN</th>
		<th>AFDELING</th>
		<th>TAHUN TANAM</th>
		<th>BLOK</th>
		<th>&nbsp;</th>
	</tr>
	<input type="hidden" name="txtComp" value="<%=v_txtComp%>">
	<tr>
	    <!--%response.write v_txtfcba %-->
	    <!--%response.write v_txtyear %-->
	    <!--%response.write v_txtmonth %-->
		<!--%response.write len(trim(request("txtYear"))) %-->
		
		<td align="center">
		<% 
			Dim mon, Yr, MinYr, MaxYr, i, MM

			Yr = txtYear
			MinYr = 2015
			MaxYr = Yr  + 2
		%>
					<select name="txtYear">
		<%
			For i = MinYr to MaxYr
		%>
				<Option Value="<%=i%>" <%If cint(i) = cint(txtYear) then%> Selected <%End If%>><%=i%></option>
		<%
			Next
		%>
					</select>
		</td>
		<td align="center">
			<select name="txtMonth">
						<%
							For i = 1 to 12
						%>
							<Option Value="<%=i%>" <%If cint(i) = cint(txtMonth) then%> Selected <%End If%>><%=MonthName(i)%></option>
						<%
							Next
						%>
					</select>
		</td>
		<td>
			<select name="txtFcba" id="vFcba" onchange="javascript:filtering()">
			<%
			Dim rsFcba
			Set rsFcba = Server.CreateObject("ADODB.Recordset")
			rsFcba.CursorLocation = adUseClient
			strsql =    " SELECT '--ALL--' as fcba from DUAL UNION " & _  
			            " select fccode fcba from businessunit " & _
						" where " & _
						" USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',fccode) IS NOT NULL " 
			  
			
			rsFcba.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
			if not rsFcba.eof then
				DO while not rsFcba.eof
					if v_txtFcba = rsFcba("fcba") then
					%>			
						<option value="<%=rsFcba("fcba")%>" Selected><%=rsFcba("fcba")%></option>
					<% Else %>	
						<option value="<%=rsFcba("fcba")%>"><%=rsFcba("fcba")%></option>
					<% End If
					
					rsFcba.movenext
				loop
			End if
			rsFcba.Close
			set rsFcba = Nothing
			%>
			</select>	
		</td>
		<td align="center">
			<select name="txtDivision">	
				<%
					Dim rsDivision
					Set rsDivision = Server.CreateObject("ADODB.Recordset")
					rsDivision.CursorLocation = adUseClient
					
					if len(trim(v_txtDivision)) > 0 then 
						strsql =  " SELECT '--ALL--' as DIVISION from DUAL UNION SELECT DISTINCT Division as Division FROM FIELD WHERE FCBA = '" & v_txtFcba & "'"
					else
						strsql =  " SELECT '--ALL--' as DIVISION from DUAL UNION SELECT DISTINCT Division as Division FROM FIELD WHERE FCBA = '" & Session("fcBA") & "'"
					end if
					
					
					rsDivision.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
					if not rsDivision.eof then
						DO while not rsDivision.eof
						if v_txtDivision = rsDivision("Division") then
				%>			<option value="<%=rsDivision("Division")%>" Selected><%=rsDivision("Division")%></option>
						<%else%>
							<option value="<%=rsDivision("Division")%>"><%=rsDivision("Division")%></option>
						<%end if%>
				<%			rsDivision.movenext
						loop
					End if
					rsDivision.Close
					set rsDivision = nothing
				%>
			</select>		
		</td>


		<td align="center">
			<select name="txtPlantingDate">	
					<%
						Dim rsPlantingDate
						Set rsPlantingDate = Server.CreateObject("ADODB.Recordset")
						rsPlantingDate.CursorLocation = adUseClient
				
						if len(trim(v_txtPlantingDate)) > 0 then 
							strsql =  " SELECT '--ALL--' as PlantingDate from DUAL UNION SELECT DISTINCT to_char(PLANTINGDATE, 'yyyy') as PlantingDate FROM FIELD WHERE FCBA = '" & v_txtFcba & "'"
						else
							strsql =  " SELECT '--ALL--' as PlantingDate from DUAL UNION SELECT DISTINCT to_char(PLANTINGDATE, 'yyyy') as PlantingDate FROM FIELD WHERE FCBA = '" & Session("fcBA") & "'"
						end if
						
						
						rsPlantingDate.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
						if not rsPlantingDate.eof then
							DO while not rsPlantingDate.eof
							if v_txtPlantingDate = rsPlantingDate("PlantingDate") then
					%>			<option value="<%=rsPlantingDate("PlantingDate")%>" Selected><%=rsPlantingDate("PlantingDate")%></option>
							<%else%>
								<option value="<%=rsPlantingDate("PlantingDate")%>"><%=rsPlantingDate("PlantingDate")%></option>
							<%end if%>
					<%			rsPlantingDate.movenext
							loop
						End if
						rsPlantingDate.Close
						set rsPlantingDate = nothing
					%>
			</select>		
		</td>
		<td align="center">
			<select name="txtBlok">	
				<%
					Dim rsBlok
					Set rsBlok = Server.CreateObject("ADODB.Recordset")
					rsBlok.CursorLocation = adUseClient

			
					if len(trim(v_txtBlok)) > 0 then 
						strsql =  " SELECT '--ALL--' as Blok from DUAL UNION SELECT DISTINCT fccode as Blok FROM FIELD WHERE FCBA = '" & v_txtFcba & "'"
					else
						strsql =  " SELECT '--ALL--' as Blok from DUAL UNION SELECT DISTINCT fccode as Blok FROM FIELD WHERE FCBA = '" & Session("fcBA") & "'"
					end if
					
					rsBlok.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
					if not rsBlok.eof then
						DO while not rsBlok.eof
						if v_txtBlok = rsBlok("Blok") then
				%>			<option value="<%=rsBlok("Blok")%>" Selected><%=rsBlok("Blok")%></option>
						<%else%>
							<option value="<%=rsBlok("Blok")%>"><%=rsBlok("Blok")%></option>
						<%end if%>
				<%			rsBlok.movenext
						loop
					End if
					rsBlok.Close
					set rsBlok = nothing
				%>
			</select>		
		</td>
		
		
		<td align="center"><input type="submit" value="View" /></td>

	</tr>
</table>



<table border="1" width="100%" id="product-table" class="xthf-site">
<thead>
<tr>
<th colspan="141" style="background:#ffffff; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
<div id="header-wrapper">
		<div id="header">
			<% if  CompanyName <>  "" then %>
				<h4>PT.<%=CompanyName%></h4> 
			<% End if%>
			<% If  FcbaName <> ""  then %>
				   <h5><%=FcbaName%></h5>   
			<% End if%>
		</div>
		<div id="subheader" class="centered">
			<h2>LAPORAN AREAL POKOK</h2>
				<h4>PERIODE&nbsp;:&nbsp; BULAN &nbsp;<%=txtMonth%>&nbsp;<%=txtYear%></h4>
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
	
<!-- Pengaturan Baris 1-->
			<tr>
				<!-- <th rowspan="4" style="background-color:#00A300;color:#fff;">TANGGAL</th> -->
				<th rowspan="5" style="background-color:#00A300;color:#fff;">KEBUN</th>
				<th rowspan="5" style="background-color:#00A300;color:#fff;">AFDELING</th>
				<th rowspan="5" style="background-color:#00A300;color:#fff;">TAHUN TANAM</th>
				<th rowspan="5" style="background-color:#00A300;color:#fff;">BLOK</th>
				<th rowspan="5" style="background-color:#00A300;color:#fff;">OWNERSHIP</th>		
				<th rowspan="5" style="background-color:#00A300;color:#fff;">STATUS</th>
				<th rowspan="5" style="background-color:#00A300;color:#fff;">LUAS<br/><font size="1">(Ha)</font></th>
				<th colspan="2" style="background-color:#00A300;color:#fff;">TITIK TANAM </th>
				<th colspan="2" style="background-color:#00A300;color:#fff;">PROGRAM SISIP</th>
				<th colspan="13" style="background-color:#00A300;color:#fff;">JANUARI</th>
				<th colspan="10" style="background-color:#00A300;color:#fff;">FEBRUARI</th>
				<th colspan="10" style="background-color:#00A300;color:#fff;">MARET</th>
				<th colspan="10" style="background-color:#00A300;color:#fff;">APRIL</th>
				<th colspan="10" style="background-color:#00A300;color:#fff;">MEI</th>
				<th colspan="10" style="background-color:#00A300;color:#fff;">JUNI</th>
				<th colspan="10" style="background-color:#00A300;color:#fff;">JULI</th>
				<th colspan="10" style="background-color:#00A300;color:#fff;">AGUSTUS</th>
				<th colspan="10" style="background-color:#00A300;color:#fff;">SEPTEMBER</th>
				<th colspan="10" style="background-color:#00A300;color:#fff;">OKTOBER</th>
				<th colspan="10" style="background-color:#00A300;color:#fff;">NOVEMBER</th>
				<th colspan="10" style="background-color:#00A300;color:#fff;">DESEMBER</th>
				<th colspan="6" style="background-color:#00A300;color:#fff;">MUTASI S/D BI</th>
				<th rowspan="5" style="background-color:#00A300;color:#fff;">STATUS</th>
			</tr>	
			
			<!-- Pengaturan Baris 2-->
			<tr>
				<!-- Januari-->
				<th rowspan="4" style="background-color:#00A300;color:#fff;">POKOK</th>
				<th rowspan="4" style="background-color:#00A300;color:#fff;">SPH</th>
				<th rowspan="4" style="background-color:#00A300;color:#fff;">POKOK</th>
				<th rowspan="4" style="background-color:#00A300;color:#fff;">SPH</th>
				<th colspan="3" rowspan="2" style="background-color:#00A300;color:#fff;">AWAL</th>
				<th colspan="7" style="background-color:#00A300;color:#fff;">MUTASI</th>
				<th colspan="3" rowspan="2" style="background-color:#00A300;color:#fff;">AKHIR</th>
				
				<!-- Februari-->
				<th colspan="7" style="background-color:#00A300;color:#fff;">MUTASI</th>
				<th colspan="3" rowspan="2" style="background-color:#00A300;color:#fff;">AKHIR</th>

				<!-- Maret-->
				<th colspan="7" style="background-color:#00A300;color:#fff;">MUTASI</th>
				<th colspan="3" rowspan="2" style="background-color:#00A300;color:#fff;">AKHIR</th>

				<!-- April-->
				<th colspan="7" style="background-color:#00A300;color:#fff;">MUTASI</th>
				<th colspan="3" rowspan="2" style="background-color:#00A300;color:#fff;">AKHIR</th>

				<!-- Mei-->
				<th colspan="7" style="background-color:#00A300;color:#fff;">MUTASI</th>
				<th colspan="3" rowspan="2" style="background-color:#00A300;color:#fff;">AKHIR</th>

				<!-- Juni-->
				<th colspan="7" style="background-color:#00A300;color:#fff;">MUTASI</th>
				<th colspan="3" rowspan="2" style="background-color:#00A300;color:#fff;">AKHIR</th>

				<!-- Juli-->
				<th colspan="7" style="background-color:#00A300;color:#fff;">MUTASI</th>
				<th colspan="3" rowspan="2" style="background-color:#00A300;color:#fff;">AKHIR</th>

				<!-- Agustus-->
				<th colspan="7" style="background-color:#00A300;color:#fff;">MUTASI</th>
				<th colspan="3" rowspan="2" style="background-color:#00A300;color:#fff;">AKHIR</th>

				<!-- September-->
				<th colspan="7" style="background-color:#00A300;color:#fff;">MUTASI</th>
				<th colspan="3" rowspan="2" style="background-color:#00A300;color:#fff;">AKHIR</th>

				<!-- Oktober-->
				<th colspan="7" style="background-color:#00A300;color:#fff;">MUTASI</th>
				<th colspan="3" rowspan="2" style="background-color:#00A300;color:#fff;">AKHIR</th>

				<!-- November-->
				<th colspan="7" style="background-color:#00A300;color:#fff;">MUTASI</th>
				<th colspan="3" rowspan="2" style="background-color:#00A300;color:#fff;">AKHIR</th>

				<!-- Desember-->
				<th colspan="7" style="background-color:#00A300;color:#fff;">MUTASI</th>
				<th colspan="3" rowspan="2" style="background-color:#00A300;color:#fff;">AKHIR</th>

				<!-- Final -->
				<th colspan="6" style="background-color:#00A300;color:#fff;">MUTASI</th>
			</tr>
			
			<!-- Pengaturan Baris 3-->
			<tr>
				<!-- Januari-->
				<th colspan="2" style="background-color:#00A300;color:#fff;">HA</th>
				<th colspan="5" style="background-color:#00A300;color:#fff;">POKOK</th>

				<!-- Februari-->
				<th colspan="2" style="background-color:#00A300;color:#fff;">HA</th>
				<th colspan="5" style="background-color:#00A300;color:#fff;">POKOK</th>

				<!-- Maret-->
				<th colspan="2" style="background-color:#00A300;color:#fff;">HA</th>
				<th colspan="5" style="background-color:#00A300;color:#fff;">POKOK</th>

				<!-- April-->
				<th colspan="2" style="background-color:#00A300;color:#fff;">HA</th>
				<th colspan="5" style="background-color:#00A300;color:#fff;">POKOK</th>

				<!-- Mei-->
				<th colspan="2" style="background-color:#00A300;color:#fff;">HA</th>
				<th colspan="5" style="background-color:#00A300;color:#fff;">POKOK</th>

				<!-- Juni-->
				<th colspan="2" style="background-color:#00A300;color:#fff;">HA</th>
				<th colspan="5" style="background-color:#00A300;color:#fff;">POKOK</th>

				<!-- Juli-->
				<th colspan="2" style="background-color:#00A300;color:#fff;">HA</th>
				<th colspan="5" style="background-color:#00A300;color:#fff;">POKOK</th>

				<!-- Agustus-->
				<th colspan="2" style="background-color:#00A300;color:#fff;">HA</th>
				<th colspan="5" style="background-color:#00A300;color:#fff;">POKOK</th>

				<!-- September-->
				<th colspan="2" style="background-color:#00A300;color:#fff;">HA</th>
				<th colspan="5" style="background-color:#00A300;color:#fff;">POKOK</th>

				<!-- Oktober-->
				<th colspan="2" style="background-color:#00A300;color:#fff;">HA</th>
				<th colspan="5" style="background-color:#00A300;color:#fff;">POKOK</th>

				<!-- November-->
				<th colspan="2" style="background-color:#00A300;color:#fff;">HA</th>
				<th colspan="5" style="background-color:#00A300;color:#fff;">POKOK</th>

				<!-- Desember-->
				<th colspan="2" style="background-color:#00A300;color:#fff;">HA</th>
				<th colspan="5" style="background-color:#00A300;color:#fff;">POKOK</th>

				<!-- Final -->
				<th colspan="2" style="background-color:#00A300;color:#fff;">HA</th>
				<th colspan="4" style="background-color:#00A300;color:#fff;">POKOK</th>
			</tr>
	
			<!-- Pengaturan Baris 4-->
			<tr>	
				<!-- Januari-->
				<th rowspan="2" style="background-color:#00A300;color:#fff;">Ha</th>
				<th rowspan="2" style="background-color:#00A300;color:#fff;">POKOK</th>
				<th rowspan="2" style="background-color:#00A300;color:#fff;">SPH</th>
				<th rowspan="2" style="background-color:#00A300;color:#fff;">TAMBAH</th>
				<th rowspan="2" style="background-color:#00A300;color:#fff;">KURANG</th>
				<th colspan="3" style="background-color:#00A300;color:#fff;">TAMBAH</th>
				<th rowspan="2" colspan="2" style="background-color:#00A300;color:#fff;">KURANG</th>
				<th rowspan="2" style="background-color:#00A300;color:#fff;">Ha</th>
				<th rowspan="2" style="background-color:#00A300;color:#fff;">POKOK</th>
				<th rowspan="2" style="background-color:#00A300;color:#fff;">SPH</th>

				<!-- Februari-->
				<th rowspan="2" style="background-color:#00A300;color:#fff;">TAMBAH</th>
				<th rowspan="2" style="background-color:#00A300;color:#fff;">KURANG</th>
				<th colspan="3" style="background-color:#00A300;color:#fff;">TAMBAH</th>
				<th rowspan="2" colspan="2" style="background-color:#00A300;color:#fff;">KURANG</th>
				<th rowspan="2" style="background-color:#00A300;color:#fff;">Ha</th>
				<th rowspan="2" style="background-color:#00A300;color:#fff;">POKOK</th>
				<th rowspan="2" style="background-color:#00A300;color:#fff;">SPH</th>

				<!-- Maret-->
				<th rowspan="2" style="background-color:#00A300;color:#fff;">TAMBAH</th>
				<th rowspan="2" style="background-color:#00A300;color:#fff;">KURANG</th>
				<th colspan="3" style="background-color:#00A300;color:#fff;">TAMBAH</th>
				<th rowspan="2" colspan="2" style="background-color:#00A300;color:#fff;">KURANG</th>
				<th rowspan="2" style="background-color:#00A300;color:#fff;">Ha</th>
				<th rowspan="2" style="background-color:#00A300;color:#fff;">POKOK</th>
				<th rowspan="2" style="background-color:#00A300;color:#fff;">SPH</th>

				<!-- April-->
				<th rowspan="2" style="background-color:#00A300;color:#fff;">TAMBAH</th>
				<th rowspan="2" style="background-color:#00A300;color:#fff;">KURANG</th>
				<th colspan="3" style="background-color:#00A300;color:#fff;">TAMBAH</th>
				<th rowspan="2" colspan="2" style="background-color:#00A300;color:#fff;">KURANG</th>
				<th rowspan="2" style="background-color:#00A300;color:#fff;">Ha</th>
				<th rowspan="2" style="background-color:#00A300;color:#fff;">POKOK</th>
				<th rowspan="2" style="background-color:#00A300;color:#fff;">SPH</th>

				<!-- Mei-->
				<th rowspan="2" style="background-color:#00A300;color:#fff;">TAMBAH</th>
				<th rowspan="2" style="background-color:#00A300;color:#fff;">KURANG</th>
				<th colspan="3" style="background-color:#00A300;color:#fff;">TAMBAH</th>
				<th rowspan="2" colspan="2" style="background-color:#00A300;color:#fff;">KURANG</th>
				<th rowspan="2" style="background-color:#00A300;color:#fff;">Ha</th>
				<th rowspan="2" style="background-color:#00A300;color:#fff;">POKOK</th>
				<th rowspan="2" style="background-color:#00A300;color:#fff;">SPH</th>

				<!-- Juni-->
				<th rowspan="2" style="background-color:#00A300;color:#fff;">TAMBAH</th>
				<th rowspan="2" style="background-color:#00A300;color:#fff;">KURANG</th>
				<th colspan="3" style="background-color:#00A300;color:#fff;">TAMBAH</th>
				<th rowspan="2" colspan="2" style="background-color:#00A300;color:#fff;">KURANG</th>
				<th rowspan="2" style="background-color:#00A300;color:#fff;">Ha</th>
				<th rowspan="2" style="background-color:#00A300;color:#fff;">POKOK</th>
				<th rowspan="2" style="background-color:#00A300;color:#fff;">SPH</th>

				<!-- Juli-->
				<th rowspan="2" style="background-color:#00A300;color:#fff;">TAMBAH</th>
				<th rowspan="2" style="background-color:#00A300;color:#fff;">KURANG</th>
				<th colspan="3" style="background-color:#00A300;color:#fff;">TAMBAH</th>
				<th rowspan="2" colspan="2" style="background-color:#00A300;color:#fff;">KURANG</th>
				<th rowspan="2" style="background-color:#00A300;color:#fff;">Ha</th>
				<th rowspan="2" style="background-color:#00A300;color:#fff;">POKOK</th>
				<th rowspan="2" style="background-color:#00A300;color:#fff;">SPH</th>

				<!-- Agustus-->
				<th rowspan="2" style="background-color:#00A300;color:#fff;">TAMBAH</th>
				<th rowspan="2" style="background-color:#00A300;color:#fff;">KURANG</th>
				<th colspan="3" style="background-color:#00A300;color:#fff;">TAMBAH</th>
				<th rowspan="2" colspan="2" style="background-color:#00A300;color:#fff;">KURANG</th>
				<th rowspan="2" style="background-color:#00A300;color:#fff;">Ha</th>
				<th rowspan="2" style="background-color:#00A300;color:#fff;">POKOK</th>
				<th rowspan="2" style="background-color:#00A300;color:#fff;">SPH</th>

				<!-- September-->
				<th rowspan="2" style="background-color:#00A300;color:#fff;">TAMBAH</th>
				<th rowspan="2" style="background-color:#00A300;color:#fff;">KURANG</th>
				<th colspan="3" style="background-color:#00A300;color:#fff;">TAMBAH</th>
				<th rowspan="2" colspan="2" style="background-color:#00A300;color:#fff;">KURANG</th>
				<th rowspan="2" style="background-color:#00A300;color:#fff;">Ha</th>
				<th rowspan="2" style="background-color:#00A300;color:#fff;">POKOK</th>
				<th rowspan="2" style="background-color:#00A300;color:#fff;">SPH</th>

				<!-- Oktober-->
				<th rowspan="2" style="background-color:#00A300;color:#fff;">TAMBAH</th>
				<th rowspan="2" style="background-color:#00A300;color:#fff;">KURANG</th>
				<th colspan="3" style="background-color:#00A300;color:#fff;">TAMBAH</th>
				<th rowspan="2" colspan="2" style="background-color:#00A300;color:#fff;">KURANG</th>
				<th rowspan="2" style="background-color:#00A300;color:#fff;">Ha</th>
				<th rowspan="2" style="background-color:#00A300;color:#fff;">POKOK</th>
				<th rowspan="2" style="background-color:#00A300;color:#fff;">SPH</th>

				<!-- November-->
				<th rowspan="2" style="background-color:#00A300;color:#fff;">TAMBAH</th>
				<th rowspan="2" style="background-color:#00A300;color:#fff;">KURANG</th>
				<th colspan="3" style="background-color:#00A300;color:#fff;">TAMBAH</th>
				<th rowspan="2" colspan="2" style="background-color:#00A300;color:#fff;">KURANG</th>
				<th rowspan="2" style="background-color:#00A300;color:#fff;">Ha</th>
				<th rowspan="2" style="background-color:#00A300;color:#fff;">POKOK</th>
				<th rowspan="2" style="background-color:#00A300;color:#fff;">SPH</th>

				<!-- Desember-->
				<th rowspan="2" style="background-color:#00A300;color:#fff;">TAMBAH</th>
				<th rowspan="2" style="background-color:#00A300;color:#fff;">KURANG</th>
				<th colspan="3" style="background-color:#00A300;color:#fff;">TAMBAH</th>
				<th rowspan="2" colspan="2" style="background-color:#00A300;color:#fff;">KURANG</th>
				<th rowspan="2" style="background-color:#00A300;color:#fff;">Ha</th>
				<th rowspan="2" style="background-color:#00A300;color:#fff;">POKOK</th>
				<th rowspan="2" style="background-color:#00A300;color:#fff;">SPH</th>

				<!-- Final-->
				<th rowspan="2" style="background-color:#00A300;color:#fff;">TAMBAH</th>
				<th rowspan="2" style="background-color:#00A300;color:#fff;">KURANG</th>
				<th colspan="3" style="background-color:#00A300;color:#fff;">TAMBAH</th>
				<th rowspan="2" style="background-color:#00A300;color:#fff;">KURANG</th>
			</tr>
		
			<!-- Pengaturan Baris 5-->
			<tr>	
				<!-- Januari-->
				<th style="background-color:#00A300;color:#fff;">SISIP</th>
				<th style="background-color:#00A300;color:#fff;">COMPLETING</th>
				<th style="background-color:#00A300;color:#fff;">TB</th>

				<!-- Februari-->
				<th style="background-color:#00A300;color:#fff;">SISIP</th>
				<th style="background-color:#00A300;color:#fff;">COMPLETING</th>
				<th style="background-color:#00A300;color:#fff;">TB</th>

				<!-- Maret-->
				<th style="background-color:#00A300;color:#fff;">SISIP</th>
				<th style="background-color:#00A300;color:#fff;">COMPLETING</th>
				<th style="background-color:#00A300;color:#fff;">TB</th>

				<!-- April-->
				<th style="background-color:#00A300;color:#fff;">SISIP</th>
				<th style="background-color:#00A300;color:#fff;">COMPLETING</th>
				<th style="background-color:#00A300;color:#fff;">TB</th>

				<!-- Mei-->
				<th style="background-color:#00A300;color:#fff;">SISIP</th>
				<th style="background-color:#00A300;color:#fff;">COMPLETING</th>
				<th style="background-color:#00A300;color:#fff;">TB</th>

				<!-- Juni-->
				<th style="background-color:#00A300;color:#fff;">SISIP</th>
				<th style="background-color:#00A300;color:#fff;">COMPLETING</th>
				<th style="background-color:#00A300;color:#fff;">TB</th>

				<!-- Juli-->
				<th style="background-color:#00A300;color:#fff;">SISIP</th>
				<th style="background-color:#00A300;color:#fff;">COMPLETING</th>
				<th style="background-color:#00A300;color:#fff;">TB</th>

				<!-- Agustus-->
				<th style="background-color:#00A300;color:#fff;">SISIP</th>
				<th style="background-color:#00A300;color:#fff;">COMPLETING</th>
				<th style="background-color:#00A300;color:#fff;">TB</th>

				<!-- September-->
				<th style="background-color:#00A300;color:#fff;">SISIP</th>
				<th style="background-color:#00A300;color:#fff;">COMPLETING</th>
				<th style="background-color:#00A300;color:#fff;">TB</th>

				<!-- Oktober-->
				<th style="background-color:#00A300;color:#fff;">SISIP</th>
				<th style="background-color:#00A300;color:#fff;">COMPLETING</th>
				<th style="background-color:#00A300;color:#fff;">TB</th>

				<!-- November-->
				<th style="background-color:#00A300;color:#fff;">SISIP</th>
				<th style="background-color:#00A300;color:#fff;">COMPLETING</th>
				<th style="background-color:#00A300;color:#fff;">TB</th>

				<!-- Desember-->
				<th style="background-color:#00A300;color:#fff;">SISIP</th>
				<th style="background-color:#00A300;color:#fff;">COMPLETING</th>
				<th style="background-color:#00A300;color:#fff;">TB</th>

				<!-- Final-->
				<th style="background-color:#00A300;color:#fff;">SISIP</th>
				<th style="background-color:#00A300;color:#fff;">COMPLETING</th>
				<th style="background-color:#00A300;color:#fff;">TB</th>


				
			</tr>
	    </thead>
 <tbody>
<%
Set rsEPS = Server.CreateObject("ADODB.Recordset")
rsEPS.Open strEPS , connstr, adOpenStatic 

Dim xNum

if rsEPS.bof and rsEPS.eof then

%>
	<tr>
		<td colspan="141">NO DATA AVAILABLE</td>
	</tr>
<%else
	 dim newcolum,tempownershipnewcolum,tempDivisionnewcolum, Total
		while not rsEPS.eof
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
			 ' Dim TotalSPH_Awal
			 ' Dim TotalHA_Akhir_Jan, TotalHA_Akhir_Feb, TotalHA_Akhir_Mar, TotalHA_Akhir_Apr, TotalHA_Akhir_Mei, TotalHA_Akhir_Jun
			 ' Dim TotalHA_Akhir_Jul, TotalHA_Akhir_Aug, TotalHA_Akhir_Sep, TotalHA_Akhir_Okt, TotalHA_Akhir_Nov, TotalHA_Akhir_Des
			 ' Dim TotalPKK_Akhir_Jan, TotalPKK_Akhir_Feb, TotalPKK_Akhir_Mar, TotalPKK_Akhir_Apr, TotalPKK_Akhir_Mei, TotalPKK_Akhir_Jun
			 ' Dim TotalPKK_Akhir_Jul, TotalPKK_Akhir_Aug, TotalPKK_Akhir_Sep, TotalPKK_Akhir_Okt, TotalPKK_Akhir_Nov, TotalPKK_Akhir_Des
			 ' Dim TotalSPH_Akhir_Jan, TotalSPH_Akhir_Feb, TotalSPH_Akhir_Mar, TotalSPH_Akhir_Apr, TotalSPH_Akhir_Mei, TotalSPH_Akhir_Jun
			 ' Dim TotalSPH_Akhir_Jul, TotalSPH_Akhir_Aug, TotalSPH_Akhir_Sep, TotalSPH_Akhir_Okt, TotalSPH_Akhir_Nov, TotalSPH_Akhir_Des
			 ' Dim TOTAL_HA, TOTAL_PKK_TANAM, TOTAL_SPH_TANAM, TOTAL_PKK_SISIP, TOTAL_SPH_SISIP
			 ' Dim TOTAL_HA_AWAL_JAN, TOTAL_PKK_AWAL_JAN, TOTAL_SPH_AWAL_JAN
			 ' Dim 
			 Dim FinalHA_Tambah, FinalSisip, FinalCompleting, FinalTB
		     newcolum = 0
			 Total = 0
	 		if Cdbl(to_nol(rsEPS("LINEHA"))) = 0 then TotalSPH_Awal = 0 else TotalSPH_Awal = Cdbl(to_nol(rsEPS("LINEPKK"))) / Cdbl(to_nol(rsEPS("LINEHA")))
			TotalHA_Akhir_Jan = Cdbl(to_nol(rsEPS("LINEHA"))) + Cdbl(to_nol(rsEPS("HECTARE_JAN")))
			TotalHA_Akhir_Feb = TotalHA_Akhir_Jan + Cdbl(to_nol(rsEPS("HECTARE_FEB")))
			TotalHA_Akhir_Mar = TotalHA_Akhir_Feb + Cdbl(to_nol(rsEPS("HECTARE_MAR")))
			TotalHA_Akhir_Apr = TotalHA_Akhir_Mar + Cdbl(to_nol(rsEPS("HECTARE_APR")))
			TotalHA_Akhir_Mei = TotalHA_Akhir_Apr + Cdbl(to_nol(rsEPS("HECTARE_MEI")))
			TotalHA_Akhir_Jun = TotalHA_Akhir_Mei + Cdbl(to_nol(rsEPS("HECTARE_JUN")))

			TotalHA_Akhir_Jul = TotalHA_Akhir_Jun + Cdbl(to_nol(rsEPS("HECTARE_JUL")))
			TotalHA_Akhir_Aug = TotalHA_Akhir_Jul + Cdbl(to_nol(rsEPS("HECTARE_AUG")))
			TotalHA_Akhir_Sep = TotalHA_Akhir_Aug + Cdbl(to_nol(rsEPS("HECTARE_SEP")))
			TotalHA_Akhir_Okt = TotalHA_Akhir_Sep + Cdbl(to_nol(rsEPS("HECTARE_OKT")))
			TotalHA_Akhir_Nov = TotalHA_Akhir_Okt + Cdbl(to_nol(rsEPS("HECTARE_NOV")))
			TotalHA_Akhir_Des = TotalHA_Akhir_Nov + Cdbl(to_nol(rsEPS("HECTARE_DES")))


			FinalHA_Tambah = TotalHA_Akhir_Jan + TotalHA_Akhir_Feb + TotalHA_Akhir_Mar + TotalHA_Akhir_Apr + TotalHA_Akhir_Mei + TotalHA_Akhir_Jun + TotalHA_Akhir_Jul + TotalHA_Akhir_Aug + TotalHA_Akhir_Sep + TotalHA_Akhir_Okt + TotalHA_Akhir_Nov + TotalHA_Akhir_Des

			FinalSisip = Cdbl(to_nol(rsEPS("SUPPLYING_JAN"))) + Cdbl(to_nol(rsEPS("SUPPLYING_FEB"))) + Cdbl(to_nol(rsEPS("SUPPLYING_MAR"))) + Cdbl(to_nol(rsEPS("SUPPLYING_APR"))) + cdbl(to_nol(rsEPS("SUPPLYING_MEI"))) + Cdbl(to_nol(rsEPS("SUPPLYING_JUN"))) + Cdbl(to_nol(rsEPS("SUPPLYING_JUL"))) + Cdbl(to_nol(rsEPS("SUPPLYING_AUG"))) + Cdbl(to_nol(rsEPS("SUPPLYING_SEP"))) + Cdbl(to_nol(rsEPS("SUPPLYING_OKT"))) + Cdbl(to_nol(rsEPS("SUPPLYING_NOV"))) + Cdbl(to_nol(rsEPS("SUPPLYING_DES"))) 

			FinalCompleting = Cdbl(to_nol(rsEPS("COMPLETING_JAN"))) + Cdbl(to_nol(rsEPS("COMPLETING_FEB"))) + Cdbl(to_nol(rsEPS("COMPLETING_MAR"))) + Cdbl(to_nol(rsEPS("COMPLETING_APR"))) + Cdbl(to_nol(rsEPS("COMPLETING_MEI"))) + Cdbl(to_nol(rsEPS("COMPLETING_JUN"))) + Cdbl(to_nol(rsEPS("COMPLETING_JUL"))) + Cdbl(to_nol(rsEPS("COMPLETING_AUG"))) + Cdbl(to_nol(rsEPS("COMPLETING_SEP"))) + Cdbl(to_nol(rsEPS("COMPLETING_OKT"))) + Cdbl(to_nol(rsEPS("COMPLETING_NOV"))) + Cdbl(to_nol(rsEPS("COMPLETING_DES")))

			FinalTB = Cdbl(to_nol(rsEPS("PLANTING_JAN"))) + Cdbl(to_nol(rsEPS("PLANTING_FEB"))) + Cdbl(to_nol(rsEPS("PLANTING_MAR"))) + Cdbl(to_nol(rsEPS("PLANTING_APR"))) + Cdbl(to_nol(rsEPS("PLANTING_MEI"))) + Cdbl(to_nol(rsEPS("PLANTING_JUN"))) + Cdbl(to_nol(rsEPS("PLANTING_JUL"))) + Cdbl(to_nol(rsEPS("PLANTING_AUG"))) + Cdbl(to_nol(rsEPS("PLANTING_SEP"))) + Cdbl(to_nol(rsEPS("PLANTING_OKT"))) + Cdbl(to_nol(rsEPS("PLANTING_NOV"))) + Cdbl(to_nol(rsEPS("PLANTING_DES")))

			TotalPKK_Akhir_Jan = Cdbl(to_nol(rsEPS("LINEPKK"))) - Cdbl(to_nol(rsEPS("DEADTREE_JAN"))) + (Cdbl(to_nol(rsEPS("SUPPLYING_JAN"))) + Cdbl(to_nol(rsEPS("COMPLETING_JAN"))) + Cdbl(to_nol(rsEPS("PLANTING_JAN"))))
			TotalPKK_Akhir_Feb = TotalPKK_Akhir_Jan - Cdbl(to_nol(rsEPS("DEADTREE_FEB"))) + (Cdbl(to_nol(rsEPS("SUPPLYING_FEB"))) + Cdbl(to_nol(rsEPS("COMPLETING_FEB"))) + Cdbl(to_nol(rsEPS("PLANTING_FEB"))))
			TotalPKK_Akhir_Mar = TotalPKK_Akhir_Feb - Cdbl(to_nol(rsEPS("DEADTREE_MAR"))) + (Cdbl(to_nol(rsEPS("SUPPLYING_MAR"))) + Cdbl(to_nol(rsEPS("COMPLETING_MAR"))) + Cdbl(to_nol(rsEPS("PLANTING_MAR"))))
			TotalPKK_Akhir_Apr = TotalPKK_Akhir_Mar - Cdbl(to_nol(rsEPS("DEADTREE_APR"))) + (Cdbl(to_nol(rsEPS("SUPPLYING_APR"))) + Cdbl(to_nol(rsEPS("COMPLETING_APR"))) + Cdbl(to_nol(rsEPS("PLANTING_APR"))))
			TotalPKK_Akhir_Mei = TotalPKK_Akhir_Apr - Cdbl(to_nol(rsEPS("DEADTREE_MEI"))) + (Cdbl(to_nol(rsEPS("SUPPLYING_MEI"))) + Cdbl(to_nol(rsEPS("COMPLETING_MEI"))) + Cdbl(to_nol(rsEPS("PLANTING_MEI"))))
			TotalPKK_Akhir_Jun = TotalPKK_Akhir_Mei - Cdbl(to_nol(rsEPS("DEADTREE_JUN"))) + (Cdbl(to_nol(rsEPS("SUPPLYING_JUN"))) + Cdbl(to_nol(rsEPS("COMPLETING_JUN"))) + Cdbl(to_nol(rsEPS("PLANTING_JUN"))))

			TotalPKK_Akhir_Jul = TotalPKK_Akhir_Jun - Cdbl(to_nol(rsEPS("DEADTREE_JUL"))) + (Cdbl(to_nol(rsEPS("SUPPLYING_JUL"))) + Cdbl(to_nol(rsEPS("COMPLETING_JUL"))) + Cdbl(to_nol(rsEPS("PLANTING_JUL"))))
			TotalPKK_Akhir_Aug = TotalPKK_Akhir_JuL - Cdbl(to_nol(rsEPS("DEADTREE_AUG"))) + (Cdbl(to_nol(rsEPS("SUPPLYING_AUG"))) + Cdbl(to_nol(rsEPS("COMPLETING_AUG"))) + Cdbl(to_nol(rsEPS("PLANTING_AUG"))))
			TotalPKK_Akhir_Sep = TotalPKK_Akhir_Aug - Cdbl(to_nol(rsEPS("DEADTREE_SEP"))) + (Cdbl(to_nol(rsEPS("SUPPLYING_SEP"))) + Cdbl(to_nol(rsEPS("COMPLETING_SEP"))) + Cdbl(to_nol(rsEPS("PLANTING_SEP"))))
			TotalPKK_Akhir_Okt = TotalPKK_Akhir_Sep - Cdbl(to_nol(rsEPS("DEADTREE_OKT"))) + (Cdbl(to_nol(rsEPS("SUPPLYING_OKT"))) + Cdbl(to_nol(rsEPS("COMPLETING_OKT"))) + Cdbl(to_nol(rsEPS("PLANTING_OKT"))))
			TotalPKK_Akhir_Nov = TotalPKK_Akhir_Okt - Cdbl(to_nol(rsEPS("DEADTREE_NOV"))) + (Cdbl(to_nol(rsEPS("SUPPLYING_NOV"))) + Cdbl(to_nol(rsEPS("COMPLETING_NOV"))) + Cdbl(to_nol(rsEPS("PLANTING_NOV"))))
			TotalPKK_Akhir_Des = TotalPKK_Akhir_Nov - Cdbl(to_nol(rsEPS("DEADTREE_DES"))) + (Cdbl(to_nol(rsEPS("SUPPLYING_DES"))) + Cdbl(to_nol(rsEPS("COMPLETING_DES"))) + Cdbl(to_nol(rsEPS("PLANTING_DES"))))

			if TotalHA_Akhir_Jan = 0 then TotalSPH_Akhir_Jan = 0 else TotalSPH_Akhir_Jan = TotalPKK_Akhir_Jan / TotalHA_Akhir_Jan
			if TotalHA_Akhir_Feb = 0 then TotalSPH_Akhir_Feb = 0 else TotalSPH_Akhir_Feb = TotalPKK_Akhir_Feb / TotalHA_Akhir_Feb
			if TotalHA_Akhir_Mar = 0 then TotalSPH_Akhir_Mar = 0 else TotalSPH_Akhir_Mar = TotalPKK_Akhir_Mei / TotalHA_Akhir_Mar
			if TotalHA_Akhir_Apr = 0 then TotalSPH_Akhir_Apr = 0 else TotalSPH_Akhir_Apr = TotalPKK_Akhir_Apr / TotalHA_Akhir_Apr
			if TotalHA_Akhir_Mei = 0 then TotalSPH_Akhir_Mei = 0 else TotalSPH_Akhir_Mei = TotalPKK_Akhir_Mei / TotalHA_Akhir_Mei
			if TotalHA_Akhir_Jun = 0 then TotalSPH_Akhir_Jun = 0 else TotalSPH_Akhir_Jun = TotalPKK_Akhir_Jun / TotalHA_Akhir_Jul

			if TotalHA_Akhir_Jul = 0 then TotalSPH_Akhir_Jul = 0 else TotalSPH_Akhir_Jul = TotalPKK_Akhir_Jul / TotalHA_Akhir_Jul
			if TotalHA_Akhir_Aug = 0 then TotalSPH_Akhir_Aug = 0 else TotalSPH_Akhir_Aug = TotalPKK_Akhir_Aug / TotalHA_Akhir_Aug
			if TotalHA_Akhir_Sep = 0 then TotalSPH_Akhir_Sep = 0 else TotalSPH_Akhir_Sep = TotalPKK_Akhir_Sep / TotalHA_Akhir_Sep
			if TotalHA_Akhir_Okt = 0 then TotalSPH_Akhir_Okt = 0 else TotalSPH_Akhir_Okt = TotalPKK_Akhir_Okt / TotalHA_Akhir_Okt
			if TotalHA_Akhir_Nov = 0 then TotalSPH_Akhir_Nov = 0 else TotalSPH_Akhir_Nov = TotalPKK_Akhir_Nov / TotalHA_Akhir_Nov
			if TotalHA_Akhir_Des = 0 then TotalSPH_Akhir_Des = 0 else TotalSPH_Akhir_Des = TotalPKK_Akhir_Des / TotalHA_Akhir_Des

			if TotalPKK_Akhir_Des = 0 then PKK_SISIP = 0 else PKK_SISIP = Cdbl(to_nol(rsEPS("plantingpoint"))) - TotalPKK_Akhir_Des  
			if Cdbl(to_nol(rsEPS("hectarageplanted"))) = 0 then SPH_SISIP = 0 else SPH_SISIP =  PKK_SISIP / Cdbl(to_nol(rsEPS("hectarageplanted")))
			if Cdbl(to_nol(rsEPS("hectarageplanted"))) = 0 then SPH_TANAM = 0 else SPH_TANAM =  Cdbl(to_nol(rsEPS("plantingpoint"))) / Cdbl(to_nol(rsEPS("hectarageplanted")))
			
			TOTAL_HA = TOTAL_HA + Cdbl(to_nol(rsEPS("hectarageplanted")))
			TOTAL_PKK_TANAM = TOTAL_PKK_TANAM + Cdbl(to_nol(rsEPS("plantingpoint")))
			if TOTAL_HA = 0 then TOTAL_SPH_TANAM = 0 else TOTAL_SPH_TANAM =  TOTAL_PKK_TANAM / TOTAL_HA
			TOTAL_PKK_SISIP = TOTAL_PKK_SISIP + PKK_SISIP
			if TOTAL_HA = 0 then TOTAL_SPH_SISIP = 0 else TOTAL_SPH_SISIP =  TOTAL_PKK_SISIP / TOTAL_HA
			TOTAL_HA_AWAL_JAN = TOTAL_HA_AWAL_JAN + Cdbl(to_nol(rsEPS("lineha")))
			TOTAL_PKK_AWAL_JAN = TOTAL_PKK_AWAL_JAN + Cdbl(to_nol(rsEPS("linepkk")))
			if TOTAL_HA_AWAL_JAN = 0 then TOTAL_SPH_AWAL_JAN = 0 else TOTAL_SPH_AWAL_JAN =  TOTAL_PKK_AWAL_JAN / TOTAL_HA_AWAL_JAN

			TOTAL_HA_TAMBAH_JAN = TOTAL_HA_TAMBAH_JAN + Cdbl(to_nol(rsEPS("HECTARE_JAN")))
			TOTAL_SISIP_POKOK_JAN = TOTAL_SISIP_POKOK_JAN + Cdbl(to_nol(rsEPS("COMPLETING_JAN")))
			TOTAL_COMP_POKOK_JAN = TOTAL_COMP_POKOK_JAN + Cdbl(to_nol(rsEPS("SUPPLYING_JAN")))
			TOTAL_TB_POKOK_JAN = TOTAL_TB_POKOK_JAN + Cdbl(to_nol(rsEPS("PLANTING_JAN")))
			TOTAL_KURANG_POKOK_JAN = TOTAL_KURANG_POKOK_JAN + Cdbl(to_nol(rsEPS("DEADTREE_JAN")))
			TOTAL_HA_AKHIR_JAN = TOTAL_HA_AKHIR_JAN + TotalHA_Akhir_Jan
			TOTAL_PKK_AKHIR_JAN = TOTAL_PKK_AKHIR_JAN + TotalPKK_Akhir_Jan
			if TOTAL_HA_AKHIR_JAN = 0 then TOTAL_SPH_AKHIR_JAN = 0 else TOTAL_SPH_AKHIR_JAN =  TOTAL_PKK_AKHIR_JAN / TOTAL_HA_AKHIR_JAN

			TOTAL_HA_TAMBAH_FEB = TOTAL_HA_TAMBAH_FEB + Cdbl(to_nol(rsEPS("HECTARE_FEB")))
			TOTAL_SISIP_POKOK_FEB = TOTAL_SISIP_POKOK_FEB + Cdbl(to_nol(rsEPS("COMPLETING_FEB")))
			TOTAL_COMP_POKOK_FEB = TOTAL_COMP_POKOK_FEB + Cdbl(to_nol(rsEPS("SUPPLYING_FEB")))
			TOTAL_TB_POKOK_FEB = TOTAL_TB_POKOK_FEB + Cdbl(to_nol(rsEPS("PLANTING_FEB")))
			TOTAL_KURANG_POKOK_FEB = TOTAL_KURANG_POKOK_FEB + Cdbl(to_nol(rsEPS("DEADTREE_FEB")))
			TOTAL_HA_AKHIR_FEB = TOTAL_HA_AKHIR_FEB + TotalHA_Akhir_FEB
			TOTAL_PKK_AKHIR_FEB = TOTAL_PKK_AKHIR_FEB + TotalPKK_Akhir_FEB
			if TOTAL_HA_AKHIR_FEB = 0 then TOTAL_SPH_AKHIR_FEB = 0 else TOTAL_SPH_AKHIR_FEB =  TOTAL_PKK_AKHIR_FEB / TOTAL_HA_AKHIR_FEB

			TOTAL_HA_TAMBAH_MAR = TOTAL_HA_TAMBAH_MAR + Cdbl(to_nol(rsEPS("HECTARE_MAR")))
			TOTAL_SISIP_POKOK_MAR = TOTAL_SISIP_POKOK_MAR + Cdbl(to_nol(rsEPS("COMPLETING_MAR")))
			TOTAL_COMP_POKOK_MAR = TOTAL_COMP_POKOK_MAR + Cdbl(to_nol(rsEPS("SUPPLYING_MAR")))
			TOTAL_TB_POKOK_MAR = TOTAL_TB_POKOK_MAR + Cdbl(to_nol(rsEPS("PLANTING_MAR")))
			TOTAL_KURANG_POKOK_MAR = TOTAL_KURANG_POKOK_MAR + Cdbl(to_nol(rsEPS("DEADTREE_MAR")))
			TOTAL_HA_AKHIR_MAR = TOTAL_HA_AKHIR_MAR + TotalHA_Akhir_MAR
			TOTAL_PKK_AKHIR_MAR = TOTAL_PKK_AKHIR_MAR + TotalPKK_Akhir_MAR
			if TOTAL_HA_AKHIR_MAR = 0 then TOTAL_SPH_AKHIR_MAR = 0 else TOTAL_SPH_AKHIR_MAR =  TOTAL_PKK_AKHIR_MAR / TOTAL_HA_AKHIR_MAR

			TOTAL_HA_TAMBAH_APR = TOTAL_HA_TAMBAH_APR + Cdbl(to_nol(rsEPS("HECTARE_APR")))
			TOTAL_SISIP_POKOK_APR = TOTAL_SISIP_POKOK_APR + Cdbl(to_nol(rsEPS("COMPLETING_APR")))
			TOTAL_COMP_POKOK_APR = TOTAL_COMP_POKOK_APR + Cdbl(to_nol(rsEPS("SUPPLYING_APR")))
			TOTAL_TB_POKOK_APR = TOTAL_TB_POKOK_APR + Cdbl(to_nol(rsEPS("PLANTING_APR")))
			TOTAL_KURANG_POKOK_APR = TOTAL_KURANG_POKOK_APR + Cdbl(to_nol(rsEPS("DEADTREE_APR")))
			TOTAL_HA_AKHIR_APR = TOTAL_HA_AKHIR_APR + TotalHA_Akhir_APR
			TOTAL_PKK_AKHIR_APR = TOTAL_PKK_AKHIR_APR + TotalPKK_Akhir_APR
			if TOTAL_HA_AKHIR_APR = 0 then TOTAL_SPH_AKHIR_APR = 0 else TOTAL_SPH_AKHIR_APR =  TOTAL_PKK_AKHIR_APR / TOTAL_HA_AKHIR_APR

			TOTAL_HA_TAMBAH_MEI = TOTAL_HA_TAMBAH_MEI + Cdbl(to_nol(rsEPS("HECTARE_MEI")))
			TOTAL_SISIP_POKOK_MEI = TOTAL_SISIP_POKOK_MEI + Cdbl(to_nol(rsEPS("COMPLETING_MEI")))
			TOTAL_COMP_POKOK_MEI = TOTAL_COMP_POKOK_MEI + Cdbl(to_nol(rsEPS("SUPPLYING_MEI")))
			TOTAL_TB_POKOK_MEI = TOTAL_TB_POKOK_MEI + Cdbl(to_nol(rsEPS("PLANTING_MEI")))
			TOTAL_KURANG_POKOK_MEI = TOTAL_KURANG_POKOK_MEI + Cdbl(to_nol(rsEPS("DEADTREE_MEI")))
			TOTAL_HA_AKHIR_MEI = TOTAL_HA_AKHIR_MEI + TotalHA_Akhir_MEI
			TOTAL_PKK_AKHIR_MEI = TOTAL_PKK_AKHIR_MEI + TotalPKK_Akhir_MEI
			if TOTAL_HA_AKHIR_MEI = 0 then TOTAL_SPH_AKHIR_MEI = 0 else TOTAL_SPH_AKHIR_MEI =  TOTAL_PKK_AKHIR_MEI / TOTAL_HA_AKHIR_MEI

			TOTAL_HA_TAMBAH_JUN = TOTAL_HA_TAMBAH_JUN + Cdbl(to_nol(rsEPS("HECTARE_JUN")))
			TOTAL_SISIP_POKOK_JUN = TOTAL_SISIP_POKOK_JUN + Cdbl(to_nol(rsEPS("COMPLETING_JUN")))
			TOTAL_COMP_POKOK_JUN = TOTAL_COMP_POKOK_JUN + Cdbl(to_nol(rsEPS("SUPPLYING_JUN")))
			TOTAL_TB_POKOK_JUN = TOTAL_TB_POKOK_JUN + Cdbl(to_nol(rsEPS("PLANTING_JUN")))
			TOTAL_KURANG_POKOK_JUN = TOTAL_KURANG_POKOK_JUN + Cdbl(to_nol(rsEPS("DEADTREE_JUN")))
			TOTAL_HA_AKHIR_JUN = TOTAL_HA_AKHIR_JUN + TotalHA_Akhir_JUN
			TOTAL_PKK_AKHIR_JUN = TOTAL_PKK_AKHIR_JUN + TotalPKK_Akhir_JUN
			if TOTAL_HA_AKHIR_JUN = 0 then TOTAL_SPH_AKHIR_JUN = 0 else TOTAL_SPH_AKHIR_JUN =  TOTAL_PKK_AKHIR_JUN / TOTAL_HA_AKHIR_JUN

			TOTAL_HA_TAMBAH_JUL = TOTAL_HA_TAMBAH_JUL + Cdbl(to_nol(rsEPS("HECTARE_JUL")))
			TOTAL_SISIP_POKOK_JUL = TOTAL_SISIP_POKOK_JUL + Cdbl(to_nol(rsEPS("COMPLETING_JUL")))
			TOTAL_COMP_POKOK_JUL = TOTAL_COMP_POKOK_JUL + Cdbl(to_nol(rsEPS("SUPPLYING_JUL")))
			TOTAL_TB_POKOK_JUL = TOTAL_TB_POKOK_JUL + Cdbl(to_nol(rsEPS("PLANTING_JUL")))
			TOTAL_KURANG_POKOK_JUL = TOTAL_KURANG_POKOK_JUL + Cdbl(to_nol(rsEPS("DEADTREE_JUL")))
			TOTAL_HA_AKHIR_JUL = TOTAL_HA_AKHIR_JUL + TotalHA_Akhir_JUL
			TOTAL_PKK_AKHIR_JUL = TOTAL_PKK_AKHIR_JUL + TotalPKK_Akhir_JUL
			if TOTAL_HA_AKHIR_JUL = 0 then TOTAL_SPH_AKHIR_JUL = 0 else TOTAL_SPH_AKHIR_JUL =  TOTAL_PKK_AKHIR_JUL / TOTAL_HA_AKHIR_JUL

			TOTAL_HA_TAMBAH_AUG = TOTAL_HA_TAMBAH_AUG + Cdbl(to_nol(rsEPS("HECTARE_AUG")))
			TOTAL_SISIP_POKOK_AUG = TOTAL_SISIP_POKOK_AUG + Cdbl(to_nol(rsEPS("COMPLETING_AUG")))
			TOTAL_COMP_POKOK_AUG = TOTAL_COMP_POKOK_AUG + Cdbl(to_nol(rsEPS("SUPPLYING_AUG")))
			TOTAL_TB_POKOK_AUG = TOTAL_TB_POKOK_AUG + Cdbl(to_nol(rsEPS("PLANTING_AUG")))
			TOTAL_KURANG_POKOK_AUG = TOTAL_KURANG_POKOK_AUG + Cdbl(to_nol(rsEPS("DEADTREE_AUG")))
			TOTAL_HA_AKHIR_AUG = TOTAL_HA_AKHIR_AUG + TotalHA_Akhir_AUG
			TOTAL_PKK_AKHIR_AUG = TOTAL_PKK_AKHIR_AUG + TotalPKK_Akhir_AUG
			if TOTAL_HA_AKHIR_AUG = 0 then TOTAL_SPH_AKHIR_AUG = 0 else TOTAL_SPH_AKHIR_AUG =  TOTAL_PKK_AKHIR_AUG / TOTAL_HA_AKHIR_AUG

			TOTAL_HA_TAMBAH_SEP = TOTAL_HA_TAMBAH_SEP + Cdbl(to_nol(rsEPS("HECTARE_SEP")))
			TOTAL_SISIP_POKOK_SEP = TOTAL_SISIP_POKOK_SEP + Cdbl(to_nol(rsEPS("COMPLETING_SEP")))
			TOTAL_COMP_POKOK_SEP = TOTAL_COMP_POKOK_SEP + Cdbl(to_nol(rsEPS("SUPPLYING_SEP")))
			TOTAL_TB_POKOK_SEP = TOTAL_TB_POKOK_SEP + Cdbl(to_nol(rsEPS("PLANTING_SEP")))
			TOTAL_KURANG_POKOK_SEP = TOTAL_KURANG_POKOK_SEP + Cdbl(to_nol(rsEPS("DEADTREE_SEP")))
			TOTAL_HA_AKHIR_SEP = TOTAL_HA_AKHIR_SEP + TotalHA_Akhir_SEP
			TOTAL_PKK_AKHIR_SEP = TOTAL_PKK_AKHIR_SEP + TotalPKK_Akhir_SEP
			if TOTAL_HA_AKHIR_SEP = 0 then TOTAL_SPH_AKHIR_SEP = 0 else TOTAL_SPH_AKHIR_SEP =  TOTAL_PKK_AKHIR_SEP / TOTAL_HA_AKHIR_SEP

			TOTAL_HA_TAMBAH_OKT = TOTAL_HA_TAMBAH_OKT + Cdbl(to_nol(rsEPS("HECTARE_OKT")))
			TOTAL_SISIP_POKOK_OKT = TOTAL_SISIP_POKOK_OKT + Cdbl(to_nol(rsEPS("COMPLETING_OKT")))
			TOTAL_COMP_POKOK_OKT = TOTAL_COMP_POKOK_OKT + Cdbl(to_nol(rsEPS("SUPPLYING_OKT")))
			TOTAL_TB_POKOK_OKT = TOTAL_TB_POKOK_OKT + Cdbl(to_nol(rsEPS("PLANTING_OKT")))
			TOTAL_KURANG_POKOK_OKT = TOTAL_KURANG_POKOK_OKT + Cdbl(to_nol(rsEPS("DEADTREE_OKT")))
			TOTAL_HA_AKHIR_OKT = TOTAL_HA_AKHIR_OKT + TotalHA_Akhir_OKT
			TOTAL_PKK_AKHIR_OKT = TOTAL_PKK_AKHIR_OKT + TotalPKK_Akhir_OKT
			if TOTAL_HA_AKHIR_OKT = 0 then TOTAL_SPH_AKHIR_OKT = 0 else TOTAL_SPH_AKHIR_OKT =  TOTAL_PKK_AKHIR_OKT / TOTAL_HA_AKHIR_OKT

			TOTAL_HA_TAMBAH_NOV = TOTAL_HA_TAMBAH_NOV + Cdbl(to_nol(rsEPS("HECTARE_NOV")))
			TOTAL_SISIP_POKOK_NOV = TOTAL_SISIP_POKOK_NOV + Cdbl(to_nol(rsEPS("COMPLETING_NOV")))
			TOTAL_COMP_POKOK_NOV = TOTAL_COMP_POKOK_NOV + Cdbl(to_nol(rsEPS("SUPPLYING_NOV")))
			TOTAL_TB_POKOK_NOV = TOTAL_TB_POKOK_NOV + Cdbl(to_nol(rsEPS("PLANTING_NOV")))
			TOTAL_KURANG_POKOK_NOV = TOTAL_KURANG_POKOK_NOV + Cdbl(to_nol(rsEPS("DEADTREE_NOV")))
			TOTAL_HA_AKHIR_NOV = TOTAL_HA_AKHIR_NOV + TotalHA_Akhir_NOV
			TOTAL_PKK_AKHIR_NOV = TOTAL_PKK_AKHIR_NOV + TotalPKK_Akhir_NOV
			if TOTAL_HA_AKHIR_NOV = 0 then TOTAL_SPH_AKHIR_NOV = 0 else TOTAL_SPH_AKHIR_NOV =  TOTAL_PKK_AKHIR_NOV / TOTAL_HA_AKHIR_NOV

			TOTAL_HA_TAMBAH_DES = TOTAL_HA_TAMBAH_DES + Cdbl(to_nol(rsEPS("HECTARE_DES")))
			TOTAL_SISIP_POKOK_DES = TOTAL_SISIP_POKOK_DES + Cdbl(to_nol(rsEPS("COMPLETING_DES")))
			TOTAL_COMP_POKOK_DES = TOTAL_COMP_POKOK_DES + Cdbl(to_nol(rsEPS("SUPPLYING_DES")))
			TOTAL_TB_POKOK_DES = TOTAL_TB_POKOK_DES + Cdbl(to_nol(rsEPS("PLANTING_DES")))
			TOTAL_KURANG_POKOK_DES = TOTAL_KURANG_POKOK_DES + Cdbl(to_nol(rsEPS("DEADTREE_DES")))
			TOTAL_HA_AKHIR_DES = TOTAL_HA_AKHIR_DES + TotalHA_Akhir_DES
			TOTAL_PKK_AKHIR_DES = TOTAL_PKK_AKHIR_DES + TotalPKK_Akhir_DES
			if TOTAL_HA_AKHIR_DES = 0 then TOTAL_SPH_AKHIR_DES = 0 else TOTAL_SPH_AKHIR_DES =  TOTAL_PKK_AKHIR_DES / TOTAL_HA_AKHIR_DES

			TOTAL_HA_TAMBAH_MUTASI = TOTAL_HA_TAMBAH_MUTASI + FinalHA_Tambah
			TOTAL_SISIP_POKOK_MUTASI = TOTAL_SISIP_POKOK_MUTASI + FinalSisip
			TOTAL_COMP_POKOK_MUTASI = TOTAL_COMP_POKOK_MUTASI + FinalCompleting
			TOTAL_TB_POKOK_MUTASI = TOTAL_TB_POKOK_MUTASI + FinalTB
			TOTAL_KURANG_POKOK_MUTASI = TOTAL_KURANG_POKOK_MUTASI + Cdbl(to_nol(rsEPS("DEADTREE_DES")))
			' Total = Cdbl(rsEPS("totalstand")) / Cdbl(rsEPS("hectarageplanted"))
%>
	<tr class="<%=cClass%>">
			<td align="center"><%=rsEPS("fcba")%></td>
			<td align="center"><%=rsEPS("Division")%></td>
			<td align="center"><%=rsEPS("plantingyear")%></td>
			<td align="center"><%=rsEPS("fieldcode")%></td>
			<td align="center"><%=rsEPS("ownership")%></td>
			<td align="center"><%=rsEPS("status")%></td>
			<td align="center"><%=rsEPS("hectarageplanted")%></td>

			<!-- Titik tanam -->
			<td align="center"><%=rsEPS("plantingpoint")%></td>
			<td align="center"><%=FormatNumber(SPH_TANAM,0)%></td>

			<!-- Program Sisip -->
			<td align="center"><%=formatnumber(PKK_SISIP,0)%></td>
			<td align="center"><%=formatnumber(SPH_SISIP,0)%></td>

			<!-- Awal -->
			<td align="center"><%=formatnumber(rsEPS("lineha"),2)%></td>
			<td align="center"><%=rsEPS("linepkk")%></td>
			<td align="center"><%=FormatNumber(TotalSPH_Awal,0)%></td>

			<!-- HA -->
			<td align="center"><%=rsEPS("HECTARE_JAN")%></td>
			<td align="center">0</td>

			<!-- Mutasi Pokok -->
			<td align="center"><%=rsEPS("SUPPLYING_JAN")%></td>
			<td align="center"><%=rsEPS("COMPLETING_JAN")%></td>
			<td align="center"><%=rsEPS("PLANTING_JAN")%></td>


			<!-- Kurang -->
			<td  colspan="2" align="center"><%=rsEPS("DEADTREE_JAN")%></td>

			<!-- Akhir -->
			<td align="center"><%=FormatNumber(TotalHA_Akhir_Jan,2)%></td>
			<td align="center"><%=formatnumber(TotalPKK_Akhir_Jan,0)%></td>
			<td align="center"><%=formatnumber(TotalSPH_Akhir_Jan,0)%></td>

			<!-- Februari -->
			<!-- Mutas HA -->
			<!-- HA -->
			<td align="center"><%=rsEPS("HECTARE_FEB")%></td>
			<td align="center">0</td>

			<!-- Mutasi Pokok -->
			<td align="center"><%=rsEPS("SUPPLYING_FEB")%></td>
			<td align="center"><%=rsEPS("COMPLETING_FEB")%></td>
			<td align="center"><%=rsEPS("PLANTING_FEB")%></td>


			<!-- Kurang -->
			<td  colspan="2" align="center"><%=rsEPS("DEADTREE_FEB")%></td>

			<!-- Akhir -->
			<td align="center"><%=FormatNumber(TotalHA_Akhir_Feb,2)%></td>
			<td align="center"><%=FormatNumber(TotalPKK_Akhir_Feb,0)%></td>
			<td align="center"><%=FormatNumber(TotalSPH_Akhir_Feb,0)%></td>

			<!-- Maret -->
			<!-- Mutas HA -->
			<!-- HA -->
			<td align="center"><%=rsEPS("HECTARE_MAR")%></td>
			<td align="center">0</td>

			<!-- Mutasi Pokok -->
			<td align="center"><%=rsEPS("SUPPLYING_MAR")%></td>
			<td align="center"><%=rsEPS("COMPLETING_MAR")%></td>
			<td align="center"><%=rsEPS("PLANTING_MAR")%></td>


			<!-- Kurang -->
			<td  colspan="2" align="center"><%=rsEPS("DEADTREE_MAR")%></td>

			<!-- Akhir -->
			<td align="center"><%=FormatNumber(TotalHA_Akhir_Mar,2)%></td>
			<td align="center"><%=FormatNumber(TotalPKK_Akhir_Mar,0)%></td>
			<td align="center"><%=FormatNumber(TotalSPH_Akhir_Mar,0)%></td>

			<!-- April -->
			<!-- Mutas HA -->
			<!-- HA -->
			<td align="center"><%=rsEPS("HECTARE_APR")%></td>
			<td align="center">0</td>

			<!-- Mutasi Pokok -->
			<td align="center"><%=rsEPS("SUPPLYING_APR")%></td>
			<td align="center"><%=rsEPS("COMPLETING_APR")%></td>
			<td align="center"><%=rsEPS("PLANTING_APR")%></td>


			<!-- Kurang -->
			<td  colspan="2" align="center"><%=rsEPS("DEADTREE_APR")%></td>

			<!-- Akhir -->
			<td align="center"><%=FormatNumber(TotalHA_Akhir_Apr,2)%></td>
			<td align="center"><%=FormatNumber(TotalPKK_Akhir_Apr,0)%></td>
			<td align="center"><%=FormatNumber(TotalSPH_Akhir_Apr,0)%></td>

			<!-- Mei -->
			<!-- Mutas HA -->
			<!-- HA -->
			<td align="center"><%=rsEPS("HECTARE_MEI")%></td>
			<td align="center">0</td>

			<!-- Mutasi Pokok -->
			<td align="center"><%=rsEPS("SUPPLYING_MEI")%></td>
			<td align="center"><%=rsEPS("COMPLETING_MEI")%></td>
			<td align="center"><%=rsEPS("PLANTING_MEI")%></td>


			<!-- Kurang -->
			<td  colspan="2" align="center"><%=rsEPS("DEADTREE_MEI")%></td>

			<!-- Akhir -->
			<td align="center"><%=FormatNumber(TotalHA_Akhir_Mei,2)%></td>
			<td align="center"><%=FormatNumber(TotalPKK_Akhir_Mei,0)%></td>
			<td align="center"><%=FormatNumber(TotalSPH_Akhir_Mei,0)%></td>

			<!-- Juni -->
			<!-- Mutas HA -->
			<!-- HA -->
			<td align="center"><%=rsEPS("HECTARE_JUN")%></td>
			<td align="center">0</td>

			<!-- Mutasi Pokok -->
			<td align="center"><%=rsEPS("SUPPLYING_JUN")%></td>
			<td align="center"><%=rsEPS("COMPLETING_JUN")%></td>
			<td align="center"><%=rsEPS("PLANTING_JUN")%></td>


			<!-- Kurang -->
			<td  colspan="2" align="center"><%=rsEPS("DEADTREE_JUN")%></td>

			<!-- Akhir -->
			<td align="center"><%=FormatNumber(TotalHA_Akhir_Jun,2)%></td>
			<td align="center"><%=FormatNumber(TotalPKK_Akhir_Jun,0)%></td>
			<td align="center"><%=FormatNumber(TotalSPH_Akhir_Jun,0)%></td>

			<!-- Juli -->
			<!-- Mutas HA -->
			<!-- HA -->
			<td align="center"><%=rsEPS("HECTARE_JUL")%></td>
			<td align="center">0</td>

			<!-- Mutasi Pokok -->
			<td align="center"><%=rsEPS("SUPPLYING_JUL")%></td>
			<td align="center"><%=rsEPS("COMPLETING_JUL")%></td>
			<td align="center"><%=rsEPS("PLANTING_JUL")%></td>


			<!-- Kurang -->
			<td  colspan="2" align="center"><%=rsEPS("DEADTREE_JUL")%></td>

			<!-- Akhir -->
			<td align="center"><%=FormatNumber(TotalHA_Akhir_Jul,2)%></td>
			<td align="center"><%=FormatNumber(TotalPKK_Akhir_Jul,0)%></td>
			<td align="center"><%=FormatNumber(TotalSPH_Akhir_Jul,0)%></td>

			<!-- Agustus -->
			<!-- Mutas HA -->
			<!-- HA -->
			<td align="center"><%=rsEPS("HECTARE_AUG")%></td>
			<td align="center">0</td>

			<!-- Mutasi Pokok -->
			<td align="center"><%=rsEPS("SUPPLYING_AUG")%></td>
			<td align="center"><%=rsEPS("COMPLETING_AUG")%></td>
			<td align="center"><%=rsEPS("PLANTING_AUG")%></td>


			<!-- Kurang -->
			<td  colspan="2" align="center"><%=rsEPS("DEADTREE_AUG")%></td>

			<!-- Akhir -->
			<td align="center"><%=FormatNumber(TotalHA_Akhir_Aug,2)%></td>
			<td align="center"><%=FormatNumber(TotalPKK_Akhir_Aug,0)%></td>
			<td align="center"><%=FormatNumber(TotalSPH_Akhir_Aug,0)%></td>

			<!-- September -->
			<!-- Mutas HA -->
			<!-- HA -->
			<td align="center"><%=rsEPS("HECTARE_SEP")%></td>
			<td align="center">0</td>

			<!-- Mutasi Pokok -->
			<td align="center"><%=rsEPS("SUPPLYING_SEP")%></td>
			<td align="center"><%=rsEPS("COMPLETING_SEP")%></td>
			<td align="center"><%=rsEPS("PLANTING_SEP")%></td>


			<!-- Kurang -->
			<td  colspan="2" align="center"><%=rsEPS("DEADTREE_SEP")%></td>

			<!-- Akhir -->
			<td align="center"><%=FormatNumber(TotalHA_Akhir_Sep,2)%></td>
			<td align="center"><%=FormatNumber(TotalPKK_Akhir_Sep,0)%></td>
			<td align="center"><%=FormatNumber(TotalSPH_Akhir_Sep,0)%></td>

			<!-- Oktober -->
			<!-- Mutas HA -->
			<!-- HA -->
			<td align="center"><%=rsEPS("HECTARE_OKT")%></td>
			<td align="center">0</td>

			<!-- Mutasi Pokok -->
			<td align="center"><%=rsEPS("SUPPLYING_OKT")%></td>
			<td align="center"><%=rsEPS("COMPLETING_OKT")%></td>
			<td align="center"><%=rsEPS("PLANTING_OKT")%></td>


			<!-- Kurang -->
			<td  colspan="2" align="center"><%=rsEPS("DEADTREE_OKT")%></td>

			<!-- Akhir -->
			<td align="center"><%=FormatNumber(TotalHA_Akhir_Okt,2)%></td>
			<td align="center"><%=FormatNumber(TotalPKK_Akhir_Okt,0)%></td>
			<td align="center"><%=FormatNumber(TotalSPH_Akhir_Okt,0)%></td>

			<!-- November -->
			<!-- Mutas HA -->
			<!-- HA -->
			<td align="center"><%=rsEPS("HECTARE_NOV")%></td>
			<td align="center">0</td>

			<!-- Mutasi Pokok -->
			<td align="center"><%=rsEPS("SUPPLYING_NOV")%></td>
			<td align="center"><%=rsEPS("COMPLETING_NOV")%></td>
			<td align="center"><%=rsEPS("PLANTING_NOV")%></td>


			<!-- Kurang -->
			<td  colspan="2" align="center"><%=rsEPS("DEADTREE_NOV")%></td>

			<!-- Akhir -->
			<td align="center"><%=FormatNumber(TotalHA_Akhir_Nov,2)%></td>
			<td align="center"><%=FormatNumber(TotalPKK_Akhir_Nov,0)%></td>
			<td align="center"><%=FormatNumber(TotalSPH_Akhir_Nov,0)%></td>

			<!-- Desember -->
			<!-- Mutas HA -->
			<!-- HA -->
			<td align="center"><%=rsEPS("HECTARE_DES")%></td>
			<td align="center">0</td>

			<!-- Mutasi Pokok -->
			<td align="center"><%=rsEPS("SUPPLYING_DES")%></td>
			<td align="center"><%=rsEPS("COMPLETING_DES")%></td>
			<td align="center"><%=rsEPS("PLANTING_DES")%></td>


			<!-- Kurang -->
			<td  colspan="2" align="center"><%=rsEPS("DEADTREE_DES")%></td>

			<!-- Akhir -->
			<td align="center"><%=FormatNumber(TotalHA_Akhir_Des,2)%></td>
			<td align="center"><%=FormatNumber(TotalPKK_Akhir_Des,0)%></td>
			<td align="center"><%=FormatNumber(TotalSPH_Akhir_Des,0)%></td>

			<!-- Final -->
			<!-- Mutas HA -->
			<!-- HA -->
			<td align="center"><%=FormatNumber(FinalHA_Tambah,2)%></td>
			<td align="center">0</td>

			<!-- Mutasi Pokok -->
			<td align="center"><%=FormatNumber(FinalSisip,0)%></td>
			<td align="center"><%=FormatNumber(FinalCompleting,0)%></td>
			<td align="center"><%=FormatNumber(FinalTB,0)%></td>


			<!-- Kurang -->
			<td align="center"><%=rsEPS("DEADTREE_DES")%></td>
			<%if cdbl(to_nol(TotalSPH_Akhir_Des)) >= cdbl(to_nol(SPH_TANAM)) then%>
				<td align="center"> FullStand</td>
			<%else%>
				<td align="center">Not FullStand</td>
			<%end if%>
	</tr>

<%
			rsEPS.MoveNext
		Wend
	end if
		
rsEPS.Close
set rsEPS = Nothing

%>
	<tr>
		<td colspan="6" align="center" style="background: #F4F4A0;">TOTAL</td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_HA,2)%></td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_PKK_TANAM,0)%></td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_SPH_TANAM,0)%></td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_PKK_SISIP,0)%></td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_SPH_SISIP,0)%></td>

		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_HA_AWAL_JAN,2)%></td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_PKK_AWAL_JAN,0)%></td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_SPH_AWAL_JAN,0)%></td>

		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_HA_TAMBAH_JAN,2)%></td>
		<td align="center" style="background: #F4F4A0;">0.00</td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_COMP_POKOK_JAN,0)%></td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_SISIP_POKOK_JAN,0)%></td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_TB_POKOK_JAN,0)%></td>
		<td colspan="2" align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_KURANG_POKOK_JAN,0)%></td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_HA_AKHIR_JAN,2)%></td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_PKK_AKHIR_JAN,0)%></td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_SPH_AKHIR_JAN,0)%></td>

		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_HA_TAMBAH_FEB,2)%></td>
		<td align="center" style="background: #F4F4A0;">0.00</td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_COMP_POKOK_FEB,0)%></td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_SISIP_POKOK_FEB,0)%></td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_TB_POKOK_FEB,0)%></td>
		<td colspan="2" align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_KURANG_POKOK_FEB,0)%></td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_HA_AKHIR_FEB,2)%></td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_PKK_AKHIR_FEB,0)%></td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_SPH_AKHIR_FEB,0)%></td>

		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_HA_TAMBAH_MAR,2)%></td>
		<td align="center" style="background: #F4F4A0;">0.00</td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_COMP_POKOK_MAR,0)%></td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_SISIP_POKOK_MAR,0)%></td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_TB_POKOK_MAR,0)%></td>
		<td colspan="2" align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_KURANG_POKOK_MAR,0)%></td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_HA_AKHIR_MAR,2)%></td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_PKK_AKHIR_MAR,0)%></td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_SPH_AKHIR_MAR,0)%></td>

		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_HA_TAMBAH_APR,2)%></td>
		<td align="center" style="background: #F4F4A0;">0.00</td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_COMP_POKOK_APR,0)%></td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_SISIP_POKOK_APR,0)%></td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_TB_POKOK_APR,0)%></td>
		<td colspan="2" align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_KURANG_POKOK_APR,0)%></td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_HA_AKHIR_APR,2)%></td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_PKK_AKHIR_APR,0)%></td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_SPH_AKHIR_APR,0)%></td>

		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_HA_TAMBAH_MEI,2)%></td>
		<td align="center" style="background: #F4F4A0;">0.00</td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_COMP_POKOK_MEI,0)%></td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_SISIP_POKOK_MEI,0)%></td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_TB_POKOK_MEI,0)%></td>
		<td colspan="2" align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_KURANG_POKOK_MEI,0)%></td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_HA_AKHIR_MEI,2)%></td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_PKK_AKHIR_MEI,0)%></td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_SPH_AKHIR_MEI,0)%></td>

		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_HA_TAMBAH_JUN,2)%></td>
		<td align="center" style="background: #F4F4A0;">0.00</td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_COMP_POKOK_JUN,0)%></td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_SISIP_POKOK_JUN,0)%></td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_TB_POKOK_JUN,0)%></td>
		<td colspan="2" align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_KURANG_POKOK_JUN,0)%></td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_HA_AKHIR_JUN,2)%></td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_PKK_AKHIR_JUN,0)%></td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_SPH_AKHIR_JUN,0)%></td>

		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_HA_TAMBAH_JUL,2)%></td>
		<td align="center" style="background: #F4F4A0;">0.00</td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_COMP_POKOK_JUL,0)%></td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_SISIP_POKOK_JUL,0)%></td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_TB_POKOK_JUL,0)%></td>
		<td colspan="2" align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_KURANG_POKOK_JUL,0)%></td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_HA_AKHIR_JUL,2)%></td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_PKK_AKHIR_JUL,0)%></td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_SPH_AKHIR_JUL,0)%></td>

		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_HA_TAMBAH_AUG,2)%></td>
		<td align="center" style="background: #F4F4A0;">0.00</td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_COMP_POKOK_AUG,0)%></td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_SISIP_POKOK_AUG,0)%></td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_TB_POKOK_AUG,0)%></td>
		<td colspan="2" align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_KURANG_POKOK_AUG,0)%></td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_HA_AKHIR_AUG,2)%></td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_PKK_AKHIR_AUG,0)%></td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_SPH_AKHIR_AUG,0)%></td>

		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_HA_TAMBAH_SEP,2)%></td>
		<td align="center" style="background: #F4F4A0;">0.00</td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_COMP_POKOK_SEP,0)%></td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_SISIP_POKOK_SEP,0)%></td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_TB_POKOK_SEP,0)%></td>
		<td colspan="2" align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_KURANG_POKOK_SEP,0)%></td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_HA_AKHIR_SEP,2)%></td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_PKK_AKHIR_SEP,0)%></td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_SPH_AKHIR_SEP,0)%></td>

		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_HA_TAMBAH_OKT,2)%></td>
		<td align="center" style="background: #F4F4A0;">0.00</td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_COMP_POKOK_OKT,0)%></td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_SISIP_POKOK_OKT,0)%></td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_TB_POKOK_OKT,0)%></td>
		<td colspan="2" align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_KURANG_POKOK_OKT,0)%></td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_HA_AKHIR_OKT,2)%></td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_PKK_AKHIR_OKT,0)%></td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_SPH_AKHIR_OKT,0)%></td>

		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_HA_TAMBAH_NOV,2)%></td>
		<td align="center" style="background: #F4F4A0;">0.00</td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_COMP_POKOK_NOV,0)%></td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_SISIP_POKOK_NOV,0)%></td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_TB_POKOK_NOV,0)%></td>
		<td colspan="2" align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_KURANG_POKOK_NOV,0)%></td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_HA_AKHIR_NOV,2)%></td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_PKK_AKHIR_NOV,0)%></td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_SPH_AKHIR_NOV,0)%></td>

		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_HA_TAMBAH_DES,2)%></td>
		<td align="center" style="background: #F4F4A0;">0.00</td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_COMP_POKOK_DES,0)%></td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_SISIP_POKOK_DES,0)%></td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_TB_POKOK_DES,0)%></td>
		<td colspan="2" align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_KURANG_POKOK_DES,0)%></td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_HA_AKHIR_DES,2)%></td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_PKK_AKHIR_DES,0)%></td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_SPH_AKHIR_DES,0)%></td>

		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_HA_TAMBAH_MUTASI,2)%></td>
		<td align="center" style="background: #F4F4A0;">0.00</td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_SISIP_POKOK_MUTASI,0)%></td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_COMP_POKOK_MUTASI,0)%></td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_TB_POKOK_MUTASI,0)%></td>
		<td align="center" style="background: #F4F4A0;"><%=FormatNumber(TOTAL_KURANG_POKOK_MUTASI,0)%></td>
		<td style="background: #F4F4A0;"></td>
	</tr>
</tbody>
</table>
</form>
</div>
</body>
</html>

