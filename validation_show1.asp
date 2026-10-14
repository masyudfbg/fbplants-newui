<%IF Request("txttable") ="Bank" then%>
	<script language="JavaScript" type="text/javascript">
	//You should create the validator only after the definition of the HTML form
	var frmvalidator  = new Validator("frmEntry");
	frmvalidator.addValidation("FCCODE","req","Please enter Code");
	frmvalidator.addValidation("FCNAME","req","Please enter Description");
	frmvalidator.addValidation("CONTROLJOB","req","Please enter Control Job");
	</script>
<%elseIF Request("txttable") ="Stock" then%>
	<script language="JavaScript" type="text/javascript">
	//You should create the validator only after the definition of the HTML form
	var frmvalidator  = new Validator("frmEntry");
	frmvalidator.addValidation("FCCODE","req","Please enter Code");
	frmvalidator.addValidation("FCNAME","req","Please enter Description");
	frmvalidator.addValidation("UNITOFMEASUREMENT","req","Please enter Unit Of Measurement");
	frmvalidator.addValidation("STOCKGROUPCODE","req","Please enter Stock Group Code");
	</script>
<%elseif Request("txttable") ="StockGroup" then%>
	<script language="JavaScript" type="text/javascript">
	//You should create the validator only after the definition of the HTML form
	var frmvalidator  = new Validator("frmEntry");
	frmvalidator.addValidation("FCCODE","req","Please enter Code");
	frmvalidator.addValidation("FCNAME","req","Please enter Description");
	frmvalidator.addValidation("CONTROLJOB","req","Please enter Control Job");
	</script>
<%elseif Request("txttable") ="Supplier" then%>
	<script language="JavaScript" type="text/javascript">
	//You should create the validator only after the definition of the HTML form
	var frmvalidator  = new Validator("frmEntry");
	frmvalidator.addValidation("FCCODE","req","Please enter Code");
	frmvalidator.addValidation("FCNAME","req","Please enter Description");
	frmvalidator.addValidation("CONTROLJOB","req","Please enter Control Job");
	</script>
<%elseif Request("txttable") ="Store" then%>
	<script language="JavaScript" type="text/javascript">
	//You should create the validator only after the definition of the HTML form
	var frmvalidator  = new Validator("frmEntry");
	frmvalidator.addValidation("FCCODE","req","Please enter Code");
	frmvalidator.addValidation("FCNAME","req","Please enter Description");
	</script>
<%elseif Request("txttable") ="Gang" then%>
	<script language="JavaScript" type="text/javascript">
	//You should create the validator only after the definition of the HTML form
	var frmvalidator  = new Validator("frmEntry");
	frmvalidator.addValidation("FCCODE","req","Please enter Code");
	frmvalidator.addValidation("FCNAME","req","Please enter Description");
	</script>
<%elseif Request("txttable") ="Workshop" then%>
	<script language="JavaScript" type="text/javascript">
	//You should create the validator only after the definition of the HTML form
	var frmvalidator  = new Validator("frmEntry");
	frmvalidator.addValidation("FCCODE","req","Please enter Code");
	frmvalidator.addValidation("FCNAME","req","Please enter Description");
	</script>
<%Elseif Request("txttable") ="Vehicle" then%>
	<script language="JavaScript" type="text/javascript">
	//You should create the validator only after the definition of the HTML form
	var frmvalidator  = new Validator("frmEntry");
	frmvalidator.addValidation("FCCODE","req","Please enter Code");
	frmvalidator.addValidation("FCNAME","req","Please enter Description");
	frmvalidator.addValidation("VEHICLEGROUPCODE","req","Please enter Vehicle Group Code");
	frmvalidator.addValidation("CONTROLJOB","req","Please enter Control Job");
	</script>
<%elseif Request("txttable") ="Contractor" then%>
	<script language="JavaScript" type="text/javascript">
	//You should create the validator only after the definition of the HTML form
	var frmvalidator  = new Validator("frmEntry");
	frmvalidator.addValidation("FCCODE","req","Please enter Code");
	frmvalidator.addValidation("FCNAME","req","Please enter Description");
	frmvalidator.addValidation("CONTROLJOB","req","Please enter Control Job");
	</script>
<%elseif Request("txttable") ="Building" then%>
	<script language="JavaScript" type="text/javascript">
	//You should create the validator only after the definition of the HTML form
	var frmvalidator  = new Validator("frmEntry");
	frmvalidator.addValidation("FCCODE","req","Please enter Code");
	frmvalidator.addValidation("FCNAME","req","Please enter Description");
	</script>
<%elseif Request("txttable") ="Station" then%>
	<script language="JavaScript" type="text/javascript">
	//You should create the validator only after the definition of the HTML form
	var frmvalidator  = new Validator("frmEntry");
	frmvalidator.addValidation("FCCODE","req","Please enter Code");
	frmvalidator.addValidation("FcName","req","Please enter Description");
	frmvalidator.addValidation("JOBCODE","req","Please enter Control Job");
	</script>
<%elseif Request("txttable") ="Machine" then%>
	<script language="JavaScript" type="text/javascript">
	//You should create the validator only after the definition of the HTML form
	var frmvalidator  = new Validator("frmEntry");
	frmvalidator.addValidation("FCCODE","req","Please enter Code");
	frmvalidator.addValidation("DESCRIPTION","req","Please enter Description");
	frmvalidator.addValidation("STATION_STATIONCODE","req","Please enter Station");
	frmvalidator.addValidation("PURCHASECOST","numeric");
	frmvalidator.addValidation("CONTROLJOB","req","Please enter Control Job");
	</script>
<%elseif Request("txttable") ="BusinessUnit" then%>
	<script language="JavaScript" type="text/javascript">
	//You should create the validator only after the definition of the HTML form
	var frmvalidator  = new Validator("frmEntry");
	frmvalidator.addValidation("FCCODE","req","Please enter Code");
	frmvalidator.addValidation("FCNAME","req","Please enter Name");
	frmvalidator.addValidation("FCCOMPANYCODE","req","Please enter Company Code");
	frmvalidator.addValidation("FCCOMPANYNAME","req","Please enter Company Name");	
	frmvalidator.addValidation("FCAREA","req","Please enter Area");		
	frmvalidator.addValidation("FCOWNERSHIP","req","Please enter OwnerShip");
	</script>
<%elseif Request("txttable") ="UserProfile" then%>
	<script language="JavaScript" type="text/javascript">
	//You should create the validator only after the definition of the HTML form
	var frmvalidator  = new Validator("frmEntry");
	frmvalidator.addValidation("FCCODE","req","Please enter Code");
	frmvalidator.addValidation("FCPASSWORD","req","Please enter Password");
	frmvalidator.addValidation("FULLNAME","req","Please enter Full Name");	
	frmvalidator.addValidation("DEPARTMENT","req","Please enter Departement");		
	frmvalidator.addValidation("ACCESSRIGHTS","req","Please enter Access Right");
	</script>
<%elseif Request("txttable") ="Parameter" then%>
	<script language="JavaScript" type="text/javascript">
	//You should create the validator only after the definition of the HTML form
	var frmvalidator  = new Validator("frmEntry");
	frmvalidator.addValidation("FCCODE","req","Please enter Code");
	frmvalidator.addValidation("PARAMETERNAME","req","Please enter Parameter Name");
	</script>
<%elseif Request("txttable") ="Trigger_Control" then%>
	<script language="JavaScript" type="text/javascript">
	//You should create the validator only after the definition of the HTML form
	var frmvalidator  = new Validator("frmEntry");
	frmvalidator.addValidation("ACTIVITY","req","Please enter Activity");
	frmvalidator.addValidation("SOURCE","req","Please enter Source");
	frmvalidator.addValidation("TARGET","req","Please enter Target");
	</script>
<%elseif Request("txttable") ="Job" then%>
	<script language="JavaScript" type="text/javascript">
	//You should create the validator only after the definition of the HTML form
	var frmvalidator  = new Validator("frmEntry");
	frmvalidator.addValidation("FCCODE","req","Please enter Code");
	</script>
<%elseif Request("txttable") ="JobGroup" then%>
	<script language="JavaScript" type="text/javascript">
	//You should create the validator only after the definition of the HTML form
	var frmvalidator  = new Validator("frmEntry");
	frmvalidator.addValidation("FCCODE","req","Please enter Code");
	</script>
<%elseif Request("txttable") ="GCMaster" then%>
	<script language="JavaScript" type="text/javascript">
	//You should create the validator only after the definition of the HTML form
	var frmvalidator  = new Validator("frmEntry");
	frmvalidator.addValidation("FCCODE","req","Please enter Type");
	frmvalidator.addValidation("ITEM","req","Please enter Item");
	</script>
<%elseif Request("txttable") ="Employee" then%>
	<script language="JavaScript" type="text/javascript">
	//You should create the validator only after the definition of the HTML form
	var frmvalidator  = new Validator("frmEntry");
	frmvalidator.addValidation("FCCODE","req","Please enter Code");
	frmvalidator.addValidation("FCNAME","req","Please enter Name");
	frmvalidator.addValidation("SECTIONNAME","req","Please enter Section Name");
	frmvalidator.addValidation("GANGCODE","req","Please enter Gang Code");
	frmvalidator.addValidation("RELIGION","req","Please enter Religion");
	frmvalidator.addValidation("RACE","req","Please enter Race");
	frmvalidator.addValidation("NATIONALITY","req","Please enter Nationality");
	frmvalidator.addValidation("MARITALSTATUS","req","Please enter Marital Status");
	frmvalidator.addValidation("JAMSOSTEKCATEGORY","req","Please enter Jamostek Category");
	frmvalidator.addValidation("POSITION","req","Please enter Position");
	frmvalidator.addValidation("PRESENTRATE","numeric");
	frmvalidator.addValidation("PREVIOUSRATE","numeric");			
	</script>
<%elseif Request("txttable") ="Allowded" then%>
	<script language="JavaScript" type="text/javascript">
	//You should create the validator only after the definition of the HTML form
	var frmvalidator  = new Validator("frmEntry");
	frmvalidator.addValidation("FCCODE","req","Please enter Code");
	
	frmvalidator.addValidation("FCNAME","req","Please enter Name");
	frmvalidator.addValidation("ALLDED_GROUP","req","Please enter Allowded Group");
	frmvalidator.addValidation("ALLDED","req","Please enter Allded");
	frmvalidator.addValidation("CROP","req","Please enter Crop");
	frmvalidator.addValidation("CONTROL_JOB","req","Please enter Control Job");
	frmvalidator.addValidation("MAN_AUTO","req","Please enter Crop");
	//frmvalidator.addValidation("TYPE_1","numeric");
	//frmvalidator.addValidation("TYPE_2","numeric");
	//frmvalidator.addValidation("TYPE_3","numeric");
	//frmvalidator.addValidation("TYPE_4","numeric");
	//frmvalidator.addValidation("MIN_AMT","numeric");
	//frmvalidator.addValidation("MAX_AMT","numeric");
	//frmvalidator.addValidation("RATE_1","numeric");
	//frmvalidator.addValidation("RATE_2","numeric");
	//frmvalidator.addValidation("RATE_3","numeric");
	//frmvalidator.addValidation("RATE_4","numeric");
	</Script>
<%elseif Request("txttable") ="Field" then%>
	<script language="JavaScript" type="text/javascript">
	//You should create the validator only after the definition of the HTML form
	var frmvalidator  = new Validator("frmEntry");
	frmvalidator.addValidation("FCCODE","req","Please enter Code");
	frmvalidator.addValidation("FCNAME","req","Please enter Name");
	frmvalidator.addValidation("DIVISION","req","Please enter Division");
	frmvalidator.addValidation("PLANTINGDATE","req","Please enter Planting Date");
	frmvalidator.addValidation("STATUS","req","Please enter Status");
	frmvalidator.addValidation("CROPTYPE","req","Please enter Crop Type");
	frmvalidator.addValidation("SOILTYPE","req","Please enter Soil Type");
	frmvalidator.addValidation("TERRAINTYPE","req","Please enter Terrain Type");
	frmvalidator.addValidation("PLANTINGMATERIAL","req","Please enter Planting Material");
	frmvalidator.addValidation("TOTALSTAND","req","Please enter Total Stand");
	frmvalidator.addValidation("OWNERSHIP","req","Please enter Ownership");
	</Script>
<%elseif Request("txttable") ="Harvesting" then%>
	<script language="JavaScript" type="text/javascript">
	//You should create the validator only after the definition of the HTML form	
	var frmvalidator  = new Validator("eForm");
	frmvalidator.addValidation("RECEPTIONDATE","req","Please enter Reception Date");
	frmvalidator.addValidation("CHITNO","req","Please enter Chitno");
	frmvalidator.addValidation("FIELDCODE","req","Please enter Field Code");
	frmvalidator.addValidation("CROPCODE","req","Please enter Crop Code");
	frmvalidator.addValidation("ProductCODE","req","Please enter Product Code");
	frmvalidator.addValidation("Vehicle","req","Please enter Vehicle");
	frmvalidator.addValidation("QUANTITY","numeric");
	frmvalidator.addValidation("ESTATE_WEIGHT","numeric");
	frmvalidator.addValidation("MILL_WEIGHT","numeric");
	</script>
<%elseif Request("txttable") ="HarvestingRound" then%>
	<script language="JavaScript" type="text/javascript">
	//You should create the validator only after the definition of the HTML form
	var frmvalidator  = new Validator("eForm");
	frmvalidator.addValidation("HRNDYEAR","req","Please enter Year");
	frmvalidator.addValidation("HRNDMONTH","req","Please enter Month");	
	frmvalidator.addValidation("FIELDCODE","req","Please enter Field Code");
	frmvalidator.addValidation("CropCode","req","Please enter Crop Code");
	frmvalidator.addValidation("HRNDYEAR","numeric");
	frmvalidator.addValidation("HRNDMONTH","numeric");
	frmvalidator.addValidation("NUMBEROFROUNDS","numeric");
	frmvalidator.addValidation("HECTARAGE","numeric");
	frmvalidator.addValidation("STANDPERHECTHARVESTED","numeric");
	</script>
<%elseif Request("txttable") ="Weather" then%>
	<script language="JavaScript" type="text/javascript">
	//You should create the validator only after the definition of the HTML form
	var frmvalidator  = new Validator("eForm");
	frmvalidator.addValidation("WIDATE","req","Please enter Date");
	frmvalidator.addValidation("WIDIVISION","req","Please enter Division");	
	//frmvalidator.addValidation("TOTRNINCHES","numeric");
	frmvalidator.addValidation("TOTRNDAYS","numeric");
	frmvalidator.addValidation("AVGTEMP","numeric");
	frmvalidator.addValidation("AVGWIND","numeric");
	</script>	
	
<%elseif Request("txttable") ="OPHarvRate" then%>
	<script language="JavaScript" type="text/javascript">
	//You should create the validator only after the definition of the HTML form
	var frmvalidator  = new Validator("eForm");
	frmvalidator.addValidation("CROPCODE","req","Please enter Crop Code");
	frmvalidator.addValidation("TERRAIN_TYPE","req","Please enter Terrain Type");	
	frmvalidator.addValidation("ABW_CENSUS","req","Please enter Cencus");	
	frmvalidator.addValidation("ABW_CENSUS","numeric");
	frmvalidator.addValidation("ABW_MIN","numeric");
	frmvalidator.addValidation("ABW_MAX","numeric");
	frmvalidator.addValidation("BASIS_QTY","numeric");
	frmvalidator.addValidation("BASIS_RATE","numeric");	
	frmvalidator.addValidation("EXTRA_BASIS_RATE","numeric");	
	</script> 
<%elseif Request("txttable") ="PRODUCT" then%>
	<script language="JavaScript" type="text/javascript">
	//You should create the validator only after the definition of the HTML form
	var frmvalidator  = new Validator("eForm");
	frmvalidator.addValidation("PRODUCTCODE","req","Please enter Code");
	frmvalidator.addValidation("PRODUCTDESCRIPTION","req","Please enter Description");	
	frmvalidator.addValidation("CROPCODE","req","Please enter Crop Code");	
	frmvalidator.addValidation("PRODUCTDESCRIPTION","req","Please enter Product Description");	
	
	</script> 
<%elseif Request("txttable") ="WEIGHBRIDGE" then%>
	<script language="JavaScript" type="text/javascript">
	//You should create the validator only after the definition of the HTML form
	var frmvalidator  = new Validator("eForm");
	frmvalidator.addValidation("WEIGHBRIDGECODE","req","Please enter Code");
	frmvalidator.addValidation("TXTDESCRIPTION","req","Please enter Description");	
	</script>

<%elseif Request("txttable") ="CUSTOMER" then%>
	<script language="JavaScript" type="text/javascript">
	//You should create the validator only after the definition of the HTML form
	var frmvalidator  = new Validator("eForm");
	frmvalidator.addValidation("CUSTOMERCODE","req","Please enter Code");
	frmvalidator.addValidation("txtDESCRIPTION","req","Please enter Description");		
	</script>
<%elseif Request("txttable") ="GRADINGITEM" then%>
	<script language="JavaScript" type="text/javascript">
	//You should create the validator only after the definition of the HTML form
	var frmvalidator  = new Validator("eForm");
	frmvalidator.addValidation("ITEMCODE","req","Please enter Item Code");
	frmvalidator.addValidation("txtDESCRIPTION","req","Please enter Description");	
	frmvalidator.addValidation("REDUCTION","numeric");		
	</script>		
<%elseif Request("txttable") ="WEIGHBRIDGETICKET" then%>,
	<script language="JavaScript" type="text/javascript">
	//You should create the validator only after the definition of the HTML form
	var frmvalidator  = new Validator("eForm");
	frmvalidator.addValidation("PRODUCT_PRODUCTCODE","req","Please enter Product Code");
	frmvalidator.addValidation("CHITNO","req","Please enter Chitno");	
	frmvalidator.addValidation("VEHICLENO","req","Please enter Vehicleno");	
	frmvalidator.addValidation("WEIGHTIN","req","Please enter Weight in");	
	frmvalidator.addValidation("WEIGHTOUT","req","Please enter Weight Out");	
	frmvalidator.addValidation("CUST_CUSTOMERCODE","req","Please enter Customer Code");	
	frmvalidator.addValidation("DATEIN","req","Please enter Date In");	
	frmvalidator.addValidation("DATEOUT","req","Please enter Date Out");	
	frmvalidator.addValidation("WBRIDGE_WEIGHBRIDGECODEIN","req","Please enter Date Weighbridge Code In");	
	frmvalidator.addValidation("WBRIDGE_WEIGHBRIDGECODEOUT","req","Please enter Date Weighbridge Code Out");	
	frmvalidator.addValidation("WEIGHTIN","numeric");
	frmvalidator.addValidation("WEIGHTOUT","numeric");	
	</script>
<%elseif Request("txttable") ="GRADINGRECORD" then%>
	<script language="JavaScript" type="text/javascript">
	//You should create the validator only after the definition of the HTML form
	var frmvalidator  = new Validator("eForm");
	frmvalidator.addValidation("GRADINGNO","req","Please enter Grading No");
	frmvalidator.addValidation("DATETIME","req","Please enter Date Time");	
	frmvalidator.addValidation("VEHICLENO","req","Please enter Vehicle No");	
	frmvalidator.addValidation("WEIGHT","req","Please enter Weight");	
	frmvalidator.addValidation("BUNCHES","req","Please enter Bunches");	
	frmvalidator.addValidation("WBT_CHITNO","req","Please enter Chitno");
	frmvalidator.addValidation("WEIGHT","numeric");
	frmvalidator.addValidation("BUNCHES","numeric");		
	</script>
<%elseif Request("txttable") ="SHIFT" then%>
	<script language="JavaScript" type="text/javascript">
	//You should create the validator only after the definition of the HTML form
	var frmvalidator  = new Validator("eForm");
	frmvalidator.addValidation("SHIFTCODE","req","Please enter Code");
	frmvalidator.addValidation("STARTTIME","req","Please enter Start Time");	
	frmvalidator.addValidation("ENDTIME","req","Please enter End Time");	
	</script>
<%elseif Request("txttable") ="SHIFTREPORT" then%>
	<script language="JavaScript" type="text/javascript">
	//You should create the validator only after the definition of the HTML form
	var frmvalidator  = new Validator("eForm");
	frmvalidator.addValidation("PRODUCTIONDATE","req","Please enter Production Date");
	frmvalidator.addValidation("ACTUALSTART","req","Please enter Actual Start");	
	frmvalidator.addValidation("ACTUALEND","req","Please enter Actual End");	
	frmvalidator.addValidation("CAGESPROCESSED","req","Please enter Cages Processed");	
	frmvalidator.addValidation("RAMPBALANCE","req","Please enter Ramp Balance");	
	frmvalidator.addValidation("CAGESBALANCE","req","Please enter Cages Balance");
	frmvalidator.addValidation("PLATFORMBALANCE","req","Please enter Platform Balance");
	frmvalidator.addValidation("SHIFT_SHIFTCODE","req","Please enter Shif Code");
	frmvalidator.addValidation("CAGEWEIGHT","numeric");
	frmvalidator.addValidation("CAGESPROCESSED","numeric");
	frmvalidator.addValidation("RAMPBALANCE","numeric");		
	frmvalidator.addValidation("CAGESBALANCE","numeric");
	frmvalidator.addValidation("PLATFORMBALANCE","numeric");		
	</script>

<%elseif Request("txttable") ="PRODUCTSTORAGE" then%>
	<script language="JavaScript" type="text/javascript">
	//You should create the validator only after the definition of the HTML form
	var frmvalidator  = new Validator("eForm");
	frmvalidator.addValidation("STORAGECODE","req","Please enter Code");
	frmvalidator.addValidation("txtDESCRIPTION","req","Please enter Description");	
	frmvalidator.addValidation("PRODUCT_PRODUCTCODE","req","Please enter Product Code");	
	frmvalidator.addValidation("MAXCAPACITY","req","Please enter Max Capacity");
	frmvalidator.addValidation("MAXCAPACITY","numeric");		
	frmvalidator.addValidation("DEADSTOCK_CAPACITY","numeric");
	
	</script>
<%elseif Request("txttable") ="PRODUCTSTORAGEREADING" then%>
	<script language="JavaScript" type="text/javascript">
	//You should create the validator only after the definition of the HTML form
	var frmvalidator  = new Validator("eForm");
	frmvalidator.addValidation("READINGDATE","req","Please enter Reading Date");
	frmvalidator.addValidation("PSTORAGE_STORAGECODE","req","Please enter Storage Code");
	frmvalidator.addValidation("QUANTITY","numeric");
	</script>
<%elseif Request("txttable") ="PRODUCTADJUSTMENT" then%>
	<script language="JavaScript" type="text/javascript">
	//You should create the validator only after the definition of the HTML form
	var frmvalidator  = new Validator("eForm");
	frmvalidator.addValidation("ADJUSTMENTNO","req","Please enter Adjusment No");
	frmvalidator.addValidation("ADJUSTMENTDATE","req","Please enter Adjusment Date");
	frmvalidator.addValidation("PSTORAGE_STORAGECODE","req","Please enter Storage Code");
	frmvalidator.addValidation("QUANTITY","numeric");	
	</script>
<%elseif Request("txttable") ="PRODUCTSTORAGETRANSFER" then%>
	<script language="JavaScript" type="text/javascript">
	//You should create the validator only after the definition of the HTML form
	var frmvalidator  = new Validator("eForm");
	frmvalidator.addValidation("TRANSFERNO","req","Please enter Transfer No");
	frmvalidator.addValidation("TRANSFERDATE","req","Please enter Transfer Date");
	frmvalidator.addValidation("PSTORAGE_STORAGECODE","req","Please enter Storage Code");	
	frmvalidator.addValidation("PSTORAGE_STORAGECODE_TO_TANK","req","Please enter Storage Code To Tank");	
	frmvalidator.addValidation("QUANTITY","numeric");
	</script>
<%elseif Request("txttable") ="DELIVERYCONTRACT" then%>
	<script language="JavaScript" type="text/javascript">
	//You should create the validator only after the definition of the HTML form
	var frmvalidator  = new Validator("eForm");
	frmvalidator.addValidation("CONTRACTCODE","req","Please enter Contract Code");
	frmvalidator.addValidation("STARTDATE","req","Please enter Start Date");
	frmvalidator.addValidation("ENDDATE","req","Please enter End Date");	
	frmvalidator.addValidation("ACTUALSTART","req","Please enter Actual Start");	
	frmvalidator.addValidation("ACTUALEND","req","Please enter Actual End");	
	frmvalidator.addValidation("TOTALWEIGHT","req","Please enter Total Weight");	
	frmvalidator.addValidation("txtSTATUS","req","Please enter Status");	
	frmvalidator.addValidation("CUST_CUSTOMERCODE","req","Please enter Customer Code");	
	frmvalidator.addValidation("PRODUCT_PRODUCTCODE","req","Please enter Product Code");
	frmvalidator.addValidation("TOTALWEIGHT","numeric");	
	</script>
<%elseif Request("txttable") ="DeliveryQuality" then%>
	<script language="JavaScript" type="text/javascript">
	//You should create the validator only after the definition of the HTML form
	var frmvalidator  = new Validator("eForm");
	frmvalidator.addValidation("DELIVERYNOTE","req","Please enter Delivery Note");
	frmvalidator.addValidation("QUALITYDATE","req","Please enter Quality Date");
	frmvalidator.addValidation("FFA","numeric");
	frmvalidator.addValidation("MOISTURE","numeric");		
	frmvalidator.addValidation("DIRT","numeric");
	frmvalidator.addValidation("DIRT_AND_MOISTURE","numeric");
	</script>
<%elseif Request("txttable") ="DELIVERYNOTE" then%>
	<script language="JavaScript" type="text/javascript">
	//You should create the validator only after the definition of the HTML form
	var frmvalidator  = new Validator("eForm");
	frmvalidator.addValidation("DELIVERYNOTECODE","req","Please enter Code");
	frmvalidator.addValidation("CONTRACTCODE","req","Please enter Contract Code");
	frmvalidator.addValidation("DNOTE_DATE","req","Please enter Date");
	frmvalidator.addValidation("TRANSPORTCO","req","Please enter Transportco");
	frmvalidator.addValidation("DRIVERNAME","req","Please enter Driver Name");
	frmvalidator.addValidation("WBT_CHITNO","req","Please enter Chitno");
	//frmvalidator.addValidation("STORAGECODE","req","Please enter Storage Code");
	frmvalidator.addValidation("QUANTITY","numeric");	
	</script>

<%elseif Request("txttable") ="DELIVERYRETURNNOTE" then%>
	<script language="JavaScript" type="text/javascript">
	//You should create the validator only after the definition of the HTML form
	var frmvalidator  = new Validator("eForm");
	frmvalidator.addValidation("RETURN_CODE","req","Please enter Code");
	frmvalidator.addValidation("RETURNDATE","req","Please enter Date");
	frmvalidator.addValidation("DNOTE_DELIVERYNOTECODE","req","Please enter Delivery Note Code");
	frmvalidator.addValidation("WBT_CHITNO","req","Please enter Chitno");
	//frmvalidator.addValidation("STORAGE_CODE","req","Please enter Storage Code");
	frmvalidator.addValidation("QUANTITY","numeric");	
	</script>

<%END IF%>