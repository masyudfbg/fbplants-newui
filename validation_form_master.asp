<%if Request("txttable") ="Employee" then%>
<script language="JavaScript" type="text/javascript">
	//You should create the validator only after the definition of the HTML form
	var frmvalidator  = new Validator("frmEntry");
	
	frmvalidator.addValidation("FCCODE","req","Please Enter Employee Code");
	frmvalidator.addValidation("FCNAME","req","Please Enter Employee Name");
	frmvalidator.addValidation("RELIGION","req","Please Enter Employee Religion");
	frmvalidator.addValidation("RACE","req","Please Enter Employee Race");
	frmvalidator.addValidation("NATIONALITY","req","Please Enter Employee Nationality");
	frmvalidator.addValidation("txtaddress","req","Alamat tidak boleh kosong..!!");
	frmvalidator.addValidation("HIGHESTEDUCATION","req","HIGHESTEDUCATION tidak boleh kosong..!!");
	frmvalidator.addValidation("MARITALSTATUS","req","Please Enter Employee Marital Status");
	frmvalidator.addValidation("IDENTITYCARD","req","Identity Card /KTP tidak boleh kosong..!!");
	frmvalidator.addValidation("SECTIONNAME","req","Please Enter Employee Section Name");
	frmvalidator.addValidation("GANGCODE","req","Please Enter Employee Gang Code");
	frmvalidator.addValidation("PRESENTGRADE","req","Please Enter Employee Present Grade");
	frmvalidator.addValidation("PRESENTRATE","req","Please Enter Employee Present Rate");
	frmvalidator.addValidation("POSITION","req","Please Enter Employee Position");
	frmvalidator.addValidation("NATURETYPE","req","Please Enter Employee Nature Type");
	frmvalidator.addValidation("subgrade","req","Please Enter Class Employee");
	frmvalidator.addValidation("JAMSOSTEKCATEGORY","req","Please Enter Employee Jamsostek Category");
	frmvalidator.addValidation("txtbpjs_category","req","Please Enter Employee BPJS Kesehatan Category");
	frmvalidator.addValidation("INCLUDEDTAXCALCULATION","req","Please Enter Employee Included Tax Calculation");
	frmvalidator.addValidation("TAXSTATUS","req","Please Enter Employee Tax Status");
	frmvalidator.addValidation("TAXPAIDBYCOMPANY","req","Please Enter Employee Tax Paid By Company");
</script>
<%end if%>