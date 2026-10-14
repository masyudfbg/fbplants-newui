
function submitData(){
 	if(isValid()==false){
		return ;
    }	
	else{
		//alert("simpan");
		document.getElementsByName("Add")[0].type = "Submit";  //Change Type=button to typeSubmit;
	}
}
		
function isValid(){
	fccode = document.getElementById('FCCODE').value;
	if(fccode =='')	{
		 alert("Field Code Harus di isi");
		 document.getElementById('FCCODE').focus();
		 return(false);
	}
	
	fcname = document.getElementById('FCNAME').value;
	if(fcname =='')	{
		 alert("Field Description Harus di isi");
		 document.getElementById('FCNAME').focus();
		 return(false);
	}
	
	division = document.getElementById('DIVISION').value;
	if(division =='')	{
		 alert("Field Division Harus di isi");
		 document.getElementById('DIVISION').focus();
		 return(false);
	}
	
	hectarageplanted = document.getElementById('HECTARAGEPLANTED').value;
	if(hectarageplanted =='')	{
		 alert("Field Hectarage Harus di isi");
		 document.getElementById('HECTARAGEPLANTED').focus();
		 return(false);
	}
	
	status = document.getElementById('STATUS').value;
	if(status =='')	{
		 alert("Field Status Harus di isi");
		 document.getElementById('STATUS').focus();
		 return(false);
	}	
	
	crop = document.getElementById('CROPTYPE').value;
	if(crop =='')	{
		 alert("Field Crop Harus di isi");
		 document.getElementById('CROPTYPE').focus();
		 return(false);
	}	
	
	soil = document.getElementById('SOILTYPE').value;
	if(soil =='')	{
		 alert("Field Soil Harus di isi");
		 document.getElementById('SOILTYPE').focus();
		 return(false);
	}	
	
	terrain = document.getElementById('TERRAINTYPE').value;
	if(terrain =='')	{
		 alert("Field Terrain Harus di isi");
		 document.getElementById('TERRAINTYPE').focus();
		 return(false);
	}
	
	plantingmaterial = document.getElementById('PLANTINGMATERIAL').value;
	if(plantingmaterial =='')	{
		 alert("Field Planting Material Harus di isi");
		 document.getElementById('PLANTINGMATERIAL').focus();
		 return(false);
	}
	
	totalstand = document.getElementById('TOTALSTAND').value;
	if(totalstand =='')	{
		 alert("Field Total Stand Harus di isi");
		 document.getElementById('TOTALSTAND').focus();
		 return(false);
	}
	
	ownership = document.getElementById('OWNERSHIP').value;
	if(ownership =='')	{
		 alert("Field Ownership Harus di isi");
		 document.getElementById('OWNERSHIP').focus();
		 return(false);
	}
	
	ts = document.getElementById('TOTALSTAND').value;
	if(ts < 0)	{
		 alert("Field TOTALSTAND Tidak Boleh Lebih Kecil Dari Nol ( 0 )");
		 document.getElementById('TOTALSTAND').focus();
		 return(false);
	}
	
	
	if (Hitung_Ha_Rounding()==false){
		return ;
	}
	
	/*
	if (Hitung_Stand_Per_Pokok()==false){
		return ;
	}
	*/
	
  return(true);
 }

function Hitung_Ha_Rounding(){
		var hap = document.getElementById('HECTARAGEPLANTED').value;
		var a1 = document.getElementById('AREAL_LC').value;
		var a2 = document.getElementById('CADANGAN').value;
		var a3 = document.getElementById('OKUPASI').value;
		var a4 = document.getElementById('BIBITAN').value;
		var a5 = document.getElementById('PABRIK').value;
		var a6 = document.getElementById('EMPLASMEN').value;
		var a7 = document.getElementById('PARIT').value;
		var a8 = document.getElementById('JALAN').value;
		var a9 = document.getElementById('PRASARANA_LAINNYA').value;
		var a10 = document.getElementById('QUARRY').value;
		var a11 = document.getElementById('TANGGUL').value;
		var a12 = document.getElementById('ENCLAVE').value;
		var a13 = document.getElementById('TAMBANG').value;
		var a14 = document.getElementById('KONSERVASI').value;
		var a15 = document.getElementById('BUKIT_SUNGAI').value;
		var a16 = document.getElementById('NO_IMAGERY').value;
		var Lg = document.getElementById('LUAS_GROSS').value;
		
		var jml = parseFloat(a1)+parseFloat(a2)+parseFloat(a3)+parseFloat(a4)+parseFloat(a5)+parseFloat(a6)+parseFloat(a7)+parseFloat(a8)+parseFloat(a9)+parseFloat(a10)+parseFloat(a11)+parseFloat(a12)+parseFloat(a13)+parseFloat(a14)+parseFloat(a15)+parseFloat(a16);
		var LandUse = parseFloat(hap)+parseFloat(jml);
		var hasil = parseFloat(Lg).toFixed(2)-parseFloat(LandUse).toFixed(2);	
		var Selisih = hasil; //parseFloat(hasil).toFixed(8);
		//alert(hasil);
		
		if (hasil != 0 || hasil < 0)
		{ 
			alert("Total Hectare Planted ("+hap+") dan Hectare Land Use ("+jml+") Tidak sama dengan Total Luas Gross ("+Lg+"), Selisih Nilai ("+Selisih+")");
			return false();
		}
		/*
		else{
			alert(hasil);
		}
		*/
		
	return(true);
 }
 
 
 /* StandPKK = Total Stand - Ha Planted*/
 /*
 function Hitung_Stand_Per_Pokok() { 	
	
	var ts = document.getElementById('TOTALSTAND').value;
	var hap = document.getElementById('HECTARAGEPLANTED').value;
	
	var spkk = parseInt(ts)-parseInt(hap);
	
	if (spkk <=0)
		{ 
			alert("Nilai Total Stand Tidak Boleh Lebih Kecil dari Nilai Hectarage");
			return false();
		}
 }
 */
 
 
function harusangka(evt) {
    var e = evt || window.event;
    var key = e.keyCode || e.which;

    if (!e.shiftKey && !e.altKey && !e.ctrlKey &&
    // numbers   
    key >= 48 && key <= 57 ||
	e.ctrlKey && key == 97 ||
    // Backspace and Tab and Enter
    key == 8 || key == 9 || key == 13 ||
    // left and right arrows
    key == 37 || key == 39 ||
    // Del and Ins
	key == 46 ) {
        // input is VALID
		}
	else {
			// input is INVALID
			e.returnValue = false;
			if (e.preventDefault) e.preventDefault();
		}
}

function isNumberKeyInt(evt){
	var evtobj=window.event? event : evt.keyCode
	var charCode = (evt.which) ? evt.which : evt.keyCode

	return !(evtobj!=46 && charCode != 118 &&  charCode > 31 && (charCode < 48 || charCode > 57)); 
}

