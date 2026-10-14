<%Option Explicit%>
<%response.buffer=True%>
<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="DailyValidation.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<html>
<head>
<title>ePlaS - integrated Plantation System</title>
<script type="text/javascript" language="JavaScript1.2" src="stlib.js" ></script>
<link rel="StyleSheet" href="dtree.css" type="text/css" />
<link rel="StyleSheet" href="style-content.css" type="text/css" />
<link rel="stylesheet" href="css/jquery-ui.css" />
<!-- HADIR/ font ICON -->
<link rel="stylesheet" href="css/font/font-awesome.css" />
<!-- END:HADIR -->
<script type="text/javascript" src="dtree.js"></script>
<script src="/js/jquery-1.9.1.js"></script>
<script src="/js/jquery-ui.js"></script>
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
<!--<script language="JavaScript" src="/js/date-picker.js"></script> -->
<script language="javascript" type="text/javascript">

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
					
				function combo_(){
					var v = document.getElementById("Vehicle");
					var c = document.getElementById("CONTRACTNUMBER");
					v.value = "";
					c.value = "";
				}
				
                function showHint(str)
					{
				var checker;
					if (str=="")
					  {
					  document.getElementById("checker").innerHTML="*";
					  $("#Vehicle").removeClass("error_val");
					  return;
					  }
					if (window.XMLHttpRequest)
					  {// code for IE7+, Firefox, Chrome, Opera, Safari
					  checker=new XMLHttpRequest();
					  }
					else
					  {// code for IE6, IE5
					  checker=new ActiveXObject("Microsoft.XMLHTTP");
					  }
						checker.onreadystatechange=function()
						  {
						  if (checker.readyState==4 && checker.status==200)
							{
								var own=document.getElementById("OWN").value; 
								var vehicle = document.getElementById("Vehicle");
								var chacker = document.getElementById("checker");
								if (own == "01" ){
									if (checker.responseText == "TRUE"){
										chacker.innerHTML="<img src='IMAGES/check.GIF'>";
										$("#Vehicle").removeClass("error_val");
									}else{
										chacker.innerHTML="<img src='IMAGES/error.GIF'>";
										vehicle.value="";
										$("#Vehicle").addClass("error_val");
										alert('" ERROR ! " Kode Vehicle ' + str + ', tidak ada di Vehicle Master');
										$("#Vehicle").focus();
									}
								}
							}
						  }
						checker.open("GET","ajax/Vehicle_ann.asp?ajaxUoM=Check_Vehicle&FcCode="+str,true);
						checker.send();
					}

</script>

<script language="JavaScript">
var newwin;
function launchwin(winurl,winname,winfeatures)
{
	//This launches a new window and then
	//focuses it if window.focus() is supported.
	newwin = window.open(winurl,winname,winfeatures);

}

function validasi(form){
 
<!-- Start additional Validation by AST -->
if (form.STACKING_HA.value == ""){
	alert("Anda belum mengisikan STACKING (Ha).");
	form.STACKING_HA.focus();
	return (false);
}
  
if (form.TANAMBARU_PKK.value == ""){
	alert("Anda belum mengisikan TANAM BARU (Pkk).");
	form.TANAMBARU_PKK.focus();
	return (false);
}
  
if (form.TANAMBARU_HA.value == ""){
	alert("Anda belum mengisikan TANAM BARU (Ha).");
	form.TANAMBARU_HA.focus();
	return (false);
} 

if (form.KOMPLETING_PKK.value == ""){
	alert("Anda belum mengisikan KOMPLETING (Pkk).");
	form.KOMPLETING_PKK.focus();
	return (false);
} 

if (form.KOMPLETING_HA.value == ""){
	alert("Anda belum mengisikan KOMPLETING (Ha).");
	form.KOMPLETING_HA.focus();
	return (false);
} 

if (form.TANAMSISIP_TBM_PKK.value == ""){
	alert("Anda belum mengisikan TANAM SISIP TBM (Pkk).");
	form.TANAMSISIP_TBM_PKK.focus();
	return (false);
} 

if (form.TANAMSISIP_TM_PKK.value == ""){
	alert("Anda belum mengisikan TANAM SISIP TM (Pkk).");
	form.TANAMSISIP_TM_PKK.focus();
	return (false);
}

if (form.KASTRASI_HA.value == ""){
	alert("Anda belum mengisikan KASTRASI (Ha).");
	form.KASTRASI_HA.focus();
	return (false);
}

if (form.SANITASI_HA.value == ""){
	alert("Anda belum mengisikan SANITASI (Ha).");
	form.SANITASI_HA.focus();
	return (false);
}
	
if (form.TK_PEMELIHARAAN_TERSEDIA.value == ""){
	alert("Anda belum mengisikan TK PEMELIHARAAN TERSEDIA.");
	form.TK_PEMELIHARAAN_TERSEDIA.focus();
	return (false);
}

if (form.TK_PEMELIHARAAN_TANAMBARU.value == ""){
	alert("Anda belum mengisikan TK PEMELIHARAAN TANAM BARU.");
	form.TK_PEMELIHARAAN_TANAMBARU.focus();
	return (false);
}

if (form.TK_PEMELIHARAAN_SISIP.value == ""){
	alert("Anda belum mengisikan TK PEMELIHARAAN SISIP.");
	form.TK_PEMELIHARAAN_SISIP.focus();
	return (false);
}

if (form.TK_PEMELIHARAAN_KOMPLETING.value == ""){
	alert("Anda belum mengisikan TK PEMELIHARAAN KOMPLETING.");
	form.TK_PEMELIHARAAN_KOMPLETING.focus();
	return (false);
}

if (form.TK_PEMELIHARAAN_PEMUPUKAN.value == ""){
	alert("Anda belum mengisikan TK PEMELIHARAAN PEMUPUKAN.");
	form.TK_PEMELIHARAAN_PEMUPUKAN.focus();
	return (false);
}

if (form.TK_PEMELIHARAAN_LAINNYA.value == ""){
	alert("Anda belum mengisikan TK PEMELIHARAAN LAINNYA.");
	form.TK_PEMELIHARAAN_LAINNYA.focus();
	return (false);
}

////
if (form.PEMELIHARAAN_JALAN_TIMBUN_MR_MTR.value == ""){
	alert("Anda belum mengisikan PEMELIHARAAN JALAN TIMBUN MR (MTR).");
	form.PEMELIHARAAN_JALAN_TIMBUN_MR_MTR.focus();
	return (false);
}

if (form.PEMELIHARAAN_JALAN_TIMBUN_CR_MTR.value == ""){
	alert("Anda belum mengisikan PEMELIHARAAN JALAN TIMBUN CR (MTR).");
	form.PEMELIHARAAN_JALAN_TIMBUN_CR_MTR.focus();
	return (false);
}

if (form.PEMELIHARAAN_JALAN_PENGERASAN_MR_MTR.value == ""){
	alert("Anda belum mengisikan PEMELIHARAAN JALAN PENGERASAN MR (MTR).");
	form.PEMELIHARAAN_JALAN_PENGERASAN_MR_MTR.focus();
	return (false);
}

if (form.PEMELIHARAAN_JALAN_PENGERASAN_CR_MTR.value == ""){
	alert("Anda belum mengisikan PEMELIHARAAN JALAN PENGERASAN CR (MTR).");
	form.PEMELIHARAAN_JALAN_PENGERASAN_CR_MTR.focus();
	return (false);
}
<!-- End additional Validation by AST -->
  
  
    
 }
</script>

</head>
<body>
<!--#include file="top.jsp"-->
	<div style=" width:21.7%; float:left;margin-left:0.3%; margin-top:0.3%; margin-bottom:0.3%;">
			<!--#include file="menu.asp"-->
		</div>
   <div class="judul_head" valign="top" bgcolor="#fff" style=" width:77.4%; margin-left:0.3%; margin-right:0.3%; margin-top:0.3%; margin-bottom:0.3%; float:left; border:none;"> 
 <div class="judul">
	   Module : <%=request("txtModule")%> / <%=request("txtTable")%>
	</div>
 <div class="box_both" style="overflow:scroll;">
 <!-------------awal!------------->
 <%
 
 Dim vSTACKING_HA, vTANAMBARU_PKK, vTANAMBARU_HA, vKOMPLETING_PKK, vKOMPLETING_HA, vTANAMSISIP_TBM_PKK
 Dim vTANAMSISIP_TM_PKK, vKASTRASI_HA, vSANITASI_HA
 Dim vTK_PEMEL_TERSEDIA, vTK_PEMEL_TANAMBARU, vTK_PEMEL_SISIP, vTK_PEMEL_KOMPLETING, vTK_PEMEL_PEMUPUKAN, vTK_PEMEL_LAINNYA
 Dim vPEMEL_JALAN_TIMBUN_MR_MTR, vPEMEL_JALAN_TIMBUN_CR_MTR, vPEMEL_JALAN_PENGERASAN_MR_MTR, vPEMEL_JALAN_PENGERASAN_CR_MTR
 Dim vID
 Dim FCENTRY, FCEDIT, FCIP, FCBA
 Dim cmdView
 Dim oSQL, oRS
 ''Dim vTRANS_DATE ''vTRANS_DATE_Update,
 
  '' Get Period_Year, Period_Month And Date
 ''=====================================
 Dim Yr, txtYear
 Dim xYr, MinYr, MaxYr
 Dim xMn, txtMonth, Mn
 Dim MaxDate, dd, txtday, tdate, xdd
 
 MaxDate = Request("MaxDate")
 txtday = Request("txtday")
 txtYear = request("txtYear")
 txtMonth = request("txtMonth")
 
 ' dd = Day(Session("v_Costbook"))
 ' If "" & vTRANS_DATE = "" Then vTRANS_DATE = 1
 
 Yr = Year(Session("v_Costbook"))
 if "" & txtYear = "" then txtYear = Yr
 
  Mn = Month(Session("v_Costbook"))
 if "" & txtMonth = "" then txtMonth = Mn
 
''======================================

 Dim Cek_IsDailyClosing : Cek_IsDailyClosing = ""
  
 vID=request("ID") '' PK
 tdate = Request("TRANS_DATE")
 
 vSTACKING_HA =request("STACKING_HA") 
 vTANAMBARU_PKK =request("TANAMBARU_PKK")
 vTANAMBARU_HA =request("TANAMBARU_HA")
 vKOMPLETING_PKK =request("KOMPLETING_PKK")
 vKOMPLETING_HA =request("KOMPLETING_HA")
 
 vTANAMSISIP_TBM_PKK =request("TANAMSISIP_TBM_PKK")
 vTANAMSISIP_TM_PKK =request("TANAMSISIP_TM_PKK")
 vKASTRASI_HA =request("KASTRASI_HA") 
 vSANITASI_HA =request("SANITASI_HA")
 
 vTK_PEMEL_TERSEDIA =request("TK_PEMELIHARAAN_TERSEDIA")
 vTK_PEMEL_TANAMBARU =request("TK_PEMELIHARAAN_TANAMBARU")
 vTK_PEMEL_SISIP =request("TK_PEMELIHARAAN_SISIP")
 vTK_PEMEL_KOMPLETING =request("TK_PEMELIHARAAN_KOMPLETING")
 vTK_PEMEL_PEMUPUKAN =request("TK_PEMELIHARAAN_PEMUPUKAN")
 vTK_PEMEL_LAINNYA =request("TK_PEMELIHARAAN_LAINNYA")
 
 vPEMEL_JALAN_TIMBUN_MR_MTR =request("PEMELIHARAAN_JALAN_TIMBUN_MR_MTR")
 vPEMEL_JALAN_TIMBUN_CR_MTR =request("PEMELIHARAAN_JALAN_TIMBUN_CR_MTR")
 vPEMEL_JALAN_PENGERASAN_MR_MTR =request("PEMELIHARAAN_JALAN_PENGERASAN_MR_MTR")
 vPEMEL_JALAN_PENGERASAN_CR_MTR =request("PEMELIHARAAN_JALAN_PENGERASAN_CR_MTR")
	
If "" & txtday = "" Then txtday = 1
tdate = txtMonth & "/" & txtday & "/" & txtYear
	
	
Sub InsertData
	q	=	" INSERT INTO LHK_NON_PROD (" & _
			"   STACKING_HA" & _
			" , TANAMBARU_PKK" & _
			" , TANAMBARU_HA" & _
			" , KOMPLETING_PKK" & _
			" , KOMPLETING_HA" & _
			" , TANAMSISIP_TBM_PKK" & _
			" , TANAMSISIP_TM_PKK" & _
			" , KASTRASI_HA" & _
			" , SANITASI_HA" & _
			" , TK_PEMELIHARAAN_TERSEDIA" & _
			" , TK_PEMELIHARAAN_TANAMBARU" & _
			" , TK_PEMELIHARAAN_SISIP" & _
			" , TK_PEMELIHARAAN_KOMPLETING" & _
			" , TK_PEMELIHARAAN_PEMUPUKAN" & _
			" , TK_PEMELIHARAAN_LAINNYA" & _
			" , PEMEL_JALAN_TIMBUN_MR_MTR" & _
			" , PEMEL_JALAN_TIMBUN_CR_MTR" & _
			" , PEMEL_JALAN_PENGERASAN_MR_MTR" & _			
			" , PEMEL_JALAN_PENGERASAN_CR_MTR" & _
			" , FCENTRY" & _
			" , FCEDIT" & _
			" , FCIP" & _
			" , FCBA" & _
			" , LASTUPDATE" & _
			" , LASTTIME" & _
			" , TRANS_DATE" & _			
			" ) values (" & _
			"   '" & vSTACKING_HA & "'" & _
			" , '" & vTANAMBARU_PKK & "'" & _
			" , '" & vTANAMBARU_HA & "'" & _
			" , '" & vKOMPLETING_PKK & "'" & _
			" , '" & vKOMPLETING_HA & "'" & _
			" , '" & vTANAMSISIP_TBM_PKK & "'" & _
			" , '" & vTANAMSISIP_TM_PKK & "'" & _
			" , '" & vKASTRASI_HA & "'" & _			
			" , '" & vSANITASI_HA & "'" & _
			" , '" & vTK_PEMEL_TERSEDIA & "'" & _
			" , '" & vTK_PEMEL_TANAMBARU & "'" & _
			" , '" & vTK_PEMEL_SISIP & "'" & _			
			" , '" & vTK_PEMEL_KOMPLETING & "'" & _			
			" , '" & vTK_PEMEL_PEMUPUKAN & "'" & _			
			" , '" & vTK_PEMEL_LAINNYA & "'" & _		
			" , '" & vPEMEL_JALAN_TIMBUN_MR_MTR & "'" & _
			" , '" & vPEMEL_JALAN_TIMBUN_CR_MTR & "'" & _
			" , '" & vPEMEL_JALAN_PENGERASAN_MR_MTR & "'" & _			
			" , '" & vPEMEL_JALAN_PENGERASAN_CR_MTR & "'" & _
			" , '" & Session("CodeName") & "'" & _
			" , '" & Session("CodeName") & "'" & _
			" , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
			" , '" & Session("fcBA") & "'" & _
			" , sysdate " & _
			" , To_Char(sysdate, 'HH24:MI')" & _	
			" , TO_DATE('" & tdate & "', 'mm/dd/yyyy')" & _			
			" )"
			
	''Response.Write q
	connStr.execute(q)
End Sub

Sub UpdateData
	q	=	" UPDATE LHK_NON_PROD SET" & _
			"   STACKING_HA = '" & vSTACKING_HA & "'" & _
			" , TANAMBARU_PKK = '" & vTANAMBARU_PKK & "'" & _
			" , TANAMBARU_HA = '" & vTANAMBARU_HA & "'" & _
			" , KOMPLETING_PKK = '" & vKOMPLETING_PKK & "'" & _
			" , KOMPLETING_HA = '" & vKOMPLETING_HA & "'" & _
			" , TANAMSISIP_TBM_PKK = '" & vTANAMSISIP_TBM_PKK & "'" & _
			" , TANAMSISIP_TM_PKK = '" & vTANAMSISIP_TM_PKK & "'" & _
			" , KASTRASI_HA = '" & vKASTRASI_HA & "'" & _
			" , SANITASI_HA = '" & vSANITASI_HA & "'" & _
			" , TK_PEMELIHARAAN_TERSEDIA = '" & vTK_PEMEL_TERSEDIA & "'" & _
			" , TK_PEMELIHARAAN_TANAMBARU = '" & vTK_PEMEL_TANAMBARU & "'" & _
			" , TK_PEMELIHARAAN_SISIP = '" & vTK_PEMEL_SISIP & "'" & _
			" , TK_PEMELIHARAAN_KOMPLETING = '" & vTK_PEMEL_KOMPLETING & "'" & _
			" , TK_PEMELIHARAAN_PEMUPUKAN = '" & vTK_PEMEL_PEMUPUKAN & "'" & _
			" , TK_PEMELIHARAAN_LAINNYA = '" & vTK_PEMEL_LAINNYA & "'" & _
			" , PEMEL_JALAN_TIMBUN_MR_MTR = '" & vPEMEL_JALAN_TIMBUN_MR_MTR & "'" & _
			" , PEMEL_JALAN_TIMBUN_CR_MTR = '" & vPEMEL_JALAN_TIMBUN_CR_MTR & "'" & _
			" , PEMEL_JALAN_PENGERASAN_MR_MTR = '" & vPEMEL_JALAN_PENGERASAN_MR_MTR & "'" & _
			" , PEMEL_JALAN_PENGERASAN_CR_MTR = '" & vPEMEL_JALAN_PENGERASAN_CR_MTR & "'" & _
			" , LASTUPDATE = sysdate " & _
			" , LASTTIME = To_Char(sysdate,'HH24:MI') " & _				
			" WHERE fcBA = '" & Session("fcBA") & "'" & _
			" AND ID = '" & vID & "'" 
			
	''Response.Write q
	connStr.execute(q)
End Sub

Sub DeleteData
	q	=	" DELETE LHK_NON_PROD " & _
			" WHERE fcBA = '" & Session("fcBA") & "'" & _
			" AND ID = '" & vID & "'" 

	'Response.Write q
	connStr.execute(q)
End Sub
 

	if Request("txtJenis")="New" or Request("txtJenis")= "Delete" or Request("txtJenis")="Edit" Then
		if Request("txtJenis")= "Delete" then
		
			Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), tdate)
			If Cek_IsDailyClosing = "FALSE" then
				Call DeleteData
			else
				err.number = 1001
				err.description = "Tanggal """ & tdate & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
			End If
			cmdView	= "True"
		
		elseif Request("txtJenis")="New" then
		
			' Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), tdate)
			' If Cek_IsDailyClosing = "FALSE" then
				 Call InsertData
			' else
				' err.number = 1001
				' err.description = "Tanggal """ & tdate & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
			' End If
			 cmdView	= "True"
			
		elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
			' vTRANS_DATE =request("TRANS_DATE") 
			' vTRANS_DATE_Update =request("TRANS_DATE") 
			 
			Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), tdate)
			If Cek_IsDailyClosing = "FALSE" then
				Call UpdateData
			else
				err.number = 1001
				err.description = "Tanggal """ & tdate & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
			End If
			cmdView	= "True"
			
		else
			cmdView	= "True"
		end if								 
		'Response.Write "sql " & q
	else
		cmdView	= request("cmdView")
	End if

Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis
txtTable =Request("txtTable")
txtModule=Request("txtModule")
txtSearch=Request("txtSearch")
txtField =Request("txtField")
txtJenis = Request("txtJenis")

	' if vTRANS_DATE <> "" then
	   ' vTRANS_DATE   =Request("TRANS_DATE")
	' else
	   ' vTRANS_DATE   =textDate(date)
	' end if



If Err.Number <> 0 then
   Response.Write "<div class = ""error"">" & Err.Description & "</div>"
End if

'' Event GetInitEdit / Mengisi Data
''============================================
if request("txtJenis")="Edit" then
 Dim rsVHx
 
q	=	" SELECT " & _
		" STACKING_HA, TANAMBARU_PKK, TANAMBARU_HA, KOMPLETING_PKK, KOMPLETING_HA, TANAMSISIP_TBM_PKK, " & _
		" TANAMSISIP_TM_PKK, KASTRASI_HA, SANITASI_HA, TK_PEMELIHARAAN_TERSEDIA, TK_PEMELIHARAAN_TANAMBARU, " & _
		" TK_PEMELIHARAAN_SISIP, TK_PEMELIHARAAN_KOMPLETING, TK_PEMELIHARAAN_PEMUPUKAN, TK_PEMELIHARAAN_LAINNYA," & _
		" PEMEL_JALAN_TIMBUN_MR_MTR, PEMEL_JALAN_TIMBUN_CR_MTR, PEMEL_JALAN_PENGERASAN_MR_MTR, PEMEL_JALAN_PENGERASAN_CR_MTR," & _
		" TRANS_DATE, ID" & _
		" FROM LHK_NON_PROD " & _
		" WHERE 1=1 AND ID ='"&vID&"'"
		
'response.Write(q)				
Set rsVHx = Server.CreateObject("ADODB.Recordset")
rsVHx.Open q, connstr

	if not rsVHx.eof then 
		While not rsVHx.eof
			txtday = day(rsVHx("TRANS_DATE"))
			txtMonth = Month(rsVHx("TRANS_DATE"))
			txtYear = Year(rsVHx("TRANS_DATE"))
			
			vSTACKING_HA =rsVHx("STACKING_HA") 
			vTANAMBARU_PKK =rsVHx("TANAMBARU_PKK")
			vTANAMBARU_HA =rsVHx("TANAMBARU_HA")
			vKOMPLETING_PKK =rsVHx("KOMPLETING_PKK")
			vKOMPLETING_HA =rsVHx("KOMPLETING_HA")

			vTANAMSISIP_TBM_PKK =rsVHx("TANAMSISIP_TBM_PKK")
			vTANAMSISIP_TM_PKK =rsVHx("TANAMSISIP_TM_PKK")
			vKASTRASI_HA =rsVHx("KASTRASI_HA") 
			vSANITASI_HA =rsVHx("SANITASI_HA")

			vTK_PEMEL_TERSEDIA =rsVHx("TK_PEMELIHARAAN_TERSEDIA")
			vTK_PEMEL_TANAMBARU =rsVHx("TK_PEMELIHARAAN_TANAMBARU")
			vTK_PEMEL_SISIP =rsVHx("TK_PEMELIHARAAN_SISIP")
			vTK_PEMEL_KOMPLETING =rsVHx("TK_PEMELIHARAAN_KOMPLETING")
			vTK_PEMEL_PEMUPUKAN =rsVHx("TK_PEMELIHARAAN_PEMUPUKAN")
			vTK_PEMEL_LAINNYA =rsVHx("TK_PEMELIHARAAN_LAINNYA")

			vPEMEL_JALAN_TIMBUN_MR_MTR =rsVHx("PEMEL_JALAN_TIMBUN_MR_MTR")
			vPEMEL_JALAN_TIMBUN_CR_MTR =rsVHx("PEMEL_JALAN_TIMBUN_CR_MTR")
			vPEMEL_JALAN_PENGERASAN_MR_MTR =rsVHx("PEMEL_JALAN_PENGERASAN_MR_MTR")
			vPEMEL_JALAN_PENGERASAN_CR_MTR =rsVHx("PEMEL_JALAN_PENGERASAN_CR_MTR")
			
		rsVHx.movenext
		Wend
		end if
	end if
	     
''============================================
'' Setting Lock Event Entry Data
''============================================
Dim val_IsTrue : val_IsTrue = "Y"  '' di Sett manual untuk kebutuhan trial di bulan februari ( harus tidak boleh )

'Dim val_IsTrue : val_IsTrue = "N"
' Dim oRecordSet, oSQLQuery
' oSQLQuery = " select 'Y' istrue" & _
            ' " from periodcontrol" & _
	        ' " where fcba = '" & Session("fcBA") & "'" & _
	        ' "       and to_number ('" & txtYear & "') = CURRENTACCOUNTINGYEAR" & _
	        ' "       and to_number ('" & txtMonth & "') >= CURRENTACCOUNTINGPERIOD" & _
	        ' "       and upper(system) = upper('Costbook')"
			
' set oRecordSet = server.CreateObject("Adodb.recordset")
' oRecordSet.Open oSQLQuery, connstr

' If Not oRecordSet.Eof then
	' val_IsTrue = oRecordSet("istrue")
' Else
	' val_IsTrue = "N"
' End if

' ''response.Write oSQLQuery

' oRecordSet.Close
' Set oRecordSet = Nothing

%>
 <!-------------Form Entry Tanggal------------->

  <form name="frmEntry" method="post" action="Lhk_NonProd.asp">
	<table class="data" width="100%" cellpadding="0" cellspacing="0">
		<tr>
			<td colspan="2">
				<div align="left"><font size="2" face="Verdana"><strong>LHK Non Produksi</strong></font></div>
			</td>
		</tr>
			
		<tr>
			<td align="left" style="width:20%;">Year</td>
			<td align="left">
				<select name="txtYear">
				<%
					MinYr = Yr -2
					MaxYr = Yr +2
					
					for xYr=MinYr to MaxYr
					if cint(txtYear) = cint(xYr) then
						response.write "<option value ='"&xYr&"' selected>"&xYr&"</option>"
					else
						response.write "<option value ='"&xYr&"'>"&xYr&"</option>"
					end if
					next
				%>
				</select>
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
		
		<tr> 		
			<td align="left" style="width:20%;">Month</td>
			<td align="left">
				<select name="txtMonth">
				<%
					for xMn=1 to 12
					if cint(txtMonth) = cint(xMn) then
						response.write "<option value ='"&xMn&"' selected>"&Bulan(xMn)&"</option>"
					else
						response.write "<option value ='"&xMn&"'>"&Bulan(xMn)&"</option>"
					end if
					next
				%>
				</select>
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>	
			
		<tr>
			<td colspan="2"  align="left">
				<input Type="Hidden" Value="True" Name="cmdView">
				<%
					MaxDate = GetLastDay(txtMonth,txtYear)						
				%>
				<input Type="Hidden" Value="MaxDate" Name="<%=MaxDate%>">
				<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
				<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
				<input Type="Submit" name="Add" value="Select">&nbsp;&nbsp;<a href="PrintLhk_NonProd.Asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&PERIOD_YEAR=<%=txtYear%>" title="Print Data" target="_blank"><img src="IMAGES/Print.gif" border="0"></a>

				<%
				dim dataModule,dataSubmodule,dataDetail
				dataModule 		= "LHK Non Produksi"
				dataSubmodule 	= "Lhk_NonProd"
				dataDetail 		= "NULL"
				%>
				<a class="jqplot-image-button" style="" href="ajax/hadirAjax.asp?dataModule=<%=dataModule%>&dataSubmodule=<%=dataSubmodule%>&dataDetail=<%=dataDetail%>"><i class="fa fa-check-square-o"></i>&nbsp;Hadir</a>

			</td>
		</tr>
		
    </table>  
	
 </form>


<%if cmdView = "True" And len(trim(txtYear)) > 0 Then
'' Query Menampilkan Data di Table Bawah
 Dim rsVH,intRecordsPerPage,cColor
 
q	=	" SELECT " & _
		" STACKING_HA, TANAMBARU_PKK, TANAMBARU_HA, KOMPLETING_PKK, KOMPLETING_HA, TANAMSISIP_TBM_PKK, " & _
		" TANAMSISIP_TM_PKK, KASTRASI_HA, SANITASI_HA, TK_PEMELIHARAAN_TERSEDIA, TK_PEMELIHARAAN_TANAMBARU, " & _
		" TK_PEMELIHARAAN_SISIP, TK_PEMELIHARAAN_KOMPLETING, TK_PEMELIHARAAN_PEMUPUKAN, TK_PEMELIHARAAN_LAINNYA," & _
		" PEMEL_JALAN_TIMBUN_MR_MTR, PEMEL_JALAN_TIMBUN_CR_MTR, PEMEL_JALAN_PENGERASAN_MR_MTR, PEMEL_JALAN_PENGERASAN_CR_MTR," & _
		" TRANS_DATE, ID" & _
		" FROM LHK_NON_PROD " & _
		" WHERE fcba = '" & Session("fcBA") & "'" & _
	    " AND TO_CHAR(TRANS_DATE, 'yyyymm') = '" & txtYear & "'||LPAD('" & txtMonth & "', 2, 0) " & _
		" ORDER BY TRANS_DATE " 
		
	''response.Write(q)				
	Set rsVH = Server.CreateObject("ADODB.Recordset")
	rsVH.Open q, connstr   
%>
	<!-------------Form Entry ------------->
    <form name="eForm" method="post" action="Lhk_NonProd.Asp" onSubmit="javascript:return validasi(this)">
     <table class="data" cellpadding="0" cellspacing="0" width="100%"> 
		
			<tr>
				<td align="left">Tanggal</td>
				<td align="left">
					<select name="txtday" <%if txtJenis = "Edit" then %> disabled <%end if%>>
						<%
							For dd = 1 to MaxDate
						%>
								<Option Value="<%=dd%>" <%If cint(dd) = cint(txtday) then%> Selected <%End If%>><%=dd%></option>
						<%
							Next
						%>
					</select>
				</td>
			</tr>
			
					
		<!--Garis Hijau /  Blok Greeen -->
		<tr bgcolor="#00a300"> 
			<td colspan="8">
				<div align="left">
					<font size="2" face="Verdana" color="#fff"><strong>Fisik</strong></font>
				</div>
			</td>
		</tr>
		
		<tr> 
			<td width="180" align="left"><input type="hidden" name="ID" id="ID" value="<%=vID%>" readonly="True" style="background:#999999"></td>
		</tr>
		<Tr> 		
            <td width="189" align="left"><font face="Verdana" size="1">Stacking</font></td>
            <td width="180" align="left"><input type="text" size="10" maxlength="30"  name="STACKING_HA" id="STACKING_HA" onkeypress="return harusangka(event);" value="<%=vSTACKING_HA%>" title="Stacking Ha 1 s/d 9"><font color="#FF0000">&nbsp Ha</font> </td>
            
			<td width="149" align="left"><font face="Verdana" size="1">Tanam Baru</font></td>
            <td width="180" align="left"><input type="text" size="10" maxlength="70" name="TANAMBARU_PKK"  id="TANAMBARU_PKK" onkeypress="return harusangka(event);" value="<%=vTANAMBARU_PKK%>"  title="Tanam Baru PKK  1 s/d 9"><font color="#FF0000">&nbsp PKK</font></td>
			
			<td width="149" align="left"><font face="Verdana" size="1">Tanam Baru</font></td>
            <td width="180" align="left"><input type="text" size="10" maxlength="70" name="TANAMBARU_HA"  id="TANAMBARU_HA" onkeypress="return harusangka(event);" value="<%=vTANAMBARU_HA%>" title="TANAM BARU HA 1 s/d 9"><font color="#FF0000">&nbsp Ha</font></td>
        </Tr>
		
		<Tr> 
            <td width="189" align="left"><font face="Verdana" size="1">Kompleting</font></td>
            <td width="180" align="left"><input type="text" size="10" maxlength="30"  name="KOMPLETING_PKK" onkeypress="return harusangka(event);" value="<%=vKOMPLETING_PKK%>" title="KOMPLETING PKK  1 s/d 9"><font color="#FF0000">&nbsp PKK</font> </td>
            
			<td width="149" align="left"><font face="Verdana" size="1">Kompleting</font></td>
            <td width="180" align="left"><input type="text" size="10" maxlength="70" name="KOMPLETING_HA"  id="KOMPLETING_HA" onkeypress="return harusangka(event);" value="<%=vKOMPLETING_HA%>" title="KOMPLETING HA  1 s/d 9"><font color="#FF0000">&nbsp Ha</font></td>
			
			<td width="149" align="left"><font face="Verdana" size="1">Tanam Sisip TBM</font></td>
            <td width="180" align="left"><input type="text" size="10" maxlength="70" name="TANAMSISIP_TBM_PKK"  id="TANAMSISIP_TBM_PKK" onkeypress="return harusangka(event);" value="<%=vTANAMSISIP_TBM_PKK%>" title="TANAMSISIP TBM PKK  1 s/d 9"><font color="#FF0000">&nbsp PKK</font></td>
        </Tr>

		<Tr> 
            <td width="189" align="left"><font face="Verdana" size="1">Tanam Sisip TM</font></td>
            <td width="180" align="left"><input type="text" size="10" maxlength="30"  name="TANAMSISIP_TM_PKK" id="TANAMSISIP_TM_PKK" onkeypress="return harusangka(event);" value="<%=vTANAMSISIP_TM_PKK%>" title="TANAMSISIP TM PKK  1 s/d 9"><font color="#FF0000">&nbsp PKK</font> </td>
            
			<td width="149" align="left"><font face="Verdana" size="1">Kastrasi</font></td>
            <td width="180" align="left"><input type="text" size="10" maxlength="70" name="KASTRASI_HA"  id="KASTRASI_HA" onkeypress="return harusangka(event);" value="<%=vKASTRASI_HA%>" title="KASTRASI HA  1 s/d 9"><font color="#FF0000">&nbsp Ha</font></td>
			
			<td width="149" align="left"><font face="Verdana" size="1">Sanitasi</font></td>
            <td width="180" align="left"><input type="text" size="10" maxlength="70" name="SANITASI_HA"  id="SANITASI_HA" onkeypress="return harusangka(event);" value="<%=vSANITASI_HA%>" title="SANITASI HA  1 s/d 9"><font color="#FF0000">&nbsp Ha</font></td>
        </Tr>
		
		<!--Garis Hijau /  Blok Greeen -->
		<tr bgcolor="#00a300"> 
			<td colspan="8">
				<div align="left">
					<font size="2" face="Verdana" color="#fff"><strong>Tenaga Kerja - Pemeliharaan</strong></font>
				</div>
			</td>
		</tr>
		
		<Tr> 
            <td width="189" align="left"><font face="Verdana" size="1">Pemeliharaan Tersedia</font></td>
            <td width="180" align="left"><input type="text" size="10" maxlength="30"  name="TK_PEMELIHARAAN_TERSEDIA" id="TK_PEMELIHARAAN_TERSEDIA" onkeypress="return harusangka(event);" value="<%=vTK_PEMEL_TERSEDIA%>" title="TK_PEMELIHARAAN_TERSEDIA Character a s/d z or 1 s/d 9"><font color="#FF0000">&nbsp Hk</font> </td>
            
			<td width="149" align="left"><font face="Verdana" size="1">Pemeliharaan Tanam Baru</font></td>
            <td width="180" align="left"><input type="text" size="10" maxlength="70" name="TK_PEMELIHARAAN_TANAMBARU"  id="TK_PEMELIHARAAN_TANAMBARU" onkeypress="return harusangka(event);" value="<%=vTK_PEMEL_TANAMBARU%>" title="TK_PEMELIHARAAN_TANAMBARU Character a s/d z or 1 s/d 9"><font color="#FF0000">&nbsp Hk</font></td>
			
			<td width="149" align="left"><font face="Verdana" size="1">Pemeliharaan Sisip</font></td>
            <td width="180" align="left"><input type="text" size="10" maxlength="70" name="TK_PEMELIHARAAN_SISIP"  id="TK_PEMELIHARAAN_SISIP" onkeypress="return harusangka(event);" value="<%=vTK_PEMEL_SISIP%>" title="TK_PEMELIHARAAN_SISIP Character a s/d z or 1 s/d 9"><font color="#FF0000">&nbsp Hk</font></td>
        </Tr>
		
		<Tr> 
            <td width="189" align="left"><font face="Verdana" size="1">Pemeliharaan Kompleting</font></td>
            <td width="180" align="left"><input type="text" size="10" maxlength="30"  name="TK_PEMELIHARAAN_KOMPLETING" id="TK_PEMELIHARAAN_KOMPLETING" onkeypress="return harusangka(event);" value="<%=vTK_PEMEL_KOMPLETING%>" title="TK_PEMELIHARAAN_KOMPLETING Character a s/d z or 1 s/d 9"><font color="#FF0000">&nbsp Hk</font> </td>
            
			<td width="149" align="left"><font face="Verdana" size="1">Pemeliharaan Pemupukan</font></td>
            <td width="180" align="left"><input type="text" size="10" maxlength="70" name="TK_PEMELIHARAAN_PEMUPUKAN"  id="TK_PEMELIHARAAN_PEMUPUKAN" onkeypress="return harusangka(event);" value="<%=vTK_PEMEL_PEMUPUKAN%>" title="TK_PEMELIHARAAN_PEMUPUKAN Character a s/d z or 1 s/d 9"><font color="#FF0000">&nbsp Hk</font></td>
			
			<td width="149" align="left"><font face="Verdana" size="1">Pemeliharaan Lainnya</font></td>
            <td width="180" align="left"><input type="text" size="10" maxlength="70" name="TK_PEMELIHARAAN_LAINNYA"  id="TK_PEMELIHARAAN_LAINNYA" onkeypress="return harusangka(event);" value="<%=vTK_PEMEL_LAINNYA%>" title="TK_PEMELIHARAAN_LAINNYA Character a s/d z or 1 s/d 9"><font color="#FF0000">&nbsp Hk</font></td>
        </Tr>
				
		<!--Garis Hijau /  Blok Greeen -->
		<tr bgcolor="#00a300"> 
			<td colspan="8">
				<div align="left">
					<font size="2" face="Verdana" color="#fff"><strong>Pemeliharaan - Jalan</strong></font>
				</div>
			</td>
		</tr>
		
		<Tr> 
            <td width="189" align="left"><font face="Verdana" size="1">Timbun MR</font></td>
            <td width="180" align="left"><input type="text" size="10" maxlength="30"  name="PEMELIHARAAN_JALAN_TIMBUN_MR_MTR" id="PEMELIHARAAN_JALAN_TIMBUN_MR_MTR" onkeypress="return harusangka(event);" value="<%=vPEMEL_JALAN_TIMBUN_MR_MTR%>" title="PEMELIHARAAN_JALAN_TIMBUN_MR_MTR Character a s/d z or 1 s/d 9"><font color="#FF0000">&nbsp MTR</font> </td>
            
			<td width="149" align="left"><font face="Verdana" size="1">Timbun CR</font></td>
            <td width="180" align="left"><input type="text" size="10" maxlength="70" name="PEMELIHARAAN_JALAN_TIMBUN_CR_MTR"  id="PEMELIHARAAN_JALAN_TIMBUN_CR_MTR" onkeypress="return harusangka(event);" value="<%=vPEMEL_JALAN_TIMBUN_CR_MTR%>" title="PEMELIHARAAN_JALAN_TIMBUN_CR_MTR Character a s/d z or 1 s/d 9"><font color="#FF0000">&nbsp MTR</font></td>
			
			<td width="149" align="left"><font face="Verdana" size="1">Pengerasan MR</font></td>
            <td width="180" align="left"><input type="text" size="10" maxlength="70" name="PEMELIHARAAN_JALAN_PENGERASAN_MR_MTR"  id="PEMELIHARAAN_JALAN_PENGERASAN_MR_MTR" onkeypress="return harusangka(event);" value="<%=vPEMEL_JALAN_PENGERASAN_MR_MTR%>" title="PEMELIHARAAN_JALAN_PENGERASAN_MR_MTR Character a s/d z or 1 s/d 9"><font color="#FF0000">&nbsp MTR</font></td>
        
		</Tr>
		
		<Tr> 
			<td width="189" align="left"><font face="Verdana" size="1">Pengerasan CR</font></td>
            <td width="180" align="left"><input type="text" size="10" maxlength="70" name="PEMELIHARAAN_JALAN_PENGERASAN_CR_MTR"  id="PEMELIHARAAN_JALAN_PENGERASAN_CR_MTR" onkeypress="return harusangka(event);" value="<%=vPEMEL_JALAN_PENGERASAN_CR_MTR%>" title="PEMELIHARAAN_JALAN_PENGERASAN_CR_MTR Character a s/d z or 1 s/d 9"><font color="#FF0000">&nbsp MTR</font></td>
        </Tr>		

		<!-- Button Eksekusi -->
	    <Tr>
			<td colspan="2" align="left">
				<input type="hidden" name="txtYear" value="<%=txtYear%>">
			    <input type="hidden" name="txtMonth" value="<%=txtMonth%>">
			  
			    <input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
			    <input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
			    <%If val_IsTrue = "Y" Then%>
					<%IF Request("txtJenis")="Edit" then%>
						<input type="hidden" name="txtday" value="<%=txtday%>">
					    <input type="hidden" name="txtJenis" value="Edit">
					    <input Type="Submit"  name="Add" value="Update">
					<%Else%>
					    <input type="hidden" name="txtJenis" value="New">
					    <input Type="Submit"  name="Add" value="Add">
					<%End if%>
			    <%End if%>
			  <input Type="Reset"  name="Reset" value="Reset">
			  <input Type="Button"  value="Back" name="B2" OnClick="javascript:history.back(1)"></td>
		</tr>
		
     </table>
	</form>
	
  <!--#include file="validation_show.Asp"--> 
   <table class="data">
	<br>
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;" >
		<TD align="left">NO</TD>
		<TD colspan="2"> ACTION&nbsp;</TD>
		<TD align="left">Stacking (Ha)</TD>	
		<TD align="left">Tanam Baru (Pkk)</TD>	
		<TD align="left">Tanam Baru (Ha)</TD>	
		<TD align="left">Kompleting (Pkk)</TD>	
		<TD align="left">Kompleting (Ha)</TD>	
		<TD align="left">Tanam Sisip TBM (Pkk)</TD>	
		<TD align="left">Tanam Sisip TM (Pkk)</TD>
		<td align="left">Kastrasi (Ha)</td>
		<td align="left">Sanitasi (Ha)</td>
		<td align="left">Pemeliharaan Tersedia</td>
		<td align="left">Pemeliharaan Tanam Baru</td>
		<TD align="left">Pemeliharaan Sisip</TD>	
		<TD align="left">Pemeliharaan Kompleting</TD>	
		<TD align="left">Pemeliharaan Pemupukan</TD>	
		<TD align="left">Pemeliharaan Lainnya</TD>	
		<TD align="left">Timbun MR (Mtr)</TD>	
		<TD align="left">Timbun CR (Mtr)</TD>	
		<TD align="left">Pengerasan MR (Mtr)</TD>	
		<TD align="left">Pengerasan CR (Mtr)</TD>	
		<TD align="left">Tanggal Transaksi</TD>	
	</tr>
   <%
   
   if Not rsVH.eof then
      nCountofField=rsVH.fields.count-1
	  no=1
	  While Not rsVH.EOF
		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if%>
			<tr bgcolor="<%=cColor%>">
			<td align="left" valign="top"><%=NO%></td>
			<%If val_IsTrue = "Y" Then%>
				<!-- Event Edit -->
				<td valign="middle" width="30">
					<a href="Lhk_NonProd.Asp?ID=<%=rsVH("ID")%>&txtJenis=Edit"><img src="IMAGES/edit.gif" alt="Edit Data : <%=rsVH(1)%>" border="0"></a>
				</td>
				<!-- Event Delete -->			
				<td valign="middle" width="30">
					<a href="Lhk_NonProd.Asp?ID=<%=rsVH("ID")%>&txtJenis=Delete"><img src="IMAGES/delete.gif" alt="Delete Data : <%=rsVH(1)%>" border="0" onClick="return confirm('Are you sure to Delete');"></a>
				</td>
			<%Else%>
				<td colspan="2"><img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18"></td>
			<%End if%>
			
			<!-- Looping Isi Data to Table  -->			
		    <%for i = 0 to nCountofField%>
				<% if rsVH(i).Name="ID" Then %>  <!-- Untuk ID sebagai PK, tidak perlu ditampilkan -->
				<%Else%>
				<td align="left" title="<%=rsVH(i).Name%>" valign="top"><%=rsVH(i)%></td>
				<%End If%>
				
		    <%next%>
		  </tr>
		  <%
		  ''response.Write(q)
		  
		  rsVH.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend
	else
	  response.Write("<td align=""center"" colspan=""21"" title="" valign=""top""> Data Is Empty</td>")	  
   end if  
   ''response.Write(q)
   rsVH.Close
   set rsVH = Nothing%>
   </TABLE>
  </div>

<%End if%>  
 
  <!-------------akhir!------------->
</div></div>
  <!--#include file="bottom.jsp"-->

</body>
</html>
