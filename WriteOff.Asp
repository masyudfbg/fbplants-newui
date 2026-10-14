<%Option Explicit%>
<%response.buffer=True%>
<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="mm.inc"-->
<!--#include file="DailyValidation.inc"-->
<!--#include file="qtyOnHandFunction.inc"-->
<html>
<head>
<title>ePlaS - Stock Write Off</title>
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

function getHarga()
{
	var WODate = $('#WODate').val();
	var xCode = $('#xCode').val();
	var itemcode = $('#ITEMCODE').val();
	var tgl = WODate.split("/");
	var bulan = tgl[0];
	var tahun = tgl[2];
	$("#harga").val(0);
		$.post("bkm_viewdata.asp?tahun=" + tahun + "&bulan=" + bulan + "&idview=" + 15 + "&storecode=" + xCode+"&itemcode=" + itemcode,function(data) {
		var json  = JSON.parse(data);
		var harga = json.results[0].amtperunit
		harga     = Math.round(harga);
		
		if(harga>0){
			$("#price").val(harga);
		}else{
			$("#price").val(0);
		}
		
	});
}

function getAmount()
{
	var qty = parseFloat($('#QUANTITY').val());
	var harga = parseFloat($('#price').val());
	var amt = qty*harga;
	$("#amount").val(amt);

}
</script>
<!--<script language="JavaScript" src="/js/date-picker.js"></script>-->
<script language="JavaScript">

var newwin;
function launchwin(winurl,winname,winfeatures)
{
	//This launches a new window and then
	//focuses it if window.focus() is supported.
	newwin = window.open(winurl,winname,winfeatures);

}
</script>
<% if Request("oldui") <> "1" then %>
<!-- Tampilan baru (css/fbp-shell.css + css/fbp-forms.css + css/fbp-docdetail.css). ?oldui=1 = tampilan lama -->
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap" rel="stylesheet">
<link rel="StyleSheet" href="css/login/font-awesome.min.css" type="text/css" />
<link rel="StyleSheet" href="css/fbp-shell.css?v=20261009d" type="text/css" />
<link rel="StyleSheet" href="css/fbp-forms.css?v=20261009e" type="text/css" />
<link rel="StyleSheet" href="css/fbp-docdetail.css?v=20261009b" type="text/css" />
<script type="text/javascript" src="js/fbp-shell.js?v=20261014a"></script>
<% end if %>
</head>
<body<% if Request("oldui") <> "1" then %> class="dash page-show page-legacy page-doc"<% end if %>>
<!--#include file="top.jsp"-->
<div style=" width:21.7%; float:left;margin-left:0.3%; margin-top:0.3%; margin-bottom:0.3%;">
 <!--#include file="menu.asp"-->
</div>
<div class="judul_head" valign="top" bgcolor="#fff" style=" width:77.4%; margin-left:0.3%; margin-right:0.3%; margin-top:0.3%; margin-bottom:0.3%; float:left; border:none;"> 
 	<div class="judul">
	   Module : <%=request("txtModule")%> / <%=request("txtTable")%>
	</div>
 <div class="box_both" style="overflow:scroll;">  
 <%
 '-------------awal-----------------------
 
 Dim xCode,fcCode,cmdView,PoCode,WODate,ITEMCODE,QUANTITY,AMOUNT, LineNoHist
 Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis
 Dim Jobcode, LocationType, LocationCode,price
 Dim qLineMax, rsqLineMax, tempMaxLine
 Dim qTotalRecordOnHand, rsTotalRecordOnHand, tempTotalRecordOnHand
 Dim qQuantityOnHand, rsQuantityOnHand, tempQuantityOnHand
 Dim done, qq  
 Dim qLessZero, rsLessZero, tempislesszero, REFERENCE, REMARKS
 Dim qLineUpper, rsqLineUpper, totalUpper, qQuantity, rsqQuantity, totalQuantity, QuantityOriginal, qDeleteHist
 Dim strLineMax, rsstrLineMax, tempMaxLineNo , LineNoKey, rsReceiveDetail, txtstorecode, txtLineNoHist, originalQuantity
 Dim txtItemCodeDelete,txtLineNoKeyDelete, v_CrossCekDetail
 Dim Cek_IsDailyClosing : Cek_IsDailyClosing = ""
 DIM qrylessstockonhand,qrylessstockledger
 Dim ExecuteLogOnHand,ExecuteLogLedger,txtModuleOnhand,txtModuleLedger
 dim ScriptName,modulFile,executeQry,rsWODetail,arr_dtlrn,x
 ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

txtModuleOnhand = "WriteOff.asp untuk update stockOnHand"
txtModuleLedger = "WriteOff.asp untuk update stockLedger"
 txtTable =Request("txtTable")
 txtModule=Request("txtModule")
 txtSearch=Request("txtSearch")
 txtField =Request("txtField")

 if request("fccOde")="" then
    fcCode="" '----- Nantinya Ambil dari Format Master
 else
    fcCode = request("fccOde")
 end if
 
 
 txtJenis = request("txtJenis") '---- Untuk Membedakan Jenis Data Yang akan diolah [insert,edit,delete]
 
 '---- Data Header
 if Request("WODate")<>"" then
   WODate   =Request("WODate")
 else
   WODate   =textDate(date)
 end if
 xCode = Request("xCode")
 
 '---- Data Details
 JobCode=Request("JobCode")
 LocationType=Request("LocationType")
 LocationCode=Request("LocationCode")
 ITEMCODE=Request("ITEMCODE")
 QUANTITY=Request("QUANTITY")
 AMOUNT =Request("AMOUNT")
 Linenohist = Request ("linenohist")
 LineNoKey = Request ("LineNoKey")
	REFERENCE=request("REFERENCE")
	REMARKS=request("REMARKS")
	price=request("price")
	
Function CheckMasterStock(initialStockCode, initialFcba)
	Dim initialAvailable
	Dim rsAvailable, strAvailable
	
	strAvailable =	" Select Count(1) Total From Stock Where" & _
					" Fccode = '" & initialStockCode & "'" & _
					" And Fcba = '" & initialFcba & "'"
							
	Set rsAvailable = Server.CreateObject("ADODB.Recordset")
	rsAvailable.Open strAvailable, connstr, 3, 3
	
	If Not rsAvailable.EOF Then
		initialAvailable = Cdbl(rsAvailable("Total"))
	Else
		initialAvailable = 0
	End If
	strAvailable = ""
	rsAvailable.Close
	set rsAvailable = Nothing

	CheckMasterStock = initialAvailable
End Function

Function CheckMasterJob(initialJobCode, initialFcba)
	Dim initialAvailable
	Dim rsAvailable, strAvailable
	
	strAvailable =	" Select Count(1) Total From Job Where" & _
					" Fccode = '" & initialJobCode & "'" & _
					" And Fcba = '" & initialFcba & "'"
							
	Set rsAvailable = Server.CreateObject("ADODB.Recordset")
	rsAvailable.Open strAvailable, connstr, 3, 3
	
	If Not rsAvailable.EOF Then
		initialAvailable = Cdbl(rsAvailable("Total"))
	Else
		initialAvailable = 0
	End If
	strAvailable = ""
	rsAvailable.Close
	set rsAvailable = Nothing

	CheckMasterJob = initialAvailable
End Function

Function CheckMasterTargetType(initialTargetType)
	Dim initialAvailable
	Dim rsAvailable, strAvailable
	
	strAvailable =	" Select Count(1) Total From Locationtype Where" & _
					" Locationtypecode = '" & initialTargetType & "'"
							
	Set rsAvailable = Server.CreateObject("ADODB.Recordset")
	rsAvailable.Open strAvailable, connstr, 3, 3
	
	If Not rsAvailable.EOF Then
		initialAvailable = Cdbl(rsAvailable("Total"))
	Else
		initialAvailable = 0
	End If
	strAvailable = ""
	rsAvailable.Close
	set rsAvailable = Nothing

	CheckMasterTargetType = initialAvailable
End Function

Function CheckTriggerControl(initialJobCode, initialSource, initialTarget)
	Dim initialAvailable
	Dim rsAvailable, strAvailable
	
	strAvailable =	" Select Count(1) Total From TRIGGER_CONTROL Where" & _
					" ACTIVITY = '" & initialJobCode & "'" & _
					" AND SOURCE = '" & initialSource & "'" & _
					" AND TARGET = '" & initialTarget & "'"
							
	Set rsAvailable = Server.CreateObject("ADODB.Recordset")
	rsAvailable.Open strAvailable, connstr, 3, 3
	
	If Not rsAvailable.EOF Then
		initialAvailable = Cdbl(rsAvailable("Total"))
	Else
		initialAvailable = 0
	End If
	strAvailable = ""
	rsAvailable.Close
	set rsAvailable = Nothing

	CheckTriggerControl = initialAvailable
End Function
 
 '--- Action ke Database di Oracle [ insert,edit,delete ]
 if Request("txtJenis")="New" or Request("txtJenis")= "Header" or Request("txtJenis")= "Details" or Request("txtJenis")="Edit" or Request("txtJenis")="Approved" Then
	if Request("txtJenis")= "Header" then
		connstr.BeginTrans
		
		Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), WODate)
		If Cek_IsDailyClosing = "FALSE" then
		
			   if err.number = 0 then
				  q="Delete from WriteOff where  WRITEOFFCODE='" & fcCode & "' and fcBA='" & Session("fcBA") & "'"
				  ExecuteAndLog "WriteOff.asp", q
				 ' connStr.execute(q)
				  if err.number=0 then
					 connStr.CommitTrans
				  end if
			   end if
			   
		else
			err.number = 1001
			err.description = "Tanggal """ & WODate & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
		End If
       cmdView	= "True"
	elseif Request("txtJenis")= "Details" then
	    connstr.BeginTrans
		
		Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), WODate)
		If Cek_IsDailyClosing = "FALSE" then
			
			qLessZero = " select islesszero from stock " & _	   
						" where fccode = '" & ITEMCODE & "'" & _
						"       and fcba = '" & session("fcBA") & "'" 

			set rsLessZero = Server.CreateObject("ADODB.Recordset")
			rsLessZero.Open qLessZero, connstr
			if not rsLessZero.eof then
				tempislesszero = rsLessZero("islesszero")
			else
				tempislesszero = "FALSE"
			end if 
			rsLessZero.Close
			set rsLessZero = Nothing
			
			qQuantity = " select quantity*-1 quantity" & _
						" from stocktranshist " & _
						" where lineno = " & linenohist & "" & _
						"       and fcba = '" & Session("fcBA") & "'"
			
			set rsqQuantity= Server.CreateObject("ADODB.Recordset")
			rsqQuantity.Open qQuantity, connstr

			if not rsqQuantity.eof then
				totalQuantity = cdbl(rsqQuantity("quantity"))
			else
				totalQuantity = 0
			end if 
			
			qQuantityOnHand = " select quantity from stockonhand " & _
							  " where storecode = '" & xCode & "'" & _
							  "       and stockcode  = '" & ITEMCODE & "'" & _
							  "       and fcba = '" & session("fcBA") &"' "
							  
			set rsQuantityOnHand = Server.CreateObject("ADODB.Recordset")
			rsQuantityOnHand.Open qQuantityOnHand, connstr
			if not rsQuantityOnHand.eof then
				QuantityOriginal = cdbl(rsQuantityOnHand("quantity"))
			else
				QuantityOriginal = 0
			end if 
		   
		   
			q = " Select storecode,wodate" & _
				" from writeoff" & _
				" where WriteOffCode = '" & fcCode & "'" & _
				"       and fcBA = '" & Session("fcBA") & "'"
				
				
			Set rsReceiveDetail = Server.CreateObject("ADODB.Recordset")
			rsReceiveDetail.Open q, connstr
					
			IF Not rsReceiveDetail.eof then
				txtStoreCode = rsReceiveDetail("STORECODE")
				wodate 		 = rsReceiveDetail("wodate")
			End if
			
			rsReceiveDetail.Close
			set rsReceiveDetail = Nothing		   
		   
		   
			q = " Select stockcode" & _
				"        , linenohist" & _
				"        , quantity" & _
				"        , LineNoKey" & _
				"        , price" & _
				" from writeoffdetails" & _
				" where WRITEOFFCODE = '" & fcCode & "'" & _
				"       And LineNoKey = '" & LineNoKey & "'" & _
				"       And StockCode = '" & ItemCode & "'" & _
				"       AND JobCode = '" & JobCode & "'" & _
				"       and Locationtype = '" & Locationtype  & "'" & _
				"       and Locationcode = '" & Locationcode & "'" & _
				"       and fcBA = '" & Session("fcBA") & "'"
			

			'Response.Write q	
			Set rsReceiveDetail = Server.CreateObject("ADODB.Recordset")
			rsReceiveDetail.Open q, connstr
			
		
					
			IF Not rsReceiveDetail.eof then
				txtLineNoHist = rsReceiveDetail("LINENOHIST")
				originalQuantity = rsReceiveDetail("QUANTITY")
				txtItemCodeDelete = rsReceiveDetail("stockcode")
				txtLineNoKeyDelete = rsReceiveDetail("LineNoKey")
				price = rsReceiveDetail("price")
			End if
			rsReceiveDetail.Close
			set rsReceiveDetail = Nothing			

			'v_CrossCekDetail = cmdCheckHist(txtLineNoHist, ItemCode)

			
		
			if tempislesszero = "TRUE" or ((QuantityOriginal + totalQuantity + Quantity)>= 0 And tempislesszero = "FALSE")  then

				'function untuk mengembalikan qty ketika ada proses delete
			qrylessstockonhand	= cmdUpdateStockOnHand(txtStoreCode,txtItemCodeDelete,originalQuantity,"WRT","Delete")
			qrylessstockledger	= cmdUpdateStockLedger(txtStoreCode,txtItemCodeDelete,wodate,originalQuantity,"WRT","Delete")
			'Response.Write qrylessstockledger

				q =  " Delete from stocktranshist " & _
					 " where lineno = " & txtLineNoHist & "" &_
					 "       and fcba = '" & Session("fcBA") & "'"
				ExecuteAndLog "WriteOff.asp", q
				'Response.Write q	
				'connStr.execute(q)		
				
									
				q = " Delete from WriteOffDetails" & _
					" where WRITEOFFCODE ='" & fcCode & "'" & _
					"       And LineNoKey = " & LineNoKey & "" & _
					"       And StockCode = '" & ItemCode & "'" & _
					"       AND JobCode = '" & JobCode & "'" & _
					"       and Locationtype = '" & Locationtype  & "'" & _
					"       and Locationcode='" & Locationcode & "'" & _
					"       and fcBA='" & Session("fcBA") & "'"
				ExecuteAndLog "WriteOff.asp", q
				'connStr.execute(q)
				if err.number = 0 then
				   connStr.CommitTrans
				end if

			else
				Err.Number = 1001
				Err.Description = "Quantity tidak boleh Negatif atau Saldo belum didefinisikan"
			end if
			
		else
			err.number = 1001
			err.description = "Tanggal """ & WODate & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
		End If
	    cmdView	= "True"
	elseif Request("txtJenis")="New" then
		connstr.BeginTrans
		
		Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), WODate)
		If Cek_IsDailyClosing = "FALSE" then
			
			Dim oRS, oSQL
			Set oRS = Server.CreateObject("ADODB.Recordset")
			oSQL	=	" SELECT WRITEOFFCODE FROM WRITEOFF" & _
						" WHERE WRITEOFFCODE = '" & fcCode & "'" & _
						"       AND FCBA = '" & Session("FCBA") & "'"
			oRS.Open oSQL, connstr
			
			IF oRS.eof then	    
				q	=	" INSERT INTO WRITEOFF (" & _
						"        WRITEOFFCODE, STORECODE, WODATE, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME, ROWSTATE" & _
						"        ) VALUES (" & _
						" '" & fcCode & "'" & _
						" , '" & xCode & "'" & _
						" , to_date('" & WODate & "','mm/dd/yyyy')" & _
						" , '" & Session("CodeName") & "'" & _
						" , '" & Session("CodeName") & "'" & _
						" , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
						" , '" & Session("fcBA") & "'" & _
						" , sysdate" & _
						" , To_Char(sysdate,'HH24:MI')" & _ 
							" , 'Planned')"
					
				connStr.execute(q)
				modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
				executeQry	= "Query executed:" & q
				ExecuteAndLog modulFile,executeQry
			End If
			oRS.Close
			set oRS = Nothing
			
			If Len(Trim(ITEMCODE)) = 0 Then
				Err.Number = 1001
				Err.Description = "Kode Stock Harus diisi, tidak boleh kosong"
			Else
				Dim val_Stock : val_Stock = 0
				val_Stock = CheckMasterStock(ITEMCODE, session("fcBA"))
				If val_Stock = 0 Then
					Err.Number = 1001
					Err.Description = "Kode Stock """ & ITEMCODE & """ tidak ada di master stock FBPlants, Koreksi kembali stock yang anda masukkan."
				End If
			End If
			
			If Len(Trim(Jobcode)) = 0 Then
				Err.Number = 1001
				Err.Description = "Jobcode Harus diisi, tidak boleh kosong"
			Else
				Dim val_Job : val_Job = 0
				val_Job = CheckMasterJob(Jobcode, session("fcBA"))
				If val_Job = 0 Then
					Err.Number = 1001
					Err.Description = "Jobcode """ & Jobcode & """ tidak ada di master Job FBPlants, Koreksi kembali Job yang anda masukkan."
				End If
			End If
			
			If Len(Trim(LOCATIONTYPE)) = 0 Then
				Err.Number = 1001
				Err.Description = "Target Type Harus diisi, tidak boleh kosong"
			Else
				Dim val_TargetType : val_TargetType = 0
				val_TargetType = CheckMasterTargetType(LOCATIONTYPE)
				If val_TargetType = 0 Then
					Err.Number = 1001
					Err.Description = "Target Type """ & LOCATIONTYPE & """ tidak ada di master Target Type FBPlants, Koreksi kembali Target Type yang anda masukkan."
				End If
			End If
			
			If val_Stock > 0 And val_Job > 0 And val_TargetType > 0 Then
				Dim val_TriggerControl : val_TriggerControl = 0
				val_TriggerControl = CheckTriggerControl(Jobcode, "ST", LOCATIONTYPE)
				If val_TriggerControl = 0 Then
					Err.Number = 1001
					Err.Description = "Jobcode """ & Jobcode & """ tidak bisa digunakan dari ""ST"" ke target """ & LOCATIONTYPE & """, Lihat kembali Trigger Management Anda !"
				Else
					If Len(Trim(QUANTITY)) = 0 Then
						Err.Number = 1001
						Err.Description = "QUANTITY tidak boleh kosong, minimal 0 !"
					ElseIf Len(Trim(AMOUNT)) = 0 Then
						Err.Number = 1001
						Err.Description = "AMOUNT tidak boleh kosong, minimal 0 !"
					Else
						If  QUANTITY < 0 and AMOUNT > 0 Then
							Err.Number = 1001
							Err.Description = "Jika Qty Negatif (-) Maka Amount Juga Seharusnya Negative (-) !"
						ElseIf  QUANTITY > 0 and AMOUNT < 0 Then
							Err.Number = 1001
							Err.Description = "Jika Qty Positive (+) Maka Amount Juga Seharusnya Positive (+) !"
						ElseIf  QUANTITY = 0 and AMOUNT = 0 Then
							Err.Number = 1001
							Err.Description = "Qty dan Amount tidak boleh sama-sama 0 !"
						Else
						
							If Err.Number = 0 Then
								qLessZero = " select islesszero from stock " & _	   
											" where fccode = '" & ITEMCODE & "'" & _
											"       and fcba = '" & session("fcBA") & "'" 

								set rsLessZero = Server.CreateObject("ADODB.Recordset")
								rsLessZero.Open qLessZero, connstr
								if not rsLessZero.eof then
									tempislesszero = rsLessZero("islesszero")
								else
									tempislesszero = "FALSE"
								end if 
								rsLessZero.Close
								set rsLessZero = Nothing
								
								qQuantityOnHand = " select nvl(quantity, 0) quantity" & _
												  " from stockonhand" & _
												  " where storecode = '" & xCode & "'" & _
												  "       and stockcode  = '" & ITEMCODE & "'" & _
												  "       and fcba = '" & session("fcBA") &"'"
												  
								set rsQuantityOnHand = Server.CreateObject("ADODB.Recordset")
								rsQuantityOnHand.Open qQuantityOnHand, connstr
								if not rsQuantityOnHand.eof then
									tempQuantityOnHand = cdbl(rsQuantityOnHand("quantity"))
								else
									tempQuantityOnHand = 0
								end if 				  
								rsQuantityOnHand.Close
								set rsQuantityOnHand = Nothing
										
								'if ((tempQuantityOnHand + Quantity >= 0) And tempislesszero = "FALSE") or tempislesszero = "TRUE" then
									qLineMax =	" select nvl(max(lineno),0) tempMaxLine from stocktranshist " & _
												" where fcba = '" & session("fcBA") & "'" 
											
									set rsqLineMax = Server.CreateObject("ADODB.Recordset")
									rsqLineMax.Open qLineMax, connstr
									if not rsqLineMax.eof then
										tempMaxLine = cdbl(rsqLineMax("tempMaxLine"))
									else
										tempMaxLine = 0
									end if 

									tempMaxLine = tempMaxLine + 1
									'Added by NAF, 2012-12-01
									strLineMax = " select nvl(max(linenokey),0) maxLineNoKey from WriteOffDetails where fcba = '" & Session("fcBA") & "'"

								  
								
									Set rsstrLineMax = Server.CreateObject("ADODB.Recordset")
									rsstrLineMax.CursorType = 3
									rsstrLineMax.CursorLocation = 3
									rsstrLineMax.Open strLineMax, connstr
									
									if not rsstrLineMax.eof then
										tempMaxLineNo = cdbl(rsstrLineMax("maxLineNoKey")) + 1
									end if
												
									rsstrLineMax.Close
									set rsstrLineMax = Nothing 
									'End of Added by NAF, 2012-12-01
									
									q	=   " Insert into WriteOffDetails (" & _
											"        WRITEOFFCODE, STOCKCODE, LOCATIONCODE, LOCATIONTYPE, QUANTITY, PRICE" & _
											"        , AMOUNT, JOBCODE, FCENTRY, FCEDIT, FCIP, FCBA" & _
											"        , LINENOHIST, REFERENCE, REMARKS, LINENOKEY, LASTUPDATE, LASTTIME" & _
											" ) VALUES (" & _
											" '" & fcCode & "'" & _
											" , '" & ITEMCODE & "'" & _
											" , '" & LOCATIONCODE & "'" & _
											" , '" & LOCATIONTYPE & "'" & _
											" , " & Quantity & " " & _
											" ,  " & price & " " & _
											" , " & AMOUNT & "" & _
											" , '" & Jobcode & "'" & _
											" , '" & Session("CodeName")  & "'" & _
											" , '" & Session("CodeName") & "'" & _
											" , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
											" , '" & Session("fcBA") & "'" & _
											" , " & tempMaxLine &  "" & _
											" , '" & REFERENCE & "'" & _
											" , '" & REMARKS & "'" & _								
											" , " & tempMaxLineNo & "" & _
											" , sysdate" & _
											" , To_Char(sysdate,'HH24:MI'))"
									connStr.execute(q)
									modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
									executeQry	= "Query executed:" & q
									ExecuteAndLog modulFile,executeQry
									
									q	=	" INSERT INTO STOCKTRANSHIST(" & _
											"     LINENO, FCBA, DOCUMENTNO, STOCKCODE, TRANSDATE, TRANSTYPE, STORECODE" & _
											"    , SOURCETYPE, SOURCECODE, TARGETTYPE, TARGETCODE, QUANTITY, LASTUPDATE, LASTTIME" & _
											" ) values (" & _
											" " & tempMaxLine & "" & _
											" , '" & session("fcBA") & "'" & _
											" , '" & fcCode & "'" & _
											" , '" & ITEMCODE & "'" & _
											" , to_date('" & WODate & "', 'mm/dd/yyyy')" & _
											" , 'ST_WRITEOFF'" & _
											" , '" & xCode & "'" & _
											" , 'ST'" & _
											" , '" & xCode & "'" & _
											" , '" & LOCATIONTYPE & "'" & _
											" , '" & LOCATIONCODE & "'" & _
											" , " & Quantity & "" & _
											" , sysdate" & _
											" , To_Char(sysdate,'HH24:MI') )"
									connStr.execute(q)
									modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
									executeQry	= "Query executed:" & q
									ExecuteAndLog modulFile,executeQry
								
									If Err.Number = 0 Then
										connStr.CommitTrans
									End If
								'else
									'Err.Number = 1001
									'Err.Description = "Quantity tidak boleh Negatif atau Saldo belum didefinisikan"
								'end if
							End If
						End If
					
					End If			
				End If
			End If
			
		else
			err.number = 1001
			err.description = "Tanggal """ & WODate & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
		End If
		
		cmdView	= "True"
	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
	    connstr.BeginTrans
		
		Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), WODate)
		If Cek_IsDailyClosing = "FALSE" then
			
			If Len(Trim(QUANTITY)) = 0 Then
				Err.Number = 1001
				Err.Description = "QUANTITY tidak boleh kosong, minimal 0 !"
			ElseIf Len(Trim(AMOUNT)) = 0 Then
				Err.Number = 1001
				Err.Description = "AMOUNT tidak boleh kosong, minimal 0 !"
			Else
			
				If  QUANTITY < 0 and AMOUNT > 0 Then
					Err.Number = 1001
					Err.Description = "Jika Qty Negatif (-) Maka Amount Juga Seharusnya Negative (-) !"
				ElseIf  QUANTITY > 0 and AMOUNT < 0 Then
					Err.Number = 1001
					Err.Description = "Jika Qty Positive (+) Maka Amount Juga Seharusnya Positive (+) !"
				ElseIf  QUANTITY = 0 and AMOUNT = 0 Then
					Err.Number = 1001
					Err.Description = "Qty dan Amount tidak boleh sama-sama 0 !"
				Else
				
					qLineUpper =   "select count(*) totalUpper from stocktranshist " & _
							   "where " & _
							   "lineno > " & linenohist & " and " & _
							   "fcba = '" & Session("fcBA") & "' and " & _
							   "stockcode = '" & ItemCode & "' " 

					set rsqLineUpper = Server.CreateObject("ADODB.Recordset")
					rsqLineUpper.Open qLineUpper, connstr
					
					if not rsqLineUpper.eof then
						totalUpper = cdbl(rsqLineUpper("totalUpper"))
					else
						totalUpper = 0
					end if 

						qQuantity = "select quantity*-1 quantity  from stocktranshist " & _
							"where " & _
							"lineno = " & linenohist & " and " & _ 
							"fcba = '" & Session("fcBA") & "'"
						
						set rsqQuantity= Server.CreateObject("ADODB.Recordset")
						rsqQuantity.Open qQuantity, connstr

						if not rsqQuantity.eof then
							totalQuantity = cdbl(rsqQuantity("quantity"))
						else
							totalQuantity = 0
						end if 
						
						
							qDeleteHist = "delete stocktranshist " & _
							"where " & _
							"lineno = " & linenohist & " and " & _ 
							"fcba = '" & Session("fcBA") & "'"
							connStr.execute(qDeleteHist)
							modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
							executeQry	= "Query executed:" & qDeleteHist
							ExecuteAndLog modulFile,executeQry
							
							q	=	" INSERT INTO STOCKTRANSHIST(" & _
									"     LINENO, FCBA, DOCUMENTNO, STOCKCODE, TRANSDATE, TRANSTYPE, STORECODE" & _
									"    , SOURCETYPE, SOURCECODE, TARGETTYPE, TARGETCODE, QUANTITY, LASTUPDATE, LASTTIME" & _
									" ) values (" & _
									" " & linenohist & "" & _
									" , '" & session("fcBA") & "'" & _
									" , '" & fcCode & "'" & _
									" , '" & ITEMCODE & "'" & _
									" , to_date('" & WODate & "', 'mm/dd/yyyy')" & _
									" , 'ST_WRITEOFF'" & _
									" , '" & xCode & "'" & _
									" , 'ST'" & _
									" , '" & xCode & "'" & _
									" , '" & LOCATIONTYPE & "'" & _
									" , '" & LOCATIONCODE & "'" & _
									" , " & Quantity & "" & _
									" , sysdate" & _
									" , To_Char(sysdate,'HH24:MI') )"
							connStr.execute(q)
							modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
							executeQry	= "Query executed:" & q
							ExecuteAndLog modulFile,executeQry
					   
							q=	" Update WriteOffDetails set Quantity = " & Quantity & "" & _
								" , JobCode = " & JobCode & "" & _
								" , Locationtype = '" & Locationtype & "'" & _
								" , Locationcode = '" & Locationcode & "'" & _
								" , price = '" & price & "'" & _
								" , Amount = " & Amount & "" & _
								" , REFERENCE = '" & REFERENCE & "'" & _
								" , REMARKS = '" & REMARKS & "'" & _
								" , LASTUPDATE = sysdate " & _
								" , LASTTIME = To_Char(sysdate,'HH24:MI') "& _
								" , fcedit = '" & Session("CodeName") & "'" & _
								" where WRITEOFFCODE = '" & fcCode & "'" & _
								"       And LineNoKey = " & LineNoKey & "" & _
								"       And StockCode = '" & ItemCode & "'" & _ 
								"       and fcBA='" & Session("fcBA") & "'"
							connStr.execute(q)
							modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
							executeQry	= "Query executed:" & q
							ExecuteAndLog modulFile,executeQry
							
							'function untuk mengembalikan qty ketika ada proses delete atau update
						  '	qrylessstockonhand	= cmdUpdateStockOnHand(xCode,ITEMCODE,QuantityOriginal,"WRT","Update")
						  '	qrylessstockonhand	= UpdateNewStockOnHand(xCode,ITEMCODE,Quantity,"WRT","Update")
						  '	qrylessstockledger 	= cmdUpdateStockLedger(xCode,ITEMCODE,WODate,QuantityOriginal,"WRT","Update")
						  '	qrylessstockledger 	= UpdateNewStockLedger(xCode,ITEMCODE,WODate,Quantity,"WRT","Update")
							
							if err.number  = 0 then
							  connStr.CommitTrans
							end if
						
				End If
			End If
			
		else
			err.number = 1001
			err.description = "Tanggal """ & WODate & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
		End If
		cmdView	= "True"
    
	elseif request("txtJenis")="Approved" then 

			'query 
			dim arr_dtlwo_eof,quantitywriteoff
			arr_dtlwo_eof =  true
			q = "Select W.WRITEOFFCODE, W.STOCKCODE, W.LOCATIONCODE,  W.QUANTITY,  " & _
			" W.PRICE, W.AMOUNT, W.JOBCODE, W.LINENOHIST, " & _
			" W.LINENOKEY,H.STORECODE " & _
			" From   WRITEOFFDETAILS W " & _
			" INNER JOIN WRITEOFF H ON H.WRITEOFFCODE = W.WRITEOFFCODE AND H.FCBA = W.FCBA " & _
			" WHERE W.WRITEOFFCODE='" & fcCode & "'"

			set rsWODetail = Server.CreateObject("ADODB.Recordset")
			rsWODetail.Open q, connstr
			if not rsWODetail.eof then 
				arr_dtlrn = rsWODetail.getrows()
				arr_dtlwo_eof = false
			End IF
			rsWODetail.close : set rsWODetail = nothing

			if not arr_dtlwo_eof then 
				for x=0 to ubound(arr_dtlrn,2)
					ITEMCODE 			= arr_dtlrn(1,x)
					xCode 				= arr_dtlrn(9,x)
					quantitywriteoff 	= arr_dtlrn(3,x)

					qTotalRecordOnHand = "select count(*) totalRecordOnHand from stockonhand " & _
											"where " & _
											"storecode = '" & xCode & "' and " & _												 
											"stockcode = '" & ITEMCODE & "' and " & _
											"fcba = '" &  session("fcBA") & "'"

					set rsTotalRecordOnHand = Server.CreateObject("ADODB.Recordset")
					rsTotalRecordOnHand.Open qTotalRecordOnHand, connstr
					if not rsTotalRecordOnHand.eof then
						tempTotalRecordOnHand = cdbl(rsTotalRecordOnHand("totalRecordOnHand"))
					else
						tempTotalRecordOnHand = 0
					end if 
					rsTotalRecordOnHand.Close
					set rsTotalRecordOnHand = Nothing
					
					if tempTotalRecordOnHand > 0 then
						'We have to Update the Quantity by Adding them with the New one.
						q = " update stockonhand" & _
							" set quantity = quantity + " & quantitywriteoff & "" & _ 
							" , LASTUPDATE = sysdate" & _
							" , LASTTIME = To_Char(sysdate,'HH24:MI')"& _
							" where storecode = '" & xCode & "'" & _
							"       and stockcode  = '" & ITEMCODE & "'" & _
							"       and fcba = '" & session("fcBA") &"'"
						connStr.execute(q)
						modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
						executeQry	= "Query executed:" & q
						ExecuteAndLog modulFile,executeQry
					Else
						q = " Insert into stockonhand ( " & _
							"     STORECODE, STOCKCODE, FCBA, QUANTITY, LASTUPDATE, LASTTIME" & _
							" ) values (" & _
							" '" & xCode & "'" & _
							" , '" & ITEMCODE & "'" & _
							" , '" & session("fcBA") & "'" & _
							" , " & quantitywriteoff & "" & _
							" , sysdate" & _
							" , To_Char(sysdate,'HH24:MI') ) "									
					connStr.execute(q)
					modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
					executeQry	= "Query executed:" & q
					ExecuteAndLog modulFile,executeQry
					End If


					'CEK ITEM DAHULU DI TABEL STOCKLEDGER
					'JIKA TIDAK ADA INSERT BARU, SELAIN ITU UPDATE
					dim qryStockLedger, rsStockLedger,temp,totalRecordStockLedger,qryinsertstockledger,qryupdateStockLedger
					qryStockLedger = "SELECT COUNT(*) totalRecordStockLedger FROM STOCKLEDGER" & _ 
					" WHERE " & _ 
					" STORECODE = '" & xCode & "' and " & _
					" ITEMCODE  = '" & ITEMCODE & "' and " & _ 
					" year  	= to_char(to_date('" & WODate & "','mm/dd/yyyy'),'yyyy') and " & _ 
					" month 	= to_char(to_date('" & WODate & "','mm/dd/yyyy'),'mm') and " & _ 
					" FCBA 		= '" & session("fcBA") & "' "
					
					set rsStockLedger = Server.CreateObject("ADODB.Recordset")
					rsStockLedger.Open qryStockLedger, connstr
					if not rsStockLedger.eof then
							totalRecordStockLedger = cdbl(rsStockLedger("totalRecordStockLedger"))
					else
							totalRecordStockLedger = 0
					end if 
				

					'function ini digunakan untuk insert/update data di tabel stockledger
					' modify jun (19 juni 2024)
					If totalRecordStockLedger = 0 then
					
						'insert stockledger
						qryinsertstockledger 	= cmdInsertStockLedger(xCode,ITEMCODE,WODate,quantitywriteoff, "WRT")
					ELSE
					
						'update stockledger
						qryupdateStockLedger 	= cmdUpdateStockLedger(xCode,ITEMCODE,WODate,quantitywriteoff, "WRT", "Insert")
					END IF
		
			NEXT 
		END IF
		
		q = "Update writeoff set rowstate='Approved',postingflg=1 " & _
		" where WRITEOFFCODE='" & fcCode & "'"
		connStr.execute(q)
		modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
		executeQry	= "Query executed:" & q
		ExecuteAndLog modulFile,executeQry

		cmdView	= "True"
	

	else
        cmdView	= "True"
	end if								 
	'Response.Write "sql " & q
 else
    cmdView	= request("cmdView")
 End if

 Dim rsPurchase,rowstate
 set rsPurchase = server.CreateObject("Adodb.recordset")
 q="select WRITEOFFCODE, STORECODE, to_char(WODATE,'mm/dd/yyyy') WODATE, FCENTRY, FCEDIT, " & _
      " FCIP, FCBA, LASTUPDATE, LASTTIME, ROWSTATE from WriteOff  " & _
  " where WRITEOFFCODE='" & fcCode & "' and " & _
  " fcBA='" & session("fcBA") & "'"
 rsPurchase.Open q, connstr
 if Not rsPurchase.Eof then
    WODate		= rsPurchase("WODate")
	xCode		= rsPurchase("StoreCode")
	rowstate	= rsPurchase("rowstate")
 End if
 if done = 1 then 
   connStr.RollbackTrans
   response.Write  "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" & err.description & "</font>"
 end if
 
 if err.number <> 0 then
   connStr.RollbackTrans
   response.Write  "<div class=""error"">" & err.description & "</div>"
 else
    if done = 1 then 
	   connStr.RollbackTrans
	   response.Write  "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>Nilai Quantity tidak boleh Negatif atau Saldo belum didefinisikan</font>"	
	end if
	if done = 2 then 
	   connStr.RollbackTrans
	   response.Write  "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>Terdapat transaksi setelah data ini, harap menghapus transaksi setelah ini</font>"	   
	end if
	
 End if
 '-------------ISI Payment Form
 

'approval user
dim isApproved,rsApproved
set rsApproved = server.CreateObject("Adodb.recordset")
q = " SELECT COUNT(1) total" & _
	" FROM authorizelevel al" & _
	" INNER JOIN authorizestructure aus ON aus.authorizelevelcode = al.fccode" & _
	" INNER JOIN authorizetemplate at ON aus.authorizetemplatecode = at.fccode" & _
	" WHERE at.document_type = 'WO' and at.fccode='509'  " & _
	"       AND al.isactive = 'TRUE'" & _
	"       AND al.startlevel = 'FALSE'" & _
	"       AND aus.fcusercode = '" & Session("CodeName") & "'" & _
	"       AND aus.fcba = '" & session("fcBA") & "'"
' response.write q
rsApproved.Open q, connstr
if Not rsApproved.Eof then
	isApproved =  cdbl(to_nol(rsApproved("total")))
else
	isApproved = 0
End if
rsApproved.Close : set rsApproved = Nothing
 %>
 <% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-trash-o" aria-hidden="true"></i>Header Write Off</div><% end if %>
 <form name="frmEntry" method="post" action="WriteOff.Asp">
	<table class="data" width="100%">
		<tr>
		<td align="left">WriteOff Code</td>
		<td align="left"><input type="text" maxlength="15" name="fcCode" size="30" value="<%=fcCode%>"><font color="#FF0000">&nbsp;*</font></td></tr><tr>
		<td align="left">Date</td>
		<td align="left"><input class="datepicker" type="text" maxlength="10" name="WODate" id="WODate" size="10" value="<%=WODate%>">
		<font color="#FF0000">&nbsp;*</font>
		</td></tr>
		<tr>
		<td align="left">Store</td>
		<td align="left">      
		<select name="xCode" id="xCode">
	    <%cmdCallCombo "Store",xCode,"",""%>
        </select><font color="#FF0000">&nbsp;*</font>
        </td>
		</td>
		</tr>
		<tr><td colspan="2"  align="left" class="fbp-doc-actions">
		<input Type="Hidden" Value="True" Name="cmdView">
		<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		<input Type="Submit"  name="Add" value="Select">&nbsp;&nbsp;
	    <a href="javascript:launchwin('PrintWriteOff.Asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&fcCode=<%=fcCode%>&WODate=<%=WODate%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><img src="IMAGES/Print.gif" border="0" title="Print Preview"><font face="Verdana" size="1">Print</font></a>&nbsp;&nbsp;
	
		<% if rowstate = "Planned" OR rowstate = "" then %>
			<a href="WriteOff.Asp?fcCode=<%=fcCode%>&WODate=<%=WODate%>&txtJenis=Header" onClick="return confirm('Are you sure to delete');"><img src="IMAGES/delete.gif" alt="Delete Data : <%=fccode%>" border="0"></a>
		<%Else%>
			<img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18">
		<%End If%>
		<!-- End Period Control In Here -->
		
		<% if rowstate = "Planned" then %>
			<%if isApproved = 0 then%>
				
			<%else%>
				<a href="?fcCode=<%=fcCode%>&WODate=<%=WODate%>&fcba=<%=session("fcBA")%>&txtJenis=Approved" onClick="return confirm('Are you sure to start authorize this data :  <%=fcCode%> ?');">Approved</a>
			<% end if%>
		<% end if%>

		<%
		dim dataModule,dataSubmodule,dataDetail
		dataModule 		= "Stores"
		dataSubmodule 	= "Write Off"
		dataDetail 		= "NULL"
		%>
		<a class="jqplot-image-button" style="" href="ajax/hadirAjax.asp?dataModule=<%=dataModule%>&dataSubmodule=<%=dataSubmodule%>&dataDetail=<%=dataDetail%>"><i class="fa fa-check-square-o"></i>&nbsp;Hadir</a>
		<!-- END:HADIR -->
		</td>
		</tr>
    </table>  
 </form>
<%if cmdView = "True" AND LEN(TRIM(fcCode)) > 4 AND Left(TRIM(xCode), 2) = "ST" Then
   Dim rsPurchaseOrder,intRecordsPerPage,cColor
   set rsPurchaseOrder = server.CreateObject("Adodb.recordset")
   q="select WRITEOFFCODE,  LINENOKEY, STOCKCODE, JOBCODE, LOCATIONTYPE, LOCATIONCODE, QUANTITY, PRICE, AMOUNT,  LINENOHIST, REFERENCE, REMARKS, FCENTRY, FCEDIT, FCIP, FCBA from WriteOffdETAILS  " & _
     " where WRITEOFFCODE='" & fcCode & "' and " & _
     " fcBA='" & session("fcBA") & "'"
   'response.Write q	
   Set rsPurchaseOrder = Server.CreateObject("ADODB.Recordset")
   rsPurchaseOrder.CursorType = 3
   rsPurchaseOrder.CursorLocation = 3
   rsPurchaseOrder.Open q, connstr%>
   <% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-edit" aria-hidden="true"></i>Detail Item</div><% end if %>
   <form action="WriteOff.Asp" method="post" name="eForm">
   <table class="data" cellpadding="0" cellspacing="0" width="100%">
		<tr><td align="left">Item Code</td>
		<td align="left">
		<%IF txtJenis="Edit" Then%>
		     <input type="text" size="15" name="ITEMCODE" id="ITEMCODE" value=<%=ITEMCODE%> readonly="True" style="background:#999999">&nbsp;<font color="#FF0000">* <a href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=Stock&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><font face="Verdana" size="1">Stock</font></a></td>
			</tr>
			<tr><td align="left">Job Code</td><td align="left"><input type="text"  size="15" name="Jobcode" value=<%=Jobcode%> >&nbsp;<font color="#FF0000">*
			<a href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=Job&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><font face="Verdana" size="1">Job</font></a></td>
			</tr>
			<tr><td align="left">TT</td><td align="left"><input type="text" size="15" name="LocationType" value=<%=LocationType%> >&nbsp;<font color="#FF0000">*<a href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=LocationType&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><font face="Verdana"   size="1"> Location Type</a></td></tr>
			<tr><td align="left">TC</td><td align="left"><input type="text" size="15" name="LocationCode" value=<%=LocationCode%> >&nbsp;<font color="#FF0000">*</td></tr>
			<input type="hidden" size="15" name="Linenohist" value="<%=Linenohist%>" readonly="True" style="background:#999999">				
			<input type="hidden" size="15" name="LineNoKey" value="<%=LineNoKey%>" readonly="True" style="background:#999999">				

		<%else%>
			<input type="text" size="15" name="ITEMCODE" id="ITEMCODE" onkeyup="getHarga();"  value=<%=ITEMCODE%>>&nbsp;<font color="#FF0000">* <a href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=Stock&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><font face="Verdana" size="1">Stock</font></a></td>
			</tr>
			<tr><td align="left">Job Code</td><td align="left"><input type="text" size="15" name="Jobcode" value=1140999>
			&nbsp;<font color="#FF0000">*
			<a href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=Job&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><font face="Verdana" size="1">Job</font></a></td>
			</tr>
			<tr><td align="left">TT</td><td align="left"><input type="text" readonly size="15" name="LocationType" value=CB>&nbsp;<font color="#FF0000">*<a href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=LocationType&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><font face="Verdana"   size="1"> Location Type</a></td></tr>
			<tr><td align="left">TC</td><td align="left"><input type="text" readonly size="15" name="LocationCode" value=BS>&nbsp;<font color="#FF0000">*</td></tr>
			<input type="hidden" size="15" name="Linenohist" value="<%=Linenohist%>" readonly="True" style="background:#999999">
			<input type="hidden" size="15" name="LineNoKey" value="<%=LineNoKey%>" readonly="True" style="background:#999999">
			

		<%End if%>
		<tr><td align="left">Harga Terakhir</td><td align="left"><input type="text" size="15"  name="price" id="price" value=<%=price%>></td></tr>
		<tr><td align="left">Quantity</td><td align="left"><input type="text" size="15" onkeyup="getAmount()" name="QUANTITY" id="QUANTITY" value=<%=QUANTITY%>>&nbsp;<font color="#FF0000">*</font> (+) Adjust Positive, (-) Adjust Negative</td></tr>
		<tr><td align="left">Amount</td><td align="left"><input type="text" size="15" name="Amount" id="amount" readonly value=<%=Amount%>>&nbsp;<font color="#FF0000">*</font> (+) Adjust Positive, (-) Adjust Negative</td></tr>
		<tr><td align="left">REFERENCE</td><td align="left"><input type="text" size="30" maxlength = "30" name="REFERENCE" value="<%=REFERENCE%>"></td></tr>
		<tr><td align="left">REMARKS</td><td align="left"><textarea  name="REMARKS" maxlength = "100" cols="50" rows="3"><%=REMARKS%></textarea></td></tr>
	    <TR><td colspan="2" align="left">
		  <!-- diambil dari form yang diatas -->
	      <input type="hidden" name="fcCode" value="<%=fcCode%>">
		  <input type="hidden" name="WODate" value="<%=WODate%>">
		  <input type="hidden" name="xCode" value="<%=xCode%>">
		  
		  <input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		  <input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		  <%IF Request("txtJenis")="Edit" then%>
		       <input type="hidden" name="txtJenis" value="Edit">
		       <input Type="Submit"  name="Add" value="Update">
		  <%Else%>
		       <input type="hidden" name="txtJenis" value="New">
		       <input Type="Submit"  name="Add" value="Add">
		  <%End if%>
		  <input Type="Reset"  name="Reset" value="Reset">
		  <input Type="Button"  value="Back" name="B2" OnClick="javascript:history.back(1)"></td>
	  </tr>
   </table>

   <% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-list" aria-hidden="true"></i>Daftar Item</div><% end if %>
   <div class="fbp-doc-scroll">
   <table class="data fbp-doc-lines">
	<thead>
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;" >
	<TD align="left">NO</TD>
	<TD colspan="2">&nbsp;</TD>	
	<TD align="left">WRITEOFF CODE</TD>	
	<TD align="left">Line No (Key)</TD>		
	<TD align="left">STOCK CODE</TD>	
	<TD align="left">JOB CODE</TD>
	<TD align="left">LOC TYPE</TD>
	<TD align="left">LOC CODE</TD>			
	<TD align="left">QUANTITY</TD>	
	<TD align="left">PRICE</TD>	
	<TD align="left">AMOUNT</TD>	
	<TD align="left">Line No (history)</TD>			
	<TD align="left">REFERENCE</TD>
	<TD align="left">REMARKS</TD>		
	</tr>
	</thead>
	<tbody>  
   <%
   if Not rsPurchaseOrder.eof then
      nCountofField=rsPurchaseOrder.fields.count-1
	  no=1
	  
      While Not rsPurchaseOrder.EOF

		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if%>
		  <tr bgcolor="<%=cColor%>">
		  <td align="left" valign="top"><%=NO%></td>
			<% if rowstate = "Planned" OR rowstate = "" then %>
				<td valign="middle" width="30">
					<a href="WriteOff.Asp?fcCode=<%=rsPurchaseOrder("WRITEOFFCODE")%>&ItemCode=<%=rsPurchaseOrder("StockCode")%>&Quantity=<%=rsPurchaseOrder("Quantity")%>&Amount=<%=rsPurchaseOrder("Amount")%>&JobCode=<%=rsPurchaseOrder("JobCode")%>&LocationType=<%=rsPurchaseOrder("LocationType")%>&LocationCode=<%=rsPurchaseOrder("LocationCode")%>&REFERENCE=<%=rsPurchaseOrder("REFERENCE")%>&REMARKS=<%=rsPurchaseOrder("REMARKS")%>&LINENOHIST=<%=rsPurchaseOrder("LINENOHIST")%>&LineNoKey=<%=rsPurchaseOrder("LineNoKey")%>&price=<%=rsPurchaseOrder("PRICE")%>&txtJenis=Edit">
					<img src="IMAGES/edit.gif" alt="Edit Data : <%=rsPurchaseOrder(1)%>" border="0"></a>
				</td>
				<td valign="middle" width="10">
					<a href="WriteOff.Asp?fcCode=<%=rsPurchaseOrder("WRITEOFFCODE")%>&ItemCode=<%=rsPurchaseOrder("StockCode")%>&Quantity=<%=rsPurchaseOrder("Quantity")%>&Amount=<%=rsPurchaseOrder("Amount")%>&JobCode=<%=rsPurchaseOrder("JobCode")%>&LocationType=<%=rsPurchaseOrder("LocationType")%>&LocationCode=<%=rsPurchaseOrder("LocationCode")%>&REFERENCE=<%=rsPurchaseOrder("REFERENCE")%>&REMARKS=<%=rsPurchaseOrder("REMARKS")%>&LINENOHIST=<%=rsPurchaseOrder("LINENOHIST")%>&LineNoKey=<%=rsPurchaseOrder("LineNoKey")%>&WODate=<%=WODate%>&txtJenis=Details"  onClick="return confirm('Are you sure to delete');">
					<img src="IMAGES/delete.gif" alt="Delete Data : <%=rsPurchaseOrder(1)%>" border="0"></a>
				</td>
			<%else%>
				<td valign="middle" width="30">
					<img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18">
				</td>
				<td valign="middle" width="10">
					<img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18">
				</td>
			<%end if%>
			<!-- End Period Control In Here -->
		    <%for i = 0 to nCountofField - 4%>
			    <td align="left" title="<%=rsPurchaseOrder(i).Name%>" valign="top"><%=rsPurchaseOrder(i)%></td>
		    <%next%>		    
		  </tr>
		  <%
		  rsPurchaseOrder.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend	   
   else
      response.Write("<tr><td colspan=""15"" class=""fbp-doc-empty"">Data is empty</td></tr>")	  
   end if 
   rsPurchaseOrder.Close
   set rsPurchaseOrder = Nothing%>
   </tbody>
   </TABLE>
   </div>
<% if Request("oldui") <> "1" then %>
<script type="text/javascript">
// Status baris jadi badge (tampilan saja)
(function () {
	var cells = document.querySelectorAll('table.fbp-doc-lines tbody td[title]');
	for (var i = 0; i < cells.length; i++) {
		if (cells[i].getAttribute('title').toUpperCase() !== 'ROWSTATEID') continue;
		var s = (cells[i].textContent || '').replace(/\s+/g, ' ').trim();
		if (!s) continue;
		cells[i].parentNode.setAttribute('data-state', s);
		cells[i].innerHTML = '';
		var b = document.createElement('span');
		b.className = 'fbp-state';
		b.textContent = s;
		cells[i].appendChild(b);
	}
})();
</script>
<% end if %>
   </form>
<%End if%>  
<%'-------------akhir----------------------%>
</div></div>
  <!--#include file="bottom.jsp"-->

</body>
</html>
