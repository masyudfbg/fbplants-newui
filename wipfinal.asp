<%
	validatesession
%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="mm.inc"-->
<!--#include file="DailyValidation.inc"-->
<html>
<head>
	<title>ePlaS - Final Work in Progress</title>
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
	//$( ".datepicker" ).datepicker();
	$("#tglawal").datepicker();
	$("#tglakhir").datepicker();
	$("#wipdate").datepicker();
	
});
</script>
<!--<script language="JavaScript" src="/js/date-picker.js"></script> -->
<script language="JavaScript">

var newwin;
function launchwin(winurl,winname,winfeatures)
{
	//This launches a new window and then
	//focuses it if window.focus() is supported.
	newwin = window.open(winurl,winname,winfeatures);

}
</script>

<script type='text/javascript'>
    //Replaces all non numeric characters
    function checkInput(element){ element.value = element.value.replace(/[^0-9\.]/g,'');}
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
		<% if Session("fcBa") <> "" then %>
         <!--#include file="menu.asp"-->
        <% end if %>
	</div>
    <div class="judul_head" valign="top" bgcolor="#fff" style=" width:77.4%; margin-left:0.3%; margin-right:0.3%; margin-top:0.3%; margin-bottom:0.3%; float:left; border:none;"> 
    <div class="judul">
	   Module : <%=request("txtModule")%> / <%=request("txtTable")%>
	</div>
    <div class="box_both" style="overflow:scroll;">
<%
'--Awal--
Dim v_txtagreementcode, v_txtmonth, v_txtyear, MinYr, MaxYr, mon, Yr, n, SudahClose, v_txtagreementfinalno, v_txttglawal,v_txttglakhir
Dim xCode,xstoreCode,fcCode,cmdView,DATECREATED,VoucherCode,QUANTITY,AMOUNT,OTHERITEMDESC,v_CrossCekPO
Dim Remarks
Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis,TOTALAMOUNT,HeaderRemarks
Dim Jobcode, LocationType, LocationCode, Reference, Remark
Dim qLineMax, rsqLineMax, maxagreementfinalno
Dim qTotalRecordOnHand, rsTotalRecordOnHand, tempTotalRecordOnHand
Dim qQuantityOnHand, rsQuantityOnHand, tempQuantityOnHand
Dim sqlSQLCheck, rsSQL, qtyRemained
Dim qCheckStore, rsqCheckStore
Dim txtLocationType	
Dim rsCheckRate, checkRate, varRate
Dim Cek_IsDailyClosing : Cek_IsDailyClosing = ""
Dim wipdate ,can_startauthorize,isprocessing 

Dim stringduedate : stringduedate = 0

Dim IsOverDuedate, varRemainDueDate, varDueDate, CompleteStringError, varErrorDueDate, stringremainqty, varRemainQty, IsOverRemainQty, varErrorRemainQty, varEditedQty

Dim qtyplanedit : qtyplanedit = Request("qtyplanedit")
Dim unitplanedit : unitplanedit = Request("unitplanedit")
Dim outputplanedit : outputplanedit = Request("outputplanedit")
Dim qtyfinaledit : qtyfinaledit = Request("qtyfinaledit")
Dim unitfinaledit : unitfinaledit = Request("unitfinaledit")
Dim outputfinaledit : outputfinaledit = Request("outputfinaledit")
Dim ratefinaledit : ratefinaledit = Request("ratefinaledit")
Dim txtagreementlineno : txtagreementlineno = Request("txtagreementlineno")

Dim txtGRNCode : txtGRNCode = 0
Dim oRs : oRs = 0
Dim oSql : oSql = 0
Dim txtLocationCode : txtLocationCode = ""
Dim ScriptName

ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

txtTable =Request("txtTable")
txtModule=Request("txtModule")
txtSearch=Request("txtSearch")
txtField =Request("txtField")

v_bappno           = Replace(Request("txtbappno")," ","")
v_txtagreementcode = Request("txtagreementcode")
v_txttglawal = Request("tglawal")
v_txttglakhir = Request("tglakhir")
v_txtmonth = Request("txtmonth")
v_txtyear = Request("txtyear")
v_txtagreementfinalno = Request("txtagreementfinalno")

 if request("wipdate") <>"" then
   wipdate   = request("wipdate")
 else
   wipdate = textDate(date)
 end if
  
 if KonversiTanggalKeDbl(wipdate) > KonversiTanggalKeDbl(textDate(date)) then
	wipdate = textDate(date)
 end if
 

Yr = Year(wipdate)
mon = Month(wipdate)
MinYr = Yr - 2
MaxYr = Yr  + 2

If len(v_txtyear)= 0 then v_txtyear = Yr
If len(v_txtmonth) = 0 then v_txtmonth = mon

'wipdate = v_txtmonth & "/" & GetLastDay(v_txtmonth, v_txtyear) & "/" & v_txtyear

txtJenis = request("txtJenis") '-- Untuk Membedakan Jenis Data Yang akan diolah [insert,edit,delete]

set rsFrez = server.CreateObject("Adodb.recordset")
q = " SELECT COUNT(1) total" & _
	" FROM authorizelevel al" & _
	" INNER JOIN authorizestructure aus ON aus.authorizelevelcode = al.fccode" & _
	" INNER JOIN authorizetemplate at ON aus.authorizetemplatecode = at.fccode" & _
	" INNER JOIN authorizelevel_users alu ON (alu.authorizelevelcode = al.fccode)" & _
	" WHERE at.document_type = 'AA' and at.fccode='504' " & _
	"       AND al.isactive = 'TRUE'" & _
	"       AND al.finallevel = 'TRUE'" & _
	"		AND alu.isactive = 'TRUE'" & _
	"       AND '"& session("CodeName") &"' = '" & session("CodeName") & "'" & _
	"       AND aus.fcusercode = '" & session("CodeName") & "'" & _
	"       AND aus.fcba = '" & session("fcba") & "'"
rsFrez.Open q, connstr
if Not rsFrez.Eof then
	can_startauthorize =  cdbl(to_nol(rsFrez("total")))
else
	can_startauthorize = 0
End if
rsFrez.Close : set rsFrez = Nothing 

Sub insertdatawipfinal_detail
	q = " insert into workinprogressfinal_new (ID,BAPPNO " & _          
		"                                  , WIPFINALDATE    " & _ 
		"                                  , YEAR            " & _ 
		"                                  , MONTH           " & _ 
		"                                  , AGREEMENTFINALNO" & _ 
		"                                  , AGREEMENTCODE   " & _ 
		"                                  , WIPDATE         " & _ 
		"                                  , AGREEMENTLINENO " & _ 
		"                                  , CHARGEUNIT      " & _ 
		"                                  , JOBCODE         " & _ 
		"                                  , LOCATIONTYPE    " & _ 
		"                                  , LOCATIONCODE    " & _ 
		"                                  , RATE_BASIC      " & _ 
		"                                  , QTY_PLAN        " & _ 
		"                                  , OUTPUT_PLAN     " & _ 
		"                                  , UNIT_PLAN       " & _ 
		"                                  , QTY_FINAL       " & _ 
		"                                  , OUTPUT_FINAL    " & _ 
		"                                  , UNIT_FINAL      " & _ 
		"                                  , RATE_FINAL      " & _ 
		"                                  , ISPROCESSING    " & _ 
		"                                  , FCENTRY         " & _ 
		"                                  , FCEDIT          " & _ 
		"                                  , FCIP            " & _ 
		"                                  , FCBA            " & _ 
		"                                  , LASTUPDATE      " & _ 
		"                                  , LASTTIME        " & _ 
		"                                  )" & _
		
		" values ( " & TRIM(n) & ",'" & v_bappno & "'" & _
		"         , to_date('" & wipdate & "', 'mm/dd/yyyy')" & _
		"         , '" & v_txtyear & "'" & _
        "         , '" & v_txtmonth & "'" & _
		"		  , '" & maxagreementfinalno & "'" & _
        "         , '" & v_txtagreementcode & "'" & _
        "         , to_date('" & request("wipdate_"&TRIM(n)) & "', 'mm/dd/yyyy')" & _
        "         , '" & request("agreementlineno_"&TRIM(n)) & "'" & _
        "         , '" & request("chargeunit_"&TRIM(n)) & "'" & _
		"         , '" & request("jobcode_"&TRIM(n)) & "'" & _
		"         , '" & request("targettype_"&TRIM(n)) & "'" & _
		"         , '" & request("targetcode_"&TRIM(n)) & "'" & _
		"         , '" & Cdbl(To_Nol(request("ctkrate_"&TRIM(n)))) & "'" & _
		"         , '" & Cdbl(To_Nol(request("qtyplan_"&TRIM(n)))) & "'" & _
		"         , '" & Cdbl(To_Nol(request("outputplan_"&TRIM(n)))) & "'" & _
	    "         , '" & Cdbl(To_Nol(request("unitplan_"&TRIM(n)))) & "'" & _
		"         , '" & Cdbl(To_Nol(request("qtyfinal_"&TRIM(n)))) & "'" & _
		"         , '" & Cdbl(To_Nol(request("outputfinal_"&TRIM(n)))) & "'" & _
		"         , '" & Cdbl(To_Nol(request("unitfinal_"&TRIM(n)))) & "'" & _
        "         , '" & Cdbl(To_Nol(request("ratefinal_"&TRIM(n)))) & "'" & _
	    "         , 'FALSE'" & _
        "         , '" & Session("CodeName") & "'" & _
        "         , '" & Session("CodeName") & "'" & _
        "         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
        "         , '" & Session("fcBA") & "'" & _
        "         , sysdate " & _
        "         , to_char(sysdate, 'hh24:mi')" & _
		" )"
			
	'Response.Write q
	'connStr.execute(q)	
	ExecuteAndLog ScriptName, q
End Sub

	
Sub insertdatawipfinal
	connstr.BeginTrans

	' Update Data to table workinprogressfinal_new
	q = "UPDATE workinprogressfinal_new SET ISPROCESSING='TRUE' " & _ 
	    "WHERE bappno = '" & v_bappno & "' and agreementcode = '" & v_txtagreementcode & "' and fcba ='" & Session("fcBA") & "'" 
	'Response.Write q
	'connStr.execute(q)
	ExecuteAndLog ScriptName, q	
		
		
	
	q=  " MERGE INTO workinprogressfinal a " & _ 
		" USING " & _ 
		" (SELECT bappno,agreementcode,agreementfinalno,wipfinaldate,wipdate,agreementlineno,year,month,fcba,chargeunit,rate_final " & _ 
		"  ,sum(qty_final) qty, sum(unit_final) unit, sum(output_final) output " & _ 
		"  FROM workinprogressfinal_new " & _ 
		"  WHERE bappno = '" & v_bappno & "' and agreementcode = '" & v_txtagreementcode & "' and fcba ='" & Session("fcBA") & "'" & _
		"  GROUP BY bappno,agreementcode,agreementfinalno,wipfinaldate,wipdate,agreementlineno,year,month,fcba,chargeunit,rate_final " & _
		" )b on (a.bappno=b.bappno and a.agreementcode=b.agreementcode and a.agreementfinalno=b.agreementfinalno and a.wipfinaldate=b.wipfinaldate and a.wipdate=b.wipdate and a.agreementlineno=b.agreementlineno and a.year=b.year and a.month=b.Month and a.fcba=b.fcba and a.chargeunit=b.chargeunit and a.rate_final=b.rate_final) " & _ 
		" WHEN MATCHED THEN UPDATE SET " & _
		" a.QTY = b.Qty, a.unit = b.unit , a.output=b.output " & _
		" WHEN NOT MATCHED THEN INSERT (a.AGREEMENTFINALNO,a.AGREEMENTCODE, a.WIPDATE,a.AGREEMENTLINENO,a.YEAR,a.MONTH,a.QTY,a.OUTPUT,a.ISPROCESSING,a.FCENTRY,a.FCEDIT, a.FCIP, a.FCBA,a.LASTUPDATE, a.LASTTIME,a.UNIT,a.CHARGEUNIT,a.WIPFINALDATE,a.BAPPNO,a.RATE_FINAL) " & _
        " VALUES ( b.AGREEMENTFINALNO,b.AGREEMENTCODE, b.WIPDATE,b.AGREEMENTLINENO,b.YEAR,b.MONTH,b.QTY,b.OUTPUT,'TRUE','" & Session("CodeName") & "','" & Session("CodeName") & "','" & Request.ServerVariables("REMOTE_ADDR") & "', b.FCBA,sysdate, to_char(sysdate, 'hh24:mi'),b.UNIT,b.CHARGEUNIT,b.WIPFINALDATE,b.BAPPNO,b.RATE_FINAL) "		
    'Response.Write q
	'connStr.execute(q)	
	ExecuteAndLog ScriptName, q
	
	
     q=  " MERGE INTO workinprogressplan a " & _ 
		" USING " & _ 
		" (SELECT bappno,agreementcode,agreementfinalno,wipfinaldate,wipdate,99999 agreementlineno,year,month,fcba,chargeunit,rate_final,jobcode,locationtype,locationcode ,99999 lineno" & _ 
		"  ,sum(qty_final-qty_plan) qty, sum(unit_final - unit_plan) unit, sum(output_final - output_plan) output " & _ 
		"  ,sum((qty_final-qty_plan) * rate_final) wipamt " & _
		"  FROM workinprogressfinal_new " & _ 
		"  WHERE bappno = '" & v_bappno & "' and agreementcode = '" & v_txtagreementcode & "' and fcba ='" & Session("fcBA") & "'" & _
		"  GROUP BY bappno,agreementcode,agreementfinalno,wipfinaldate,wipdate,year,month,fcba,chargeunit,rate_final,jobcode,locationtype,locationcode " & _
		" )b on (a.lineno=b.lineno and a.agreementcode=b.agreementcode and a.agreementfinalno=b.agreementfinalno and a.wipdate=b.wipdate and a.agreementlineno=b.agreementlineno  and a.fcba=b.fcba and a.chargeunit=b.chargeunit and a.jobcode=b.jobcode and a.locationtype=b.locationtype and a.locationcode=b.locationcode ) " & _ 
		" WHEN MATCHED THEN UPDATE SET " & _
		" a.QTY = b.Qty, a.unit = b.unit , a.output=b.output,a.wipamt=b.wipamt " & _
		" WHEN NOT MATCHED THEN INSERT (a.agreementcode,a.wipdate,a.agreementlineno,a.lineno,a.qty, a.unit,a.output,a.postedflag,a.isgenerated,a.fcentry,a.fcedit,a.fcip,a.fcba,a.wipamt,a.reference,a.remarks,a.lastupdate,a.lasttime,a.agreementfinalno,a.jobcode,a.locationtype,a.locationcode,a.chargeunit) " & _
	    " VALUES ( b.agreementcode,b.wipdate,b.agreementlineno,b.agreementlineno ,b.qty, b.unit,b.output,0,'TRUE','" & Session("CodeName") & "','" & Session("CodeName") & "','" & Request.ServerVariables("REMOTE_ADDR") & "',b.fcba,b.wipamt,null,null,sysdate, to_char(sysdate, 'hh24:mi'),b.agreementfinalno,b.jobcode,b.locationtype,b.locationcode,b.chargeunit) "		
    'Response.Write q
	'connStr.execute(q)	
	ExecuteAndLog ScriptName, q
	
	
	 q=  " MERGE INTO workinprogress a " & _ 
		" USING " & _ 
		" (SELECT bappno,agreementcode,agreementfinalno,wipfinaldate,wipdate,agreementlineno,year,month,fcba,chargeunit,rate_final" & _ 
		"  ,sum(qty_final) qty, sum(unit_final) unit, sum(output_final) output " & _ 
		"  ,sum((qty_final) * rate_final) wipamt " & _
		"  FROM workinprogressfinal_new " & _ 
		"  WHERE bappno = '" & v_bappno & "' and agreementcode = '" & v_txtagreementcode & "' and fcba ='" & Session("fcBA") & "'" & _
		"  GROUP BY bappno,agreementcode,agreementfinalno,wipfinaldate,wipdate,agreementlineno,year,month,fcba,chargeunit,rate_final " & _
		" )b on (a.agreementcode=b.agreementcode and a.wipdate=b.wipdate  and a.agreementlineno=b.agreementlineno  and  a.agreementfinalno=b.agreementfinalno  and a.fcba=b.fcba and a.chargeunit=b.chargeunit) " & _ 
		" WHEN MATCHED THEN UPDATE SET " & _
		" a.QTY = b.Qty, a.unit = b.unit , a.output=b.output,a.wipamt=b.wipamt " & _
		" WHEN NOT MATCHED THEN INSERT (a.agreementcode,a.wipdate,a.agreementlineno,a.qty, a.postedflag,a.progresspaymentdate,a.fcentry,a.fcedit,a.fcip,a.fcba,a.wipamt,a.paid,a.netbalance,a.reference,a.remarks,a.lastupdate,a.lasttime,a.agreementfinalno,a.chargeunit)" & _
	    " VALUES ( b.agreementcode,b.wipdate,b.agreementlineno,b.qty,0,b.wipdate,'" & Session("CodeName") & "','" & Session("CodeName") & "','" & Request.ServerVariables("REMOTE_ADDR") & "',b.fcba,b.wipamt,0,0,null,null,sysdate, to_char(sysdate, 'hh24:mi'),b.agreementfinalno,b.chargeunit) "		
    'Response.Write q
	'connStr.execute(q)	
	ExecuteAndLog ScriptName, q
	
	' ' Insert Data to table workinprogress
	' qq = " insert into workinprogress (agreementcode" & _
	    ' "                             , wipdate" & _
	    ' "                             , agreementlineno" & _
	    ' "                             , qty" & _
	    ' "                             , postedflag" & _
	    ' "                             , progresspaymentdate" & _
	    ' "                             , fcentry" & _
	    ' "                             , fcedit" & _
	    ' "                             , fcip" & _
	    ' "                             , fcba" & _
	    ' "                             , wipamt" & _
	    ' "                             , paid" & _
	    ' "                             , netbalance" & _
	    ' "                             , reference" & _
	    ' "                             , remarks" & _
	    ' "                             , lastupdate" & _
	    ' "                             , lasttime" & _
	    ' "                             , unit" & _
		' "                             , output" & _
	    ' "                             , agreementfinalno" & _
		' "                             , chargeunit)" & _
		' " values ('" & v_txtagreementcode & "'" & _
        ' "         , to_date('" & request("wipdate_"&TRIM(n)) & "', 'mm/dd/yyyy')" & _
        ' "         , '" & request("agreementlineno_"&TRIM(n)) & "'" & _
        ' "         , '" & Cdbl(To_Nol(request("qtyfinal_"&TRIM(n)))) & "'" & _
		' "         , 0" & _
		' "         , to_date('" & request("wipdate_"&TRIM(n)) & "', 'mm/dd/yyyy')" & _
        ' "         , '" & Session("CodeName") & "'" & _
        ' "         , '" & Session("CodeName") & "'" & _
        ' "         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
        ' "         , '" & Session("fcBA") & "'" & _
		' "         , '" & Cdbl(To_Nol(request("qtyfinal_"&TRIM(n)))) * varRate & "'" & _
		' "         , 0" & _
		' "         , 0" & _
		' "         , null" & _
		' "         , null" & _
        ' "         , sysdate " & _
        ' "         , to_char(sysdate, 'hh24:mi')" & _
        ' "         , '" & Cdbl(To_Nol(request("unitfinal_"&TRIM(n)))) & "'" & _
		' "         , '" & Cdbl(To_Nol(request("outputfinal_"&TRIM(n)))) & "'" & _
		' "         , '" & maxagreementfinalno & "'" & _
		' "         , '" & request("chargeunit_"&TRIM(n)) & "'" & _
        ' " )"

		' 'Response.Write qq
		' connStr.execute(qq)


		' update is final pada tabel workinprogressplan 
		'q = "update workinprogressplan  set isfinal=1 where agreementcode='" & v_txtagreementcode & "' and fcba='" & Session("fcBA") & "' and jobcode='" & request("jobcode_"&TRIM(n)) & "' AND locationcode='" & request("targetcode_"&TRIM(n)) & "'"
		
	q=  " MERGE INTO workinprogressplan a " & _ 
		" USING " & _ 
		" (SELECT 1 isfinal,agreementcode,agreementfinalno,agreementlineno,fcba,wipdate,Jobcode,locationtype,locationcode,chargeunit " & _ 
		"  FROM workinprogressfinal_new " & _ 
		"  WHERE bappno = '" & v_bappno & "' and agreementcode = '" & v_txtagreementcode & "' and fcba ='" & Session("fcBA") & "'" & _
		" )b on (a.agreementcode=b.agreementcode and a.agreementfinalno=b.agreementfinalno and a.wipdate=b.wipdate and a.agreementlineno=b.agreementlineno and a.fcba=b.fcba and a.chargeunit=b.chargeunit and a.Jobcode=b.Jobcode and a.LocationCode=b.LocationCode and a.LocationType=b.LocationType) " & _ 
		" WHEN MATCHED THEN UPDATE SET " & _
		" a.isfinal = b.isfinal " 
		
	'Response.Write q
	'connStr.execute(q)
	ExecuteAndLog ScriptName, q
	
	If Err.Number = 0 Then
			connstr.CommitTrans
			'Response.write("<font face=Verdana size=1 color=Green>Normalisasi employee_transit_tab, Success! " & " Time: " & time() & " ..<br/><br/></font>")
			'Response.Flush
		Else
			connstr.RollbackTrans
			err.number = 1001
			err.description = "Error Save ["&ExecuteSQL&"] ("&Err.Description&")"
			'Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
			'Response.Flush
		End If
End Sub

Sub Updatedatawipfinal
	oSql = " select rate from contractagreementdetail " & _
				" where fcba = '" & Session("fcBA") & "'" & _
				"       and agreementcode = '" & v_txtagreementcode & "'" & _
				"       and lineno = '" & txtagreementlineno & "'"
				
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open oSql, connstr
	if Not oRs.Eof then
		varRate = cdbl(to_nol(oRs("rate")))
	Else
		varRate = 0
	End if	

	oRs.Close
	set oRs = Nothing
	
	'Update data from table workinprogress
	q = " Update workinprogressfinal_new" & _
	    " set qty_final = '" & Cdbl(To_Nol(qtyfinaledit)) & "'" & _
		"     , unit_final = '" & Cdbl(To_Nol(unitfinaledit)) & "'" & _
		"     , OUTPUT_FINAL = '" & Cdbl(To_Nol(outputfinaledit)) & "'" & _
		"     , rate_final =  '" & Cdbl(To_Nol(ratefinaledit)) & "'" & _
		" where fcba = '" & Session("fcBA") & "'" & _
        "       and agreementfinalno = '" & v_txtagreementfinalno & "'" & _ 
		"       and bappno = '" & v_bappno & "'" & _ 
		"       and id = '" & Request("ID") & "'"
	'Response.Write q
	'connStr.execute(q)
	ExecuteAndLog ScriptName, q
	
	
	q = " Update workinprogress" & _
	    " set qty = '" & Cdbl(To_Nol(qtyfinaledit)) & "'" & _
		"     , unit = '" & Cdbl(To_Nol(unitfinaledit)) & "'" & _
		"     , output = '" & Cdbl(To_Nol(outputfinaledit)) & "'" & _
		"     , wipamt = '" & Cdbl(To_Nol(qtyfinaledit)) * varRate & "'" & _
        " where fcba = '" & Session("fcBA") & "'" & _
        "       and agreementfinalno = '" & v_txtagreementfinalno & "'"

	'Response.Write q
	'connStr.execute(q)
	
	'Update data from table workinprogressplan
	q = " Update workinprogressplan" & _
	    " set qty = '" & Cdbl(To_Nol(qtyfinaledit)) - Cdbl(To_Nol(qtyplanedit)) & "'" & _
		"     , unit = '" & Cdbl(To_Nol(unitfinaledit)) - Cdbl(To_Nol(unitplanedit)) & "'" & _
		"     , output = '" & Cdbl(To_Nol(outputfinaledit)) - Cdbl(To_Nol(outputplanedit)) & "'" & _
		"     , wipamt = '" & ( Cdbl(To_Nol(qtyfinaledit)) - Cdbl(To_Nol(qtyplanedit)) ) * varRate & "'" & _
        " where fcba = '" & Session("fcBA") & "'" & _
        "       and agreementfinalno = '" & v_txtagreementfinalno & "'"

	'Response.Write q
	'connStr.execute(q)
	
	'Update data from table workinprogressfinal
	q = " Update workinprogressfinal" & _
	    " set qty = '" & Cdbl(To_Nol(qtyfinaledit)) & "'" & _
		"     , unit = '" & Cdbl(To_Nol(unitfinaledit)) & "'" & _
		"     , output = '" & Cdbl(To_Nol(outputfinaledit)) & "'" & _
        " where fcba = '" & Session("fcBA") & "'" & _
        "       and agreementfinalno = '" & v_txtagreementfinalno & "'" & _ 
		"       and bappno = '" & v_bappno & "'"

	'Response.Write q
	'connStr.execute(q)
End Sub

Sub deletedatawipfinal
	'Delete data from table workinprogress
	q = " delete from workinprogress" & _
        " where fcba = '" & Session("fcBA") & "'" & _
        "       and agreementfinalno = '" & v_txtagreementfinalno & "'"

	'Response.Write q
	'connStr.execute(q)
	ExecuteAndLog ScriptName, q
	
	'Delete data from table workinprogressplan
	q = " delete from workinprogressplan" & _
        " where fcba = '" & Session("fcBA") & "'" & _
        "       and agreementfinalno = '" & v_txtagreementfinalno & "'"

	'Response.Write q
	'connStr.execute(q)
	ExecuteAndLog ScriptName, q
	
	'Delete data from table workinprogressfinal
	q = " delete from workinprogressfinal" & _
        " where fcba = '" & Session("fcBA") & "'" & _
        "       and agreementfinalno = '" & v_txtagreementfinalno & "'"

	'Response.Write q
	'connStr.execute(q)
	ExecuteAndLog ScriptName, q

	q = " update workinprogressplan set isfinal=0 where agreementcode='" & v_txtagreementcode & "'"
	'connStr.execute(q)
	ExecuteAndLog ScriptName, q
End Sub

 			'--- Action ke Database di Oracle [ insert,edit,delete ]
Cek_IsAccClosing = StatusAccClosingTgl(wipdate,session("fcba"))

if 	Cek_IsAccClosing = "FALSE" then 
if Request("txtJenis")="New" or Request("txtJenis")= "Header" or Request("txtJenis")= "Details" or Request("txtJenis")="Edit" or Request("txtJenis")="Confirm" Then
	if Request("txtJenis")= "Header" then
		Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), wipdate)
		If Cek_IsDailyClosing = "FALSE" then
			connstr.BeginTrans
			
			set oRs = server.CreateObject("Adodb.recordset")
			oSql = " select bappno,agreementcode,agreementfinalno,to_char(wipdate,'MM/DD/YYYY') wipdate " & _
				   " from workinprogressfinal_new " & _
				   " where fcba = '" & Session("fcBA") & "'" & _
				   "       and agreementcode = '" & v_txtagreementcode & "'" & _
				   "       and bappno = '" & v_bappno & "'" & _ 
				   "  AND wipdate between to_date('" & v_txttglawal & "','mm/dd/yyyy') and to_date('" & v_txttglakhir & "','mm/dd/yyyy')"
			

			oRs.Open oSql, connstr
			do while not oRs.Eof
				'Delete data from table workinprogress
				q = " delete from workinprogressfinal_new" & _
					" where fcba = '" & Session("fcBA") & "'" & _
					"       and bappno = '" & ors("bappno") & "'" & _ 
					"       and agreementfinalno = '" & ors("agreementfinalno") & "'"

			'	Response.Write q
				'connStr.execute(q)
				ExecuteAndLog ScriptName, q
				
				'Delete data from table workinprogress
				q = " delete from workinprogress" & _
					" where fcba = '" & Session("fcBA") & "'" & _
					"       and agreementfinalno = '" & ors("agreementfinalno") & "'"

			'	Response.Write q
				'connStr.execute(q)
				ExecuteAndLog ScriptName, q
				
				'Delete data from table workinprogressplan
				q = " delete from workinprogressplan" & _
					" where lineno=99999 and fcba = '" & Session("fcBA") & "'" & _
					"       and agreementfinalno = '" & ors("agreementfinalno") & "'"

				'Response.Write q
				'connStr.execute(q)
				ExecuteAndLog ScriptName, q
				
				'Delete data from table workinprogressfinal
				q = " delete from workinprogressfinal" & _
					" where fcba = '" & Session("fcBA") & "'" & _
					"       and agreementfinalno = '" & ors("agreementfinalno") & "'"

				'Response.Write q
				'connStr.execute(q)
				ExecuteAndLog ScriptName, q

				'Delete data from table workinprogressfinal
				q = " update workinprogressplan set isfinal=0 " & _
					" where fcba = '" & Session("fcBA") & "'" & _
					"       and agreementcode = '" & v_txtagreementcode & "'" & _ 
					"       and to_char(wipdate,'MM/DD/YYYY') = '" & ors("wipdate") & "'"

				'Response.Write q
				'connStr.execute(q)
				ExecuteAndLog ScriptName, q
				
				oRs.movenext
			loop
			oRs.Close
			Set oRs = Nothing

			if err.number = 0 then
				connStr.CommitTrans
			end if
		else
			CompleteStringError = "Tanggal """ & wipdate & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
		End If
		cmdView	= "True"
	elseif Request("txtJenis")= "Details" then
		Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), wipdate)
		If Cek_IsDailyClosing = "FALSE" then
			connstr.BeginTrans
			
			Call deletedatawipfinal
		
			if err.number=0 then
				connStr.CommitTrans
			End if
		else
			CompleteStringError = "Tanggal """ & wipdate & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
		End If
		cmdView	= "True"					
	elseif Request("txtJenis")="New" then
		Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), wipdate)
		If Cek_IsDailyClosing = "FALSE" then
			connstr.BeginTrans
			
			stringduedate = Cstr("0")
			stringremainqty = Cstr("0")
			varErrorDueDate = ""
			
			For n = 1 to (cdbl(request("no")) - 1)
			
				'Cek Remain Qty
				oSql = " select fcba" & _
					   "        , agreementcode" & _
					   "        , lineno" & _
					   "        , sum (totalagreement) totalagreement" & _
					   "        , sum (totalwip) totalwip" & _
					   "        , sum (totalagreement) - sum (totalwip) selisih" & _
					   " from (select fcba" & _
					   "              , agreementcode" & _
					   "              , lineno" & _
					   "              , nvl(qty,0) totalagreement" & _
					   "              , 0 totalwip" & _
					   "       from contractagreementdetail" & _
					   "       where fcba = '" & session("fcBA") & "'" & _
					   "             and agreementcode = '" & v_txtagreementcode & "'" & _
					   "             and lineno = '" & request("agreementlineno_"&TRIM(n)) & "'" & _
					   "       union all" & _
					   "       select fcba" & _
					   "              , agreementcode" & _
					   "              , agreementlineno lineno" & _
					   "              , 0 totalagreement" & _
					   "              , sum(nvl(qty_final,0)) totalwip" & _
					   "       from workinprogressfinal_new" & _
					   "       where fcba = '" & session("fcBA") & "'" & _
					   "             and agreementcode = '" & v_txtagreementcode & "'" & _
					   "             and agreementlineno = '" & request("agreementlineno_"&TRIM(n)) & "'" & _
					   "       Group BY fcba,agreementcode,agreementlineno " & _ 
					   " ) dataall" & _
					   " where 1=1" & _
					   " group by fcba" & _
					   "          , agreementcode" & _
					   "          , lineno"
				
				Set oRs = Server.CreateObject("ADODB.Recordset")
				oRs.Open oSql, connstr
				if Not oRs.Eof then
					varRemainQty = cdbl(to_nol(oRs("selisih")))
				Else
					varRemainQty = 0		   
				End if
				oRs.Close
				Set oRs = Nothing
				
				
				if ( varRemainQty - Cdbl(To_Nol(request("qtyfinal_"&TRIM(n)))) ) >= 0 then								
				stringremainqty = stringremainqty & Cstr("0")
				else
					stringremainqty = stringremainqty & Cstr("1")
					if len(trim(varErrorRemainQty)) = 0 then
						varErrorRemainQty = "&nbsp;&nbsp;&nbsp;&nbsp;agreementcode-->" & v_txtagreementcode & ", lineno-->" & request("agreementlineno_"&TRIM(n)) & ", Quantity Maximal-->" & varRemainQty & "<br/>"
					else
						varErrorRemainQty = varErrorRemainQty & "&nbsp;&nbsp;&nbsp;&nbsp;agreementcode-->" & v_txtagreementcode & ", lineno-->" & request("agreementlineno_"&TRIM(n)) & ", Quantity Maximal-->" & varRemainQty & "<br/>"
					end if
				end if
				
				set oRs = server.CreateObject("Adodb.recordset")
				oSql = " SELECT dtl.rate, hdr.finishdate FROM contractagreement hdr INNER JOIN contractagreementdetail dtl ON (dtl.fcba = hdr.fcba AND dtl.agreementcode = hdr.agreementcode)" & _
					   " WHERE dtl.fcba = '" & Session("fcBA") & "'" & _
					   "       and dtl.agreementcode = '" & v_txtagreementcode & "'" & _
					   "       and dtl.lineno = '" & request("agreementlineno_"&TRIM(n)) & "'"

				oRs.Open oSql, connstr
				if Not oRs.Eof then
					varRate = cdbl(oRs("rate"))
					varDueDate = oRs("finishdate")
				Else
					varRate = 0
				End if
				oRs.Close
				Set oRs = Nothing
				
				set oRs = server.CreateObject("Adodb.recordset")
				oSql = "select nvl(to_number(to_date('" & wipdate & "','mm/dd/yyyy') - to_date('" &  varDueDate & "','mm/dd/yyyy')),0) as remaindate from dual "

				oRs.Open oSql, connstr
				if Not oRs.Eof then
					varRemainDueDate = cdbl(oRs("RemainDate"))
				Else
					varRemainDueDate = 0
				End if
				oRs.Close
				Set oRs = Nothing
				
				if varRemainDueDate <= 0 then
					stringduedate = stringduedate & Cstr("0")
				else
					stringduedate = stringduedate & Cstr("1")
					if len(trim(varErrorDueDate)) = 0 then
						varErrorDueDate = "&nbsp;&nbsp;&nbsp;&nbsp;agreementcode-->" & v_txtagreementcode & ", WIPDate" & textdate(wipdate) & " lineno-->" & request("agreementlineno_"&TRIM(n)) & ", Due Date-->" & textdate(varDueDate) & "<br/>"
					else
						varErrorDueDate = varErrorDueDate & "&nbsp;&nbsp;&nbsp;&nbsp;agreementcode-->" & v_txtagreementcode & ", lineno-->" & request("agreementlineno_"&TRIM(n)) & ", Due Date-->" & textdate(varDueDate) & "<br/>"
					end if
				end if
				
			Next
			
			set oRs = server.CreateObject("Adodb.recordset")
			oSql = "select count(1) total from dual where instr('" & stringduedate & "', '1') <> 0"
			oRs.Open oSql, connstr
			if Not oRs.Eof then
				IsOverDuedate = cdbl(oRs("total"))
			Else
				IsOverDuedate = 1
			End if
			oRs.Close
			Set oRs = Nothing
			
			set oRs = server.CreateObject("Adodb.recordset")
			oSql = "select count(1) total from dual where instr('" & stringremainqty & "', '1') <> 0"
			oRs.Open oSql, connstr
			if Not oRs.Eof then
				IsOverRemainQty = cdbl(oRs("total"))
			Else
				IsOverRemainQty = 1
			End if
			oRs.Close
			Set oRs = Nothing
			
			CompleteStringError = ""
			if trim(v_bappno)= "" then 
				cmdView	= "True"
				CompleteStringError = "No.BAPP tidak ada !!" & "<br/>"
			elseif (IsOverRemainQty = 0 And IsOverDuedate = 0) then
					set oRs = server.CreateObject("Adodb.recordset")
					oSql = " select nvl(max(agreementfinalno), 0) maxagreementfinalno" & _
						   " from workinprogressfinal " & _
						   " where fcba = '" & session("fcBA") & "'"

					oRs.Open oSql, connstr
					if Not oRs.Eof then
						maxagreementfinalno = cdbl(oRs("maxagreementfinalno"))
					else
						maxagreementfinalno = 0
					End if
					oRs.Close
					Set oRs = Nothing
					
					For n = 1 to (cdbl(request("no")) - 1)
						maxagreementfinalno = maxagreementfinalno + 1
						Call insertdatawipfinal_detail
						'Call insertdatawipfinal
					Next
			elseif (IsOverRemainQty = 0 And IsOverDuedate = 1) then
				cmdView	= "True"
				CompleteStringError = "Tanggal BAPP harus lebih kecil dari Tanggal DUEDATE pada Contract Aggreement" & "<br/>"
				CompleteStringError = CompleteStringError & varErrorDueDate
			elseif (IsOverRemainQty = 1 And IsOverDuedate = 0) then
				cmdView	= "True"
				CompleteStringError = "Total Quantity yang anda masukkan melebihi dari kontrak yang disepakati" & "<br/>"
				CompleteStringError = CompleteStringError & varErrorRemainQty
			else
				cmdView	= "True"
				CompleteStringError = "Total Quantity yang anda masukkan melebihi dari kontrak yang disepakati" & "<br/>"
				CompleteStringError = CompleteStringError & varErrorRemainQty
				CompleteStringError = CompleteStringError & "<br/>" & "Tanggal BAPP harus lebih kecil dari Tanggal DUEDATE pada Contract Aggreement" & "<br/>"
				CompleteStringError = CompleteStringError & varErrorDueDate
			end if

			if err.number=0 then
				connStr.CommitTrans
			End if
		else
			CompleteStringError = "Tanggal """ & wipdate & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
		End If
		cmdView	= "True"
	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
		Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), wipdate)
		If Cek_IsDailyClosing = "FALSE" then
			'Cek Remain Qty
			oSql = " select fcba" & _
				   "        , agreementcode" & _
				   "        , lineno" & _
				   "        , sum (totalagreement) totalagreement" & _
				   "        , sum (totalwip) totalwip" & _
				   "        , sum (totalagreement) - sum (totalwip) selisih" & _
				   " from (select fcba" & _
				   "              , agreementcode" & _
				   "              , lineno" & _
				   "              , nvl(qty, 0) totalagreement" & _
				   "              , 0 totalwip" & _
				   "       from contractagreementdetail" & _
				   "       where fcba = '" & session("fcBA") & "'" & _
				   "             and agreementcode = '" & v_txtagreementcode & "'" & _
				   "             and lineno = '" & txtagreementlineno & "'" & _
				   "       union all" & _
				   "       select fcba" & _
				   "              , agreementcode" & _
				   "              , agreementlineno lineno" & _
				   "              , 0 totalagreement" & _
				   "              , nvl(qty, 0) totalwip" & _
				   "       from workinprogress" & _
				   "       where fcba = '" & session("fcBA") & "'" & _
				   "             and agreementcode = '" & v_txtagreementcode & "'" & _
				   "             and agreementlineno = '" & txtagreementlineno & "'" & _
				   " ) dataall" & _
				   " where 1=1" & _
				   " group by fcba" & _
				   "          , agreementcode" & _
				   "          , lineno"
			
			Set oRs = Server.CreateObject("ADODB.Recordset")
			oRs.Open oSql, connstr
			if Not oRs.Eof then
				varRemainQty = cdbl(oRs("selisih"))
			Else
				varRemainQty = 0		   
			End if
			oRs.Close
			Set oRs = Nothing
			
			'Cek Edited Qty
			oSql = " select nvl(sum (qty), 0) totalwipedit" & _
				   " from workinprogress" & _
				   " where fcba = '" & session("fcBA") & "'" & _
				   "       and agreementfinalno = '" & v_txtagreementfinalno & "'"
			
			Set oRs = Server.CreateObject("ADODB.Recordset")
			oRs.Open oSql, connstr
			if Not oRs.Eof then
				varEditedQty = cdbl(oRs("totalwipedit"))
			Else
				varEditedQty = 0		   
			End if
			oRs.Close
			Set oRs = Nothing
			
			if ((varRemainQty + varEditedQty) - Cdbl(To_Nol(qtyfinaledit)) ) >=  0 then
				connstr.BeginTrans
				Call Updatedatawipfinal
				if err.number=0 then
					connStr.CommitTrans
				End if
				cmdView	= "True"
				Response.Redirect "/wipfinal.asp?txtbappno=" & v_bappno &"&txtagreementcode=" & v_txtagreementcode & "&txtyear=" & v_txtyear & "&txtmonth=" & v_txtmonth & "&cmdView=True&SudahClose=Y"
			else
				cmdView	= "True"
				CompleteStringError = "Total Quantity yang anda masukan melebihi dari kontrak yang disepakati" & "<br/>"
				CompleteStringError = CompleteStringError & "&nbsp;&nbsp;&nbsp;&nbsp;Quantity Maximal-->" & (varRemainQty + varEditedQty)
			end if
		else
			CompleteStringError = "Tanggal """ & wipdate & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
			'Response.Redirect "/wipfinal.asp?txtagreementcode=" & v_txtagreementcode & "&txtyear=" & v_txtyear & "&txtmonth=" & v_txtmonth & "&cmdView=True&SudahClose=Y"
		End If
		cmdView	= "True"
	elseif Request("txtJenis")="Confirm"  then
		' Update isProsess True, dan update ke table lain
		Call insertdatawipfinal
	    cmdView	= "True"
    else
	
        cmdView	= "True"
	end if								 
	'Response.Write "sql " & q
else
    cmdView	= request("cmdView")		    
End if
else
  CompleteStringError = "Accounting Period  " & Yr & "/" &  mon & "  LOCKED !!!"
	cmdView	= request("cmdView")
End if

SudahClose = "N"
isprocessing = "FALSE"

set oRs = server.CreateObject("Adodb.recordset")
oSql = " SELECT  " & _
	   " bappno,year,month           " & _
	   " ,to_char(wipfinaldate,'MM/DD/YYYY') wipfinaldate " & _
	   " ,agreementcode    " & _
	   " , to_char(min(wipdate),'MM/DD/YYYY') tglawal " & _ 
	   " , to_char(max(wipdate),'MM/DD/YYYY') tglakhir,isprocessing " & _ 
	   " FROM workinprogressfinal_new" & _
       " where fcba = '" & Session("fcba") & "'" & _
	   "       and bappno='" & v_bappno & "'" & _
       "       /*and agreementcode = '" & v_txtagreementcode & "'*/" & _
       " /*AND wipdate between to_date('" & v_txttglawal & "','mm/dd/yyyy') and to_date('" & v_txttglakhir & "','mm/dd/yyyy')*/" & _ 
	   " GROUP BY fcba,year,month,BAPPNO,wipfinaldate,agreementcode,isprocessing "

oRs.Open oSql, connstr
if Not oRs.Eof then 
v_txtagreementcode = oRs("agreementcode")
wipdate            = oRs("wipfinaldate")
v_txttglawal       = oRs("tglawal")
v_txttglakhir      = oRs("tglakhir")
v_txtmonth         = oRs("month")
v_txtyear          = oRs("year")
isprocessing       = oRs("isprocessing")
SudahClose = "Y"
End if
oRs.Close
Set oRs = Nothing

If Len(Trim(CompleteStringError)) > 0 Then
	response.Write  "<div class = ""error"">" & CompleteStringError & "</div>"
End If


Dim val_IsTrue : val_IsTrue = "Y"

 %>
<% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-check-square-o" aria-hidden="true"></i>Header WIP Final</div><% end if %>
<form name="frmEntry" method="post" action="wipfinal.asp">
	<table class="data" width="100%">
		<tr>
			<td align="left">No.BAPP</td>
			<td align="left">
				
				<input type="text" maxlength="120" name="txtbappno" id="txtbappno" onkeyup="hurufbesar(this)" size="50" value="<%=v_bappno%>" <% if isprocessing="TRUE" then %>readonly<%end if%>  required>
				<font color="#FF0000">&nbsp;*</font>&nbsp;<%if isprocessing="FALSE" then response.write "PLANNED" else response.write "APPROVED" End IF%>
			</td>
		</tr>
		<tr>
			<td align="left">Tgl.BAPP</td>
			<td align="left">
			    <input type="text" class="datepicker" name="wipdate" id="wipdate" value="<%=wipdate%>" > 
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
		<tr>
			<td align="left">No.Kontrak</td>
			<td align="left">
			    <select name="txtagreementcode" id="txtagreementcode" style="width:200px"  >
									<option></option>
									<%
									if trim(v_txtagreementcode) <> "" then
									qLov =  " SELECT agreementcode fccode,agreementcode ||' ('||b.fcname||')' fcname " & _
											" FROM contractagreement a INNER JOIN contractor b ON (a.contractorcode=b.fccode AND a.fcba=b.fcba) " & _
											" WHERE a.fcba='" & Session("fcBA") & "' " & _
											" and agreementcode='" & v_txtagreementcode & "'" & _
											" ORDER BY b.fcname,agreementcode " 
									else
									qLov =  " SELECT agreementcode fccode,agreementcode ||' ('||b.fcname||')' fcname " & _
											" FROM contractagreement a INNER JOIN contractor b ON (a.contractorcode=b.fccode AND a.fcba=b.fcba) " & _
											" WHERE a.fcba='" & Session("fcBA") & "' " & _
											" ORDER BY agreementcode,b.fcname " 
									
									end if

						   
									Set rsLov = Server.CreateObject("ADODB.Recordset")
									rsLov.Open qLov, connstr
									if not rsLov.eof then
										while not rsLov.eof
									%>
									<option value="<%=rsLov("fccode")%>" <% if cstr(""&rsLov("fccode")) = cstr(""&v_txtagreementcode) then%> selected <%end if%>><%=rsLov("fcname")%></option>
									<%
									rsLov.movenext
									wend
									end if
									%>
				</select>
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
		
		<tr>
			<td align="left">Periode WIP.Daily</td>
			<td align="left">
				<input type="text" class="datepicker" name="tglawal" id="tglawal" value="<%=v_txttglawal%>" <% if isprocessing="TRUE" then %>readonly<%end if%> > s.d <input type="text" class="datepicker" name="tglakhir" id="tglakhir" value="<%=v_txttglakhir%>" <% if isprocessing="TRUE" then %>readonly<%end if%> >
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
		<tr>
			<td colspan="2"  align="left" class="fbp-doc-actions"> 
				<input Type="Hidden" Value="True" Name="cmdView">  
				<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">  
				<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">   	 
				<input Type="Submit" name="Add" value="Select">&nbsp;&nbsp;  
				<a href="print_wipfinal.asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&txtagreementcode=<%=v_txtagreementcode%>" target="_blank"><img src="IMAGES/Print.gif" border="0" title="Print Preview"><font face="Verdana" size="1">Print</font></a>&nbsp;&nbsp;
				<%If val_IsTrue = "Y" and (isprocessing= "FALSE" or isprocessing="" ) Then%>
					<a href="wipfinal.asp?txtbappno=<%=v_bappno%>&txtagreementcode=<%=v_txtagreementcode%>&txtyear=<%=v_txtyear%>&txtmonth=<%=v_txtmonth%>&wipdate=<%=wipdate%>&txtJenis=Header&tglawal=<%=v_txttglawal%>&tglakhir=<%=v_txttglakhir%>"><img src="IMAGES/delete.gif" alt="Delete Data : <%=v_bappno%>" border="0" onClick="return confirm('Are you sure to delete');"></a>
				<%Else%>
					<img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18">
				<%End If%>
				
				<%If trim(v_bappno) <> "" and isprocessing="FALSE" Then %>
						<a href="wipfinal.asp?txtbappno=<%=v_bappno%>&txtagreementcode=<%=v_txtagreementcode%>&txtyear=<%=v_txtyear%>&txtmonth=<%=v_txtmonth%>&wipdate=<%=wipdate%>&txtJenis=Confirm&tglawal=<%=v_txttglawal%>&tglakhir=<%=v_txttglakhir%>" onClick="return confirm('Are you sure to Confirm this data :  <%=v_bappno%> ?');">Confirm</a>
				<%End If%>
				
				<!-- HADIR -->
				<%
				dim dataModule,dataSubmodule,dataDetail
				dataModule 		= "Contract"
				dataSubmodule 	= "Work in Progress Final"
				dataDetail 		= "NULL"
				%>
				
				
				
				<a class="jqplot-image-button" style="" href="ajax/hadirAjax.asp?dataModule=<%=dataModule%>&dataSubmodule=<%=dataSubmodule%>&dataDetail=<%=dataDetail%>"><i class="fa fa-check-square-o"></i>&nbsp;Hadir</a>
				<!-- END:HADIR -->
			</td> 
		</tr> 
	</table>    
</form>
<%
If cmdView = "True" and Len(Trim(v_txtagreementcode)) > 0 Then
	Dim rswipfinal, intRecordsPerPage, cColor
	set rswipfinal = server.CreateObject("Adodb.recordset")
	If SudahClose <> "Y" Then
		q	= " select wip.fcba" & _
              "        , wip.agreementcode" & _
			  "        , to_char(wip.wipdate,'mm/dd/yyyy') wipdate" & _
              "        , wip.agreementlineno" & _
              "        , wip.jobcode" & _
              "        , jobmaster.fcname jobname" & _
			  "        , ctk.remarks ctkjobname " & _
			  "        , ctk.qty ctkqty " & _
			  "        , ctk.rate ctkrate " & _
			  "        , wip.chargeunit" & _
              "        , wip.locationtype targettype" & _
              "        , wip.locationcode targetcode" & _
              "        , nvl(sum(wip.qty), 0) qtyplan" & _
			  "        , nvl(sum(wip.unit), 0) unitplan" & _
              "        , nvl(sum(wip.output), 0) outputplan" & _
              " from workinprogressplan wip" & _
			  " inner join contractagreementdetail ctk on (wip.agreementcode=ctk.agreementcode and wip.fcba=ctk.fcba and wip.agreementlineno=ctk.lineno)" & _
              " left join jobmaster on (jobmaster.fccode = wip.jobcode)" & _
              " where wip.fcba = '" & Session("fcba") & "'" & _
              "       and wip.agreementcode = '" & v_txtagreementcode & "'" & _
			  "       and  isgenerated IS NULL " & _
              " AND wip.wipdate between to_date('" & v_txttglawal & "','mm/dd/yyyy') and to_date('" & v_txttglakhir & "','mm/dd/yyyy')  AND wip.isfinal=0 " & _
              " group by wip.fcba" & _
              "          , wip.agreementcode" & _
              "          , wip.agreementlineno" & _
			  "          , wip.jobcode" & _
              "          , jobmaster.fcname" & _
			  "			 , ctk.remarks, ctk.qty, ctk.rate " & _
			  "          , wip.chargeunit" & _
              "          , wip.locationtype" & _
              "          , wip.locationcode" & _
              "          , wip.wipdate" & _
			  " ORDER BY wip.fcba,wip.agreementlineno,wip.wipdate "

	Else	 
	      q = " select wipf.id,wipf.fcba,wipf.bappno " & _
			  "        , wipf.agreementcode,to_char(wipf.wipfinaldate,'mm/dd/yyyy') wipfinaldate ,wipf.Year,wipf.month,wipf.agreementfinalno" & _
			  "        , to_char(wipf.wipdate,'mm/dd/yyyy') wipdate" & _
              "        , wipf.agreementlineno" & _
              "        , wipf.jobcode" & _
              "        , jobmaster.fcname jobname" & _
			  "        , ctk.remarks ctkjobname " & _
			  "        , ctk.qty ctkqty " & _
			  "        , wipf.rate_basic ctkrate " & _
			  "        , wipf.chargeunit" & _
              "        , wipf.locationtype targettype" & _
              "        , wipf.locationcode targetcode" & _
              "        , nvl(wipf.qty_plan, 0) qtyplan" & _
			  "        , nvl(wipf.unit_plan, 0) unitplan" & _
              "        , nvl(wipf.output_plan, 0) outputplan" & _
			  "        , nvl(wipf.qty_final, 0) qtyfinal" & _
			  "        , nvl(wipf.unit_final, 0) unitfinal" & _
              "        , nvl(wipf.output_final, 0) outputfinal" & _
			  "        , nvl(wipf.rate_final, 0) ratefinal" & _
			  "        , (wipf.qty_final * wipf.rate_final) amnt " & _
			  " from workinprogressfinal_new wipf" & _
			  " inner join contractagreementdetail ctk on (wipf.agreementcode=ctk.agreementcode and wipf.fcba=ctk.fcba and wipf.agreementlineno=ctk.lineno)" & _
              " left join jobmaster on (jobmaster.fccode = wipf.jobcode)" & _
			  " where " & _
			  " bappno='" & v_bappno & "' " & _ 
			  " Order by ID "
			  
		
   End if

   If Request("txtJenis")="Edit" Then
		q=  " select " & _ 
			" a.ID,              " & _ 
			" a.BAPPNO,          " & _ 
			" a.WIPFINALDATE,    " & _ 
			" a.YEAR,            " & _ 
			" a.MONTH,           " & _ 
			" a.AGREEMENTFINALNO," & _ 
			" a.AGREEMENTCODE,   " & _ 
			" a.WIPDATE,         " & _ 
			" a.AGREEMENTLINENO, " & _ 
			" a.CHARGEUNIT,      " & _ 
			" a.JOBCODE,         " & _ 
			" a.LOCATIONTYPE,    " & _ 
			" a.LOCATIONCODE,    " & _ 
			" a.RATE_BASIC,      " & _ 
			" a.QTY_PLAN,        " & _ 
			" a.OUTPUT_PLAN,     " & _ 
			" a.UNIT_PLAN,       " & _ 
			" a.QTY_FINAL,       " & _ 
			" a.OUTPUT_FINAL,    " & _ 
			" a.UNIT_FINAL,      " & _ 
			" a.RATE_FINAL,      " & _ 
			" a.ISPROCESSING,    " & _ 
			" a.FCBA,            " & _ 
			" a.ISINVOICE,        " & _
			" Jobmaster.fcname jobname," & _
			" ctk.remarks ctkjobname, " & _
			" ctk.qty ctkqty, " & _
			" ctk.rate ctkrate, " & _
			" (a.Unit_FINAL * a.RATE_FINAL) amnt " & _
			" FROM workinprogressfinal_new a" & _ 
			" inner join contractagreementdetail ctk on (a.agreementcode=ctk.agreementcode and a.fcba=ctk.fcba and a.agreementlineno=ctk.lineno)" & _
            " left join jobmaster on (jobmaster.fccode = a.jobcode)" & _
			" WHERE " & _
			"       a.fcba = '" & Session("fcba") & "'" & _
			"   and a.bappno = '" & v_bappno & "'" & _
			"   and a.agreementcode = '" & v_txtagreementcode & "'" & _
			"   and a.agreementfinalno = '" & v_txtagreementfinalno & "'" & _ 
			"   and a.id = " & request("id") & "" & _
			" ORDER By ID "
			
	
	End If
			  
  ' response.Write q	
   rswipfinal.Open q, connstr%>
   <% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-edit" aria-hidden="true"></i>Detail Item</div><% end if %>
   <form action="wipfinal.asp" method="post" name="eForm">
   <table class="data" border="1" width="100%" >
   <%if SudahClose <> "Y" Then%>
	   <tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;">
		   <td align="center">No.</td>
		   <td align="center">Line No</td>
		   <td align="center">Tanggal</td>
		   <td align="center">Charge.Unit</td>
		   <td align="center">Ctk.Remark</td>		
		   <td align="center">Ctk.Qty</td>	
		   <td align="center">Ctk.Rate</td>	
		   <td align="center">Rate (Final)</td>	  
		   <td align="center">Jobcode</td>
		   <td align="center">Jobname</td>		
		   <td align="center">Target Type</td>
		   <td align="center">Target Code</td>
		   <td align="center">Quantity (Plan)</td>
		   <td align="center">Unit (Plan)</td>
		   <td align="center">Output.Ha (Plan)</td>
		   <td align="center">Quantity (Final)</td>
		   <td align="center">Unit (Final)</td>
		   <td align="center">Output.Ha (Final)</td>
		  
		   
   	   </tr>
   <%Else%>
		<%if Request("txtJenis")="Edit" Then%>
   
			<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;">
			    <td align="center">No.</td>
				<td align="center">Line No</td>
				<td align="center">Tanggal</td>
				<td align="center">Charge Unit</td>
				<td align="center">Ctk.Remark</td>		
			    <td align="center">Ctk.Qty</td>	
			    <td align="center">Ctk.Rate</td>	
				<td align="center">Rate (Final)</td>
				<td align="center">Jobcode</td>
				<td align="center">Jobname</td>
				<td align="center">Target Type</td>
				<td align="center">Target Code</td>
				<td align="center">Quantity (Plan)</td>
				<td align="center">Unit (Plan)</td>
				<td align="center">Output.Ha (Plan)</td>
				<td align="center">Quantity (Final)</td>
				<td align="center">Unit (Final)</td>
				<td align="center">Output.Ha (Final)</td>
		   </tr>
		<%else%>
	   
			<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;">
				<td align="center" colspan="2">&nbsp;</td>
				<td align="center">No.</td>
				<td align="center">Line No</td>
				<td align="center">Tanggal</td>
				<td align="center">Charge Unit</td>
				<td align="center">Ctk.Remark</td>		
			    <td align="center">Ctk.Qty</td>	
			    <td align="center">Ctk.Rate</td>
				<td align="center">Rate (Final)</td>
				<td align="center">Jobcode</td>
				<td align="center">Jobname</td>
				<td align="center">Target Type</td>
				<td align="center">Target Code</td>
				<td align="center">Quantity (Plan)</td>
				<td align="center">Unit (Plan)</td>
				<td align="center">Output.Ha (Plan)</td>
				<td align="center">Quantity (Final)</td>
				<td align="center">Unit (Final)</td>
				<td align="center">Output.Ha (Final)</td>
				<td align="center">Amount</td>
			</tr>
		<%End if%>

   <%End if
   if Not rswipfinal.eof then
      nCountofField=rswipfinal.fields.count-1
	  no = 1
	  While Not rswipfinal.EOF

		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if%>
			<tr bgcolor="<%=cColor%>">
				<%IF SudahClose = "Y" then%>

					<%if Request("txtJenis")="Edit" Then%>
						<td align="left" title="No" valign="top"><%=rswipfinal("id")%></td>
						<td align="left" title="Agreement Line No" valign="top"><%=rswipfinal("agreementlineno")%></td>
						<td align="left" title="Tanggal" valign="top"><%=rswipfinal("wipdate")%></td>
						<td align="left" title="Chargeunit" valign="top"><%=rswipfinal("chargeunit")%></td>
						<td align="left" title="ctkjobname" valign="top"><%=rswipfinal("ctkjobname")%></td>
						<td align="Right" title="ctkqty" valign="top"><%=FbpNum2(rswipfinal("ctkqty"))%></td>
						<td align="Right" title="ctkrate" valign="top"><%=FbpNum2(rswipfinal("ctkrate"))%></td>
						<td align="left" >
							<input type="text" name="ratefinaledit"  value="<%=rswipfinal("rate_final")%>" onkeyup="checkInput(this);">
						</td>
						<td align="left" title="Job Code" valign="top"><%=rswipfinal("jobcode")%></td>
						<td align="left" title="Job Name" valign="top"><%=rswipfinal("jobname")%></td>
						<td align="left" title="Target Type" valign="top"><%=rswipfinal("LocationType")%></td>
						<td align="left" title="Target Code" valign="top"><%=rswipfinal("LocationCode")%></td>
						<td align="Right" title="Qty Plan" valign="top"><%=FbpNum2(rswipfinal("qty_plan"))%></td>
						<td align="Right" title="Unit Plan" valign="top"><%=FbpNum2(rswipfinal("unit_plan"))%></td>
						<td align="Right" title="Output Plan" valign="top"><%=FbpNum2(rswipfinal("output_plan"))%></td>
						<td align="left" width="2">
							<input type="text" name="qtyfinaledit" width="2" value="<%=rswipfinal("qty_final")%>" onkeyup="checkInput(this);">
						</td>
						<td align="left" width="2">
							<input type="text" name="unitfinaledit" width="2" value="<%=rswipfinal("unit_final")%>" onkeyup="checkInput(this);">
						</td>
						<td align="left" width="2">
							<input type="text" name="outputfinaledit" width="2" value="<%=rswipfinal("output_final")%>" onkeyup="checkInput(this);">
						</td>
						<input type="hidden" name="txtagreementfinalno" value="<%=rswipfinal("agreementfinalno")%>">
						<input type="hidden" name="txtagreementlineno" value="<%=rswipfinal("agreementlineno")%>">
						<input type="hidden" name="qtyplanedit" value="<%=rswipfinal("qty_plan")%>">
						<input type="hidden" name="unitplanedit" value="<%=rswipfinal("unit_plan")%>">
						<input type="hidden" name="outputplanedit" value="<%=rswipfinal("output_plan")%>">
						<input type="hidden" name="chargeunitedit" value="<%=rswipfinal("chargeunit")%>">
						<input type="hidden" name="jobcodeedit" value="<%=rswipfinal("jobcode")%>">
						<input type="hidden" name="targettypeedit" value="<%=rswipfinal("LocationType")%>">
						<input type="hidden" name="targetcodeedit" value="<%=rswipfinal("LocationCode")%>">
						<input type="hidden" name="ID" value="<%=rswipfinal("id")%>">
						
					<%Else%>
				
						<%If isprocessing = "FALSE" Then%>
						<td valign="middle" width="10">
							<a href="wipfinal.asp?ID=<%=rswipfinal("id")%>&txtbappno=<%=rswipfinal("bappno")%>&txtagreementcode=<%=rswipfinal("agreementcode")%>&txtagreementfinalno=<%=rswipfinal("agreementfinalno")%>&txtyear=<%=rswipfinal("year")%>&txtmonth=<%=rswipfinal("month")%>&txtJenis=Edit">
							<img src="IMAGES/edit.gif" alt="Edit Data : <%=rswipfinal(1)%>" border="0"></a>
						</td>
						<td valign="middle" width="10">
							<a href="wipfinal.asp?ID==<%=rswipfinal("id")%>&txtbappno=<%=rswipfinal("bappno")%>&txtagreementcode=<%=rswipfinal("agreementcode")%>&txtagreementfinalno=<%=rswipfinal("agreementfinalno")%>&txtyear=<%=rswipfinal("year")%>&txtmonth=<%=rswipfinal("month")%>&wipdate=<%=wipdate%>&txtJenis=Details" onClick="return confirm('Are you sure to delete');">
							<img src="IMAGES/delete.gif" alt="Delete Data : <%=rswipfinal(1)%>" border="0"></a>
						</td>
						<%Else%>
						<td colspan="2">
							<img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18">
						</td>
						<%End If%>
						
						<td align="left" valign="top"><%=rswipfinal("id")%></td>
						<td align="left" title="agreementlineno" valign="top"><%=rswipfinal("agreementlineno")%></td>
						<td align="left" title="wipdate" valign="top"><%=rswipfinal("wipdate")%></td>
						<td align="left" title="chargeunit" valign="top"><%=rswipfinal("chargeunit")%></td>
						<td align="left" title="ctkjobname" valign="top"><%=rswipfinal("ctkjobname")%></td>
						<td align="Right" title="ctkqty" valign="top"><%=FbpNum2(rswipfinal("ctkqty"))%></td>
						<td align="Right" title="rate_basic" valign="top"><%=FbpNum2(rswipfinal("ctkrate"))%></td>
						<td align="Right" title="rate_final" valign="top"><%=FbpNum2(rswipfinal("ratefinal"))%></td>
						<td align="left" title="Job Code" valign="top"><%=rswipfinal("jobcode")%></td>
						<td align="left" title="Job Name" valign="top"><%=rswipfinal("jobname")%></td>
						<td align="left" title="Target Type" valign="top"><%=rswipfinal("targettype")%></td>
						<td align="left" title="Target Code" valign="top"><%=rswipfinal("targetcode")%></td>
						<td align="Right" title="Quantity Plan" valign="top"><%=FbpNum2(rswipfinal("qtyplan"))%></td>
						<td align="Right" title="Unit Plan" valign="top"><%=FbpNum2(rswipfinal("unitplan"))%></td>
						<td align="Right" title="Output Plan" valign="top"><%=FbpNum2(rswipfinal("outputplan"))%></td>
						<td align="Right" title="Quantity Plan" valign="top"><%=FbpNum2(rswipfinal("qtyfinal"))%></td>
						<td align="Right" title="Unit Plan" valign="top"><%=FbpNum2(rswipfinal("unitfinal"))%></td>
						<td align="Right" title="Output Plan" valign="top"><%=FbpNum2(rswipfinal("outputfinal"))%></td>
						<td align="Right" title="Output Plan" valign="top"><%=FbpNum2(rswipfinal("amnt"))%></td>
					
					<%End If%>
					
				<%End if%>
				
				<%if SudahClose <> "Y" Then%>
					<td align="left" title="No" valign="top"><%=no%></td>
					<td align="left" title="Agreement Line No" valign="top"><%=rswipfinal("agreementlineno")%></td>
					<td align="left" title="Tanggal" valign="top"><%=rswipfinal("wipdate")%></td>
					<td align="left" title="Chargeunit" valign="top"><%=rswipfinal("chargeunit")%></td>
					
					<td align="left" title="ctkjobname" valign="top"><%=rswipfinal("ctkjobname")%></td>
					<td align="Right" title="ctkqty" valign="top"><%=FbpNum2(rswipfinal("ctkqty"))%></td>
					<td align="Right" title="ctkrate" valign="top"><%=FbpNum2(rswipfinal("ctkrate"))%></td>
					<td align="left" >
						<input type="text" name="ratefinal_<%=no%>"  value="<%=rswipfinal("ctkrate")%>" onkeyup="checkInput(this);">
					</td>
				
					<td align="left" title="Job Code" valign="top"><%=rswipfinal("jobcode")%></td>
					<td align="left" title="Job Name" valign="top"><%=rswipfinal("jobname")%></td>
					<td align="left" title="Target Type" valign="top"><%=rswipfinal("targettype")%></td>
					<td align="left" title="Target Code" valign="top"><%=rswipfinal("targetcode")%></td>
					<td align="Right" title="Quantity Plan" valign="top"><%=FbpNum2(rswipfinal("qtyplan"))%></td>
					<td align="Right" title="Unit Plan" valign="top"><%=FbpNum2(rswipfinal("unitplan"))%></td>
					<td align="Right" title="Output Plan" valign="top"><%=FbpNum2(rswipfinal("outputplan"))%></td>
					<td align="left" >
						<input type="text" name="qtyfinal_<%=no%>"  value="<%=rswipfinal("qtyplan")%>" onkeyup="checkInput(this);">
					</td>
					<td align="left" >
						<input type="text" name="unitfinal_<%=no%>"  value="<%=rswipfinal("unitplan")%>" onkeyup="checkInput(this);">
					</td>
					<td align="left" >
						<input type="text" name="outputfinal_<%=no%>" value="<%=rswipfinal("outputplan")%>" onkeyup="checkInput(this);">
					</td>
					
					
						<input type="hidden" name="wipdate_<%=no%>" value="<%=rswipfinal("wipdate")%>">
						<input type="hidden" name="agreementlineno_<%=no%>" value="<%=rswipfinal("agreementlineno")%>">
						<input type="hidden" name="ctkrate_<%=no%>" value="<%=rswipfinal("ctkrate")%>">
						<input type="hidden" name="qtyplan_<%=no%>" value="<%=rswipfinal("qtyplan")%>">
						<input type="hidden" name="unitplan_<%=no%>" value="<%=rswipfinal("unitplan")%>">
						<input type="hidden" name="outputplan_<%=no%>" value="<%=rswipfinal("outputplan")%>">
						<input type="hidden" name="chargeunit_<%=no%>" value="<%=rswipfinal("chargeunit")%>">
						<input type="hidden" name="jobcode_<%=no%>" value="<%=rswipfinal("jobcode")%>">
						<input type="hidden" name="targettype_<%=no%>" value="<%=rswipfinal("targettype")%>">
						<input type="hidden" name="targetcode_<%=no%>" value="<%=rswipfinal("targetcode")%>">
				<%End if%>
		  </tr>
		  <%
		  rswipfinal.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no = no + 1
	  Wend
   else
      response.Write("Data is empty") 	  
   end if 
   rswipfinal.Close
   set rswipfinal = Nothing%>
		<tr>
			<td colspan="12" align="left">
			<!-- diambil dari form yang diatas -->
			<input type="hidden" name="txtbappno" value="<%=v_bappno%>">
			<input type="hidden" name="txtagreementcode" value="<%=v_txtagreementcode%>">
			<input type="hidden" name="txtyear" value="<%=v_txtyear%>">
			<input type="hidden" name="tglawal" value="<%=v_txttglawal%>">
			<input type="hidden" name="tglakhir" value="<%=v_txttglakhir%>">
			<input type="hidden" name="txtmonth" value="<%=v_txtmonth%>">
			<input type="hidden" name="wipdate" value="<%=wipdate%>">
			<input type="hidden" name="No" value="<%=No%>">
			<input type="hidden" name="fcba" id="fcbaData" value="<%=Session("fcBa")%>">
			<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
			<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
			<%If SudahClose = "Y" Then%>
				<input type="hidden" name="txtJenis" value="Edit">
				<input Type="Submit"  name="Add" value="Update">
			<%End if%>
			
			<%If SudahClose <> "Y" Then%>
				<input type="hidden" name="txtJenis" value="New">
				<input Type="Submit"  name="Add" value="Submit">
			<%End if%>
			<input Type="Reset"  name="Reset" value="Reset">
			<input Type="Button"  value="Back" name="B2" OnClick="javascript:history.back(1)"></td>
		</tr>
	
	</TABLE>
</form>
<%End if%>
<!---------------akhir!--------------->
</div>
</div>
<!--#include file="bottom.jsp"-->
</body>
</html>
<script>
	$("#wipdate").on("change", function(e) {
		var wipDateVal = $("#wipdate").val(); // format MM/DD/YYYY
		var fcbaData = $("#fcbaData").val();
		$.post("cektanggalwip.asp?fcba=" + fcbaData + "&wipdate=" + wipDateVal, function(data) {
			var json = JSON.parse(data);
			var jumlview = json.results.length;

			if (jumlview > 0) {
				var result = json.results[0];

				// Parse tanggal dari format DD/MM/YYYY (dari JSON)
				function parseDMY(str) {
					var parts = str.split("/");
					return new Date(parts[2], parts[1] - 1, parts[0]);
				}

				// Parse tanggal dari format MM/DD/YYYY (dari input)
				function parseMDY(str) {
					var parts = str.split("/");
					return new Date(parts[2], parts[0] - 1, parts[1]);
				}

				var startDate    = parseDMY(result.start_date);
				var endDate      = parseDMY(result.end_date);
				var selectedDate = parseMDY(wipDateVal); // <-- pakai parseMDY

				if (selectedDate >= startDate && selectedDate <= endDate) {
					console.log("Tanggal valid, boleh input.");
				} else {
					alert("Tanggal yang dipilih diluar rentang yang diizinkan!");
					$("#wipdate").val("");
				}
			}
		});
	});

	function hurufbesar(obj) {
		var strval	= $(obj).val();
		strval		= strval.toUpperCase();
		strval		= strval.replace(/\s/g,"");
		$(obj).val(strval);
	}
</script>