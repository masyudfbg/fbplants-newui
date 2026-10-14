<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->  
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>ePlaS - Print Stock On Hand</title>
<link rel="stylesheet" href="css/stylereport2.css" type="text/css" />
<link rel="stylesheet" href="css/jquery-ui.css" />
<script type="text/javascript" language="JavaScript1.2" src="stlib.js"></script>
<script language="JavaScript" src='js/x.js'></script>
<script language="JavaScript" src='js/xtableheaderfixed.js'></script>
        <script language="JavaScript">
        xAddEventListener(window, 'load',
          function() {
            new xTableHeaderFixed('xthf-site', window);
            new xTableHeaderFixed('xthf-green', window);
            new xTableHeaderFixed('xthf-blue', window);
          }, false
        );
        </script>
    
  <script language="JavaScript" src="js/jquery.printPage.js"></script> 
  <script>
          function toggle(id){
            var e = document.getElementById(id);
            if(e.style.display == 'inline-block'){
              e.style.display = 'none';
            }else if(e.style.display == 'none'){
              e.style.display = 'inline-block';
            }
          }
  </script>
  <script>
   function printDiv(divName) {
   var printContents = document.getElementById(divName).innerHTML;
     var originalContents = document.body.innerHTML;

     document.body.innerHTML = printContents;

     window.print();

     document.body.innerHTML = originalContents;
}
</script>
<style type="text/css">
@media all {
#product-table td{padding:5px;}
.page-break { display: block; }
a.button_print { display:block; } 
#wrapper{background: #e2e2e2; /* Old browsers */
background: -moz-linear-gradient(-45deg,  #e2e2e2 0%, #dbdbdb 50%, #d1d1d1 51%, #fefefe 100%); /* FF3.6+ */
background: -webkit-gradient(linear, left top, right bottom, color-stop(0%,#e2e2e2), color-stop(50%,#dbdbdb), color-stop(51%,#d1d1d1), color-stop(100%,#fefefe)); /* Chrome,Safari4+ */
background: -webkit-linear-gradient(-45deg,  #e2e2e2 0%,#dbdbdb 50%,#d1d1d1 51%,#fefefe 100%); /* Chrome10+,Safari5.1+ */
background: -o-linear-gradient(-45deg,  #e2e2e2 0%,#dbdbdb 50%,#d1d1d1 51%,#fefefe 100%); /* Opera 11.10+ */
background: -ms-linear-gradient(-45deg,  #e2e2e2 0%,#dbdbdb 50%,#d1d1d1 51%,#fefefe 100%); /* IE10+ */
background: linear-gradient(135deg,  #e2e2e2 0%,#dbdbdb 50%,#d1d1d1 51%,#fefefe 100%); /* W3C */
filter: progid:DXImageTransform.Microsoft.gradient( startColorstr='#e2e2e2', endColorstr='#fefefe',GradientType=1 ); /* IE6-9 fallback on horizontal gradient */
}
a.button_print {display: inline-block;}
.letter { 
background-color:#fff; 
box-shadow:0px 0px 4px rgba(0,0,0,0.4); 
border-radius:4px;
margin: auto;
width: 1200px;
padding-top:30px;
padding-bottom:30px;
padding-left:30px;
padding-right:30px;
}
#header-wrapper {display:block} 
} 
</style>    
</head>

<body>
<%
cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"

  v_txtFcba = Request("fcba") 
  v_txtACCOUNTINGYEARPERIOD=trim(request("ACCOUNTINGYEARPERIOD"))
  invoicecode = trim(request("invoicecode")) 

  varDefaultBa = Session("fcBA")   
  v_txtdatefrom = Request("txtdatefrom")
  v_txtdateuntil = Request("txtdateuntil")
  If (v_txtdatefrom = "" Or Len(Trim(v_txtdatefrom)) = 0) Then v_txtdatefrom = textdate(date)
  If (v_txtdateuntil = "" Or Len(Trim(v_txtdateuntil)) = 0) Then v_txtdateuntil = textdate(date)

  If "" & txttdate = "" Then txttdate = textDate(date)
  txttdate = txtdatefrom & "" & txtdateuntil

  If (v_txtFcba = "" Or Len(Trim(v_txtFcba)) = 0) Then
    v_txtFcba = varDefaultBa
  End If  
    
    if  v_txtfcba <> "" then
       whereCons = whereCons & "  AND fp.fcba = '" & v_txtfcba & "'"
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

  Dim tglbulanprev,bulanprevsplit,tglprevsplit,thnprevsplit,v_txthi,strStockOnHand,rsGrnInv
  dim openingqty,receivenote,issue,transportask,receivetask,qtyonhand
  Dim bulannow,tahunnow

  v_txthi         = date()
  IssueDate       = Request("IssueDate")
  
  'tglbulanprev    = DateAdd("m", -1, IssueDate)
  tglbulanprev    = IssueDate
  
  bulanprevsplit  = month(tglbulanprev)
  tglprevsplit    = day(tglbulanprev)
  thnprevsplit    = year(tglbulanprev)
  
  bulannow        = left(IssueDate,2)
  tahunnow        = right(IssueDate,4)
  tglawal 		    = bulannow & "/" & "01" & "/" & tahunnow


	'Response.Write tglbulanprev

	
	strStockOnHand		=	"select t.storecode,t.storename,t.itemcode,t.stockname,t.uom,sum(nvl(t.openingbalanceqty,0)) as openingbalanceqty,sum(nvl(t.receivenoteqty,0)) as receivenoteqty, " & _
" sum(nvl(t.issueqty,0)) as issueqty,sum(nvl(t.quantitytransfered,0)) as quantitytransfered,sum(nvl(t.quantityreceived,0)) as quantityreceived,sum(nvl(t.qtywriteoff,0)) as qtywriteoff, " & _
" (sum(nvl(t.openingbalanceqty,0))+sum(nvl(t.receivenoteqty,0))-sum(nvl(t.issueqty,0))-sum(nvl(t.quantitytransfered,0))+sum(nvl(t.quantityreceived,0))+sum(nvl(t.qtywriteoff,0))) as qty " & _
" from( " & _
    " select l.STORECODE,store_api.get_fcname(storecode, fcba) storename,l.itemcode, " & _
    " stock_api.get_stockname(itemcode, fcba) stockname,stock_api.get_uom(itemcode, fcba) uom,l.openingbalanceqty, " & _
    " 0 as receivenoteqty,0 as issueqty,0 as quantitytransfered,0 as quantityreceived,0 as qtywriteoff " & _
    " from stockledger l " & _
    " where l.YEAR='" & thnprevsplit & "' and l.MONTH='" & bulanprevsplit & "' and l.fcba='" & v_txtfcba & "' " & _
    "      union all " & _
    " select r.storecode,store_api.get_fcname(r.storecode, r.fcba) storename,rd.purchaseitemcode, " & _
    " stock_api.get_stockname(rd.purchaseitemcode, r.fcba) stockname, stock_api.get_uom(rd.purchaseitemcode, r.fcba) uom, " & _
    " 0 as openingbalanceqty,rd.quantity,0 as issueqty,0 as quantitytransfered,0 as quantityreceived,0 as qtywriteoff " & _
    " from receivenote r,receivenotedetail rd " & _
    " where r.receivenotecode=rd.receivenotecode and r.rndate between to_date('" & tglawal & "','mm/dd/yyyy') and to_date('" & IssueDate & "','mm/dd/yyyy') " & _
    " and r.fcba='" & v_txtfcba & "' " & _
    "      union all " & _
    " select h.storecode,store_api.get_fcname(h.storecode, h.fcba) storename,d.STOCKCODE,stock_api.get_stockname(d.STOCKCODE, d.fcba) stockname, " & _
    " stock_api.get_uom(d.STOCKCODE, d.fcba) uom,0 as openingbalanceqty,0 as receivenoteqty,d.quantity,0 as quantitytransfered, " & _
    " 0 as quantityreceived,0 as qtywriteoff " & _
    " from issue h,issuedetails d " & _
    " where h.sivcode=d.sivcode and h.fcba='" & v_txtfcba & "' and h.sivdate between to_date('" & tglawal & "','mm/dd/yyyy') and to_date('" & IssueDate & "','mm/dd/yyyy') " & _
    "      union all " & _
    " select d.storecode,store_api.get_fcname(d.storecode, d.fcba) storename,d.STOCKCODE,stock_api.get_stockname(d.STOCKCODE, d.fcba) stockname, " & _
    " stock_api.get_uom(d.STOCKCODE, d.fcba) uom,0 as openingbalanceqty,0 as receivenoteqty,0 as issueqty,d.quantitytransfered,0 as quantityreceived,0 as qtywriteoff " & _
    " from transporttaskdtl_tab d,transporttask_tab h " & _
    " where h.orderno=d.orderno and h.fcba='" & v_txtfcba & "' AND h.orderdate between to_date('" & tglawal & "','mm/dd/yyyy') and to_date('" & IssueDate & "','mm/dd/yyyy') " & _
    "     union all " & _
    " select d.locationcode,store_api.get_fcname(d.locationcode, d.fcba) storename,d.STOCKCODE,stock_api.get_stockname(d.STOCKCODE, d.fcba) stockname, " & _
    " stock_api.get_uom(d.STOCKCODE, d.fcba) uom,0 as openingbalanceqty,0 as receivenoteqty,0 as issueqty,0 as quantitytransfered,d.quantityreceived,0 as qtywriteoff  " & _
    " from receivedtask_tab h,receivedtaskdtl_tab d " & _
    " where h.orderno=d.orderno and h.orderdate between to_date('" & tglawal & "','mm/dd/yyyy') and to_date('" & IssueDate & "','mm/dd/yyyy')  and h.fcba='" & v_txtfcba & "' " & _
    "  UNION ALL " & _
    " select h.STORECODE locationcode,store_api.get_fcname(d.locationcode, d.fcba) storename,d.stockcode,stock_api.get_stockname(d.STOCKCODE, d.fcba) stockname, " & _
    " stock_api.get_uom(d.STOCKCODE, d.fcba) uom,0 as openingbalanceqty,0 as receivenoteqty,0 as issueqty,0 as quantitytransfered, " & _
    " 0 as quantityreceived,d.quantity as qtywriteoff  " & _
    " from writeoff h,writeoffdetails d " & _
    " where h.writeoffcode=d.writeoffcode and h.wodate between to_date('" & tglawal & "','mm/dd/yyyy') and to_date('" & IssueDate & "','mm/dd/yyyy') and h.fcba='" & v_txtfcba & "' " & _ 
" )t " & _
" group by t.storecode,t.storename,t.itemcode,t.stockname,t.uom order by storecode" 


 'response.write strStockOnHand  
  Set rsGrnInv = Server.CreateObject("ADODB.Recordset")
  rsGrnInv.Open strStockOnHand, connstr, adOpenStatic 

%>
<div id="wrapper">
<div class="navi_to_print">

</div>
<div class="letter">
<div id="note">
<style>
@media all {
td{padding:5px;}
table.PAYSLIP td{padding:5px;}
.page-break hr {display:block;border: 1px dashed red;}
.PAYSLIP td {font-family:verdana,sans-serif; font-size:8px; padding:0px;}
#note {margin-left:auto;padding:5px;}
#info_note .judul { text-align:center; font-size:16px;padding-top:0px;padding-bottom:0px;}
#info_note .left { float: left; width:35%;}
#info_note .right {  float: right; width:35%;}
#info_note .point { font-size:10px; font-weight:bold; width:100%;  margin-top:5px;}
#info_note .point span { font-size:10px; font-style:italic; display:block; border-top:1px #000 solid;}
#info_note .point .name { width:100px; float:left; margin-top :10px;}
#info_note .point .isi  { float:left; height:26px; line-height:25px;margin-top:7px; margin-left:10px;font-size:10px;}

#info_note .point_right { font-size:10px; font-weight:bold; width:100%;  margin-top:5px; text-align:left;}
#info_note .point_right span { font-size:10px; font-style:italic; display:block; border-top:1px #000 solid;}
#info_note .point_right .name {width:25%; float:left;  margin-top :10px;}
#info_note .point_right .isi  { float:left; height:26px; line-height:25px; margin-left:10px;margin-top:7px;font-size:10px;}
.clear { clear:both; content:'';}
table .long {width:400px;}
table .small {width:40px;}
table th span { font-style:italic; font-weight:normal;}
.line { height:1px; background:#000; content:' ';margin: auto;}
.line2 { width:90%; height:1px; background:#000000; content:' '; margin:auto;}
#product-table th { background-color:#fff; color:#000;}
#product-table td { border-color:#000; color:#000; font-size:9px;}
#product-table td.none { border: none;}
#footer_note .block { font-size:10px; float:left; width:25%; height:90px; background-color:#fff; text-align:center;}
input[type="text"].isian {border:0px;box-shadow:none; padding:0px;font-size: 10px;color:#000;height: 20px;width:90%; text-align:center;}
hr{border-top:1px #000000 solid; border-bottom:0px; border-left:0px; border-right:0px;}
.terbilang{font-size:10px;float:left; text-align:left;font-style: italic;margin-left:20px; padding-top:10px;}
.title_terbilang {width:100px;float:left; text-align:left;font-style: italic;font-size:10px;}
.page{width:100%;}
}
@media print
{
td{padding:5px;}
table.PAYSLIP td{padding:5px;}
.page-break hr {display:block;border: 1px dashed red;}
.PAYSLIP td {font-family:verdana,sans-serif; font-size:8px; padding:0px;}
#note {margin-left:auto;padding:5px;}
#info_note .judul { text-align:center; font-size:16px;padding-top:0px;padding-bottom:0px;}
#info_note .left { float: left; width:35%;}
#info_note .right {  float: right; width:35%;}
#info_note .point { font-size:10px; font-weight:bold; width:100%;  margin-top:5px;}
#info_note .point span { font-size:10px; font-style:italic; display:block; border-top:1px #000 solid;}
#info_note .point .name { width:100px; float:left; margin-top :10px;}
#info_note .point .isi  { float:left; height:26px; line-height:25px;margin-top:7px; margin-left:10px;font-size:10px;}

#info_note .point_right { font-size:10px; font-weight:bold; width:100%;  margin-top:5px; text-align:left;}
#info_note .point_right span { font-size:10px; font-style:italic; display:block; border-top:1px #000 solid;}
#info_note .point_right .name {width:25%; float:left;  margin-top :10px;}
#info_note .point_right .isi  { float:left; height:26px; line-height:25px; margin-left:10px;margin-top:7px;font-size:10px;}
.clear { clear:both; content:'';}
table .long {width:400px;}
table .small {width:40px;}
table th span { font-style:italic; font-weight:normal;}
.line { height:1px; background:#000; content:' ';margin: auto;}
.line2 { width:90%; height:1px; background:#000000; content:' '; margin:auto;}
#product-table th { background-color:#fff; color:#000;}
#product-table td { border-color:#000; color:#000; font-size:9px;}
#product-table td.none { border: none;}
#footer_note .block { font-size:10px; float:left; width:19%; height:90px; background-color:#fff; margin-left:1%; text-align:center;}
input[type="text"].isian {border:0px;box-shadow:none; padding:0px;font-size: 10px;color:#000;height: 20px;width:100%; text-align:center;}
hr{border-top:1px #000000 solid; border-bottom:0px; border-left:0px; border-right:0px;}
.terbilang{font-size:10px;float:left; text-align:left;font-style: italic;  padding-top:10px;}
.title_terbilang {width:100px;float:left; text-align:left;font-style: italic;font-size:10px;}
.page{width:98.5%;}
}

</style>
<div class="page" >  
  <table border="0" width="100%" id="product-table">
    <thead>
      <tr>
        <th colspan="14" style=" background-color:#FFFFFF; color:#000000; border:none">
          <div id="header-wrapper">
             <div id="head_judul" style="float:left; text-align:left;">
              <h4>PT. <%=CompanyName%></h4>
              <h5><%=FcbaName%></h5>
            </div>
            <div id="subheader" class="centered" style="width:50%; margin:auto;"> 
              <h1 class="judul">Stock On Hand</h1> 
            </div>  
            <div style="clear:both;"></div>
          </div>
        </th>
      </tr> 
      <!--<div class="printhide">--> 
      <tr>
        <th colspan="14" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></th>
      </tr>
      <!--</div>--> 
        <th style="background-color: #0c5b24; color: #fff;">Store Code</th>
        <th style="background-color: #0c5b24; color: #fff;">Stock Code</th>
        <th style="background-color: #0c5b24; color: #fff;">Stock Name</th> 
        <th style="background-color: #0c5b24; color: #fff;">UOM</th>
        <th style="background-color: #0c5b24; color: #fff;">Quantity</th>  
      </tr>
    </thead>
  <tbody>
<%
   
  if Not rsGrnInv.eof then
      nCountofField=rsGrnInv.fields.count-1
    v_deliveredqty=0
    v_acceptedqtybruto=0
    v_varianceqty = 0
    v_amountbeforeloss = 0
    v_losspercentage = 0
    v_maxtolerance= 0
    v_penaltyqty= 0
    v_penaltyamount= 0
    v_amountafterloss= 0
  end if 

  while not rsGrnInv.eof 
      if no  mod 2 = 1 then
        cClass = "odd"
      else
        cClass = "default"
      end if

      

      openingqty  = rsGrnInv("openingbalanceqty")
      receivenote = rsGrnInv("receivenoteqty")
      issue		    = rsGrnInv("issueqty")
      transportask= rsGrnInv("quantitytransfered")
      receivetask	= rsGrnInv("quantityreceived")
      qtywriteoff	= rsGrnInv("qtywriteoff")
      qtyonhand	= cdbl(openingqty)+cdbl(receivenote)-cdbl(issue)-cdbl(transportask)+cdbl(receivetask)+cdbl(qtywriteoff)


  %> 
    <tr class="<%=cClass%>">
      <td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsGrnInv("storecode")%></td>
      <td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsGrnInv("itemcode")%></td>
      <td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsGrnInv("stockname")%></td> 
      <td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsGrnInv("uom")%></td>
      <td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(cdbl(to_nol(qtyonhand)),2)%></td>
    </tr>
<%      
      rsGrnInv.MoveNext
    Wend
  rsGrnInv.Close
  set rsGrnInv = Nothing
%> 
</tbody>  
</table> 
</body>
</html>






