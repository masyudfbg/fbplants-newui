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
   
</head>
<style type="text/css">


#btnselect {
  background-color: #4CAF50; /* Green */
  border: none;
  color: white;
  padding: 3.5px 8px;
  text-align: center;
  text-decoration: none;
  display: inline-block;
  font-size: 12px;
}


</style> 
<body>
<%


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
    " where h.orderno=d.orderno and h.rowstateid not in('Rejected') and h.fcba='" & v_txtfcba & "' AND h.orderdate between to_date('" & tglawal & "','mm/dd/yyyy') and to_date('" & IssueDate & "','mm/dd/yyyy') " & _
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
        <th style="background-color: #0c5b24; color: #fff;">Action</th>
        <th style="background-color: #0c5b24; color: #fff;">Store Code</th>
        <th style="background-color: #0c5b24; color: #fff;">Stock Code</th>
        <th style="background-color: #0c5b24; color: #fff;">Stock Name</th> 
        <th style="background-color: #0c5b24; color: #fff;">UOM</th>
        <th style="background-color: #0c5b24; color: #fff;">Quantity</th>  
      </tr>
    </thead>
  <tbody>
<%
   
 
  dim no 
  no = 1
  while not rsGrnInv.eof 
     

      openingqty  = rsGrnInv("openingbalanceqty")
      receivenote = rsGrnInv("receivenoteqty")
      issue		    = rsGrnInv("issueqty")
      transportask= rsGrnInv("quantitytransfered")
      receivetask	= rsGrnInv("quantityreceived")
      qtywriteoff	= rsGrnInv("qtywriteoff")
      qtyonhand	= cdbl(openingqty)+cdbl(receivenote)-cdbl(issue)-cdbl(transportask)+cdbl(receivetask)+cdbl(qtywriteoff)


  %> 
    <tr class="<%=cClass%>">
      <td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><button id="btnselect" onclick="getSelected(<%=no%>)" >Select</button></td>
      <td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsGrnInv("storecode")%></td>
      <td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsGrnInv("itemcode")%>
        <input type=hidden id='itemcode<%=no%>' value="<%=rsGrnInv("itemcode")%>"></input>
      </td>
      <td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsGrnInv("stockname")%>
        <input type=hidden id='StockName<%=no%>' value="<%=rsGrnInv("stockname")%>"></input>
      </td> 
      <td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsGrnInv("uom")%></td>
      <td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(cdbl(to_nol(qtyonhand)),2)%></td>
    </tr>
<%    
      no = no + 1  
      rsGrnInv.MoveNext
    Wend
  rsGrnInv.Close
  set rsGrnInv = Nothing
%> 
</tbody>  
</table> 
</body>
</html>
<script src="/js/jquery-1.9.1.js"></script>
<script>
 
  function getSelected(a) {

      var itemcode = $('#itemcode' + a).val();
      var StockName = $('#StockName' + a).val();
      window.opener.parent.document.getElementById("itemcode").value = itemcode;
      window.opener.parent.document.getElementById("StockName").innerHTML = StockName;
      window.close();
     //alert(itemcode+ '/' + StockName);
      
	}
</script>






