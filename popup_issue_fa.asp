<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->  
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>FB-Plants - ASSET ISSUE</title>
<link rel="stylesheet" href="css/stylereport2.css" type="text/css" />
<link rel="stylesheet" href="css/jquery-ui.css" />

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
</head>

<body>
<%

Dim sqljob,rsjob,v_txtFcba,txtSearch,txtColom,txtdate

  v_txtFcba     = request("fcba") 
 		
  q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & v_txtFcba & "' "
  set rsCat = connStr.Execute(q)
  if not rsCat.eof then
     fcbaname = rsCat("fcName")
     CompanyName=rsCat("FCCOMPANYNAME")
     MgrName = rsCat("FCMANAGER")
     KTUName = rsCat("FCKTU")
     CropName=rsCat("FCCROP")
     FcbaType=rsCat("fcType")

  end if

  txtColom = " SELECT a.stockcode,c.fcname stockname,c.unitofmeasurement uom,a.remarks,a.amount,a.sivcode,a.LINENOHIST FROM issuedetails a " & _
			 " INNER JOIN issue b ON (a.sivcode=b.sivcode AND a.fcba=b.fcba) " & _
			 " INNER JOIN bankstock c ON (a.stockcode=c.fccode) " & _
			 " INNER JOIN stockgroup d ON (c.stockgroupcode=d.fccode AND a.fcba=d.fcba) " & _
			 " LEFT JOIN fixedasset e ON (a.sivcode=e.sivcode AND a.stockcode=e.stockcode AND a.fcba=e.fcba) " & _
			 " WHERE a.fcba='" & v_txtFcba & "' AND d.grouptype='DIRECTCOST' AND d.activation='Y' " & _
			 " AND e.sivcode IS NULL AND a.LINENOHIST IS NOT NULL" & _
			 " ORDER BY sivcode "

  Set rsjob = Server.CreateObject("ADODB.Recordset")
  rsjob.Open txtColom, connstr, adOpenStatic 

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
              <h1 class="judul">Data Issue Asset</h1> 
            </div>  
            <div style="clear:both;"></div>
          </div>
        </th>
      </tr> 
     
      <tr>
        <th colspan="14" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></th>
      </tr>
     
        <th style="background-color: #32CD32; color: #fff;">Action</th>
        <th style="background-color: #32CD32; color: #fff;">SivCode</th>
        <th style="background-color: #32CD32; color: #fff;">StockCode</th>
        <th style="background-color: #32CD32; color: #fff;">StockName</th>
        <th style="background-color: #32CD32; color: #fff;">UOM</th>
        <th style="background-color: #32CD32; color: #fff;">Description</th>
        <th style="background-color: #32CD32; color: #fff;">Amount</th>
      </tr>
    </thead>
  <tbody>
<%
   dim no 
   no =1
  while not rsjob.eof 

  %> 
    <tr class="<%=cClass%>">
      <td align="Center" ><font face="Verdana" size="1"><button id="btnselect" onclick="getSelected(<%=no%>)" >Select</button>
        <input type=hidden id='sivcode<%=no%>' value="<%=rsjob("sivcode")%>"></input>
        <input type=hidden id='stockcode<%=no%>' value="<%=rsjob("stockcode")%>"></input>
        <input type=hidden id='stockname<%=no%>' value="<%=rsjob("stockname")%>"></input>
        <input type=hidden id='uom<%=no%>' value="<%=rsjob("uom")%>"></input>
        <input type=hidden id='remarks<%=no%>' value="<%=rsjob("remarks")%>"></input>
        <input type=hidden id='amount<%=no%>' value="<%=rsjob("amount")%>"></input>
      </td>
      <td align="Center" ><font face="Verdana" size="1"><%=rsjob("sivcode")%></td>
      <td align="Center" ><font face="Verdana" size="1"><%=rsjob("stockcode")%></td> 
      <td align="Center" ><font face="Verdana" size="1"><%=rsjob("stockname")%></td> 
      <td align="Center" ><font face="Verdana" size="1"><%=rsjob("uom")%></td>
      <td align="Center" ><font face="Verdana" size="1"><%=rsjob("remarks")%></td>
      <td align="Center" ><font face="Verdana" size="1"><%=rsjob("amount")%></td>
    </tr>
<%      
      rsjob.MoveNext
      no = no +1
    Wend
  rsjob.Close
  set rsjob = Nothing
%> 
</tbody>  
</table> 
</body>
</html>
<script src="/js/jquery-1.9.1.js"></script>
<script>
 
  function getSelected(a) {

    var sivcode   = $('#sivcode' + a).val();
    var stockcode = $('#stockcode' + a).val();
	  var stockname = $('#stockname' + a).val();
	  var uom       = $('#uom' + a).val();
	  var remarks   = $('#remarks' + a).val();
	  var amount    = $('#amount' + a).val();
	  
    window.opener.parent.document.getElementById("sivcode").value = sivcode;
    window.opener.parent.document.getElementById("kodebarang").value = stockcode;
	  window.opener.parent.document.getElementById("namabarang").value = stockname;
	  window.opener.parent.document.getElementById("fcname").value = stockname;
	  window.opener.parent.document.getElementById("description").value = remarks;
	  window.opener.parent.document.getElementById("amount").value = amount;
	  window.opener.parent.document.getElementById("additionalvalue").value = 0;
	  window.opener.parent.document.getElementById("total").value = amount;
	  
	  
	  
	  
	  
	  
      window.close();
	}
</script>





