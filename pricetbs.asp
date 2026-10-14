<%Option Explicit%>
<%response.buffer=True%>
<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<html>
<head>
<title>ePlaS - Price TBS</title>
<script type="text/javascript" language="JavaScript1.2" src="stlib.js" ></script>
<link rel="StyleSheet" href="dtree.css" type="text/css" />
<link rel="StyleSheet" href="style-content.css" type="text/css" />
<link rel="stylesheet" href="css/jquery-ui.css" />
<script type="text/javascript" src="dtree.js"></script>
<script src="/js/jquery-1.9.1.js"></script>
<script src="/js/jquery-ui.js"></script>
<link href="/js/select2/select2.min.css" rel="stylesheet" type="text/css"/>
<script src="/js/select2/select2.min.js" type="text/javascript"></script>
<script language="JavaScript">
</script>
 <script>
$(function() {
$( ".datepicker" ).datepicker();
	getSupplier();
	getProduct();
});

function getSupplier()
{
	var fcba = '<%= Session("fcBA") %>';
	$('#supplier').select2({
		ajax: {
			delay: 1000,
			url: 'getlistsupplierinvajax.asp?jnsinvoice=100&fcba=' + fcba,
			dataType: 'json',
			type: "POST",
			allowClear: true,
			data: function (params) {
					var query = {
									keyword: params.term
								}
					return query;
			}
		}
	});
}

function getProduct()
{
	var product = 200;
	$('#product').select2({
		ajax: {
			delay: 1000,
			url : 'getlistsupplierinvajax.asp?jnsinvoice=' + product,
			dataType: 'json',
			type: "POST",
			allowClear: true,
			data: function (params) {
					var query = {
									keyword: params.term
								}
					return query;
			}
		}
	});
}

function hanyaAngka(evt) {
		evt.value = evt.value.replace(/[^0-9,.-]/g,"");
}
</script>

<%
dim fcba
fcba 		= Session("fcBA")
%>
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

   <% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-tags" aria-hidden="true"></i>Harga TBS</div><% end if %>
   <form action="prosesdata.Asp" method="post" name="eForm">
   <table class="data" cellpadding="0" cellspacing="0" width="100%">
		<tr>
			<td align="left" >Ondate</td>
			<td align="left">       
				<input class="datepicker" type="text" name="ondate" id="ondate" placeholder="dd/mm/yyyy"></td>
			</td>
		</tr>
		
		<tr>
			<td align="left" width="20%">Supplier / Customer</td>
			<td align="left">
				<select type="text" name="supplier" id="supplier" style="width: 250px;">
				</select>
			</td> 
		</tr>
				
		<tr>
		    <td align="left" width="20%">Product</td>
			<td align="left">
				<select type="text" name="product" id="product" style="width: 250px;">
				</select>
			</td>
	    </tr>
		<tr>
		    <td align="left" width="20%">Price</td>
			<td align="left">
				<input style="text-align: right;" onkeyup="hanyaAngka(this);" type="text" size="10" name="price" value="0">
			</td>
	    </tr>
		<tr>
		<td colspan="2" align="left">
		       <input Type="Submit"  name="Add" value="Add">
		  	<input Type="Reset"  name="Reset" value="Reset">
			  <input type="hidden" size="10" name="prosesdata" value="prosespricetbs">
		</td>
	  </tr>
   </table>

   <table class="data">
	<BR>
	<%
	Dim q,rsprice,no
	q=" select p.id, p.ondate, p.suppliercode || ' - ' || s.description as supplier, productcode, price " & _
	 " from INVOICE_PRICETBS p, CUSTOMER s " & _
     " where p.suppliercode=s.customercode and UPPER(s.SUPPLIER) ='Y' AND p.fcBA=s.fcBA and p.fcBA='" & session("fcBA") & "' " & _ 
	 " order by p.ondate desc FETCH FIRST 200 ROWS ONLY"
   Set rsprice = Server.CreateObject("ADODB.Recordset")
   rsprice.CursorType = 3
	rsprice.CursorLocation = 3
	rsprice.Open q, connstr
	%>
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;" >
		
		<TD align="center">NO</TD>
		<TD align="left">ONDATE</TD>	
		<TD align="left">SUPPLIER / CUSTOMER</TD>	
		<TD align="left">PRODUCT</TD>	
		<TD align="left">PRICE</TD>	
		<TD colspan="2">&nbsp;</TD>
	</tr>
	<%
	if Not rsprice.eof then
	  no=1
	  While Not rsprice.EOF %>
	  <tr>
		<TD align="center"><%=no%></TD>
		<TD align="left"><%=rsprice("ondate")%></TD>	
		<TD align="left"><%=rsprice("supplier")%></TD>	
		<TD align="left"><%=rsprice("productcode")%></TD>	
		<TD align="right"><%=formatnumber(rsprice("price"),0)%></TD>	
		<TD align="left">
			<!-- <a href="pricetbs.Asp?ondate=<%=rsprice("id")%>txtJenis=Edit">
				<img src="IMAGES/edit.gif" alt="Edit Data : 02-PAYMENT" border="0">
			</a> -->
		</TD>
	</tr>

	<% 
	rsprice.MoveNext
	no = no +1 
	wend
	rsprice.Close
   	set rsprice = Nothing 
	end if
	%>
</div>
   </TABLE>
   </form>

 
  <!-------------akhir!------------->
</div>
  <!--#include file="bottom.jsp"-->

</body>
</html>
