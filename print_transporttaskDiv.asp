<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>ePlaS - Transport Task</title>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<link rel="stylesheet" href="/css/jquery-ui.css" />
<script type="text/javascript" language="JavaScript1.2" src="stlib.js"></script>
<script language="JavaScript" src='/js/x.js'></script>
<script language="JavaScript" src='/js/xtableheaderfixed.js'></script>
        <script language="JavaScript">
        xAddEventListener(window, 'load',
          function() {
            new xTableHeaderFixed('xthf-site', window);
            new xTableHeaderFixed('xthf-green', window);
            new xTableHeaderFixed('xthf-blue', window);
          }, false
        );
        </script>
</head>

<body>
<style>
.note {margin:auto;padding:5px;}
#info_note .judul { text-align:center; font-size:20px;}
#info_note .left { float: left; width:35%;}
#info_note .right {  float: right; width:35%;}
#info_note .point { font-size:12px; font-weight:bold; width:100%;  margin-top:5px;}
#info_note .point span { font-size:11px; font-style:italic; display:block; border-top:1px #000 solid;padding-top:2px;margin-top:2px;}
#info_note .point .name { width:150px; float:left; margin-top: 5px;}
#info_note .point .isi  { float:left; height:10px; line-height:35px; margin-left:20px;}

#info_note .point_right { font-size:12px; font-weight:bold; width:100%;  margin-top:5px; text-align:left;}
#info_note .point_right span { font-size:11px; font-style:italic; display:block; border-top:1px #000 solid;padding-top:2px;margin-top:2px;}
#info_note .point_right .name { width:100px; float:left; margin-top: 5px;}
#info_note .point_right .isi  { float:left; height:10px; line-height:35px; margin-left:20px;}
.clear { clear:both; content:'';}
table .long {width:400px;}
table .small {width:40px;}
table th span { font-style:italic; font-weight:normal;}
.line { height:1px; background:#000; content:' ';}
.line2 { width:90%; height:1px; background:#000000; content:' '; margin:auto;}
#product-table th { background-color:#fff; color:#000;}
#product-table td { border-color:#000; color:#000;}
#product-table td.none { border: none;}
#footer_note .block { font-size:10px; float:left; width:19%; height:200px; background-color:#fff; margin-left:1%; text-align:center;}
input[type="text"].isian {border:0px;box-shadow:none; padding:0px;font-size: 10px;color:#000;height: 20px;width:100%; text-align:center;}
hr{border-top:1px #000000 solid; border-bottom:0px; border-left:0px; border-right:0px;}
</style>
<%
Dim fccode, nCountofField, NO
Dim oRS, oSQL, NoUrut, nTotalValue, grnno

orderno = Request("orderno")

oSQL = " SELECT DISTINCT id.orderno" & _
       "        , id.fcba" & _
	   "        , id.orderdate" & _
       "        , TO_CHAR(id.orderdate, 'DD') orderdate_dd" & _
       "        , TO_CHAR(id.orderdate, 'MM') orderdate_mm" & _
       "        , TO_CHAR(id.orderdate, 'YYYY') orderdate_yy" & _
       "        , id.remarks" & _
	   "        , td.storecodefrom" & _
       "        , store_api.get_fcname(td.storecodefrom,td.fcba) storenamefrom" & _
       "        , td.storecodeto " & _
       "        , store_api.get_fcname(td.storecodeto,td.fcba) storenameto " & _ 
	   " FROM transporttaskdiv_tab id " & _
	   " INNER JOIN transporttaskdivdtl_tab td ON (td.orderno = id.orderno and td.fcba = id.fcba) " & _
	   " WHERE id.fcba = '" & session("fcBA") & "'" & _
       "       AND id.orderno = '" & orderno & "'"

' response.Write(oSQL)
response.Flush()

Set rsCat = Server.CreateObject("ADODB.Recordset")
rsCat.Open oSQL, connstr, adLockOptimistic, adOpenStatic

%>
<div class="note" >

	<header id="info_note" >
		<div class="judul"><strong>TRANSPORT TASK DIVISI</strong></div>
			<table width="100%">
			<%IF NOT rsCat.EOF THEN%>
			<!-- Left -->
			<td valign="top">
				<div class="point">
					<div class="name">
					Store Asal
					</div>
					<div class="isi">
					: <%=rsCat("storecodefrom")%> - <%=rsCat("storenamefrom")%>
					</div>
				</div>
				<div class="clear"></div>
				<div class="point">
					<div class="name">
					Store Tujuan
					</div>
					<div class="isi">
					: <%=rsCat("storecodeto")%> - <%=rsCat("storenameto")%>
					</div>
				</div>
				<div class="clear"></div>

			<!-- right -->
			<td width="400px" valign="top">			
				<div class="point_right">
					<div class="name">
					Nomor
					</div>
					<div class="isi">
					: <%=rsCat("orderno")%>
					</div>
				</div>
				<div class="clear"></div>
				<div class="point_right">
					<div class="name">
					Tanggal
					</div>
					<div class="isi">
					: <%=rsCat("orderdate_dd")%>&nbsp;<%=Bulan(rsCat("orderdate_mm"))%>&nbsp;<%=rsCat("orderdate_yy")%>
					</div>
				</div>
				<div class="clear"></div>
			</td>
			<%ELSE%>
			<!-- Left -->
			<td valign="top">
				<div class="point">
					<div class="name">
					Store Asal
					</div>
					<div class="isi">
					: 
					</div>
				</div>
				<div class="clear"></div>
				<div class="point">
					<div class="name">
					Store Tujuan
					</div>
					<div class="isi">
					: 
					</div>
				</div>
				<div class="clear"></div>

			<!-- right -->
			<td width="400px" valign="top">			
				<div class="point_right">
					<div class="name">
					Nomor
					</div>
					<div class="isi">
					: 
					</div>
				</div>
				<div class="clear"></div>
				<div class="point_right">
					<div class="name">
					Tanggal
					</div>
					<div class="isi">
					: 
					</div>
				</div>
				<div class="clear"></div>
			</td>
			<%END IF%>
		</table>
		

	</header>

<div class="clear"></div>
<br/>
<table border="0" width="100%" id="product-table" >
<thead>
    <tr>
        <th rowspan="2" class="small">Nomor<div class="line"></div><span>Item</span></th>
		<th rowspan="2">Stock Code</th>
        <th rowspan="2" class="long">Description</th>
		<th rowspan="2" style="font-size:10px">Satuan<div class="line"></div><span>UOM</span></th>
        <th rowspan="2" style="font-size:10px">Dikeluarkan<div class="line"></div><span>Qty</th>
        <th rowspan="2">Keterangan<div class="line"></div><span>Remark</span></th>
    </tr>
</thead >

<tbody>
<%
oSQL = " SELECT " & _
       "         id.stockcode" & _
       "        , stock_api.get_stockname(id.stockcode, id.fcba) stockname" & _
       "        , stock_api.get_uom(id.stockcode, id.fcba) uom" & _
       "        , NVL(id.quantitytransfered, 0) AS quantitytransfered" & _
       "        , id.remarks" & _
       " FROM transporttaskdivdtl_tab id " & _
       " WHERE id.fcba = '" & session("fcBA") & "'" & _
       "       AND id.orderno = '" & orderno & "'"

'response.Write(oSQL)
response.Flush()

Set rsCat = Server.CreateObject("ADODB.Recordset")
rsCat.Open oSQL, connstr, adLockOptimistic, adOpenStatic

If rsCat.BOF And rsCat.EOF Then
%>
    <tr>
        <td Colspan="11">NO DATA AVAILABLE</td>
    </tr>
<%Else
	While Not rsCat.EOF
		nCountofField=rsCat.fields.count-1
		NoUrut	=	NoUrut + 1
%>
		<tr>
			<td align="center"><%=NoUrut%></td>
			<td align="center" style="width:100px;"><%=rsCat("stockcode")%></td>
			<td align="left"><%=rsCat("stockname")%></td>
			<td align="center" style="width:50px;"><%=rsCat("uom")%></td>
			<td align="center" style="width:80px;"><%=rsCat("quantitytransfered")%></td>
			<td align="left"><%=rsCat("remarks")%></td>
		</tr>
<%
		rsCat.MoveNext
	Wend
	rsCat.close()
	rsCat=null
%>
   
<%End If%>

</tbody>
</table>
<br/>
    <div id="footer_note">
    	<div class="block" style="margin:0px;">
        	<table style="text-align:left; width:100%;">
            	<tr><td>Asli</td><td>- Untuk Accounting</td></tr>
                <tr><td>Copy 1</td><td>- Untuk Gudang Kirim</td></tr>
                <tr><td>Copy 2</td><td>- Untuk Gudang Terima</td></tr>
	        </table>
        
        </div>
        <div class="block">Received by,
        <br/><br/><br/><br/>
		<input class="isian" type="text" value="">
        <hr/>
        Mandor
        </div>
		<div class="block">Approved by,
        <br/><br/><br/><br/>
        <input class="isian" type="text" value="<%=Session("FCMANAGER")%>" readonly='True'>
        <hr/>
        Asisten
        </div>
		<div class="block">Issued by,
        <br/><br/><br/><br/>
		<input class="isian" type="text" value="">
        <hr/>
        Ka.Gudang/Store Keeper
        </div>
        <div class="block"></div>
    </div>
    
    
    <div class="clear"></div>
</div>




</body>
</html>
