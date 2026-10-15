<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>ePlaS - Store Issue Voucher</title>
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
#info_note .point .name { width:150px; float:left;}
#info_note .point .isi  { float:left; height:36px; line-height:35px; margin-left:20px;}

#info_note .point_right { font-size:12px; font-weight:bold; width:100%;  margin-top:5px; text-align:left;}
#info_note .point_right span { font-size:11px; font-style:italic; display:block; border-top:1px #000 solid;padding-top:2px;margin-top:2px;}
#info_note .point_right .name { width:100px; float:left;}
#info_note .point_right .isi  { float:left; height:36px; line-height:35px; margin-left:20px;}
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
DIM v_txtSIVDATE_From,v_txtSIVDATE_To,v_txtSIVCODE,v_txtSTOCKCODE,v_txtJOBCODE,v_txtLOCATIONTYPE,v_txtLOCATIONCODE,txtSelect,nCountofField,no,v_txtSIVDATE
v_txtSIVCODE=Trim(REQUEST("fcCode"))
v_txtSIVDATE = request("IssueDate")
Dim oRS, oSQL, NoUrut, nTotalValue, grnno

grnno = Request("grnno")

oSQL	="SELECT issue.sivdate,issue.storecode,store_api.get_fcname(issue.storecode,iss.fcba) STORENAME,iss.sivcode, iss.stockcode, stock.fcname,STOCK_API.get_uom(iss.stockcode,iss.fcba) uom ,iss.jobcode, iss.locationcode," & _
		 "iss.locationtype, nvl(sum(iss.quantity),0) as quantity, nvl(sum(iss.price),0) as price, nvl(sum(iss.amount),0) as amount, job.fcname as jobname, iss.reference as reference, iss.remarks as remarks, iss.linenohist" & _
		 " FROM issuedetails iss, issue issue, stock stock, job job" & _
		 " WHERE (((iss.sivcode = issue.sivcode) and (iss.fcBA = issue.fcBA))" & _
		 " AND ((iss.stockcode = stock.fccode) and (iss.fcBA= stock.fcBa)) " & _
		 " AND ((iss.jobcode = job.fccode) and (iss.fcBA = job.fcBA)) )" & _
		 " AND iss.fcba='" & session("fcBA") & "'"

if  len(trim(v_txtSIVDATE)) > 0 then
    oSQL = oSQL & " and issue.sivdate = to_Date('" & v_txtSIVDATE & "','MM/DD/YYYY')"
	txtSelect = " From " & FormatDateTime(v_txtSIVDATE,1)
end if

if  len(trim(v_txtSIVCODE)) > 0 then
    oSQL = oSQL & " and iss.sivcode='" & v_txtSIVCODE & "'"
	txtSelect = txtSelect & "  SIV Code " & v_txtSIVCODE
end if

oSQL=(oSQL) & 	" group by issue.sivdate,issue.storecode,store_api.get_fcname(issue.storecode,iss.fcba),iss.sivcode, iss.stockcode, stock.fcname,STOCK_API.get_uom(iss.stockcode,iss.fcba), iss.jobcode," & _
					" iss.locationcode, iss.locationtype, job.fcname, iss.reference, iss.remarks, iss.linenohist  order by issue.sivdate,iss.sivcode,iss.stockcode"
if len(trim(txtSelect)) > 0 then
   txtSelect = "Selected  by : " & txtSelect
else
   txtSelect = "Selected by All Data"
end if

' response.Write(oSQL)
response.Flush()

Set rsCat = Server.CreateObject("ADODB.Recordset")
rsCat.Open oSQL, connstr, adLockOptimistic, adOpenStatic

%>
<div class="note" >

	<header id="info_note" >
		<div class="judul"><strong>STORE ISSUE VOUCHER</strong></div>
			<table width="100%">
			<!-- Left -->
			<td valign="top">
				<div class="point">
					<div class="name">
					Kebun/Pabrik<span>Business Unit</span>
					</div>
					<div class="isi">
					: <%=session("FCBA")%> - <%=Session("BA")%>
					</div>
				</div>
				<div class="clear"></div>
				<div class="point">
					<div class="name">
					Gudang<span>Store</span>
					</div>
					<div class="isi">
					: <%=rsCat("STORECODE")%> - <%=rsCat("STORENAME")%>
					</div>
				</div>
				<div class="clear"></div>

			<!-- right -->
			<td width="400px" valign="top">			
				<div class="point_right">
					<div class="name">
					Nomor<span>Number</span>
					</div>
					<div class="isi">
					: <%=rsCat("SIVCODE")%>
					</div>
				</div>
				<div class="clear"></div>
				<div class="point_right">
					<div class="name">
					Tanggal<span>Date</span>
					</div>
					<div class="isi">
					: <%=rsCat("SIVDATE")%>
					</div>
				</div>
				<div class="clear"></div>
			</td>
		</table>
		

	</header>

<div class="clear"></div>
<br/><br/>
<table border="0" width="100%" id="product-table" >
<thead>
    <tr>
        <th rowspan="2" class="small">Nomor<div class="line"></div><span>Item</span></th>
		<th rowspan="2">Stock Code</th>
        <th rowspan="2" class="long">Description</th>
        <th rowspan="2" style="font-size:10px">Dikeluarkan<div class="line"></div><span>Qty</th>
        <th rowspan="2" style="font-size:10px">Satuan<div class="line"></div><span>UOM</span></th>
        <th rowspan="2">Keterangan<div class="line"></div><span>Remark</span></th>
    </tr>
</thead >

<tbody>
<%	
if Not rsCat.eof then
	nCountofField=rsCat.fields.count-1
end if
NO=1%>
<%
Dim v_Quantity,v_price,v_amount
v_Quantity=0
v_price=0
v_amount=0		
%>
<%If rsCat.BOF And rsCat.EOF Then%>
    <tr>
        <td Colspan="14">NO DATA AVAILABLE</td>
    </tr>
<%Else
	While Not rsCat.EOF
	if cdbl(rscat("quantity")) <>0 and not isnull(rscat("quantity")) then
	v_price=cdbl(rscat("amount")) / cdbl(rscat("quantity"))
	else
	v_price=0   
	end if   
		NoUrut	=	NoUrut + 1%>
		<tr>
			<td align="Right"><%=NoUrut%></td>
			<td align="CENTER" style="width:100px;"><%=rsCat("STOCKCODE")%></td>
			<td align="LEFT"><%=rsCat("FCNAME")%></td>
			<td align="Right" style="width:80px;"><%=FbpNum2(rsCat("QUANTITY"))%></td>
			<td align="CENTER" style="width:50px;"><%=rsCat("uom")%></td>
			<td align="LEFT"><%=rsCat("REMARKS")%></td>
		</tr>
	  <%
	'v_Quantity=v_Quantity+trim(cdbl(rsCat("quantity")))
	v_Quantity  =v_Quantity+cdbl(rsCat("quantity"))
	v_amount =v_amount+cdbl(rsCat("amount"))
	'response.write cdbl(trim(rsCat("MTHOPENBALANCE")))+1
	rsCat.MoveNext
	'response.Flush()
	no=no+1
	Wend
	rsCat.close()
	rsCat=null%>
   
<%End If%>

</tbody>
</table>
<br/>
    <div id="footer_note">
    	<div class="block" style="margin:0px;">
        	<table style="text-align:left; width:100%;">
            	<tr><td>Asli</td><td>- Untuk Accounting</td></tr>
                <tr><td>Copy 1</td><td>- Untuk Gudang</td></tr>
            </table>
        
        </div>
        <div class="block">Requested by,
        <br/><br/><br/><br/>
		<input class="isian" type="text" value="">
        <hr/>
        Pemohon
        </div>
        <div class="block">Issued by,
        <br/><br/><br/><br/>
		<input class="isian" type="text" value="">
        <hr/>
        Ka.Gudang/Store Keeper
        </div>
        <div class="block">Approved by,
        <br/><br/><br/><br/>
        <input class="isian" type="text" value="<%=Session("FCMANAGER")%>" readonly='True'>
        <hr/>
        Manager
        </div>
        <div class="block">Received by,
        <br/><br/><br/><br/>
		<input class="isian" type="text" value="">
        <hr/>
        Penerima
        </div>
        <div class="block"></div>
    </div>
    
    
    <div class="clear"></div>
</div>




</body>
</html>
