<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<title>ePlaS Reporting - Return Stock Issue</title>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<link rel="stylesheet" href="/css/jquery-ui.css" />
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
<script language="JavaScript"> 
var newwin;
function launchwin(winurl,winname,winfeatures)
{
	//This launches a new window and then
	//focuses it if window.focus() is supported.
	newwin = window.open(winurl,winname,winfeatures);

}
</script>
</head>

<%
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"

v_txtORDERDATE_From=Trim(REQUEST("ORDERDATE_From"))
v_txtORDERDATE_To=Trim(REQUEST("ORDERDATE_To"))
v_txtORDERNO=Trim(REQUEST("ORDERNO"))
v_txtSTOCKCODE=Trim(REQUEST("STOCKCODE"))
businessunit = request("businessunit")
fcbadistination = request("fcbadistination")

		q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & businessunit & "' "
			' response.write q
			set rsCat = connStr.Execute(q)
			if not rsCat.eof then
			   fcbaname = rsCat("fcName")
			   CompanyName=rsCat("FCCOMPANYNAME")
			   MgrName = rsCat("FCMANAGER")
			   KTUName = rsCat("FCKTU")
			   CropName=rsCat("FCCROP")
			   FcbaType=rsCat("fcType")
			end if

txtSQL = " SELECT TO_CHAR(H.RETURNDATE,'MM/DD/YYYY') RETURNDATE,H.RETURNNO,D.ITEMCODE,S.FCNAME,D.UOM,NVL(D.QTYISSUE,0)QTYISSUE,NVL(D.QTYRETURN,0)QTYRETURN,H.REMARK,H.POSTINGFLG,H.FCBA " & _
" FROM RETURNISSUEHDR H " & _
" LEFT JOIN RETURNISSUEDTL D ON H.RETURNNO=D.RETURNNO AND H.FCBA=D.FCBA " & _
" LEFT JOIN STOCK S ON S.FCCODE=D.ITEMCODE AND S.FCBA=D.FCBA " & _
" WHERE  H.FCBA = '" & businessunit & "'"

if  len(trim(v_txtORDERDATE_From)) > 0 then
    txtSQL = txtSQL & " AND H.RETURNDATE >=to_Date('" & v_txtORDERDATE_From & "','MM/DD/YYYY')"
	txtSelect = " From " & FormatDateTime(v_txtORDERDATE_From,1)
end if
if  len(trim(v_txtORDERDATE_To)) > 0 then
    txtSQL =txtSQL &  " AND H.RETURNDATE <=to_Date('" & v_txtORDERDATE_To & "','MM/DD/YYYY')"
	txtSelect = txtSelect & "  To " & FormatDateTime(v_txtORDERDATE_To,1)
end if
if  len(trim(v_txtORDERNO)) > 0 then
    txtSQL = txtSQL & " AND H.RETURNNO='" & v_txtORDERNO & "'"
	txtSelect = txtSelect & "  ORDER No " & v_txtORDERNO
end if
if  len(trim(v_txtSTOCKCODE)) > 0 then
    txtSQL =txtSQL &  " AND D.ITEMCODE='" & v_txtSTOCKCODE & "'"
	txtSelect = txtSelect & "  Stock Code " & v_txtSTOCKCODE
end if 


txtsql=(txtsql) & 	" ORDER BY H.RETURNDATE,H.RETURNNO"
if len(trim(txtSelect)) > 0 then
   txtSelect = "Selected  by : " & txtSelect
else
   txtSelect = "Selected by All Data"
end if

' response.Write(txtsql)
response.Flush()
%>
	<body>
		<div id="wrapper">
        <!--#include file="print_excel.asp"-->
        <table border="0" width="100%" id="product-table" class="xthf-site">
         <thead>
         <tr>
         <th colspan="21" style=" background-color:#FFFFFF; color:#000000; border:none">
			<div id="header-wrapper">
				<div id="head_judul" style="float:left; text-align:left;">
					<h4>PT. <%=CompanyName%></h4>
					<h5><%=fcbaname%></h5>
				</div>
				<div id="subheader" class="centered" style="width:50%; margin:auto;">
					<h2>LAPORAN RETURN STOCK ISSUE</h2>
				</div>
			</div>
				</th>
			</tr>
						<!--<div class="printhide">-->
						<tr>
							<th colspan="21" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></th>
						</tr>
						<!--</div>-->
							<tr>
							<%	
							Set rsCat = connStr.Execute(txtSQL)
							if Not rsCat.eof then
							   nCountofField=rsCat.fields.count-1
							end if
							NO=1%>
							</tr>
							<tr>
								<th>Action</th>
								<th>Date</th>
								<th>Order No</th>
								<th>Stock Code</th>
								<th>Stock Desc.</th>
								<th>UOM</th>
								<th>Qty Issue</th>
								<th>Qty Return</th>
								<th>Remark</th>
								<th>Posting</th>
								
							</tr>
                            </thead>
                            <tbody>
							<%
								Dim v_Quantity,v_price,v_amount,POSTINGVIEW
								v_Quantity=0
								v_price=0
								v_amount=0
								While Not rsCat.EOF
	    
								POSTINGVIEW = rscat("POSTINGFLG")
								IF POSTINGVIEW = "1" THEN
									POSTINGVIEW = "YA"
								ELSE 
									POSTINGVIEW = "TIDAK"
								END IF
							%>
							<tr>
								
								<td align="center" style="<%=cmdstyle%>"><a href="stockreturn_edit.Asp?returnno=<%=rscat("RETURNNO")%>&returndate=<%=rscat("RETURNDATE")%>&fcba=<%=rscat("FCBA")%>&txtJenis=Edit">
								<img src="IMAGES/edit.gif" alt="Edit Data : <%=rscat(1)%>" border="0">
							</a></td>
								<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("RETURNDATE")%></font></td>
								<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("RETURNNO")%></font></td>
								<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("ITEMCODE")%></font></td>
								<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("FCNAME")%></font></td>
								
								<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("UOM")%></font></td>
								<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("QTYISSUE"))%></font></td>
								<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("QTYRETURN"))%></font></td>
								<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("REMARK")%></font></td>
								<td align="LEFT" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=POSTINGVIEW%></font></td>

							</tr>
							  <%
							  'v_Quantity=v_Quantity+trim(cdbl(rsCat("quantity")))
							  v_Quantity  =v_Quantity+cdbl(rsCat("QTYRETURN"))
							
							  rsCat.MoveNext
							  'response.Flush()
							  no=no+1
							Wend
							rsCat.close()
							rsCat=null%>
							<input  type="hidden" value="<%=txtTable%>" name="txtTable">
							<input  type="hidden" value="<%=txtModule%>" name="txtModule">
							</form>
							<tr bgcolor="#F4F4A0">
							<td align="Center" style="<%=cmdstyleLeft%>" colspan="7"> <font face="Verdana" size="1"><b>T o t a l</b></font></td>
							<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><b><%=FormatNumber(v_Quantity,2)%></b></font></td>
							
							</tr>
                            </tbody>
						</table>
						<!--#include file="Bottom_Report.asp"-->
			
		</DIV>
	</BODY>
</HTML>