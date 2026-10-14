<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<title>ePlaS Reporting - Transport Task</title>
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

txtSQL = " SELECT ID.ORDERNO , TT.ORDERDATE, ID.STOCKCODE ,STOCK_API.get_stockName(ID.STOCKCODE,ID.Fcba) StockNAME, " & _
         " NVL(ID.QUANTITYTRANSFERED,0) AS QUANTITYTRANSFERED , " & _
		 " ID.REMARKS ,  STORECODEFROM  , STORECODETO, TT.ROWSTATE,TT.POSTINGFLG " & _
		 "  FROM TRANSPORTTASKDIVDTL_TAB  ID  " & _
		 "  INNER JOIN TRANSPORTTASKDIV_TAB  TT ON (TT.ORDERNO = ID.ORDERNO AND TT.FCBA = ID.FCBA) " & _
         " WHERE  ID.FCBA = '" & businessunit & "'"

if  len(trim(v_txtORDERDATE_From)) > 0 then
    txtSQL = txtSQL & " AND TT.ORDERDATE >=to_Date('" & v_txtORDERDATE_From & "','MM/DD/YYYY')"
	txtSelect = " From " & FormatDateTime(v_txtORDERDATE_From,1)
end if
if  len(trim(v_txtORDERDATE_To)) > 0 then
    txtSQL =txtSQL &  " AND TT.ORDERDATE <=to_Date('" & v_txtORDERDATE_To & "','MM/DD/YYYY')"
	txtSelect = txtSelect & "  To " & FormatDateTime(v_txtORDERDATE_To,1)
end if
if  len(trim(v_txtORDERNO)) > 0 then
    txtSQL = txtSQL & " AND ID.ORDERNO='" & v_txtORDERNO & "'"
	txtSelect = txtSelect & "  ORDER No " & v_txtORDERNO
end if
if  len(trim(v_txtSTOCKCODE)) > 0 then
    txtSQL =txtSQL &  " AND ID.stockcode='" & v_txtSTOCKCODE & "'"
	txtSelect = txtSelect & "  Stock Code " & v_txtSTOCKCODE
end if 


txtsql=(txtsql) & 	" ORDER BY TT.ORDERDATE,ID.ORDERNO"
if len(trim(txtSelect)) > 0 then
   txtSelect = "Selected  by : " & txtSelect
else
   txtSelect = "Selected by All Data"
end if

'' response.Write(txtsql)
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
					<h2>TRANSPORT TASK DIVISI</h2>
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
								<th></th>
								<th>Date</th>
								<th>Order No</th>
								<th>Stock Code</th>
								<th>Stock Desc.</th>
								<th>Remarks</th>
								<th>Store From</th>
								<th>Store To</th>
								<th>Quantity</th>
								<th>Rowstate</th>
								<th>Posting</th>
								
							</tr>
                            </thead>
                            <tbody>
							<%
								Dim v_Quantity,v_price,v_amount
								v_Quantity=0
								v_price=0
								v_amount=0
								While Not rsCat.EOF
	    
								if rscat("POSTINGFLG") = "0"  then
								   postingflg="Tidak"
								else
								   postingflg="YA"
								end if   
							%>
							<tr>
								<td>
									<% if Session("CodeName") = "MUKTI_ACCT_01" Or Session("CodeName") = "CORNELIUS_ACCT" Or Session("CodeName") = "RENRI_ACCT" Or Session("CodeName") = "STAFF_OFFICER_ACCT" Or Session("CodeName") = "SR_OFFICER_ACCT_01" Then%>
										<%if rscat("rowstate") = "Approved" then%>
											<center>
												<a href="javascript:if(confirm('Apakah Anda yakin ingin melakukan Replanned untuk Order No <%=rscat("ORDERNO")%>?')) { launchwin('ResetTransportTaskDiv.Asp?OrderNo=<%=rscat("ORDERNO")%>','newwindow','width='+(400)+',height='+(200)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=0'); }">
													<img src="https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcQ2X7OYNwyKXW8HKDtL3yBGG972_OEYpiY89g&s" alt="Reset Posting : <%=rscat("ORDERNO")%>" border="0" width="20" height="20">
												</a>		
											</center>
										<%Else%>
											<center>-</center>
										<%End If%>
									<%Else%>
										<center>-</center>
									<%End If%>
								</td>
								<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("ORDERDATE")%></font></td>
								<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("ORDERNO")%></font></td>
								<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("stockcode")%></font></td>
								<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("StockNAME")%></font></td>
								
								<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("remarks")%></font></td>
								<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("STORECODEFROM")%></font></td>
								<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("STORECODETO")%></font></td>
								<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("QUANTITYTRANSFERED"),3)%></font></td>
								<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("rowstate")%></font></td>
								<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=postingflg%></font></td>

							</tr>
							  <%
							  'v_Quantity=v_Quantity+trim(cdbl(rsCat("quantity")))
							  v_Quantity  =v_Quantity+cdbl(rsCat("QUANTITYTRANSFERED"))
							
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