<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
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
<title>ePlaS Reporting - Shift Report</title>
</head>
<%
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"

DATE_From= trim(request("DATE_From"))
DATE_To= trim(request("DATE_To"))


txtSQL = "SELECT ONDATE, PRODUCTCODE, PRICE " & _
		 " FROM PRICELISTTBS" & _
		 " where fcba='" & session("fcBA") & "'"

if  len(trim(DATE_From)) > 0 then
    txtSQL = txtSQL & " and ONDATE >=to_Date('" & DATE_From & "','MM/DD/YYYY')"
	txtSelect = " From " & FormatDateTime(DATE_From,1)
end if
if  len(trim(PRODUCTCODE)) > 0 then
    txtSQL =txtSQL &  " and PRODUCTCODE =PRODUCTCODE"
	xtSelect = txtSelect & "  to " & FormatDateTime(v_txtPRODUCTIONDATE_To,1)
end if

txtsql= txtsql & " Order By  ONDATE"
if len(trim(txtSelect)) > 0 then
   txtSelect = "Selected  by : " & txtSelect
else
   txtSelect = "Selected by All Data"
end if
'& " group by CHITNO, FIELDCODE, CROPCODE, ProductCODE, COLLECTIONDATE,HARVESTINGDATE,RECEPTIONDATE, " & _
'                     "VEHICLE, LOADERCODE, RAMPATTENDANTCODE,WEIGHT"

'response.Write(txtsql)
'response.Flush()
%>
 <body>
		<div id="wrapper">
        <!--#include file="print_excel.asp"-->
			 <table border="0" width="100%" id="product-table" class="xthf-site">
         <thead>
         <tr>
         <th colspan="13" style=" background-color:#FFFFFF; color:#000000; border:none">
			<div id="header-wrapper">
				<div id="head_judul" style="float:left; text-align:left;">
					<h4>PT. <%=Session("COMPANY")%></h4>
					<h5><%=Session("BA")%></h5>
				</div>
				<div id="subheader" class="centered" style="width:50%; margin:auto;">
					<h2>TBS PRICE LIST</h2>
				</div>
			</div>
</th>
			</tr>
						<!--<div class="printhide">-->
						<tr>
							<td colspan="13" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></td>
				</tr>
				<!--</div>-->
				<%	
				Set rsCat = connStr.Execute(txtSQL)
				if Not rsCat.eof then
					nCountofField=rsCat.fields.count-1
				end if
				NO=1%>
				<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;" >
					<TD align="left">PRODUCT CODE</TD>	
					<TD align="left">DATE</TD>	
					<TD align="left">PRICE</TD>	
				</tr>
                </thead>
               <tbody>
	<%
	Dim v_Quantity
	v_Quantity=0
	While Not rsCat.EOF%>
		<tr>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("PRODUCTCODE")%></td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("ONDATE")%>&nbsp;</td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("PRICE")%>&nbsp;</td>
		</tr>
	  <%
	  'v_Quantity=v_Quantity+trim(cdbl(rsCat("quantity")))
	  'response.write cdbl(trim(rsCat("MTHOPENBALANCE")))+1
	  rsCat.MoveNext
	  'response.Flush()
	  no=no+1
	Wend
	rsCat.close()
	rsCat=null%>
	<input  type="hidden" value="<%=txtTable%>" name="txtTable">
	<input  type="hidden" value="<%=txtModule%>" name="txtModule">
	</form>
    </tbody>
</table>
<!--#include file="Bottom_Report.asp"-->