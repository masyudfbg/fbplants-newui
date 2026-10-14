<%Option Explicit%>
<%response.buffer=True%>
<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include virtual="txtFunction.inc"-->
<html>
<head>
<title>RW Plant</title>
<script type="text/javascript" language="JavaScript1.2" src="stlib.js"></script>
<script language="JavaScript" src="date-picker.js"></script> 
</head>
<body>
<!--#include file="top.jsp"-->
<table width="100%">
<tr>
 <td width="22%" valign="top">
 <!--#include file="menu.asp"-->
 </td>
 <td width="78%" valign="top">
 <%'-------------awal-----------------------%>
 <%Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,fddate,fccode
 fddate   =Request("fddate")
 fccode   =Request("fccode")
 txtTable =Request("txtTable")
 txtModule=Request("txtModule")
 txtSearch=Request("txtSearch")
 txtField =Request("txtField")
 '-------------ISI Payment Form%>
 <form name="frmEntry" method="post" action="PaymentVoucher.asp">
	<table class="data">
		<tr><td align="left">Date : </td>
		<td align="left"><input type="text" maxlength="10" name="fdDate" size="10" value="<%=fdDate%>">
		<a href="javascript:show_calendar('frmEntry.fdDate');" onMouseOver="window.status='Date Picker';return true;" 
		onmouseout="window.status='';return true;">
		<img src="Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
		</td></tr>
		<tr><td  align="left">Voucher No : </td><td><input type="text" value="<%=fcCode%>" name="fcCode"></td></tr>
		<tr><td colspan="2"  align="left">
		<input Type="Hidden" Value="True" Name="cmdView">
		<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		<input Type="Submit" style="background-color: #298CF7; color: #F7C610;
		font-weight: bold; font-size: 8pt; font-family:Verdana" name="Add" value="Select">
		</td>
		</tr>
    </table>  
 </form>
<%if request("cmdView") = "True" Then%>
   <%
   Dim rsVoucher,intRecordsPerPage,cColor
   set rsVoucher = server.CreateObject("Adodb.recordset")
   q="select a.VOUCHERCODE, a.LOCATIONTYPE, a.LOCATIONCODE, a.JOBCODE, a.AMOUNT, a.REFERENCE, a.REMARKS" & _
     " from PaymentVoucherDetail a,PAYMENTVOUCHER b where " & _
	 " a.VOUCHERCODE = b.VOUCHERCODE and " & _
	 " a.VOUCHERCODE='" & trim(fccode) & "'" & _
     " and b.datecreated=to_date('" & trim(request("fddate")) & "','" & "MM/DD/YYYY" & "')"
   'response.Write q	
   Set rsVoucher = Server.CreateObject("ADODB.Recordset")
   rsVoucher.CursorType = 3
   rsVoucher.CursorLocation = 3
   rsVoucher.Open q, connstr%>
   <form action="SimpanTransaction.Asp" method="post" name="eForm">
   <table class="data">
	 <tr align="center"><td>Code</td><td>TT</td><td>TC</td><td>Job</td>
		 <td>Amount</td><td>Ref#</td><td>Remark</td>
	 <tr><td><input type="text" size="10" name="fcCode"></td>
		 <td><input type="text" size="10" name="fcLocationType"></td>
		 <td><input type="text" size="10" name="fcLocationCode"></td>
		 <td><input type="text" size="10" name="fcCode"></td>
		 <td><input type="text" size="10" name="fnAmount"></td>
		 <td><input type="text" size="10" name="fcReferenceno"></td>
		 <td><input type="text" size="10" name="fcRemark"></td>
	  </tr>	 
	  <tr>
	  <td colspan="7">
		  <input Type="Submit" style="background-color: #298CF7; color: #F7C610; font-weight: 
		  bold; font-size: 8pt; font-family:Verdana" name="Add" value="Add">
		  <input Type="Reset" style="background-color: #298CF7; color: #F7C610; font-weight:
		  bold; font-size: 8pt; font-family:Verdana" name="Reset" value="Reset">
		  <input Type="Button" style="background-color: #298CF7; color: #F7C610; font-weight: 
		  bold; font-size: 8pt; font-family:Verdana" value="Back" name="B2" OnClick="javascript:history.back(1)"></td>
	  </tr>
   </table>

   <table class="data">
   <%
   if Not rsVoucher.eof then
       nCountofField=rsVoucher.fields.count-1
	  no=1
	  While Not rsVoucher.EOF
		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if%>
		  <tr bgcolor="<%=cColor%>">
		  <td align="left" valign="top"><%=NO%></td>
			<td valign="middle" width="30">
			<a href="EditData.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>&txtCode=<%=rsVoucher(0)%>&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>"><img src="IMAGES/edit.gif" alt="Edit Data : <%=rsVoucher(1)%>" border="0"></a><a href="DeleteData.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>&txtCode=<%=rsVoucher(0)%>&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&iPageno=<%=Request("iPageNo")%>"><img src="IMAGES/delete.gif" alt="Delete Data : <%=rsVoucher(1)%>" border="0"></a></td>
		  <%for i = 0 to nCountofField%>
			<td align="left" title="<%=rsVoucher(i).Name%>" valign="top"><%=rsVoucher(i)%></td>
		  <%next%>
		  </tr>
		  <%
		  rsVoucher.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend%>
	  </TABLE>
   <%
   end if 
   rsVoucher.Close
   set rsVoucher = Nothing%>
   </form>
<%End if%>  
 
  <%'-------------akhir----------------------%>
 </td>
</tr>
<tr>
  <td colspan="2">
  <!--#include file="bottom.jsp"-->
  </td>
</tr>
</table>
</body>
</html>