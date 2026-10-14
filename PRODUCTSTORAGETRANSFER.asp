<%Option Explicit%>
<%response.buffer=True%>
<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="DailyValidation.inc"-->
<html>
<head>
<title>ePlaS - Product Storage Transfer</title>
<script type="text/javascript" language="JavaScript1.2" src="stlib.js" ></script>
<link rel="StyleSheet" href="dtree.css" type="text/css" />
<link rel="StyleSheet" href="style-content.css" type="text/css" />
<link rel="stylesheet" href="css/jquery-ui.css" />
<script type="text/javascript" src="dtree.js"></script>
<script src="/js/jquery-1.9.1.js"></script>
<script src="/js/jquery-ui.js"></script>
<script language="JavaScript">
xAddEventListener(window, 'load',
  function() {
    new xTableHeaderFixed('xthf-site', window);
    new xTableHeaderFixed('xthf-green', window);
    new xTableHeaderFixed('xthf-blue', window);
  }, false
);
</script>
 <script>
$(function() {
$( ".datepicker" ).datepicker();
});
</script>
<!--<script language="JavaScript" src="/js/date-picker.js"></script> -->
<script language="JavaScript">

var newwin;
function launchwin(winurl,winname,winfeatures)
{
	//This launches a new window and then
	//focuses it if window.focus() is supported.
	newwin = window.open(winurl,winname,winfeatures);

}
</script>
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
 <%
 '-------------awal-----------------------
 %>
 <%


Dim FCENTRY, FCEDIT, FCIP, FCBA
Dim TRANSFERNO, TRANSFERDATE, PSTORAGE_STORAGECODE, PSTORAGE_STORAGECODE_TO_TANK, QUANTITY
Dim cmdView,rsPT
dim Perioddate,txt_TRANSFERNO,VEHICLENO, PercentLosses
Dim tempQuantityStorageIn, tempQuantityLosses, Product_ProductCode
Dim Cek_IsDailyClosing : Cek_IsDailyClosing = ""

txttable=request("txttable")

	Perioddate=request("Perioddate")
	txt_TRANSFERNO=request("txt_TRANSFERNO")
	TRANSFERNO=request("TRANSFERNO")
	VEHICLENO=request("VEHICLENO")
	TRANSFERDATE=request("TRANSFERDATE")
	PSTORAGE_STORAGECODE=request("PSTORAGE_STORAGECODE")
	PSTORAGE_STORAGECODE_TO_TANK=request("PSTORAGE_STORAGECODE_TO_TANK")
	QUANTITY=request("QUANTITY")
	PERCENTLOSSES=request("PERCENTLOSSES")
	PRODUCT_PRODUCTCODE = request("PRODUCT_PRODUCTCODE")
	
 if Request("txtJenis")="New" or Request("txtJenis")= "Delete" or Request("txtJenis")="Edit" Then
	if Request("txtJenis")= "Delete" then
		Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), TRANSFERDATE)
		If Cek_IsDailyClosing = "FALSE" then
			q="Delete from PRODUCTSTORAGETRANSFER where TRANSFERNO='" & TRANSFERNO & "' And TRANSFERDATE=to_Date('" & TRANSFERDATE & "','MM/DD/YYYY') And PSTORAGE_STORAGECODE='" & PSTORAGE_STORAGECODE & "' And fcBA='" & Session("fcBA") & "'"
			connStr.execute(q)
		else
			err.number = 1001
			err.description = "Tanggal """ & TRANSFERDATE & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
		End If
		cmdView	= "True"
	elseif Request("txtJenis")="New" then
		Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), TRANSFERDATE)
		If Cek_IsDailyClosing = "FALSE" then
			tempQuantityStorageIn = Quantity * (1 - PercentLosses / 100)
			tempQuantityLosses = Quantity - tempQuantityStorageIn 
			q="Insert into PRODUCTSTORAGETRANSFER (TRANSFERNO, VEHICLENO,TRANSFERDATE,PRODUCT_PRODUCTCODE,PSTORAGE_STORAGECODE, PSTORAGE_STORAGECODE_TO_TANK, " & _
		     " QUANTITY, PERCENTAGE_LOSSES, QUANTITY_STORAGE_OUT, QUANTITY_STORAGE_IN, QUANTITY_LOSSES, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) " & _
				"values ('" & TRANSFERNO & "','" & VEHICLENO & "',to_date('" & Request("TRANSFERDATE") & "','mm/dd/yyyy') " & ",'" & PRODUCT_PRODUCTCODE & "','" &  PSTORAGE_STORAGECODE & "','" &  PSTORAGE_STORAGECODE_TO_TANK & "'," &  QUANTITY & "," & _
				PercentLosses & "," & Quantity & "," & tempQuantityStorageIn & "," & tempQuantityLosses & _
				",'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
				Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, To_Char(sysdate,'HH24:MI') )"
				'response.Write(q)
				'response.Flush()
				connStr.execute(q)
		else
			err.number = 1001
			err.description = "Tanggal """ & TRANSFERDATE & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
		End If
		cmdView	= "True"
	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
		Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), TRANSFERDATE)
		If Cek_IsDailyClosing = "FALSE" then
			tempQuantityStorageIn = Quantity * (1 - PercentLosses / 100)
			tempQuantityLosses = Quantity - tempQuantityStorageIn 

		   q=" Update PRODUCTSTORAGETRANSFER set " & _
			" VEHICLENO ='" & VEHICLENO & "'" & _
			" ,TRANSFERDATE =to_date('" & TRANSFERDATE & "','mm/dd/yyyy')" & _
			" ,PRODUCT_PRODUCTCODE='" & PRODUCT_PRODUCTCODE & "'" & _
			" ,PSTORAGE_STORAGECODE='" & PSTORAGE_STORAGECODE & "'" & _
			" ,PSTORAGE_STORAGECODE_TO_TANK ='" & PSTORAGE_STORAGECODE_TO_TANK & "'" & _
			" ,QUANTITY ='" & QUANTITY & "'" & _
			" ,PERCENTAGE_LOSSES = " & PercentLosses  & _
			" ,QUANTITY_STORAGE_OUT = " & Quantity  & _
			" ,QUANTITY_STORAGE_IN = " & tempQuantityStorageIn  & _
			" ,QUANTITY_LOSSES = " & tempQuantityLosses  & _
			" , LASTUPDATE = sysdate " & _
			" , LASTTIME = To_Char(sysdate,'HH24:MI') " & _
			" where TRANSFERNO='" & TRANSFERNO & "'" & _ 
			" And TRANSFERDATE=to_Date('" & TRANSFERDATE & "','MM/DD/YYYY')" & _ 
			" And PSTORAGE_STORAGECODE='" & PSTORAGE_STORAGECODE & "'" & _ 
			" And fcBA='" & Session("fcBA") & "'"
			connStr.execute(q)
			'response.Write(q)
			'response.Flush()
			connStr.execute("commit")
			'RESPONSE.Write "SQLnsddsdya : " & Q
		else
			err.number = 1001
			err.description = "Tanggal """ & TRANSFERDATE & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
		End If
        cmdView	= "True"
    else
        cmdView	= "True"
	end if								 
	'Response.Write "sql " & q
else
    cmdView	= request("cmdView")
End if

Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis

txtTable =Request("txtTable")
txtModule=Request("txtModule")
txtSearch=Request("txtSearch")
txtField =Request("txtField")
'fcCode=request("fcCode")
txtJenis = request("txtJenis")
If Err.Number <> 0 then
   response.Write  "<div class = ""error"">" & err.description & "</div>"
End if

 '-------------ISI Payment Form
%>
 <% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-exchange" aria-hidden="true"></i>Header Product Storage Transfer</div><% end if %>
 <form name="frmEntry" method="post" action="PRODUCTSTORAGETRANSFER.asp">
	<table class="data" width="70%" cellpadding="0" cellspacing="0">
        <TR> 
            <td colspan="2"><div align="left"><font size="2" face="Verdana"><strong>MAINTENANCE PRODUCT STORAGE TRANSFER</strong></font></div></td>
        </TR>
		<tr>
		    <td align="left">Date</td>
			<td align="left"><input class="datepicker" type="text" maxlength="10" name="Perioddate" size="10" value="<%=Perioddate%>"></td>
		</tr>	
		<tr>
			<td align="left">Transfer No</td><td align="left"><input type="text" size="15" maxlength="15" name="txt_TRANSFERNO" value="<%=REQUEST("txt_TRANSFERNO")%>"></td>
		</tr>

		<tr><td colspan="2"  align="left" class="fbp-doc-actions">
		<input Type="Hidden" Value="True" Name="cmdView">
		<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		<input Type="Submit"  name="Add" value="Select">&nbsp;&nbsp;
		<a href="PrintPRODUCTSTORAGETRANSFER.Asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&Perioddate=<%=Perioddate%>&txt_STORAGECODE=<%=REQUEST("txt_STORAGECODE")%>&txt_ADJUSTMENTNO=<%=REQUEST("txt_ADJUSTMENTNO")%>" 
		title="Print Data" target="_blank"><img src="IMAGES/Print.gif" border="0"></a>

		</td>
		</tr>
    </table>  
 </form>
 <%if cmdView = "True" Then
   Dim rsGAD,intRecordsPerPage,cColor
   set rsGAD = server.CreateObject("Adodb.recordset")
   'Field Numeric Jangan di Sum karena ditampilkan mesti level detail
   q="select TRANSFERDATE, TRANSFERNO,VEHICLENO, PSTORAGE_STORAGECODE, PSTORAGE_STORAGECODE_TO_TANK, PRODUCT_PRODUCTCODE, QUANTITY, PERCENTAGE_LOSSES, QUANTITY_STORAGE_OUT, " & _
     " QUANTITY_STORAGE_IN, QUANTITY_LOSSES, FCENTRY, FCEDIT, FCIP, FCBA from PRODUCTSTORAGETRANSFER where " & _
     " fcBA='" & session("fcBA") & "'"
	if  len(trim(request("Perioddate"))) > 0 then
		q = q & " and to_char(TRANSFERDATE,'mm/yyyy')=to_char(to_date('" & request("Perioddate") & "','" & "MM/DD/YYYY" & "'),'mm/yyyy')"
	end if
	if  len(trim(request("txt_TRANSFERNO"))) > 0 then
		q = q & " and TRANSFERNO='" & request("txt_TRANSFERNO") & "'"
	end if
	q = q
  
   'response.Write(q)
  'response.Flush()
   rsGAD.CursorType = 3
   rsGAD.CursorLocation = 3
   rsGAD.Open q, connstr %>
   <% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-edit" aria-hidden="true"></i>Detail Item</div><% end if %>
   <form action="PRODUCTSTORAGETRANSFER.asp" method="post" name="eForm">
   <table class="data" cellpadding="0" cellspacing="0" width="50%">
		<tr>
			<td align="left">Transfer Date</td>
			<td align="left">
				<%IF Request("txtJenis")="Edit" then%>
				<input type="text" maxlength="10" name="TRANSFERDATE" size="10" value="<%=TRANSFERDATE%>" readonly="True" style="background:#999999">
				<%Else%>
				<input class="datepicker" type="text" maxlength="10" name="TRANSFERDATE" size="10" value="<%=TRANSFERDATE%>">
				<%End IF%>
				<font color="#FF0000">*</font>
			</td>
		</tr>
		<tr>
			<td align="left">Transfer No</td>
			<td align="left">
				<%IF Request("txtJenis")="Edit" then%>
				<input type="text" size="20" maxlength="20" name="TRANSFERNO" value="<%=TRANSFERNO%>" readonly="True" style="background:#999999">
				<%Else%>
				<input type="text" size="20" maxlength="20" name="TRANSFERNO" value="<%=TRANSFERNO%>">
				<%End IF%>				
				<font color="#FF0000">*</font>
			</td>
		</tr>
		<tr>
			<td align="left">Vehicle No</td>
			<td align="left">
				<%IF Request("txtJenis")="Edit" then%>
				<input type="text" size="20" maxlength="20" name="VEHICLENO" value="<%=VEHICLENO%>" readonly="True" style="background:#999999">
				<%Else%>
				<input type="text" size="20" maxlength="20" name="VEHICLENO" value="<%=VEHICLENO%>">
				<%End IF%>				
				<font color="#FF0000">*</font>
			</td>
		</tr>
		<TR>
		<%IF Request("txtJenis")="Edit" then%>		
            <td align="left"><font face="Verdana" size="1">Storage Code (From)</font></td>
            <td align="left"> <select name="PSTORAGE_STORAGECODE">
                <%cmdCallComboSetup "PRODUCTSTORAGE",PSTORAGE_STORAGECODE,"fcba",session("fcba")%>
              </select> <font color="#FF0000">*</font> </td>
		<%else%>
				<td align="left"><font face="Verdana" size="1">Storage Code (From)</font></td>
				<td align="left"> <select name="PSTORAGE_STORAGECODE">
				<%cmdCallComboSetup "PRODUCTSTORAGE",PSTORAGE_STORAGECODE,"fcba",session("fcba")%>
				</select> <font color="#FF0000">*</font> </td>
		<%end if%>
		</tr>
		<TR>
		<%IF Request("txtJenis")="Edit" then%>		
            <td align="left"><font face="Verdana" size="1">Storage Code (To)</font></td>
            <td align="left"> <select name="PSTORAGE_STORAGECODE_TO_TANK">
                <%cmdCallComboSetup "PRODUCTSTORAGE",PSTORAGE_STORAGECODE_TO_TANK,"fcba",session("fcba")%>
              </select> <font color="#FF0000">*</font> </td>
		<%else%>
				<td align="left"><font face="Verdana" size="1">Storage Code (To)</font></td>
				<td align="left"> <select name="PSTORAGE_STORAGECODE_TO_TANK">
				<%cmdCallComboSetup "PRODUCTSTORAGE",PSTORAGE_STORAGECODE_TO_TANK,"fcba",session("fcba")%>
				</select> <font color="#FF0000">*</font> </td>
		<%end if%>
		</tr>
		
		<tr>
		<%IF Request("txtJenis")="Edit" then%>
			<td align="left">Product Storage Code</td>
			<% 	Dim rsProductA
			q = "select P.PRODUCTCODE,P.PRODUCTDESCRIPTION from  PRODUCT P where fcba = '" & Session("fcba") & "'"
			Set rsProductA = Server.CreateObject("ADODB.Recordset")
			rsProductA.CursorType = 3
			rsProductA.CursorLocation = 3
			rsProductA.Open q, connstr
			if not rsProductA.eof then%>
			<td align="left"> <select name="PRODUCT_PRODUCTCODE" title="PRODUCT_PRODUCTCODE Character a s/d z or 1 s/d 9">         
			<option value=""></option> 
			<%
			Do While Not rsProductA.eoF%>
			<%IF Request("PRODUCT_PRODUCTCODE")=rsProductA("PRODUCTCODE") Then%>
				<option value="<%=rsProductA("PRODUCTCODE")%>" selected><%=rsProductA("PRODUCTDESCRIPTION")%></option> 
			<%Else%>
				<option value="<%=rsProductA("PRODUCTCODE")%>"><%=rsProductA("PRODUCTDESCRIPTION")%></option> 
			<%End if%>
			<%rsProductA.movenext
			loop
			End if
			rsProductA.Close
			set rsProductA = Nothing
			q=""
			%>
			</select> <font color="#FF0000">*</font></td>
		<%Else%>			
			<td align="left">Product Storage Code</td>
			<% 	Dim rsProductc
			q = "select P.PRODUCTCODE,P.PRODUCTDESCRIPTION from  PRODUCT P where p.fcba= '"&session("fcba")&"'"
			Set rsProductc = Server.CreateObject("ADODB.Recordset")
			rsProductc.CursorType = 3
			rsProductc.CursorLocation = 3
			rsProductc.Open q, connstr
			if not rsProductc.eof then%>
			<td align="left"> <select name="PRODUCT_PRODUCTCODE" title="PRODUCT_PRODUCTCODE Character a s/d z or 1 s/d 9">         
			<option value=""></option> 
			<%
			Do While Not rsProductc.eoF%>
			<%IF Request("PRODUCT_PRODUCTCODE")=rsProductc("PRODUCTCODE") Then%>
				<option value="<%=rsProductc("PRODUCTCODE")%>" selected><%=rsProductc("PRODUCTDESCRIPTION")%></option> 
			<%Else%>
				<option value="<%=rsProductc("PRODUCTCODE")%>"><%=rsProductc("PRODUCTDESCRIPTION")%></option> 
			<%End if%>
			<%rsProductc.movenext
			loop
			End if
			rsProductc.Close
			set rsProductc = Nothing
			q=""
			%>
			</select> <font color="#FF0000">*</font></td>
		<%End if%>
		</tr>
		
		
		<tr>
		<td align="left">Quantity (Kg)</td><td align="left"><input type="text" size="15" name="QUANTITY" value="<%=QUANTITY%>">
		<font color="#FF0000">*</font></td>
		</tr>
		<tr>
		<td align="left">Percentage Losses (%) </td><td align="left"><input type="text" size="15" name="PERCENTLOSSES" value="<%=PERCENTLOSSES%>">
		<font color="#FF0000">*</font></td>
		</tr>
	    <TR><td colspan="2" align="left">
		  <input Type="Hidden" Value="<%=Perioddate%>" Name="Perioddate">
		  <input Type="Hidden" Value="<%=txt_TRANSFERNO%>" Name="txt_TRANSFERNO">
		  <input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		  <input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		  <%IF Request("txtJenis")="Edit" then%>
		       <input type="hidden" name="txtJenis" value="Edit">
		       <input Type="Submit"  name="Add" value="Update">
		  <%Else%>
		       <input type="hidden" name="txtJenis" value="New">
		       <input Type="Submit"  name="Add" value="Add">
		  <%End if%>
		  <input Type="Reset"  name="Reset" value="Reset">
		  <input Type="Button"  value="Back" name="B2" OnClick="javascript:history.back(1)"></td>
	  </tr>
   </table>
   </form>
   <!--#include file="validation_show.Asp"--> 
<%
	Dim intNumberOfPages,intRecords,nPage,howmanyfields,JumRec
	howmanyfields=rsGAD.fields.count-1
	
	JumRec=1
	intRecordsPerPage = 20
	
	rsGAD.PageSize = intRecordsPerPage
	intNumberOfPages = 0
	intNumberOfPages =  rsGAD.PageCount
	intRecords = rsGAD.RecordCount
	'Set our AbsolutePage property.
	if not rsGAD.EOF then
		If (isEmpty(Request("iPageNo")) or trim(Request("iPageNo"))="") Then
		  rsGAD.AbsolutePage = 1
		Else
		  rsGAD.AbsolutePage = Request("iPageNo")
		End If  
	end if
	if Request("iPageNo") = "" then
		nPage = 1
	Else
		nPage = Request("iPageNo")
	End if
   rsGAD.PageSize = intRecordsPerPage
   intNumberOfPages = 0
   intNumberOfPages =  rsGAD.PageCount
   intRecords = rsGAD.RecordCount   
   %>
   <% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-list" aria-hidden="true"></i>Daftar Data</div><% end if %>
   <div class="fbp-doc-scroll">
   <table class="data fbp-doc-lines" >
	<thead>
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;" >
	<TD align="left">NO</TD>
	<TD colspan="2">&nbsp;</TD>
	<TD align="left">DATE</TD>	
	<TD align="left">TRANSFER NO</TD>
    <TD align="left">VEHICLE NO</TD>	
	<TD align="left">STORAGE CODE</TD>	
	<TD align="left">STORAGE CODE TO TANK</TD>	
	<TD align="left">PRODUCT</TD>	
	<TD align="left">QUANTITY (Kg)</TD>	
	<TD align="left">% LOSSES</TD>	
	<TD align="left">QTY OUT</TD>	
	<TD align="left">QTY IN</TD>	
	<TD align="left">QTY LOSSES</TD>	
	
	</tr>
	</thead>
	<tbody>

        <%
   if Not rsGAD.eof then
      nCountofField=rsGAD.fields.count-5
	  no=1
	  While Not rsGAD.EOF
		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if%>
		  <tr bgcolor="<%=cColor%>">
		  <td width="79" align="left" valign="top"><%=NO%></td>
			<td valign="middle" width="21">
			<a href="PRODUCTSTORAGETRANSFER.asp?Perioddate=<%=Perioddate%>&txt_TRANSFERNO=<%=txt_TRANSFERNO%>&PRODUCT_PRODUCTCODE=<%=rsGAD("PRODUCT_PRODUCTCODE")%>&TRANSFERNO=<%=rsGAD("TRANSFERNO")%>&VEHICLENO=<%=rsGAD("VEHICLENO")%>&txttable=<%=request("txttable")%>&TRANSFERDATE=<%=rsGAD("TRANSFERDATE")%>&PSTORAGE_STORAGECODE=<%=rsGAD("PSTORAGE_STORAGECODE")%>&PSTORAGE_STORAGECODE_TO_TANK=<%=rsGAD("PSTORAGE_STORAGECODE_TO_TANK")%>&QUANTITY=<%=rsGAD("QUANTITY")%>&PERCENTLOSSES=<%=rsGAD("PERCENTAGE_LOSSES")%>&txtJenis=Edit"><img src="IMAGES/edit.gif" alt="Edit Data : <%=rsGAD(1)%>" border="0"></a>
			</td>
			<td valign="middle" width="20">
			<a href="PRODUCTSTORAGETRANSFER.asp?Perioddate=<%=Perioddate%>&txt_TRANSFERNO=<%=txt_TRANSFERNO%>&TRANSFERNO=<%=rsGAD("TRANSFERNO")%>&txttable=<%=request("txttable")%>&TRANSFERDATE=<%=rsGAD("TRANSFERDATE")%>&PSTORAGE_STORAGECODE=<%=rsGAD("PSTORAGE_STORAGECODE")%>&txtJenis=Delete" onClick="return confirm('Are you sure to Delete');"><img src="IMAGES/delete.gif" alt="Delete Data : <%=rsGAD(1)%>" border="0"></a>
			</td>
		    <%for i = 0 to nCountofField%>
			    <td width="597" align="left" valign="top" title="<%=rsGAD(i).Name%>"><%=rsGAD(i)%></td>
		    <%next%>
		  </tr>
		  <%
		  rsGAD.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend
   end if 
   rsGAD.Close
   set rsGAD = Nothing%>
   </tbody>
   </TABLE>
   </div>
<%End if%>  
 
  <%'-------------akhir----------------------%>
</div></div>
  <!--#include file="bottom.jsp"-->

</body>
</html>
