<%Option Explicit%>
<%'On Error resume Next%>
<%response.buffer=True%>
<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<html>
<head>
<title>ePlaS - Delivery Contract</title>
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
</head>
<body>
<!--#include file="top.jsp"-->
	<div style=" width:21.7%; float:left;margin-left:0.3%; margin-top:0.3%; margin-bottom:0.3%;">
			<!--#include file="menu.asp"-->
		</div>
   <div class="judul_head" valign="top" bgcolor="#fff" style=" width:77.4%; margin-left:0.3%; margin-right:0.3%; margin-top:0.3%; margin-bottom:0.3%; float:left; border:none;"> 
 <div class="judul">
	   Module : <%=request("txtModule")%> / <%=request("txtTable")%>
	</div>
 <div class="box_both" style="overflow:scroll;">
 <%'-------------awal-----------------------%>
 <%


Dim FCENTRY, FCEDIT, FCIP, FCBA
Dim CONTRACTCODE, STARTDATE, ENDDATE, ACTUALSTART, ACTUALEND, TOTALWEIGHT, txtSTATUS, REFERENCE, CUST_CUSTOMERCODE, PRODUCT_PRODUCTCODE
Dim cmdView,txt_CONTRACTCODE, TEXTFIELDID
txttable=request("txttable")
response.write Request("txtJenis")
 if Request("txtJenis")="New" or Request("txtJenis")= "Delete" or Request("txtJenis")="Edit" Then
 	txt_CONTRACTCODE=request("txt_CONTRACTCODE")
 	CONTRACTCODE=request("CONTRACTCODE")
	CUST_CUSTOMERCODE=request("CUST_CUSTOMERCODE")
	STARTDATE=request("STARTDATE")
	ENDDATE=request("ENDDATE")
	ACTUALSTART=request("ACTUALSTART")
	ACTUALEND=request("ACTUALEND")
	PRODUCT_PRODUCTCODE=request("PRODUCT_PRODUCTCODE")
	TOTALWEIGHT=request("TOTALWEIGHT")
	txtSTATUS=request("txtSTATUS")
	REFERENCE=request("REFERENCE")
	
	if Request("txtJenis")= "Delete" then
       q="Delete from DELIVERYCONTRACT where CONTRACTCODE='" & CONTRACTCODE & "' And STARTDATE=to_Date('" & STARTDATE & "','MM/DD/YYYY') And ENDDATE=to_Date('" & ENDDATE & "','MM/DD/YYYY') And fcBA='" & Session("fcBA") & "'"
   	   connStr.execute(q)
       cmdView	= "True"
	   
	elseif Request("txtJenis")="New" then
			
	  q="Insert into DELIVERYCONTRACT (FIELDID,CONTRACTCODE, CUST_CUSTOMERCODE, STARTDATE, ENDDATE, ACTUALSTART, ACTUALEND, PRODUCT_PRODUCTCODE, TOTALWEIGHT, RowStateId, REFERENCE, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) " & _
		"values (Random_GUID,'" & CONTRACTCODE & "','" &  CUST_CUSTOMERCODE & "',to_date('" & Request("STARTDATE") & "','mm/dd/yyyy') " & ",to_date('" & Request("ENDDATE") & "','mm/dd/yyyy') " & "," & _
		" to_date('" & Request("ACTUALSTART") & "','mm/dd/yyyy') " & ",to_date('" & Request("ACTUALEND") & "','mm/dd/yyyy') " & "," & _
		"'" &  PRODUCT_PRODUCTCODE & "','" &  TOTALWEIGHT & "','Planned','" &  REFERENCE & "'," & _
		"'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
		Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, To_Char(sysdate,'HH24:MI') )"
		'response.Write(q)
		'response.Flush()
	    connStr.execute(q)
		connStr.execute("commit")
        cmdView	= "True"
	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
       q="Update DELIVERYCONTRACT set " & _
	    " CONTRACTCODE ='" & CONTRACTCODE & "'" & _
		" ,STARTDATE =to_date('" & Request("STARTDATE") & "','mm/dd/yyyy') " & " " & _
		" ,ENDDATE=to_date('" & Request("ENDDATE") & "','mm/dd/yyyy') " & " " & _
		" ,ACTUALSTART=to_date('" & Request("ACTUALSTART") & "','mm/dd/yyyy') " & " " & _
		" ,ACTUALEND=to_date('" & Request("ACTUALEND") & "','mm/dd/yyyy') " & " " & _
		" ,TOTALWEIGHT='" & TOTALWEIGHT & "'" & _
		" ,STATUS ='" & txtSTATUS & "'" & _
		" ,REFERENCE ='" & REFERENCE & "'" & _
		" ,CUST_CUSTOMERCODE ='" & CUST_CUSTOMERCODE & "'" & _
		" ,PRODUCT_PRODUCTCODE ='" & PRODUCT_PRODUCTCODE & "'" & _
		" , LASTUPDATE = sysdate " & _
		" , LASTTIME = To_Char(sysdate,'HH24:MI') " & _
		" where CONTRACTCODE='" & CONTRACTCODE & "'" & _ 
		" And STARTDATE=to_Date('" & STARTDATE & "','MM/DD/YYYY')" & _ 
		" And ENDDATE=to_Date('" & ENDDATE & "','MM/DD/YYYY') " & _
		" And fcBA='" & Session("fcBA") & "'"
		connStr.execute(q)
		'response.Write(q)
		'response.Flush()
		connStr.execute("commit")
		'RESPONSE.Write "SQLnsddsdya : " & Q
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
   Response.Write "<font color=red>" & "Error Number : " & Err.Number &  " Desc. " & Err.Description & "</font>"
End if

 '-------------ISI Payment Form%>
<form name="frmEntry" method="post" action="DELIVERYCONTRACT.asp">
	<table class="data" width="70%" cellpadding="0" cellspacing="0">
           <TR> 
            <td colspan="2"><div align="left"><font size="2" face="Verdana"><strong>MAINTENANCE DELIVERY CONTRACT</strong></font></div></td>
          </TR>
		<tr>
			<td align="left">Contract Code</td><td align="left"><input type="text" size="25" name="txt_CONTRACTCODE" value="<%=REQUEST("txt_CONTRACTCODE")%>"  ></td>
		</tr>	
		<tr><td colspan="2"  align="left">
		<input Type="Hidden" Value="True" Name="cmdView">
		<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		<input Type="Submit"  name="Add" value="Select">&nbsp;
		<a href="PrintDELIVERYCONTRACT.Asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&txt_CUSTOMERCODE=<%=REQUEST("txt_CONTRACTCODE")%>" 
		title="Print Data" target="_blank"><img src="IMAGES/Print.gif" border="0"></a>		
		
		</td>
		</tr>
    </table>  
 </form>
  <%if cmdView = "True" Then
   Dim rsGAD,intRecordsPerPage,cColor
   set rsGAD = server.CreateObject("Adodb.recordset")
   'Field Numeric Jangan di Sum karena ditampilkan mesti level detail
   q="select CONTRACTCODE, CUST_CUSTOMERCODE,customer_api.get_description(cust_customercode,fcba) Custname, STARTDATE, ENDDATE, ACTUALSTART, ACTUALEND, " & _
     "PRODUCT_PRODUCTCODE, TOTALWEIGHT, RowStateId, REFERENCE, FCENTRY, FCEDIT, FCIP, FCBA,FIELDID from DELIVERYCONTRACT where " & _
     " fcBA='" & session("fcBA") & "'"
   Set rsGAD = Server.CreateObject("ADODB.Recordset")
	if  len(trim(request("txt_CONTRACTCODE"))) > 0 then
		q = q & " and CONTRACTCODE='" & request("txt_CONTRACTCODE") & "'"
	end if
	q = q
   
   'response.Write(q)
   'response.Flush()
   rsGAD.CursorType = 3
   rsGAD.CursorLocation = 3
   rsGAD.Open q, connstr %>
   <form action="DELIVERYCONTRACT.asp" method="post" name="eForm">
   <table class="data" cellpadding="0" cellspacing="0" width="50%">


		<%IF Request("txtJenis")="Edit" then%>
			<td align="left">Contract Code</td><td align="left"><input type="text" size="25" maxlength="25" name="CONTRACTCODE" value="<%=CONTRACTCODE%>" readonly="True" style="background:#999999"><font color="#FF0000">*</font></td>
		<%else%>
			<td align="left">Contract Code</td><td align="left"><input type="text" size="25" maxlength="25" name="CONTRACTCODE" value="<%=CONTRACTCODE%>"  ><font color="#FF0000">*</font></td>
		<%end if%>
		</tr>
		<tr>
		<%IF txtJenis="Edit" Then%>
			<td align="left">Customer Code</td>
			<td align="left"> <select name="CUST_CUSTOMERCODE">
			<%cmdCallComboSetup "CUSTOMER",request("CUST_CUSTOMERCODE"),"fcba",session("fcba")%>
			</select> <font color="#FF0000">*</font></td>
		<%Else%>
			<td align="left">Customer Code</td>
			<td align="left"> <select name="CUST_CUSTOMERCODE">
			<%cmdCallComboSetup "CUSTOMER","CUSTOMERCODE","fcba",session("fcba")%>
			</select> <font color="#FF0000">*</font></td>
		<%End if%>
		</tr>
		<tr>
			<td align="left">Start Date</td>
			<td align="left"><input class="datepicker" type="text" maxlength="10" name="STARTDATE" size="10" value="<%=STARTDATE%>">
			<font color="#FF0000">*</font>
			</td>
		</tr>
		<tr>
			<td align="left">End Date</td>
			<td align="left"><input class="datepicker" type="text" maxlength="10" name="ENDDATE" size="10" value="<%=ENDDATE%>">
			<font color="#FF0000">*</font>
			</td>
		</tr>
		<tr>
			<td align="left">Actual Start</td>
			<td align="left"><input class="datepicker" type="text" maxlength="10" name="ACTUALSTART" size="10" value="<%=ACTUALSTART%>">
			<font color="#FF0000">*</font>
			</td>
		</tr>
		<tr>
			<td align="left">Actual End</td>
			<td align="left"><input class="datepicker" type="text" maxlength="10" name="ACTUALEND" size="10" value="<%=ACTUALEND%>">
			<font color="#FF0000">*</font>
			</td>
		</tr>
		<tr>
		<%IF txtJenis="Edit" Then%>
			<td align="left">Product Code</td>
			<td align="left"> <select name="PRODUCT_PRODUCTCODE">
			<%cmdCallComboSetup "PRODUCT",request("PRODUCT_PRODUCTCODE"),"fcba",session("fcba")%>
			</select> <font color="#FF0000">*</font> </td>
		<%Else%>
			<td align="left">Product Code</td>
			<td align="left"> <select name="PRODUCT_PRODUCTCODE">
			<%cmdCallComboSetup "PRODUCT","ProductCODE","fcba",session("fcba")%>
			</select> <font color="#FF0000">*</font> </td>
		<%End if%>
	 </tr>		
		<tr>
			<td align="left">Total Weight</td><td align="left"><input type="text" size="15" name="TOTALWEIGHT" value="<%=TOTALWEIGHT%>"><font color="#FF0000">*</font></td>
		</tr>
		<tr>
			<td align="left">Status</td><td align="left"><input type="text" size="10" maxlength="10" name="txtSTATUS" value="<%=txtSTATUS%>"><font color="#FF0000">*</font></td>
		</tr>
		<tr>
			<td align="left">Reference</td><td align="left"><input type="text" size="50" maxlength="100" name="REFERENCE" value="<%=REFERENCE%>"></td>
		</tr>
	    <TR><td colspan="2" align="left">
		  <input type="hidden" name="txt_CONTRACTCODE" value="<%=REQUEST("txt_CONTRACTCODE")%>">				
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
   <table width="737" class="data" >
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;" >
	<TD align="left">NO</TD>
	<TD colspan="2">&nbsp;</TD>
	<TD align="left">CONTRACT</TD>	
	<TD align="left">CUSTOMER CODE</TD>	
	<TD align="left">CUSTOMER NAME</TD>	
	<TD align="left">START DATE</TD>	
	<TD align="left">END DATE</TD>	
	<TD align="left">ACTUAL START</TD>	
	<TD align="left">ACTUAL END</TD>	
	<TD align="left">PRODUCT</TD>	
	<TD align="left">TOTAL WEIGHT</TD>	
	<TD align="left">STATUS</TD>	
	<TD align="left">REFERENCES</TD>	
	</tr>

        <%
   if Not rsGAD.eof then
      nCountofField=rsGAD.fields.count-6
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
			<a href="DELIVERYCONTRACT.asp?CONTRACTCODE=<%=rsGAD("CONTRACTCODE")%>&txtField=<%=rsGAD("FIELDID")%>&txttable=<%=request("txttable")%>&STARTDATE=<%=rsGAD("STARTDATE")%>&ENDDATE=<%=rsGAD("ENDDATE")%>&ACTUALSTART=<%=rsGAD("ACTUALSTART")%>&ACTUALEND=<%=rsGAD("ACTUALEND")%>&TOTALWEIGHT=<%=rsGAD("TOTALWEIGHT")%>&txtSTATUS=<%=rsGAD("RowStateId")%>&REFERENCE=<%=rsGAD("REFERENCE")%>&CUST_CUSTOMERCODE=<%=rsGAD("CUST_CUSTOMERCODE")%>&PRODUCT_PRODUCTCODE=<%=rsGAD("PRODUCT_PRODUCTCODE")%>&txtJenis=Edit"><img src="IMAGES/edit.gif" alt="Edit Data : <%=rsGAD(1)%>" border="0"></a>
			</td>
			<td valign="middle" width="20">
			<a href="DELIVERYCONTRACT.asp?CONTRACTCODE=<%=rsGAD("CONTRACTCODE")%>&txtField=<%=rsGAD("FIELDID")%>&txttable=<%=request("txttable")%>&STARTDATE=<%=rsGAD("STARTDATE")%>&ENDDATE=<%=rsGAD("ENDDATE")%>&txtJenis=Delete"><img src="IMAGES/delete.gif" alt="Delete Data : <%=rsGAD(1)%>" border="0" onClick="return confirm('Are you sure to Delete');"></a>
			</td>
		    <%for i = 0 to nCountofField%>
			    <% If i = 0 then %>
						<td valign="middle" width="597">
							<a target="_blank" href="DeliveryContractDetails.Asp?CONTRACTCODE=<%=rsGAD("CONTRACTCODE")%>&STARTDATE=<%=rsGAD("STARTDATE")%>&customercode=<%=rsGAD("CUST_CUSTOMERCODE")%>&customername=<%=rsGAD("Custname")%>&productcode=<%=rsGAD("PRODUCT_PRODUCTCODE")%>&quantitycontract=<%=rsGAD("TOTALWEIGHT")%>
							"<%=rsGAD(i).Name%>"><%=rsGAD(i)%>
							</a>						
	                    </td>
				<% else %>
						<td width="597" align="left" valign="top" title="<%=rsGAD(i).Name%>"><%=rsGAD(i)%>
				</td>
				<% End If%>
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
	</div> 
   </TABLE>
<%End if%>  
 
  <%'-------------akhir----------------------%>
</div></div>
  <!--#include file="bottom.jsp"-->

</body>
</html>
