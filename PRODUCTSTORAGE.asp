<%Option Explicit%>
<%response.buffer=True%>
<%validatesession%>

<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<html>
<head>
<title>ePlaS -  Product Storage</title>
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
<script type="text/javascript" src="js/fbp-shell.js?v=20261014b"></script>
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
<!---------------awal----------------------->
 <%
Dim FCENTRY, FCEDIT, FCIP, FCBA
Dim STORAGECODE, MAXCAPACITY, DESCRIPTION, PRODUCT_PRODUCTCODE, DEADSTOCK_CAPACITY,ABLETODIRECTDELIVERY,ISMOBILESTORAGE,ISAUTOSOUNDING,REMARKS
Dim cmdView,txt_STORAGECODE,ISQTYPROD
txttable=request("txttable")
 if Request("txtJenis")="New" or Request("txtJenis")= "Delete" or Request("txtJenis")="Edit" Then
 	txt_STORAGECODE=request("txt_STORAGECODE")
	STORAGECODE=request("STORAGECODE")
	MAXCAPACITY=request("MAXCAPACITY")
	DESCRIPTION=request("DESCRIPTION")
	PRODUCT_PRODUCTCODE=request("PRODUCT_PRODUCTCODE")
	DEADSTOCK_CAPACITY=request("DEADSTOCK_CAPACITY")
	ABLETODIRECTDELIVERY=request("ABLETODIRECTDELIVERY")
	ISMOBILESTORAGE=request("ISMOBILESTORAGE")
	ISAUTOSOUNDING=request("ISAUTOSOUNDING")
	ISQTYPROD=request("ISQTYPROD")
	REMARKS=request("REMARKS")
	
	
	

	if Request("txtJenis")= "Delete" then
       q="Delete from PRODUCTSTORAGE where STORAGECODE='" & STORAGECODE & "' And fcBA='" & Session("fcBA") & "'"
   	   connStr.execute(q)
       cmdView	= "True"
	   
	elseif Request("txtJenis")="New" then
			
	  q="Insert into PRODUCTSTORAGE (STORAGECODE, MAXCAPACITY, DESCRIPTION, PRODUCT_PRODUCTCODE, DEADSTOCK_CAPACITY,REMARKS,FCENTRY, FCEDIT, " & _
	    "FCIP, FCBA, LASTUPDATE, LASTTIME, ABLETODIRECTDELIVERY, ISMOBILESTORAGE, ISAUTOSOUNDING,ISQTYPROD) " & _
		"values ('" & STORAGECODE & "','" & MAXCAPACITY & "','" &  DESCRIPTION & "','" &  PRODUCT_PRODUCTCODE & "','" &  DEADSTOCK_CAPACITY & "','" &  REMARKS & "'," & _
		"'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
		Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, To_Char(sysdate,'HH24:MI'),'" & ABLETODIRECTDELIVERY & "','" & ISMOBILESTORAGE & "','" & ISAUTOSOUNDING & "','" & ISQTYPROD & "' )"
		'response.Write(q)
		response.Flush()
	    connStr.execute(q)
        cmdView	= "True"

	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
       q="Update PRODUCTSTORAGE set " & _
	    " STORAGECODE ='" & STORAGECODE & "'" & _
		" ,MAXCAPACITY ='" & MAXCAPACITY & "'" & _
		" ,DESCRIPTION='" & DESCRIPTION & "'" & _
		" ,PRODUCT_PRODUCTCODE ='" & PRODUCT_PRODUCTCODE & "'" & _
		" ,DEADSTOCK_CAPACITY ='" & DEADSTOCK_CAPACITY & "'" & _
		" ,ABLETODIRECTDELIVERY ='" & ABLETODIRECTDELIVERY & "'" & _
		" ,ISMOBILESTORAGE ='" & ISMOBILESTORAGE & "'" & _
		" ,ISAUTOSOUNDING ='" & ISAUTOSOUNDING & "'" & _
		" ,ISQTYPROD ='" & ISQTYPROD  & "'" & _
		" ,REMARKS ='" & REMARKS & "'" & _
		" , LASTUPDATE = sysdate " & _
		" , LASTTIME = To_Char(sysdate,'HH24:MI') " & _
		" where STORAGECODE='" & STORAGECODE & "'" & _ 
		" And fcBA='" & Session("fcBA") & "'"
		'RESPONSE.Write(Q)
		
		connStr.execute (q)
		connStr.execute("commit")
		'
		'RESPONSE.Flush()
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
%>
 <!--------------- ISI Payment Form ------------------>
<% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-archive" aria-hidden="true"></i>Header Product Storage</div><% end if %>
<form name="frmEntry" method="post" action="PRODUCTSTORAGE.asp">
	<table class="data" width="70%" cellpadding="0" cellspacing="0">
          <TR> 
            <td colspan="2"><div align="left"><font size="2" face="Verdana"><strong>MAINTENANCE PRODUCT STORAGE</strong></font></div></td>
          </TR>
		<tr>
			<td align="left">Storage Code</td><td align="left"><input type="text" size="15" name="txt_STORAGECODE" value="<%=REQUEST("txt_STORAGECODE")%>"  ></td>
		</tr>	
		<tr><td colspan="2"  align="left" class="fbp-doc-actions">
		<input Type="Hidden" Value="True" Name="cmdView">
		<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		<input Type="Submit"  name="Add" value="Select">&nbsp;
		<a href="PrintPRODUCTSTORAGE.Asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&txt_CUSTOMERCODE=<%=REQUEST("txt_STORAGECODE")%>" 
		title="Print Data" target="_blank"><img src="IMAGES/Print.gif" border="0"></a>		
		
		</td>
		</tr>
    </table>  
 </form>
 <%if cmdView = "True" Then
   Dim rsGAD,intRecordsPerPage,cColor
   set rsGAD = server.CreateObject("Adodb.recordset")
   'Field Numeric Jangan di Sum karena ditampilkan mesti level detail
   q="select STORAGECODE, DESCRIPTION, PRODUCT_PRODUCTCODE, MAXCAPACITY, DEADSTOCK_CAPACITY, REMARKS,ABLETODIRECTDELIVERY,ISMOBILESTORAGE,ISAUTOSOUNDING,ISQTYPROD, FCENTRY, FCEDIT, FCIP, FCBA from PRODUCTSTORAGE where " & _
     " fcBA='" & session("fcBA") & "'"
   Set rsGAD = Server.CreateObject("ADODB.Recordset")
	if  len(trim(request("txt_STORAGECODE"))) > 0 then
		q = q & " and STORAGECODE='" & request("txt_STORAGECODE") & "'"
	end if
	q = q
   
   
   'response.Write(q)
   'response.Flush()
   rsGAD.CursorType = 3
   rsGAD.CursorLocation = 3
   rsGAD.Open q, connstr %>
   
   <% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-edit" aria-hidden="true"></i>Detail Item</div><% end if %>
   <form action="PRODUCTSTORAGE.asp" method="post" name="eForm">
   <table class="data" cellpadding="0" cellspacing="0" width="50%">

		<%IF Request("txtJenis")="Edit" then%>
			<td align="left">Storage Code</td><td align="left"><input type="text" size="15" maxlength="15" name="STORAGECODE" value="<%=STORAGECODE%>" readonly="True" style="background:#999999"><font color="#FF0000">*</font></td>
		<%else%>
			<td align="left">Storage Code</td><td align="left"><input type="text" size="15" maxlength="15" name="STORAGECODE" value="<%=STORAGECODE%>"  ><font color="#FF0000">*</font></td>
		<%end if%>
		</tr>
		<tr>
		<td align="left">Description</td><td align="left"><input type="text" size="30" maxlength="50" name="DESCRIPTION" value="<%=DESCRIPTION%>"><font color="#FF0000">*</font></td>
		</TR>
		<tr>
		<%IF Request("txtJenis")="Edit" then%>
			<td align="left">Product Storage Code</td>
			<% 	Dim rsProductA
			q = "select P.PRODUCTCODE,P.PRODUCTDESCRIPTION from  PRODUCT P"
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
		<td align="left">Maximum Capacity (Kg)</td><td align="left"><input type="text" size="15" name="MAXCAPACITY" value="<%=MAXCAPACITY%>"><font color="#FF0000">*</font></td>
		</tr>
		<tr>
		<td align="left">Deadstock Capacity (Kg)</td><td align="left"><input type="text" size="15" name="DEADSTOCK_CAPACITY" value="<%=DEADSTOCK_CAPACITY%>"></td>
		</tr>
		<tr>
		<td align="left">Direct Delivery</td><td align="left">
		<select name="ABLETODIRECTDELIVERY" title="Mobile Storage Character a s/d z or 1 s/d 9">  
		<% IF ABLETODIRECTDELIVERY = "TRUE" then%>
		<option value="" disabled>-- No Selected --</option>
		<option value="TRUE" selected>TRUE</option>
		<option value="FALSE">FALSE</option>
		<%elseif ABLETODIRECTDELIVERY = "FALSE" then%>
		<option value="" disabled>-- No Selected --</option>
		<option value="TRUE">TRUE</option>
		<option value="FALSE" selected>FALSE</option>
		<%else%>
		<option value="" selected disabled>-- No Selected --</option>
		<option value="TRUE">TRUE</option>
		<option value="FALSE">FALSE</option>
		<%end if%>
		</select>
		</td>
		</tr>
		<tr>
		<td align="left">Mobile Storage</td><td align="left">
		<select name="ISMOBILESTORAGE" title="Mobile Storage Character a s/d z or 1 s/d 9">         
		<% IF ISMOBILESTORAGE = "TRUE" then%>
		<option value="" DISABLED>-- No Selected --</option>
		<option value="TRUE" SELECTED>TRUE</option>
		<option value="FALSE">FALSE</option>
		<%elseif ISMOBILESTORAGE = "FALSE" then%>
		<option value="" DISABLED>-- No Selected --</option>
		<option value="TRUE">TRUE</option>
		<option value="FALSE" SELECTED">FALSE</option>
		<%else%>
		<option value="" SELECTED DISABLED>-- No Selected --</option>
		<option value="TRUE">TRUE</option>
		<option value="FALSE">FALSE</option>
		<%end if%>
		</select>
		</td>
		</tr>
		<tr>
		<td align="left">Auto Sounding</td><td align="left">
		<select name="ISAUTOSOUNDING" title="Mobile Storage Character a s/d z or 1 s/d 9">         
		<% IF ISAUTOSOUNDING = "TRUE" then%>
		<option value="" DISABLED>-- No Selected --</option>
		<option value="TRUE" SELECTED>TRUE</option>
		<option value="FALSE">FALSE</option>
		<%elseif ISAUTOSOUNDING = "FALSE" then%>
		<option value="" DISABLED>-- No Selected --</option>
		<option value="TRUE">TRUE</option>
		<option value="FALSE" SELECTED">FALSE</option>
		<%else%>
		<option value="" SELECTED DISABLED>-- No Selected --</option>
		<option value="TRUE">TRUE</option>
		<option value="FALSE">FALSE</option>
		<%end if%>
		</select></td>
		</tr>
		<tr>
		<tr>
		<td align="left">Calculate Production</td><td align="left">
		<select name="ISQTYPROD" title="Mobile Storage Character a s/d z or 1 s/d 9">         
		<% IF ISQTYPROD = "Y" then%>
		<option value="" DISABLED>-- No Selected --</option>
		<option value="Y" SELECTED>Y</option>
		<option value="N">N</option>
		<%elseif ISQTYPROD = "N" then%>
		<option value="" DISABLED>-- No Selected --</option>
		<option value="Y">Y</option>
		<option value="N" SELECTED">N</option>
		<%else%>
		<option value="" SELECTED DISABLED>-- No Selected --</option>
		<option value="Y">Y</option>
		<option value="N">N</option>
		<%end if%>
		</select></td>
		</tr>
		<tr>
		<td align="left">Remarks</td><td align="left"><input type="text" size="15" maxlength="100" name="REMARKS" value="<%=REMARKS%>"></td>
		</tr>
	    <TR><td colspan="2">
		  <input type="hidden" name="txt_STORAGECODE" value="<%=REQUEST("txt_STORAGECODE")%>">		
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
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;">
	<TD align="left">NO</TD>
	<TD colspan="2">&nbsp;</TD>
	<TD align="left">STORAGE CODE</TD>
	<TD align="left">DESCRIPTION</TD>	
	<TD align="left">PRODUCT CODE</TD>		
	<TD align="left">MAX CAPACITY</TD>		
	<TD align="left">DEADSTOCK CAPACITY</TD>
	<TD align="left">REMARKS</TD>		
	<TD align="left">DIRECT DELIVERY</TD>	
	<TD align="left">MOBILE STORAGE</TD>	
	<TD align="left">AUTO SOUNDING</TD>	
	<TD align="left">CALCULATE PRODUCTION</TD>	
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

			<a href="PRODUCTSTORAGE.asp?txttable=<%=request("txttable")%>&
			<%for i = 0 to nCountofField %>
			<%=rsGAD(i).Name%>=<%=rsGAD(i)%>&
			<%next%>
			txtJenis=Edit"><img src="IMAGES/edit.gif" alt="Edit Data : <%=rsGAD(1)%>" border="0"></a>
			</td>
			<td valign="middle" width="20">
			<a href="PRODUCTSTORAGE.asp?STORAGECODE=<%=rsGAD("STORAGECODE")%>&txttable=<%=request("txttable")%>&txtJenis=Delete"><img src="IMAGES/delete.gif" alt="Delete Data : <%=rsGAD(1)%>" border="0" onClick="return confirm('Are you sure to Delete');"></a>
			</td>
		    <%for i = 0 to nCountofField %>
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
