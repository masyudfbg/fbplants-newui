<%Option Explicit%>
<%response.buffer=True%>
<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<!--#include file="DailyValidation.inc"-->
<html>
<head>
<title>ePlaS - Konstanta Ratio</title>
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
<!--<script language="JavaScript" src="/js/date-picker.js"></script>-->
<script language="JavaScript">

var newwin;
function launchwin(winurl,winname,winfeatures)
{
	//This launches a new window and then
	//focuses it if window.focus() is supported.
	newwin = window.open(winurl,winname,winfeatures);

}
</script>
<script type='text/javascript'>
    //Replaces all non numeric characters
    function checkInput(element){ element.value = element.value.replace(/[^0-9\.\-]/g,'');}
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
<%
'-------------awal-----------------------
Dim datefrom, dateuntil, productcode, ownership, ratio, txtLineNoKey
Dim cmdView, rsBA, qBA
Dim qCheckStatus, txtClosingStatus, rsCheckStatus
Dim valueAorD, objectRS, objectSQL
Dim Cek_IsDailyClosing : Cek_IsDailyClosing = ""
Dim rsmaxlinenokey, sqlmaxlinenokey
Dim rsdatefrom, sqldatefrom, cekdatefrom, rsdateuntil, sqldateuntil, cekdateuntil
 
Dim v_txtYear : v_txtYear = Request("txtYear")

productcode = request("productcode")


Dim MinYr, MaxYr, mon, Yr
Yr = Year(Session("v_Mill_Production"))
mon = Month(Session("v_Mill_Production"))
MinYr = Yr - 2
MaxYr = Yr  + 2

If len(v_txtyear)= 0 then v_txtyear = Yr
	
datefrom = request("datefrom")
dateuntil = request("dateuntil")
productcode = request("productcode")
ownership = request("ownership")
ratio = request("ratio")

txtLineNoKey = request("txtLineNoKey")
 
if Request("txtJenis")="New" or Request("txtJenis")= "Delete" or Request("txtJenis")="Edit" Then
	if Request("txtJenis")= "Delete" then		
		q = " delete from ratio_constanta_tab" & _
			" where fcba = '" & Session("fcBA") & "'" & _
			"       and linenokey = '" & txtLineNoKey & "'"
		connStr.execute(q)
		cmdView	= "True"
	elseif Request("txtJenis")="New" then
		sqlmaxlinenokey = " select nvl(max(linenokey), 0) linenokey" & _
		                  " from ratio_constanta_tab" & _
						  " where fcba = '" & Session("FCBA") & "'"
		
		Set rsmaxlinenokey = Server.CreateObject("ADODB.Recordset")
		rsmaxlinenokey.Open sqlmaxlinenokey, connstr
		
		If Not rsmaxlinenokey.Eof Then
			txtLineNoKey = Cdbl(rsmaxlinenokey("linenokey"))
		Else
			txtLineNoKey = 0
		End If
		rsmaxlinenokey.Close
		set rsmaxlinenokey = Nothing
		
		sqldatefrom	= " select count(1) record from ratio_constanta_tab" & _
					  " where fcba = '" & Session("FCBA") & "'" & _
					  "       and to_date('" & datefrom & "', 'mm/dd/yyyy') >= date_from" & _
					  "       and to_date('" & datefrom & "', 'mm/dd/yyyy') <= date_until"
		
		Set rsdatefrom = Server.CreateObject("ADODB.Recordset")
		rsdatefrom.Open sqldatefrom, connstr
		
		If Not rsdatefrom.Eof Then
			cekdatefrom = Cdbl(rsdatefrom("record"))
		Else
			cekdatefrom = 0
		End If
		rsdatefrom.Close
		set rsdatefrom = Nothing
		
		sqldateuntil = " select count(1) record from ratio_constanta_tab" & _
					   " where fcba = '" & Session("FCBA") & "'" & _
					   "       and to_date('" & dateuntil & "', 'mm/dd/yyyy') >= date_from" & _
					   "       and to_date('" & dateuntil & "', 'mm/dd/yyyy') <= date_until"
		
		Set rsdateuntil = Server.CreateObject("ADODB.Recordset")
		rsdateuntil.Open sqldateuntil, connstr
		
		If Not rsdateuntil.Eof Then
			cekdateuntil = Cdbl(rsdateuntil("record"))
		Else
			cekdateuntil = 0
		End If
		rsdateuntil.Close
		set rsdateuntil = Nothing
		
		if cekdatefrom = 0 then
			if cekdateuntil = 0 then

				txtLineNoKey = txtLineNoKey + 1
				
				q = " Insert into ratio_constanta_tab (linenokey" & _
					"                                  , date_from" & _
					"                                  , date_until" & _
					"                                  , productcode" & _
					"                                  , ownership" & _
					"                                  , ratio" & _
					"                                  , fcentry" & _
					"                                  , fcedit" & _
					"                                  , fcip" & _
					"                                  , fcba" & _
					"                                  , lastupdate" & _
					"                                  , lasttime) " & _
					" values ('" & txtLineNoKey & "'" & _
					"         , to_date('" & datefrom & "','mm/dd/yyyy') " & _
					"         , to_date('" & dateuntil & "','mm/dd/yyyy') " & _
					"         , '" & productcode & "'" & _
					"         , '" & ownership & "'" & _
					"         , " & ratio & "" & _
					"         , '" & Session("CodeName") & "'" & _
					"         , '" & Session("CodeName") & "'" & _
					"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
					"         , '" & Session("fcBA") & "'" & _
					"         , sysdate" & _
					"         , To_Char(sysdate,'HH24:MI') )"
						
				'Response.Write q
				connStr.execute(q)
			else
				err.number = 1001
				err.description = "Tanggal """ & dateuntil & """, sudah ada dalam record, Silahkan rubah datanya!"
			end if
			
		else
			err.number = 1001
			err.description = "Tanggal """ & datefrom & """, sudah ada dalam record, Silahkan rubah datanya!"
		end if
		
		cmdView	= "True"
		
	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
			
		q = " Update ratio_constanta_tab" & _
			" set ratio = " & ratio & "" & _
			"     , fcedit = '" & Session("CodeName") & "'" & _
			"     , lastupdate = sysdate " & _
			"     , lasttime = to_char(sysdate,'hh24:mi') "& _
			" where fcba = '" & Session("fcBA") & "'" & _
			"       and linenokey = '" & txtLineNoKey & "'"
		
		'Response.Write q
		connStr.execute(q)
		connStr.execute("commit")
		cmdView	= "True"
	else
		cmdView	= "True"
	end if
else
    cmdView	= request("cmdView")
End if


Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis
if len(trim(datefrom)) <> 0 then
	datefrom = Request("datefrom")
else
	datefrom = textdate(date)
end if

if len(trim(dateuntil)) <> 0 then
	dateuntil = Request("dateuntil")
else
	dateuntil = textdate(date)
end if

if Request("ratio")<>"" then
   ratio = Request("ratio")
else
   ratio = 0
end if

txtTable = Request("txtTable")
txtModule = Request("txtModule")
txtSearch = Request("txtSearch")
txtField = Request("txtField")
txtJenis = request("txtJenis")

if err.number <> 0 then
   response.Write  "<div class = ""error"">" & err.description & "</div>"
End if
 '-------------ISI Payment Form

%>
 <form name="frmEntry" method="post" action="ratio_constanta.asp">
	<table class="data" width="100%">
		<tr>
			<td align="left" width="150">Year/ Tahun</td>
			<td align="left">
				<select name="txtyear">
				<%For i = MinYr to MaxYr%>
					<Option Value="<%=i%>" <%If cint(i) = cint(v_txtyear) then%> Selected <%End If%>><%=i%></option>
				<%Next%>
				</select>
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
		<tr>
			<td  align="left">Product</td>
			<td align="left">
				<select name="productcode" >
					<option value=""></option>
					<%
					Dim rsproductcode, strproduct
					strproduct = " select productcode fccode, productdescription fcname" & _
                                 " from product" & _
                                 " where fcba = '" & Session("fcBA") & "'" & _
                                 " order by productdescription asc"
					
					Set rsproductcode = Server.CreateObject("ADODB.Recordset")
					rsproductcode.Open strproduct, connstr, adOpenStatic 
					
					if not rsproductcode.eof then
					Do While Not rsproductcode.eof
					%>
						<option value="<%=rsproductcode("fccode")%>" <%if rsproductcode("fccode") = productcode then%> Selected <%end if%>><%=rsproductcode("fccode")%> : <%=rsproductcode("fcname")%></option>
					<%
						rsproductcode.movenext
					loop
					end if
					rsproductcode.Close
					set rsproductcode = Nothing
					%>
				</select>
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
		<tr>
			<td  align="left">Ownership</td>
			<td align="left">
				<select name="ownership" >
					<option value = "" <%If Len(Trim(ownership)) = 0 Then %> Selected <%End If%> ></option>
					<option value = "INTI" <%If ownership = "INTI" Then %> Selected <%End If%> >INTI</option>
					<option value = "PLASMA" <%If ownership = "PLASMA" Then %> Selected <%End If%> >PLASMA</option>
					<option value = "LUAR" <%If ownership = "LUAR" Then %> Selected <%End If%> >LUAR</option>
				</select>
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
		<tr>
			<td colspan="2"  align="left">
				<input Type="Hidden" Value="True" Name="cmdView">
				<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
				<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
				<input Type="Submit"  name="Add" value="Select">&nbsp;&nbsp;
				<a href="printratio_constanta.asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&txtYear=<%=v_txtYear%>&productcode=<%=productcode%>" title="Print Data" target="_blank"><img src="IMAGES/Print.gif" border="0"></a>
			</td>
		</tr>
    </table>  
 </form>
<%
if cmdView = "True" and len(trim(productcode)) <> 0 and len(trim(ownership)) <> 0 Then
	Dim rsVH, intRecordsPerPage, cColor
	set rsVH = server.CreateObject("Adodb.recordset")
	'Field Numeric Jangan di Sum karena ditampilkan mesti level detail
	q = " select linenokey" & _
	    "        , date_from" & _
	    "        , date_until" & _
	    "        , productcode" & _
	    "        , ownership" & _
	    "        , ratio" & _
	    "        , fcentry" & _
	    "        , fcedit" & _
	    "        , fcip" & _
	    "        , fcba" & _
	    "        , lastupdate" & _
	    "        , lasttime" & _
        " from ratio_constanta_tab" & _
		" where fcba = '" & Session("Fcba") & "'" & _
		"       and to_number(to_char(date_until,'yyyy')) = '" & v_txtYear & "'" & _
		"       and productcode = '" & productcode & "'" & _
		"       and ownership = '" & ownership & "'"

	Set rsVH = Server.CreateObject("ADODB.Recordset")
	rsVH.CursorType = 3
	rsVH.CursorLocation = 3
	rsVH.Open q, connstr
%>
	<form action="ratio_constanta.asp" method="post" name="eForm">
	<table class="data" cellpadding="0" cellspacing="0" width="100%">
		<tr>
			<td align="left" width="150">Date From</td>
			<td align="left">
			<%if txtJenis = "Edit" then%>
				<input type="text" size="15" name="datefrom" value="<%=datefrom%>" readonly="True" style="background:#999999">
			<%else%>
				<input class="datepicker" type="text" size="15" name="datefrom" value="<%=datefrom%>">
			<%end if%>
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
		<tr>
			<td align="left">Date Until</td>
			<td align="left">
			<%if txtJenis = "Edit" then%>
				<input type="text" size="15" name="dateuntil" value="<%=dateuntil%>" readonly="True" style="background:#999999">
			<%else%>
				<input class="datepicker" type="text" size="15" name="dateuntil" value="<%=dateuntil%>">
			<%end if%>
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
		<tr>
			<td align="left">Ratio</td>
			<td align="left">
				<input type="text" size="15" name="ratio" value="<%=ratio%>" onkeyup="checkInput(this);"><font color="#FF0000">&nbsp;*</font>
			</td>
			<%if txtJenis = "Edit" then%>
				<input type="hidden" name="txtLineNoKey" value="<%=txtLineNoKey%>">
			<%end if%>
		</tr>
	    <tr>
			<td colspan="2" align="left">
			<input type="hidden" name="txtyear" value="<%=v_txtYear%>">
			<input type="hidden" name="productcode" value="<%=productcode%>">
			<input type="hidden" name="ownership" value="<%=ownership%>">			
			<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
			<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
			<%IF Request("txtJenis")="Edit" then%>
				<input type="hidden" name="txtJenis" value="Edit">
				<input Type="Submit" style="background-color: #298CF7; color: #F7C610; font-weight: bold; font-size: 8pt; font-family:Verdana" name="Add" value="Update">
			<%Else%>
			   <input type="hidden" name="txtJenis" value="New">
			   <input Type="Submit"  name="Add" value="Add">
			<%End if%>
			<input Type="Reset"  name="Reset" value="Reset">
			<input Type="Button" value="Back" name="B2" OnClick="javascript:history.back(1)"></td>
		</tr>
   </table>

   <table class="data" width="100%">
 <table class="data">
	<BR>
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;" >
	<TD align="left">NO</TD>
	<TD colspan="2">&nbsp;</TD>
	<TD align="left">DATE FROM</TD>	
	<TD align="left">DATE UNTIL</TD>	
	<TD align="left">PRODUCT CODE</TD>	
	<TD align="left">OWNERSHIP</TD>	
	<TD align="left">RATIO</TD>	
	</tr>   
   <%
   if Not rsVH.eof then
      nCountofField = rsVH.fields.count-1
	  no=1
	  While Not rsVH.EOF
		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if%>
		  <tr bgcolor="<%=cColor%>">
		  <td align="Right" valign="top" width="30"><%=NO%></td>
				<td valign="middle" width="10">
					<a href="ratio_constanta.asp?txtYear=<%=v_txtYear%>&productcode=<%=productcode%>&ownership=<%=ownership%>&txtLineNoKey=<%=rsVH("linenokey")%>&datefrom=<%=rsVH("date_from")%>&dateuntil=<%=rsVH("date_until")%>&ratio=<%=rsVH("ratio")%>&txtJenis=Edit">
					<img src="IMAGES/edit.gif" alt="Edit Data : <%=rsVH(1)%>" border="0"></a>
				</td>
				<td width="10">
					<a href="ratio_constanta.asp?txtYear=<%=v_txtYear%>&productcode=<%=productcode%>&ownership=<%=ownership%>&txtLineNoKey=<%=rsVH("linenokey")%>&datefrom=<%=rsVH("date_from")%>&dateuntil=<%=rsVH("date_until")%>&txtJenis=Delete" onClick="return confirm('Are you sure to delete');">
					<img src="IMAGES/delete.gif" alt="Delete Data : <%=rsVH(1)%>" border="0"></a>
				</td>

		    <%for i = 1 to nCountofField - 6%>
			    <td align="left" title="<%=rsVH(i).Name%>" valign="top"><%=rsVH(i)%></td>
		    <%next%>
		  </tr>
		  <%
		  rsVH.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend
   end if 
   rsVH.Close
   set rsVH = Nothing%>
   </div>
   </TABLE>
   </form>
<%End if%>  
 
  <%'-------------akhir----------------------%>
</div></div>
  <!--#include file="bottom.jsp"-->

</body>
</html>