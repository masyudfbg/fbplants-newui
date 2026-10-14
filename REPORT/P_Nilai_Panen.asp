<%@Language=VBScript%>
<% Option Explicit %>

<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->

<%
    Dim rsNilaiPanen,strNilaiPanen, strGroup, strOrder, strQuery
	Dim v_txtYear, v_txtMonth, v_txtAfdeling, cClass,  v_txtcustcode ,v_txturaian
	Dim nSpanTotal,ScriptName
	Dim txtYear, txtMonth, txtAfdeling, txtcustcode ,txturaian
	Dim sprevFcba, sprevAfdeling
    Dim SprevCompany,SprevRegion,SprevGroupRegion,SprevPlantation,sprevcustcode
	Dim temptxtYear, temptxtMonth, temptxtAfdeling, temptxtcustcode ,temptxturaian
	Dim tempfcba,tempCompany,tempRegion,tempGroupRegion,tempPlantation
	Dim ExtendedQuery, supportExtendedQuery, strExtendedQuery,strEPS2 
    Dim selectFields,strFieldCons,selectCons,whereCons,orderCons
	DIM q,FcbaName,CompanyName,MgrName,KTUName,CropName,FcbaType
	Dim no
		
	 nSpanTotal = 10
	
	v_txtGroupPlantation   = Request("txtGroupPlantation")
	v_txtPlantation  	   = Request("txtPlantation")
	v_txtGroupRegion  	   = Request("txtGroupRegion")
	v_txtRegion  		   = Request("txtRegion")
	v_txtcompany  		   = Request("txtCompany")
	v_txtfcba 			   = Request("txtfcba")
	v_txtAfdeling		   = Request("txtAfdeling")
	v_txtYear			   = Request("txtYear")
    v_txtMonth			   = Request("txtMonth")
	v_txtcustcode     	   = Request("txtcustcode")
	v_txturaian			   = Request("txturaian") 	
	
	ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

	       				   
			wherecons = " where company is not null"
			
			if v_txtGroupPlantation <> "--SUMMARY--"  and v_txtGroupPlantation <> "--ALL--" and v_txtGroupPlantation <> "" then
			   whereCons = wherecons & " And GroupPlantation = '" & v_txtGroupPlantation & "' "
			 
			end if
			if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then
			    if v_txtPlantation <> "--ALL--" then
				   whereCons = whereCons & " And  Plantation = '" & v_txtPlantation & "' "
				end if   
				selectCons = selectCons & ",Plantation"
				orderCons  = orderCons & ", PLANTATION_API.get_OrderData(Plantation) "
			else
			    nSpanTotal = nSpanTotal - 1 
			end if
			if v_txtGroupRegion <> "--SUMMARY--"   and v_txtGroupRegion <> "" then
			    if v_txtGroupRegion <> "--ALL--" then
				   whereCons = whereCons & " And  GroupRegion = '" & v_txtGroupRegion & "' "
				end if
				selectCons = selectCons & " ,GroupRegion"
				orderCons  = orderCons & ", GROUPREGION_API.get_OrderData(GroupRegion) "
			else
			    nSpanTotal = nSpanTotal - 1 
			end if
			if  v_txtRegion <> "--SUMMARY--"  and v_txtRegion <> "" then
			    if v_txtRegion <> "--ALL--" then
				   whereCons = whereCons & " And  Region = '" & v_txtRegion & "' "
				end if
				selectCons = selectCons & " ,Region"
				orderCons  = orderCons & ", REGION_API.get_OrderData(Region) "
			else
			    nSpanTotal = nSpanTotal - 1 
			end if
			if  v_txtCompany <> "--SUMMARY--" and v_txtCompany <> "" then
			    if v_txtCompany  <> "--ALL--" then
				   whereCons = whereCons & " And  company = '" & v_txtCompany & "' "
				end if
				selectCons = selectCons & " ,company"
			    orderCons  = orderCons & ", COMPANY_API.get_OrderData(company) "
				q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where FCCOMPANYCODE = '" & v_txtCompany & "'"	
				set rsCat = connStr.Execute(q)
				if not rsCat.eof then
				   CompanyName=rsCat("FCCOMPANYNAME")			   
			    end if
			else
			    nSpanTotal = nSpanTotal - 1 	
			end if
	        
			
			if  v_txtfcba <> "--SUMMARY--"  and v_txtfcba <> "" then
			    if v_txtfcba  <> "--ALL--" then
			       whereCons = whereCons & " And  fcba = '" & v_txtfcba & "' "
				end if
			    selectCons = selectCons & ",fcba"
			    orderCons  = orderCons & ",BUSINESSUNIT_API.get_OrderData(fcba) "
				q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & v_txtfcba & "' "
				set rsCat = connStr.Execute(q)
				if not rsCat.eof then
				   fcbaname = rsCat("fcName")
				   CompanyName=rsCat("FCCOMPANYNAME")
				   MgrName = rsCat("FCMANAGER")
				   KTUName = rsCat("FCKTU")
				   CropName=rsCat("FCCROP")
				   FcbaType=rsCat("fcType")
			    end if
			else
			    nSpanTotal = nSpanTotal - 1 
			end if
			'response.write v_txtcustcode
			if  v_txtcustcode <> "ALL"   then
			      if v_txtcustcode <> "ALL"  then
					whereCons = whereCons & " AND CUSTCODE = '" & v_txtcustcode & "'"
			      end if
                 ' selectCons = selectCons &  " , CUSTCODE,CUSTOMER_API.get_description(CUSTCODE,FCBA) CUSTNAME "
				  orderCons  = orderCons & ",CUSTCODE "
			else
			  'selectCons = selectCons &  " , CUSTCODE , CUSTOMER_API.get_description(CUSTCODE,FCBA) CUSTNAME"
			  orderCons  = orderCons & ",CUSTCODE "
			   nSpanTotal = nSpanTotal - 1 
			end if
			
			
			if  v_txtAfdeling <> "--SUMMARY--"   then
			      if v_txtAfdeling <> "--ALL--"  then
					whereCons = whereCons & " AND DIVISION = '" & v_txtAfdeling & "'"
			      end if
                  selectCons = selectCons &  " , DIVISION "
				  orderCons  = orderCons & ",DIVISION "
			else
			   nSpanTotal = nSpanTotal - 1 
			end if
            
			

	             strNilaiPanen    =   " select GroupPlantation ,CUSTCODE,CUSTOMER_API.get_description(CUSTCODE,FCBA) CUSTNAME" & selectCons & ", uraian ,to_char(t.fddate,'YYYY') thn,to_char(t.fddate,'MM') bln, " & _ 
									  "	   sum(decode(to_char(t.fddate,'DD'),'01',nvalue,0)) DAY1, " & _ 
									  "	   sum(decode(to_char(t.fddate,'DD'),'02',nvalue,0)) DAY2, " & _ 
									  "	   sum(decode(to_char(t.fddate,'DD'),'03',nvalue,0)) DAY3, " & _ 
									  "	   sum(decode(to_char(t.fddate,'DD'),'04',nvalue,0)) DAY4, " & _ 
									  "	   sum(decode(to_char(t.fddate,'DD'),'05',nvalue,0)) DAY5, " & _ 
									  "	   sum(decode(to_char(t.fddate,'DD'),'06',nvalue,0)) DAY6, " & _ 
									  "	   sum(decode(to_char(t.fddate,'DD'),'07',nvalue,0)) DAY7, " & _ 
									  "	   sum(decode(to_char(t.fddate,'DD'),'08',nvalue,0)) DAY8, " & _ 
									  "	   sum(decode(to_char(t.fddate,'DD'),'09',nvalue,0)) DAY9, " & _ 
									  "	   sum(decode(to_char(t.fddate,'DD'),'10',nvalue,0)) DAY10, " & _ 
									  "	   sum(decode(to_char(t.fddate,'DD'),'11',nvalue,0)) DAY11, " & _ 
									  "	   sum(decode(to_char(t.fddate,'DD'),'12',nvalue,0)) DAY12, " & _ 
									  "    sum(decode(to_char(t.fddate,'DD'),'13',nvalue,0)) DAY13, " & _ 
									  "	   sum(decode(to_char(t.fddate,'DD'),'14',nvalue,0)) DAY14, " & _ 
									  "	   sum(decode(to_char(t.fddate,'DD'),'15',nvalue,0)) DAY15, " & _ 
									  "	   sum(decode(to_char(t.fddate,'DD'),'16',nvalue,0)) DAY16, " & _ 
									  "	   sum(decode(to_char(t.fddate,'DD'),'17',nvalue,0)) DAY17, " & _ 
									  "	   sum(decode(to_char(t.fddate,'DD'),'18',nvalue,0)) DAY18, " & _ 
									  "	   sum(decode(to_char(t.fddate,'DD'),'19',nvalue,0)) DAY19, " & _ 
									  "	   sum(decode(to_char(t.fddate,'DD'),'20',nvalue,0)) DAY20, " & _ 
									  "	   sum(decode(to_char(t.fddate,'DD'),'21',nvalue,0)) DAY21, " & _ 
									  "	   sum(decode(to_char(t.fddate,'DD'),'22',nvalue,0)) DAY22, " & _ 
									  "	   sum(decode(to_char(t.fddate,'DD'),'23',nvalue,0)) DAY23, " & _ 
									  "	   sum(decode(to_char(t.fddate,'DD'),'24',nvalue,0)) DAY24, " & _ 
									  "	   sum(decode(to_char(t.fddate,'DD'),'25',nvalue,0)) DAY25, " & _ 
									  "	   sum(decode(to_char(t.fddate,'DD'),'26',nvalue,0)) DAY26, " & _ 
									  "	   sum(decode(to_char(t.fddate,'DD'),'27',nvalue,0)) DAY27, " & _ 
									  "	   sum(decode(to_char(t.fddate,'DD'),'28',nvalue,0)) DAY28, " & _ 
									  "	   sum(decode(to_char(t.fddate,'DD'),'29',nvalue,0)) DAY29, " & _ 
									  "	   sum(decode(to_char(t.fddate,'DD'),'30',nvalue,0)) DAY30, " & _ 
									  "	   sum(decode(to_char(t.fddate,'DD'),'31',nvalue,0)) DAY31  " & _ 
									  "  from HARVESTING_VALUE_REPORT_TAB t " &  whereCons  & " and USER_COMPANY_API.GetfcbaType('" & Session("CodeName") & "',fcba) = 'M' " & _ 
									  "  group by  GroupPlantation, to_char(t.fddate,'YYYY') ,to_char(t.fddate,'MM'),FCBA,CUSTCODE " & selectCons & " ,uraian" & _
									  "  order by GroupPlantation " & orderCons
									 
	' strMerumputTM =  "select param_2,param_2_name,GroupPlantation " & selectCons & strEPS2 & "  From (" & strMerumputTM & " )" & whereCons & " and USER_COMPANY_API.GetfcbaType('" & Session("CodeName") & "',fcba) = 'E' group by  param_2,param_2_name,GroupPlantation" & selectCons & "  Order by param_2,param_2_name,GroupPlantation" & orderCons  
	
	
	 'Set rsNilaiPanen= Server.CreateObject("ADODB.Recordset")
	 'response.write  strNilaiPanen 
		
	'rsNilaiPanen.Open strNilaiPanen, connstr, adOpenStatic 	

	
%>
<html>

<head>
<title>ePlaS Reporting - Nilai Panen</title>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<link rel="stylesheet" href="/css/jquery-ui.css" />
<script src="/js/jquery-1.9.1.js"></script>
<script src="/js/jquery-ui.js"></script>
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
 <script>
$(function() {
$( ".datepicker" ).datepicker();
});
</script>
<!--<script language="JavaScript" src="/js/date-picker.js"></script> 
<script language="JavaScript">

var newwin;
function launchwin(winurl,winname,winfeatures)
{
	//This launches a new window and then
	//focuses it if window.focus() is supported.
	newwin = window.open(winurl,winname,winfeatures);

}
</script>-->

<style type="text/css">
@media print
{
.printhide{display:none;}
}
</style>
</head>

<body>
	<div id="wrapper">
	 <!--#include file="../print_excel.asp"-->
		<div id="header-wrapper">
			<div class="printhide">
				<!--#include file="menutop.asp"-->
			</div>
			<form action="<%=ScriptName%>" method="post" name="frmEntry">
				<div class="printhide">
				   <!--#include file="FormSelectNewSumM.html"-->
					<table id="form-table" class="centered-margin">	
						<tr>
							<td><font face="verdana" size="1">TAHUN</font></td>
									<% 
										Dim mon, Yr, MinYr, MaxYr, i, MM
										mon = Month(Session("v_Costbook"))
										If len(v_txtMonth)=0 then v_txtMonth=mon

										Yr = Year(Session("v_Costbook"))
										MinYr = Yr - 2
										MaxYr = Yr  + 2
										If len(v_txtYear)= 0 then v_txtYear = Yr
									%>
							<td width="4">
								<select name="txtYear">
											<%
												For i = MinYr to MaxYr
											%>
										<Option Value="<%=i%>" <%If cint(i) = cint(v_txtYear) then%> Selected <%End If%>><%=i%></option>
									<%
										Next
									%>
								</select>
							</td>
							
							<td><font face="verdana" size="1">BULAN</font></td>
							<td width="2">
								<select name="txtMonth">
									<%
										For i = 1 to 12
									%>
											<Option Value="<%=i%>" <%If cint(i) = cint(v_txtMonth) then%> Selected <%End If%>><%=Bulan(i)%></option>
									<%
										Next
									%>
								</select>
							</td>
							
							
							
							<td>
								<font face="verdana" size="1">CUSTOMER</font>
							</td>
							<td>
								<select name="txtcustcode" >
									<option value="ALL">--ALL--</option>
									<%
									Dim rsCust, strCust
									strCust = " Select  DISTINCT T.CUSTCODE, CUSTOMER_API.get_description(T.CUSTCODE,T.FCBA) CUSTNAME from HARVESTING_VALUE_REPORT_TAB t "
									
									Set rsCust = Server.CreateObject("ADODB.Recordset")
									rsCust.Open strCust, connstr, adOpenStatic 
									
									if not rsCust.eof then
									Do While Not rsCust.eof
									if rsCust("CUSTCODE") = v_txtcustcode then%>
									<option value="<%=rsCust("CUSTCODE")%>" Selected><%=rsCust("CUSTCODE")%> - <%=rsCust("CUSTNAME")%></option> 
									<%Else%>
									<option value="<%=rsCust("CUSTCODE")%>"><%=rsCust("CUSTCODE")%> - <%=rsCust("CUSTNAME")%></option>  
									<%End if
										rsCust.movenext
									loop
									end if
									rsCust.Close
									set rsCust = Nothing
									%>
								</select>
							</td>
							
							<td>
								<font face="verdana" size="1">AFDELING</font>
							</td>
							<td>
								<select name="txtAfdeling" >
									<option value="--SUMMARY--">--SUMMARY--</option>
									<%
									Dim rsAfdeling, strAfdeling
									strAfdeling	=	" select '--ALL--' DIVISION from dual union Select  DISTINCT T.DIVISION from HARVESTING_VALUE_REPORT_TAB t  "
									
									Set rsAfdeling = Server.CreateObject("ADODB.Recordset")
									rsAfdeling.Open strAfdeling, connstr, adOpenStatic 
									
									if not rsAfdeling.eof then
									Do While Not rsAfdeling.eof
									if rsAfdeling("DIVISION") = v_txtAfdeling then%>
									<option value="<%=rsAfdeling("DIVISION")%>" Selected><%=rsAfdeling("DIVISION")%></option> 
									<%Else%>
									<option value="<%=rsAfdeling("DIVISION")%>"><%=rsAfdeling("DIVISION")%></option>  
									<%End if
										rsAfdeling.movenext
									loop
									end if
									rsAfdeling.Close
									set rsAfdeling = Nothing
									%>
								</select>
							</td>
							 
							 <td>
								<font face="verdana" size="1">URAIAN</font>
							</td>
							<td>
								<select name="txturaian" >
									<option value="ALL">--ALL--</option>
									<%
									Dim rsuraian, struraian
									struraian	=	" Select  DISTINCT T.Uraian from HARVESTING_VALUE_REPORT_TAB t WHERE FCBA='PSA'"
									
									Set rsuraian = Server.CreateObject("ADODB.Recordset")
									rsuraian.Open struraian, connstr, adOpenStatic 
									
									if not rsuraian.eof then
									Do While Not rsuraian.eof
									if rsuraian("Uraian") = v_txturaian then%>
									<option value="<%=rsuraian("Uraian")%>" Selected><%=rsuraian("Uraian")%></option> 
									<%Else%>
									<option value="<%=rsuraian("Uraian")%>"><%=rsuraian("Uraian")%></option>  
									<%End if
										rsuraian.movenext
									loop
									end if
									rsuraian.Close
									set rsuraian = Nothing
									%>
								</select>
							</td>
							 
							<td><input type="submit" value="View" /></td>
						</tr>				
					</table>
				</div>
			</form>
		</div>
<table border="0" width="100%" id="product-table" class="xthf-site">
	<thead style="border-left:1px #000000 solid;">
	<tr>
         <th colspan="<%=5+GetLastDay(v_txtMonth, v_txtYear)%>" style="background:#FFFFFF; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
		 <div id="header-wrapper">
			<div id="head_judul">
				<% if  CompanyName <>  "" then %>
				       <h4>PT.<%=CompanyName%></h4> 
				 <% End if%>
				 <% If  FcbaName <> ""  then %>
					   <h5><%=FcbaName%></h5>   
				 <% End if%>	
			</div>
			<div id="subheader" class="centered">
				<h2>LAPORAN HARIAN NILAI PANEN</h3>
				<br/>
				<h4>PERIODE&nbsp;:&nbsp;<%=Bulan(v_txtMonth)%>&nbsp;<%=v_txtYear%></h4>
			</div>
          </div>  
		 </th>
	</tr>
	<tr>
	    <% if v_txtfcba <> "--SUMMARY--" AND v_txtfcba <> "" then %>
		      <th rowspan="3" style="background-color:#00A300;color:#fff;" >FCBA</th>
		<% end if %>
		
		
		<th rowspan="3" style="background-color:#00A300;color:#fff;" >KEBUN</th>
		
		<% if v_txtAfdeling <> "--SUMMARY--" AND v_txtAfdeling <> "" then %>
		      <th rowspan="3" style="background-color:#00A300;color:#fff;" >AFDELING</th>
		<% end if %>
		<th rowspan="2" style="background-color:#00A300;color:#fff;">URAIAN</th>
		<th colspan="<%=GetLastDay(v_txtMonth, v_txtYear)%>" style="background-color:#00A300;color:#fff;">TANGGAL</th>
	</tr>
	<tr>
	
	
<%
Dim rsGetColorTanggal, strGetColorTanggal
Dim val_FirstDay, val_LastDay
Dim val_Background

val_FirstDay = v_txtMonth & "/" & 1 & "/" & v_txtYear
val_LastDay = v_txtMonth & "/" & GetLastDay(v_txtMonth, v_txtYear) & "/" & v_txtYear

strGetColorTanggal =	" SELECT" & _
						"     FCBA" & _
						"     , TO_NUMBER(TO_CHAR(TANGGAL, 'DD')) LIST_DAY" & _
						"     , DECODE(" & _
						"         BUDGET_TRANSACTION_WORKDAY_API.GETQTYWORKINGDAY_HI(FCBA, TO_NUMBER(TO_CHAR(TANGGAL, 'YYYY')), 1, 1,TANGGAL)" & _
						"         , 0,'LIBUR', 'NORMAL') KETERANGAN" & _
						" FROM (" & _
						"     SELECT" & _
						"         FCCODE FCBA" & _
						"         , TANGGAL" & _
						"     FROM (" & _
						"         SELECT TO_DATE('" & val_FirstDay & "' , 'MM/DD/YYYY') + ROWNUM - 1 TANGGAL" & _
						"         FROM ALL_OBJECTS" & _
						"         WHERE ROWNUM <= TO_DATE('" & val_LastDay & "','MM/DD/YYYY') - TO_DATE('" & val_FirstDay & "' , 'MM/DD/YYYY') + 1" & _
						"         )" & _
						"     , BUSINESSUNIT" & _
						"     ) WHERE FCBA = '" & v_txtFcba & "' ORDER BY TO_NUMBER(TO_CHAR(TANGGAL, 'DD'))"

Set rsGetColorTanggal = Server.CreateObject("ADODB.Recordset")
rsGetColorTanggal.Open strGetColorTanggal, connstr, adOpenStatic

If Not rsGetColorTanggal.EOF Then
	While Not rsGetColorTanggal.EOF
		If rsGetColorTanggal("KETERANGAN") = "LIBUR" Then
			val_Background = "background-color:#e7dc14;color:#fff;"
		Else
			val_Background = "background-color:#00A300;color:#fff;"
		End If
%>
	<th style="<%=val_Background%>"><%=rsGetColorTanggal("LIST_DAY")%></th>
<%
		rsGetColorTanggal.MoveNext
	Wend
End If
%>
	</tr>
    </thead>
	<tbody>
<%
'response.write  strNilaiPanen 
Set rsNilaiPanen = Server.CreateObject("ADODB.Recordset")
rsNilaiPanen.Open strNilaiPanen, connstr, adOpenStatic

Dim xNum

if rsNilaiPanen.bof and rsNilaiPanen.eof then
%>
	<tr>
		<td colspan="<%=5+GetLastDay(v_txtMonth, v_txtYear)%>">NO DATA AVAILABLE</td>
	</tr>
<%end if


		while not rsNilaiPanen.eof
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
%>
			<tr class="<%=cClass%>">
			
			     <% if v_txtfcba <> "--SUMMARY--" AND v_txtfcba <> "" then %>
					  
					  <% if sprevFcba <> rsNilaiPanen("FCBA") then %>
						<td align="center">
							<%=rsNilaiPanen("FCBA")%>	
						</td>
					       <% sprevFcba = rsNilaiPanen("FCBA") %>	
					 <% else	%>
						 <td align="center"></td>					
					 <% end if %>	
				<% end if %>
				
				 <% if v_txtcustcode <> "--SUMMARY--" AND v_txtcustcode <> "" then %>
					  <% if sprevcustcode <> rsNilaiPanen("CUSTNAME") then %>
						<td align="center">
							<%=rsNilaiPanen("CUSTNAME")%>	
						</td>
					       <% sprevcustcode = rsNilaiPanen("CUSTNAME") %>	
					 <% else	%>
						 <td align="center"></td>					
					 <% end if %>	
				<% end if %>
				
				 <% if v_txtAfdeling <> "--SUMMARY--" AND v_txtAfdeling <> "" then %>
					  <% if sprevAfdeling <> rsNilaiPanen("DIVISION") then %>
						<td align="center">
							<%=rsNilaiPanen("DIVISION")%>	
						</td>
					       <% sprevAfdeling = rsNilaiPanen("DIVISION") %>	
					 <% else	%>
						 <td align="center"></td>					
					 <% end if %>	
				<% end if %>
				
				<td align="center"><%=rsNilaiPanen("URAIAN")%></td>
				
				<%For xNum = 1 to GetLastDay(v_txtMonth, v_txtYear)%>
				
				<td align="right" ><%=FormatNumber(rsNilaiPanen("day"&xNum&""),1)%></td>
				<%Next%>
				
				
			</tr>


<%
			rsNilaiPanen.MoveNext
		Wend
rsNilaiPanen.Close
set rsNilaiPanen = Nothing
%>
</tbody>
</table>

   <table width="90%" class="data">
	   <tr>
 		<td width="50%" align="center"><font face="Verdana" size="1">Approved by,</font></td>
		<td width="50%" align="center"><font face="Verdana" size="1">Prepared by,</font></td>
	   </tr>
	   <tr> 
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	   </tr>
	   <tr> 
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	   </tr>
	   <tr> 
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	   </tr>
	   <tr> 
		<td width="50%" align="center"><font face="Verdana" size="2"><strong><u><%=Session("FCMANAGER")%></u></strong></td>
		<td width="50%" align="center"><font face="Verdana" size="2"><strong><u><%=Session("KTU")%></u></strong></td>
	   </tr>
	   	  <tr> 
		<td width="50%" align="center"><font face="Verdana" size="1"><strong><i>MANAGER</i></strong></td>
		<td width="50%" align="center"><font face="Verdana" size="1"><strong><i>KTU</i></strong></td>
	   </tr>
   </table>
   </body>
   </html>
