<%@Language=VBScript%>
<% Option Explicit %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%
	Dim rsMerumputTM, strHeadcountTK, strGroup, strOrder, strQuery
	Dim  MM, yyyy, mon, Yr, MinYr, MaxYr, i,v_txtyear, v_txtmonth,v_txtdatefrom, v_txtdateto, v_txtSectionName, v_txtBlok, v_txtPlantingDate, NO, cClass, v_txtGangCode, v_txtjobcode, v_txtGroup
	Dim nSpanTotal
	Dim txtdatefrom, txtdateto, txtSectionName, txtBlok, txtPlantingDate, txtGangCode, txtjobcode, txtGroup
	Dim sprevFcba, sprevDivision, sprevPlantingYear, sprevFieldCode, sprevStatus, sprevJobCode, sprevGroup, sprevSectionName
    Dim SprevCompany,SprevRegion,SprevGroupRegion,SprevPlantation
	Dim tempSectionName, tempPlantingDate,  tempBlok,  tempStatus,  tempJobcode,  tempGroup 
	Dim tempfcba,tempCompany,tempRegion,tempGroupRegion,tempPlantation, tempGangCode
	Dim ExtendedQuery, supportExtendedQuery, strExtendedQuery,strEPS2 
    'Dim rscomp , strcomp,txtcompany ,txcompanyname,txtComp,v_txtComp
	DIM q,FcbaName,CompanyName,MgrName,KTUName,CropName,FcbaType
	Dim sectionname
		
	 nSpanTotal = 5
	
	v_txtGroupPlantation   = Request("txtGroupPlantation")
	v_txtPlantation  	   = Request("txtPlantation")
	v_txtGroupRegion  	   = Request("txtGroupRegion")
	v_txtRegion  		   = Request("txtRegion")
	v_txtcompany  		   = Request("txtCompany")
	v_txtfcba 			   = Request("txtfcba")
	v_txtSectionName	   = Request("txtSectionName")
	v_txtBlok			   = Request("txtBlok")
	v_txtPlantingDate	   = Request("txtPlantingDate")
	v_txtdatefrom		   = Request("txtdatefrom")
	v_txtdateto			   = Request("txtdateto")
	v_txtGangCode          = Request("txtGangCode")
	v_txtjobcode     	   = Request("txtjobcode")
	v_txtGroup			   = Request("txtGroup")
	sectionname			   = Request("sectionname")
	'response.write v_txtGroup	
	if v_txtdateto ="" then
	   v_txtdatefrom = textdate(Now())
	   v_txtdateto = textdate(Now())
    end if	
	
	        v_txtyear 		= Request("txtyear")
			v_txtmonth 		= lpad(Request("txtmonth"),"0",2)

	        'response.write  "GP " & v_txtGroupPlantation & "  P " & v_txtPlantation & "  GR " & v_txtGroupRegion & " R " & v_txtRegion & " C " & v_txtcompany & " S " & v_txtfcba  &  " D " & v_txtSectionName & " TT" & v_txtPlantingDate
			'response.write  "TM/TBM " & v_txtGangCode & "  jobcode " & v_txtjobcode & "  Group " &  v_txtGroup

					
			Dim selectFields,strFieldCons,selectCons,whereCons,orderCons 
			strFieldCons = " USER_COMPANY_API.GetGroupPlantationCode('" & Session("CodeName") & "',fcba) GroupPlantation,USER_COMPANY_API.GetPlantationCode('" & Session("CodeName") & "',fcba) Plantation, " & _
						   " USER_COMPANY_API.GetGroupRegionCode('" & Session("CodeName") & "',fcba) groupregion ,USER_COMPANY_API.GetRegionCode('" & Session("CodeName") & "',fcba) region,USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',fcba) company, BUSINESSUNIT_API.get_OrderData(fcba) norut, fcba, " 

						   
			wherecons = " where company is not null "
			if v_txtGroup <> "--SUMMARY--" And v_txtGroup <> "--ALL--" And v_txtGroup <> ""  then
			   whereCons = wherecons & " And param_2= '" & v_txtGroup & "' "  
			end if
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
				set rsMerumputTM = connStr.Execute(q)
				if not rsMerumputTM.eof then
				   CompanyName=rsMerumputTM("FCCOMPANYNAME")			   
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
				set rsMerumputTM = connStr.Execute(q)
				' response.write q
				if not rsMerumputTM.eof then
				   fcbaname = rsMerumputTM("fcName")
				   CompanyName=rsMerumputTM("FCCOMPANYNAME")
				   MgrName = rsMerumputTM("FCMANAGER")
				   KTUName = rsMerumputTM("FCKTU")
				   CropName=rsMerumputTM("FCCROP")
				   FcbaType=rsMerumputTM("fcType")
			    end if
			else
			    nSpanTotal = nSpanTotal - 1 
			end if
			if  v_txtSectionName <> "--SUMMARY--" AND v_txtSectionName <> "" then
			      if v_txtSectionName <>  "--ALL--"  then
					whereCons = whereCons & " AND sectionname = '" & v_txtSectionName & "'"
			      end if
                  selectCons = selectCons &  " , sectionname "
				  orderCons  = orderCons & ",sectionname "
			else
			   nSpanTotal = nSpanTotal - 1 
			end if
            ' response.write v_txtSectionName
			' if v_txtGangCode <> "--SUMMARY--"  AND  v_txtGangCode <> "" then
			' 	   if v_txtGangCode <>  "--ALL--" then
			' 		  whereCons = whereCons & " AND GANGCODE = '" & v_txtGangCode & "'"
			' 	   end if
			' 	   selectCons = selectCons &  " , GANGCODE "
			' 	   orderCons  = orderCons & ",GANGCODE "
			' else
			'    nSpanTotal = nSpanTotal - 1 
			' end if
			' response.write v_txtGangCode
			' if  v_txtBlok <> "--SUMMARY--" AND v_txtBlok <> "" then
			' 	   if v_txtBlok <>  "--ALL--" then
			' 		  whereCons = whereCons & " AND FIELDCODE = '" & v_txtBlok & "'"
			' 	   end if
			' 	   selectCons = selectCons & " , FIELDCODE "
   '                 orderCons  = orderCons & ",FIELDCODE "				   
			' else
			'    nSpanTotal = nSpanTotal - 1 
   '          end if
			
	        
			' if v_txtjobcode <> "--SUMMARY--" AND v_txtjobcode <> "" then
			' 	   if v_txtjobcode <> "--ALL--" then
			' 		  whereCons = whereCons &  " AND  jobcode = '" & v_txtjobcode & "'"
			' 	   end if
			'       selectCons = selectCons &  " , jobcode"				   
			' else
			'    nSpanTotal = nSpanTotal - 1 
			' end if	
	
    'strHeadcountTK = " SELECT c.param_2, a.status "
    ' strHeadcountTK = " SELECT c.param_2, c.param_2_name "
	
	' supportExtendedQuery = 0

	

	strHeadcountTK 		=     " select * from EMPLOYEE_HIST_TAB t " & _ 
                              " where t.year = '" & v_txtyear &"' and t.month = '"& v_txtmonth &"' and USER_COMPANY_API.IsAuthorizedFcba('" & Session("CodeName") & "',fcba) ='TRUE'  " 
							  
	
                      strEPS2    =  " a2.fcba,a2.sectionname,a2.presentgrade,a2.gangcode, " & _
                      				" b.fccode, b.fcname, b.fccompanycode, b.fccompanyname, b.fcarea, b.fcownership, b.fcmanager, b.groupplantation, b.plantation, b.groupregion, b.region,"& _
                      				" sum(decode(presentgrade,'KBT',Jumlah_TK_PNN,0)) KBT_PNN," & _
									" sum(decode(presentgrade,'KHT',Jumlah_TK_PNN,0)) KHT_PNN," & _
									" sum(decode(presentgrade,'KHL',Jumlah_TK_PNN,0)) KHL_PNN," & _
									" sum(decode(presentgrade,'KBT',Jumlah_TK_MD,0)) KBT_MD," & _
									" sum(decode(presentgrade,'KHT',Jumlah_TK_MD,0)) KHT_MD," & _
									" sum(decode(presentgrade,'KHL',Jumlah_TK_MD,0)) KHL_MD," & _
									" sum(decode(presentgrade,'KBT',Jumlah_TK_TR,0)) KBT_TR," & _
									" sum(decode(presentgrade,'KHT',Jumlah_TK_TR,0)) KHT_TR," & _
									" sum(decode(presentgrade,'KHL',Jumlah_TK_TR,0)) KHL_TR," & _
									" sum(decode(presentgrade,'KBT',Jumlah_TK_KU,0)) KBT_KU," & _
									" sum(decode(presentgrade,'KHT',Jumlah_TK_KU,0)) KHT_KU," & _
									" sum(decode(presentgrade,'KHL',Jumlah_TK_KU,0)) KHL_KU," & _
									" sum(decode(presentgrade,'KBT',Jumlah_TK_PML,0)) KBT_PML," & _
									" sum(decode(presentgrade,'KHT',Jumlah_TK_PML,0)) KHT_PML," & _
									" sum(decode(presentgrade,'KHL',Jumlah_TK_PML,0)) KHL_PML" & _
									" from (" & _
									        " select a.fcba,a.sectionname,a.presentgrade,a.gangcode, sum(a.Jumlah_TK_PNN) Jumlah_TK_PNN, sum(a.Jumlah_TK_MD) Jumlah_TK_MD, " & _
									        " sum(a.Jumlah_TK_TR) Jumlah_TK_TR, sum(a.Jumlah_TK_KU) Jumlah_TK_KU, sum(a.Jumlah_TK_PML) Jumlah_TK_PML" & _
									        " from ( " & _
									        		" select fcba,sectionname,gangcode,presentgrade,count(presentgrade) as Jumlah_TK_PNN,0 Jumlah_TK_MD ,0 Jumlah_TK_TR,0 Jumlah_TK_KU,0 Jumlah_TK_PML" & _
									                " from iplasprod.employee_hist_tab" & _
									                " where fcba = '"& Session  ("fcba") &"' and dateterminate is null" & _
									                " and substr(gangcode,1,2) in ('PN')" & _
									                " and year = '"& v_txtyear & "'" & _
													" and month = '"& v_txtmonth & "'" & _
									                " group by fcba,sectionname,gangcode,presentgrade" & _
									                " UNION ALL" & _
									                " select fcba,sectionname,gangcode,presentgrade,0 Jumlah_TK_PNN, count(presentgrade) as Jumlah_TK_MD,0 Jumlah_TK_TR,0 Jumlah_TK_KU,0 Jumlah_TK_PML" & _
									                " from iplasprod.employee_hist_tab" & _
									                " where fcba = '"& Session  ("fcba") &"' and dateterminate is null" & _
									                " and substr(gangcode,1,2) in ('MD')" & _
									                " and year = '"& v_txtyear & "'" & _
													" and month = '"& v_txtmonth & "'" & _
									                " group by fcba,sectionname,gangcode,presentgrade " & _
									                " UNION ALL" & _
									                " select fcba,sectionname,gangcode,presentgrade,0 Jumlah_TK_PNN, 0 Jumlah_TK_MD, count(presentgrade) as Jumlah_TK_TR,0 Jumlah_TK_KU,0 Jumlah_TK_PML" & _
									                " from iplasprod.employee_hist_tab" & _
									                " where fcba = '"& Session  ("fcba") &"' and dateterminate is null" & _
									                " and substr(gangcode,1,2) in ('OS','MT','TB')" & _
									                " and year = '"& v_txtyear & "'" & _
													" and month = '"& v_txtmonth & "'" & _
									                " group by fcba,sectionname,gangcode,presentgrade " & _
									                " UNION ALL" & _
									                " select fcba,sectionname,gangcode,presentgrade,0 Jumlah_TK_PNN,0 Jumlah_TK_MD, 0 Jumlah_TK_TR, count(presentgrade) as Jumlah_TK_KU,0 Jumlah_TK_PML" & _
									                " from iplasprod.employee_hist_tab" & _
									                " where fcba = '"& Session  ("fcba") &"' and dateterminate is null" & _
									                " and substr(gangcode,1,2) in ('KT','KM','UM')" & _
									                " and year = '"& v_txtyear & "'" & _
													" and month = '"& v_txtmonth & "'" & _
									                " group by fcba,sectionname,gangcode,presentgrade" & _
									                " UNION ALL" & _
										            " select fcba,sectionname,gangcode,presentgrade,0 Jumlah_TK_PNN,0 Jumlah_TK_MD,0 Jumlah_TK_TR,0 Jumlah_TK_KU, count(presentgrade) as Jumlah_TK_PML" & _
									                " from iplasprod.employee_hist_tab" & _
									                " where fcba = '"& Session  ("fcba") &"' and dateterminate is null" & _
									                " and substr(gangcode,1,2) not in ('PN','KT','KM','UM','OS','MT','TB','MD')" & _
									                " and year = '"& v_txtyear & "'" & _
													" and month = '"& v_txtmonth & "'" & _
									                " group by fcba,sectionname,gangcode,presentgrade" & _
									             " ) a" & _
									            " group by a.fcba,a.sectionname,a.presentgrade,a.gangcode" & _
									" ) a2  left join businessunit b on (b.fccode = a2.fcba)" & _
									" where a2.sectionname = '"& v_txtSectionName &"' " & _
									" group by a2.fcba,a2.sectionname, a2.presentgrade, a2.gangcode, b.fccode, b.fcname, b.fccompanycode, b.fccompanyname, b.fcarea, b.fcownership, b.fcmanager, b.groupplantation, b.plantation, b.groupregion, b.region" & _
									" order by a2.fcba,a2.sectionname, a2.presentgrade, a2.gangcode, b.fccode, b.fcname, b.fccompanycode, b.fccompanyname, b.fcarea, b.fcownership, b.fcmanager, b.groupplantation, b.plantation, b.groupregion, b.region"
									 
	strHeadcountTK = " select  " & strEPS2 
					 ' " From (" & strHeadcountTK & " )" & whereCons & " and USER_COMPANY_API.GetfcbaType('" & Session("CodeName") & "',fcba) = 'E' group by  fccode, fcname, gangcode" & selectCons & "  Order fccode, fcname, gangcode" & orderCons  
	
	
	Set rsMerumputTM = Server.CreateObject("ADODB.Recordset")
	' response.write  strHeadcountTK
		
	rsMerumputTM.Open strHeadcountTK, connstr, adOpenStatic 	 	
	
%>


<html>

<head>
<title>iPS Reporting - Laporan Headcount Tenaga Kerja</title>
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

<!--<script type="text/javascript" src="http://jqueryjs.googlecode.com/files/jquery-1.3.1.min.js" > </script> -->
<script type="text/javascript">

    function PrintElem(elem)
    {
        Popup($(elem).html());
    }

    function Popup(data) 
    {
        var mywindow = window.open('', 'my div', 'height=400,width=600');
        mywindow.document.write('<html><head><title>my div</title>');
        mywindow.document.write('<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />');
        mywindow.document.write('</head><body >');
        mywindow.document.write(data);
        mywindow.document.write('</body></html>');

        mywindow.print();
        mywindow.close();

        return true;
    }

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
</head>


<body>
<div id="wrapper">
	<!--#include file="../print_excel.asp"-->
	</div>
<div class="top_select" style=" position:absolute;  top:0; left:0px; width:100%;"> 
<form name="frmEntry" method="post" action="p_laporanheadcount.asp">

	<div id="header-wrapper">
	<table id="menu-table" class="centered-margin">
		<tr>
			<td align="center"><a href="/login.asp">Back to Home</a></td>
			<td align="center"><a href="/report/Dashboard.asp">Back to Report</a></td>
		</tr>
	</table>
	
	</div>
<!--#include file="FormSelectNewSum.html"-->	
<table id="form-table" class="centered-margin">
	<tr>
	    <th>&nbsp;</th>
		<th>BULAN</th>
		<th>TAHUN</th>
		<th>DIVISION</th>
		<th>&nbsp;</th>
	</tr>
	
	<tr>
		<td align="center">
		<% 
					MM=Request("txtMonth")
					yyyy=Request("txtYear")

					mon = Month(Session("v_Costbook"))
					If len(MM)=0 then MM=mon

					Yr = Year(Session("v_Costbook"))
					MinYr = Yr - 2
					MaxYr = Yr  + 2
					If len(yyyy)= 0 then yyyy = Yr
				%>
		
		   <td>
				<font size="2" color="#0000FF" ></font>
					<select name="txtMonth">
						<%
							For i = 1 to 12
						%>
							<Option Value="<%=i%>" <%If cint(i) = cint(MM) then%> Selected <%End If%>><%=MonthName(i)%></option>
						<%
							Next
						%>
					</select>
		  </td>	
		   <td>
				<font size="2"  color="#0000FF" ></font>
				
				<select name="txtYear">
					<%
						For i = MinYr to MaxYr
					%>
							<Option Value="<%=i%>" <%If cint(i) = cint(yyyy) then%> Selected <%End If%>><%=i%></option>
					<%
						Next
					%>
				</select>
			</td>

		<td align="center">
		<select name="txtSectionName">	
			<option value="--SUMMARY--">--SUMMARY--</option>
		<%
			Dim rsDivision
			Set rsDivision = Server.CreateObject("ADODB.Recordset")
			rsDivision.CursorLocation = adUseClient

			if len(trim(txtSectionName)) > 0 then 
				strsql =  " SELECT '--ALL--' as DIVISION from DUAL UNION SELECT DISTINCT sectionname as Division FROM EMPLOYEE_HIST_TAB WHERE FCBA = '" & v_txtFcba & "'"
			else
				strsql =  " SELECT '--ALL--' as DIVISION from DUAL UNION SELECT DISTINCT sectionname as Division FROM EMPLOYEE_HIST_TAB "
			end if
			
			rsDivision.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	

			if not rsDivision.eof then
				DO while not rsDivision.eof
				if v_txtSectionName = rsDivision("Division") then
		%>			<option value="<%=rsDivision("Division")%>" Selected><%=rsDivision("Division")%></option>
				<%else%>
					<option value="<%=rsDivision("Division")%>"><%=rsDivision("Division")%></option>
				<%end if%>
		<%			rsDivision.movenext
				loop
			End if
			rsDivision.Close
			set rsDivision = nothing
		%>
		</select>		
		</td>

		

		<td align="center"><input type="submit" value="View" /></td>
	</tr>
    <tr>
<td colspan="140" style="background:#ffffff; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
<div id="header-wrapper">
				    <div id="header">
						 <% if  CompanyName <>  "" then %>
								<h4>PT.<%=CompanyName%></h4> 
						 <% End if%>
						 <% If  FcbaName <> ""  then %>
								   <h5><%=FcbaName%></h5>   
						 <% End if%>
					</div>
					<div id="subheader" class="centered">
						<h2>LAPORAN HEADCOUNT TENAGA KERJA</h2>
						<h5>PERIODE : <%=v_txtmonth%> - <%=v_txtyear%> </h5>
					</div>
				</div>

</td>
</tr>
</table>
</div>

<div id="mydiv" style="margin-top:280px;">
<table border="0" width="100%" id="product-table" class="xthf-site" style="background-color:#FFFFFF;">
<thead>
	<tr>
		
				<th rowspan="3" style="background-color:#00A300;color:#fff;">AFDELING</th>
		<% if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then %>
		      <th rowspan="3" style="background-color:#00A300;color:#fff;" ><%=titlePlantation%></th>
		<% end if %>
		
		<% if v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  "" then %>
		      <th rowspan="3" style="background-color:#00A300;color:#fff;" ><%=titleGroupRegion%></th>
		<% end if %>
		
		<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
		      <th rowspan="3" style="background-color:#00A300;color:#fff;" ><%=titleRegion%></th>
		<% end if %>
		
		<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then %>
		       <th rowspan="3" style="background-color:#00A300;color:#fff;" ><%=titleCompany%></th>
		<% end if %>
		
		<% if v_txtfcba <>  "--SUMMARY--" and v_txtfcba <>  "" then %>
		       <th rowspan="3" style="background-color:#00A300;color:#fff;"  ><%=titleSite%></th>
		<% end if %>

		<th colspan="4">PEMELIHARAAN</th>
		<th colspan="4">PANEN</th>
		<th colspan="4">SUPERVISI</th>
		<th colspan="4">TRAKSI</th>
		<th colspan="4">KANTOR & UMUM</th>
		<th colspan="4">TOTAL</th>
		
	</tr>
	<tr>
		<% DIM K
		   FOR K = 1 TO 6 %>
		<th>KBT</th>
		<th>KHT</th>
		<th>KHL</th>
		<th>TOTAL</th>
		<% NEXT %>
	</tr>
	
</thead>
<tbody>
<%
' Dim nsubsku, nsubkhl, nsubtotal, nsubha_harian, nsubbo_harian,nbsubbudHK,nbsubbudHA,nbsubbudRasio,nbsubbudGLI,nbsubbudPAR,nbsubbudMET,nbsubbudFLO
' Dim nsubtotal_bo_harian, nsubha_per_hk, nsubgli_ltr, nsubgli_ltr_ha, nsubpar_ltr
' Dim nsubpar_ltr_ha, nsubmet_kg, nsubmet_kg_ha, nsubflo_kg, nsubflo_kg_ha, nsubamount_hi_per_ha
' Dim ntotal_01,ntotal_bo_01,nhk_ha_01,nAmount_01,nAmount_ha_01,ntotal_02,ntotal_bo_02,nhk_ha_02,nAmount_02,nAmount_ha_02,ntotal_03,ntotal_bo_03,nhk_ha_03,nAmount_03,nAmount_ha_03
' Dim	ntotal_04,ntotal_bo_04,nhk_ha_04,nAmount_04,nAmount_ha_04,ntotal_05,ntotal_bo_05,nhk_ha_05,nAmount_05,nAmount_ha_05,ntotal_06,ntotal_bo_06,nhk_ha_06,nAmount_06,nAmount_ha_06
' Dim	ntotal_07,ntotal_bo_07,nhk_ha_07,nAmount_07,nAmount_ha_07,ntotal_08,ntotal_bo_08,nhk_ha_08,nAmount_08,nAmount_ha_08,ntotal_09,ntotal_bo_09,nhk_ha_09,nAmount_09,nAmount_ha_09
' Dim	ntotal_10,ntotal_bo_10,nhk_ha_10,nAmount_10,nAmount_ha_10,ntotal_11,ntotal_bo_11,nhk_ha_11,nAmount_11,nAmount_ha_11,ntotal_12,ntotal_bo_12,nhk_ha_12,nAmount_12,nAmount_ha_12
' Dim	ntotal_sbi,ntotal_bo_sbi,nhk_ha_sbi,nAmount_sbi,nAmount_ha_sbi,ntotal_bud,ntotal_bo_bud,nhk_ha_bud,nAmount_bud,nAmount_ha_bud
Dim KBT_PML, KHT_PML, KHL_PML, KBT_PNN, KHT_PNN, KHL_PNN, KBT_MD, KHT_MD, KHL_MD, KBT_KU, KHT_KU, KHL_KU, KBT_TR, KHT_TR, KHL_TR 
Dim TOTALPML ,TOTALPNN ,TOTALMD ,TOTALTR ,TOTALKU ,TOTALALLKBT ,TOTALALLKHT ,TOTALALLKHL ,TOTALALL
Dim TOTALKBT_PML, TOTALKHT_PML, TOTALKHL_PML, TOTALKBT_PNN, TOTALKHT_PNN, TOTALKHL_PNN, TOTALKBT_MD, TOTALKHT_MD, TOTALKHL_MD, TOTALKBT_TR, TOTALKHT_TR, TOTALKHL_TR
Dim TOTALKBT_KU, TOTALKHT_KU, TOTALKHL_KU, TOTALKBT, TOTALKHL, TOTALKHT, TOTALALLSEC
Dim TOTALPMLALL, TOTALPNNALL, TOTALMDALL, TOTALTRALL, TOTALKUALL, TOTALSECALL
KBT_PML = Request("KBT_PML")
KHT_PML = Request("KHT_PML")
KHL_PML = Request("KHL_PML")
KBT_PNN = Request("KBT_PNN")
KHT_PNN = Request("KHT_PNN")
KHL_PNN = Request("KHL_PNN")
KBT_MD = Request("KBT_MD")
KHT_MD = Request("KHT_MD")
KHL_MD = Request("KHL_MD")
KBT_KU = Request("KBT_KU")
KHT_KU = Request("KHT_KU")
KHL_KU = Request("KHL_KU")
KBT_TR = Request("KBT_TR")
KHT_TR = Request("KHT_TR")
KHL_TR = Request("KHL_TR")
TOTALPML = Request("TOTALPML")
TOTALPNN = Request("TOTALPNN")
TOTALMD = Request("TOTALMD")
TOTALTR = Request("TOTALTR")
TOTALKU = Request("TOTALKU")
TOTALALLKBT = Request("TOTALALLKBT")
TOTALALLKHT = Request("TOTALALLKHT")
TOTALALLKHL = Request("TOTALALLKHL")
TOTALALL = Request("TOTALALL")
TOTALKBT_PML = Request("TOTALKBT_PML")
TOTALKHT_PML = Request("TOTALKHT_PML")
TOTALKHL_PML = Request("TOTALKHL_PML")
TOTALKBT_PNN = Request("TOTALKBT_PNN")
TOTALKHT_PNN = Request("TOTALKHT_PNN")
TOTALKHL_PNN = Request("TOTALKHL_PNN")
TOTALKBT_MD = Request("TOTALKBT_MD")
TOTALKHT_MD = Request("TOTALKHT_MD")
TOTALKHL_MD = Request("TOTALKHL_MD")
TOTALKBT_TR = Request("TOTALKBT_TR")
TOTALKHT_TR = Request("TOTALKHT_TR")
TOTALKHL_TR = Request("TOTALKHL_TR")
TOTALKBT_KU = Request("TOTALKBT_KU")
TOTALKHT_KU = Request("TOTALKHT_KU")
TOTALKHL_KU = Request("TOTALKHL_KU")
TOTALKBT = Request("TOTALKBT")
TOTALKHL = Request("TOTALKHL")
TOTALKHT = Request("TOTALKHT")
TOTALALLSEC = Request("TOTALALLSEC")
TOTALPMLALL = Request("TOTALPMLALL")
TOTALPNNALL = Request("TOTALPNNALL")
TOTALMDALL = Request("TOTALMDALL")
TOTALTRALL = Request("TOTALTRALL")
TOTALKUALL = Request("TOTALKUALL")
TOTALSECALL = Request("TOTALSECALL")

	if not rsMerumputTM.EOF then
		NO = 1
		KBT_PML = 0
		KHT_PML = 0
		KHL_PML = 0
		KBT_PNN = 0
		KHT_PNN = 0
		KHL_PNN = 0
		KBT_MD = 0
		KHT_MD = 0
		KHL_MD = 0
		KBT_KU = 0
		KHT_KU = 0
		KHL_KU = 0
		KBT_TR = 0
		KHT_TR = 0
		KHL_TR = 0
		TOTALKBT_PML = 0
		TOTALKHT_PML = 0 
		TOTALKHL_PML = 0 
		TOTALKBT_PNN = 0
		TOTALKHT_PNN = 0 
		TOTALKHL_PNN = 0 
		TOTALKBT_MD = 0
		TOTALKHT_MD = 0 
		TOTALKHL_MD = 0 
		TOTALKBT_TR = 0
		TOTALKHT_TR = 0 
		TOTALKHL_TR = 0 
		TOTALKBT_KU = 0
		TOTALKHT_KU = 0 
		TOTALKHL_KU = 0 
		TOTALKBT = 0
		TOTALKHT = 0 
		TOTALKHL = 0  
		TOTALALLSEC = 0 
		TOTALPMLALL = 0 
		TOTALPNNALL = 0 
		TOTALMDALL = 0
		TOTALTRALL = 0 
		TOTALKUALL = 0  
		TOTALSECALL = 0 
		TOTALPML = 0 
		TOTALPNN = 0
		TOTALMD = 0
		TOTALTR = 0
		TOTALKU = 0

		TOTALPML = cdbl(to_nol(rsMerumputTM("KBT_PML"))) + cdbl(to_nol(rsMerumputTM("KHT_PML"))) + cdbl(to_nol(rsMerumputTM("KHL_PML"))) 
		TOTALPNN = cdbl(to_nol(rsMerumputTM("KBT_PNN"))) + cdbl(to_nol(rsMerumputTM("KHT_PNN"))) + cdbl(to_nol(rsMerumputTM("KHL_PNN"))) 
		TOTALMD = cdbl(to_nol(rsMerumputTM("KBT_MD"))) + cdbl(to_nol(rsMerumputTM("KHT_MD"))) + cdbl(to_nol(rsMerumputTM("KHL_MD"))) 
		TOTALTR = cdbl(to_nol(rsMerumputTM("KBT_TR"))) + cdbl(to_nol(rsMerumputTM("KHT_TR"))) + cdbl(to_nol(rsMerumputTM("KHL_TR"))) 
		TOTALKU = cdbl(to_nol(rsMerumputTM("KBT_KU"))) + cdbl(to_nol(rsMerumputTM("KHT_KU"))) + cdbl(to_nol(rsMerumputTM("KHL_KU"))) 
		TOTALALLKBT = cdbl(to_nol(rsMerumputTM("KBT_PML"))) + cdbl(to_nol(rsMerumputTM("KBT_PNN"))) + cdbl(to_nol(rsMerumputTM("KBT_MD"))) + cdbl(to_nol(rsMerumputTM("KBT_TR"))) + cdbl(to_nol(rsMerumputTM("KBT_KU"))) 
		TOTALALLKHT = cdbl(to_nol(rsMerumputTM("KHT_PML"))) + cdbl(to_nol(rsMerumputTM("KHT_PNN"))) + cdbl(to_nol(rsMerumputTM("KHT_MD"))) + cdbl(to_nol(rsMerumputTM("KHT_TR"))) + cdbl(to_nol(rsMerumputTM("KHT_KU"))) 
		TOTALALLKHL = cdbl(to_nol(rsMerumputTM("KHL_PML"))) + cdbl(to_nol(rsMerumputTM("KHL_PNN"))) + cdbl(to_nol(rsMerumputTM("KHL_MD"))) + cdbl(to_nol(rsMerumputTM("KHL_TR"))) + cdbl(to_nol(rsMerumputTM("KHL_KU"))) 
		TOTALALL = TOTALALLKBT + TOTALALLKHT + TOTALALLKHL
		while not rsMerumputTM.EOF
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
%>
			<tr class="<%=cClass%>">
				<% if sprevSectionName <> rsMerumputTM("sectionname") then %>
							<td align="left" >
								<a href="p_laporanheadcount.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtSectionName=<%=tempSectionName%>&txtGangCode=<%=tempGangCode%>&txtGroup=<%=tempGroup%>">
							<%=rsMerumputTM("sectionname")%>
							</a>
							</td>
							<% sprevSectionName = rsMerumputTM("sectionname") %>	
						<% else	%>
							<td align="left" ></td>					
						<% end if  %>

				<% if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  "" then %>
					<%     tempPlantation = rsMerumputTM("Plantation")
						   tempGroupRegion = "--ALL--"
						   tempRegion = "--SUMMARY--"
						   tempCompany ="--SUMMARY--"
						   tempfcba = "--SUMMARY--"
						   tempSectionName = "--SUMMARY--"
						   
					%>
					<% if SprevPlantation <> rsMerumputTM("Plantation") then %>
						<td align="center">
							<a href="p_laporanheadcount.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtSectionName=<%=tempSectionName%>&txtGangCode=<%=tempGangCode%>&txtGroup=<%=tempGroup%>">
							<%=rsMerumputTM("Plantation")%>
							</a>
						</td>
						<% SprevPlantation = rsMerumputTM("plantation") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  "" then %>
				    <%  
					    if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then
							tempPlantation = rsMerumputTM("Plantation")
						else
							tempPlantation = ""
						end if
					   tempGroupRegion = rsMerumputTM("GroupRegion")
					   tempRegion = "--ALL--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					   tempSectionName = "--SUMMARY--"
					  
					%>
					
					<% if SprevGroupRegion <> rsMerumputTM("GroupRegion") then %>
						<td align="center">
							<a href="p_laporanheadcount.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtSectionName=<%=tempSectionName%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>&txtGangCode=<%=tempStatus%>&txtjobcode=<%=tempjobcode%>&txtGroup=<%=tempGroup%>">
							<%=rsMerumputTM("groupregion")%>
							</a>
						</td>
						<% SprevGroupRegion = rsMerumputTM("GroupRegion") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
				    <%  
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
					      tempPlantation = rsMerumputTM("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
					       tempGroupRegion = rsMerumputTM("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   tempRegion = rsMerumputTM("Region") 
					   tempCompany ="--ALL--"
					   tempfcba = "--SUMMARY--"
					   tempSectionName = "--SUMMARY--"
					  
					%>
					
					<% if SprevRegion <> rsMerumputTM("Region") then %>
						<td align="center">
							<a href="p_laporanheadcount.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtSectionName=<%=tempSectionName%>&txtGroup=<%=tempGroup%>">
							<%=rsMerumputTM("region")%>
							</a>
						</td>
						<% SprevRegion = rsMerumputTM("region") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then %>
				    <% 
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
					      tempPlantation = rsMerumputTM("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
					       tempGroupRegion = rsMerumputTM("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
					       tempRegion = rsMerumputTM("Region") 
					   else
					       tempRegion = ""
					   end if
					   tempCompany =rsMerumputTM("company")
					   tempfcba = "--ALL--"
					   tempSectionName = "--SUMMARY--"
					      
					 %>
					
					<% if SprevCompany <> rsMerumputTM("company") then %>
						<td align="center">
							<a href="p_laporanheadcount.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtSectionName=<%=tempSectionName%>&txtGangCode=<%=tempGangCode%>&txtGroup=<%=tempGroup%>">
							<%=rsMerumputTM("company")%>
							</a>
						</td>
						<% SprevCompany = rsMerumputTM("company") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	

				<% end if %>
				
				<% if v_txtFcba <>  "--SUMMARY--" and v_txtFcba <>  "" then %>
				    <% 
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
					      tempPlantation = rsMerumputTM("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
					       tempGroupRegion = rsMerumputTM("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
					       tempRegion = rsMerumputTM("Region") 
					   else
					       tempRegion = ""
					   end if
					   if  v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then
					       tempCompany =rsMerumputTM("company")
					   else
					       tempCompany =""
					   end if
					   tempfcba = rsMerumputTM("fcba")
					   tempSectionName = "--ALL--"
					  	   
					 %>
					
					<% if sprevFcba <> rsMerumputTM("fcba") then %>
						<td align="center">
							<a href="p_laporanheadcount.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtSectionName=<%=tempSectionName%>&txtGangCode=<%=tempStatus%>&txtGroup=<%=tempGroup%>">
						    <%=rsMerumputTM("fcba")%>
							</a>
						</td>
						<% sprevFcba = rsMerumputTM("fcba") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	

				<% end if %>
				
				<% if v_txtSectionName <> "--SUMMARY--" and v_txtSectionName <> "" then %>
				   <%  
				   	   
				       if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
					      tempPlantation = rsMerumputTM("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
					       tempGroupRegion = rsMerumputTM("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
					       tempRegion = rsMerumputTM("Region") 
					   else
					       tempRegion = ""
					   end if
					   if  v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then
					       tempCompany =rsMerumputTM("company")
					   else
					       tempCompany =""
					   end if
					   if   v_txtFcba <>  "--SUMMARY--" and v_txtFcba <>  "" then
					       tempfcba = rsMerumputTM("fcba")
					   else
					       tempfcba = ""
					   end if
					   if v_txtSectionName <>  "--SUMMARY--" and v_txtSectionName  <>  "" then 
					      tempSectionName = rsMerumputTM("sectionname")
					   else
					      tempSectionName = ""
                       end if					   
					  		   
					 %>						 
				<% end if %>
				
				<td align="right"><%=FbpNum2(rsMerumputTM("KBT_PML"))%></td>
				<td align="right"><%=FbpNum2(rsMerumputTM("KHT_PML"))%></td>
				<td align="right"><%=FbpNum2(rsMerumputTM("KHL_PML"))%></td>
				<td align="right"><%=FbpNum2(TOTALPML)%></td>

				<td align="right"><%=FbpNum2(rsMerumputTM("KBT_PNN"))%></td>
				<td align="right"><%=FbpNum2(rsMerumputTM("KHT_PNN"))%></td>
				<td align="right"><%=FbpNum2(rsMerumputTM("KHL_PNN"))%></td>
				<td align="right"><%=FbpNum2(TOTALPNN)%></td>

				<td align="right"><%=FbpNum2(rsMerumputTM("KBT_MD"))%></td>
				<td align="right"><%=FbpNum2(rsMerumputTM("KHT_MD"))%></td>
				<td align="right"><%=FbpNum2(rsMerumputTM("KHL_MD"))%></td>
				<td align="right"><%=FbpNum2(TOTALMD)%></td>

				<td align="right"><%=FbpNum2(rsMerumputTM("KBT_TR"))%></td>
				<td align="right"><%=FbpNum2(rsMerumputTM("KHT_TR"))%></td>
				<td align="right"><%=FbpNum2(rsMerumputTM("KHL_TR"))%></td>
				<td align="right"><%=FbpNum2(TOTALTR)%></td>

				<td align="right"><%=FbpNum2(rsMerumputTM("KBT_KU"))%></td>
				<td align="right"><%=FbpNum2(rsMerumputTM("KHT_KU"))%></td>
				<td align="right"><%=FbpNum2(rsMerumputTM("KHL_KU"))%></td>
				<td align="right"><%=FbpNum2(TOTALKU)%></td>

				<td align="right"><%=FbpNum2(TOTALALLKBT)%></td>
				<td align="right"><%=FbpNum2(TOTALALLKHT)%></td>
				<td align="right"><%=FbpNum2(TOTALALLKHL)%></td>
				<td align="right"><%=FbpNum2(TOTALALL)%></td>
			</tr>
<%
			NO = NO + 1			
			' ntotal_01 = ntotal_01 + cdbl(rsMerumputTM("total_01"))
			' ntotal_bo_01 = ntotal_bo_01 + cdbl(rsMerumputTM("total_bo_01"))
			' nhk_ha_01 = nhk_ha_01 + cdbl(rsMerumputTM("hk_ha_01"))
			' nAmount_01 = nAmount_01 + cdbl(rsMerumputTM("Amount_01"))
			' nAmount_ha_01 = nAmount_ha_01 + cdbl(rsMerumputTM("Amount_ha_01"))
			
			' ntotal_02 = ntotal_02 + cdbl(rsMerumputTM("total_02"))
			' ntotal_bo_02 = ntotal_bo_02 + cdbl(rsMerumputTM("total_bo_02"))
			' nhk_ha_02 = nhk_ha_02 + cdbl(rsMerumputTM("hk_ha_02"))
			' nAmount_02 = nAmount_02 + cdbl(rsMerumputTM("Amount_02"))
			' nAmount_ha_02 = nAmount_ha_02 + cdbl(rsMerumputTM("Amount_ha_02"))
            
			' ntotal_03 = ntotal_03 + cdbl(rsMerumputTM("total_03"))
			' ntotal_bo_03 = ntotal_bo_03 + cdbl(rsMerumputTM("total_bo_03"))
			' nhk_ha_03 = nhk_ha_03 + cdbl(rsMerumputTM("hk_ha_03"))
			' nAmount_03 = nAmount_03 + cdbl(rsMerumputTM("Amount_03"))
			' nAmount_ha_03 = nAmount_ha_03 + cdbl(rsMerumputTM("Amount_ha_03"))
			TOTALKBT_PML = TOTALKBT_PML + cdbl(to_nol(rsMerumputTM("KBT_PML")))
			TOTALKHT_PML = TOTALKHT_PML + cdbl(to_nol(rsMerumputTM("KHT_PML")))
			TOTALKHL_PML = TOTALKHL_PML + cdbl(to_nol(rsMerumputTM("KHL_PML")))
			TOTALPMLALL = TOTALPMLALL + TOTALPML
			
			TOTALKBT_PNN = TOTALKBT_PNN + cdbl(to_nol(rsMerumputTM("KBT_PNN")))
			TOTALKHT_PNN = TOTALKHT_PNN + cdbl(to_nol(rsMerumputTM("KHT_PNN")))
			TOTALKHL_PNN = TOTALKHL_PNN + cdbl(to_nol(rsMerumputTM("KHL_PNN")))
			TOTALPNNALL = TOTALPNNALL + TOTALPNN

			TOTALKBT_MD = TOTALKBT_MD + cdbl(to_nol(rsMerumputTM("KBT_MD")))
			TOTALKHT_MD = TOTALKHT_MD + cdbl(to_nol(rsMerumputTM("KHT_MD")))
			TOTALKHL_MD = TOTALKHL_MD + cdbl(to_nol(rsMerumputTM("KHL_MD")))
			TOTALMDALL = TOTALMDALL + TOTALMD

			TOTALKBT_TR = TOTALKBT_TR + cdbl(to_nol(rsMerumputTM("KBT_TR")))
			TOTALKHT_TR = TOTALKHT_TR + cdbl(to_nol(rsMerumputTM("KHT_TR")))
			TOTALKHL_TR = TOTALKHL_TR + cdbl(to_nol(rsMerumputTM("KHL_TR")))
			TOTALTRALL = TOTALTRALL + TOTALTR

			TOTALKBT_KU = TOTALKBT_KU + cdbl(to_nol(rsMerumputTM("KBT_KU")))
			TOTALKHT_KU = TOTALKHT_KU + cdbl(to_nol(rsMerumputTM("KHT_KU")))
			TOTALKHL_KU = TOTALKHL_KU + cdbl(to_nol(rsMerumputTM("KHL_KU")))
			TOTALKUALL = TOTALKUALL + TOTALKU

			TOTALKBT = TOTALKBT + TOTALALLKBT
			TOTALKHT = TOTALKHT + TOTALALLKHT
			TOTALKHL = TOTALKHL + TOTALALLKHL
			TOTALALLSEC = TOTALALLSEC + TOTALALL

			rsMerumputTM.MoveNext
		Wend
	end if
	rsMerumputTM.Close
	set rsMerumputTM = Nothing
%>
			<tr bgcolor="#F4F4A0">
				<td align="center"  colspan="<%= 1+ nSpanTotal %>">TOTAL</td>
				<td align="right"><%=FormatNumber(TOTALKBT_PML,0)%></td>
				<td align="right"><%=FormatNumber(TOTALKHT_PML,0)%></td>
				<td align="right"><%=FormatNumber(TOTALKHL_PML,0)%></td>
				<td align="right"><%=FormatNumber(TOTALPMLALL,0)%></td>

				<td align="right"><%=FormatNumber(TOTALKBT_PNN,0)%></td>
				<td align="right"><%=FormatNumber(TOTALKHT_PNN,0)%></td>
				<td align="right"><%=FormatNumber(TOTALKHL_PNN,0)%></td>
				<td align="right"><%=FormatNumber(TOTALPNNALL,0)%></td>
			    
			    <td align="right"><%=FormatNumber(TOTALKBT_MD,0)%></td>
				<td align="right"><%=FormatNumber(TOTALKHT_MD,0)%></td>
				<td align="right"><%=FormatNumber(TOTALKHL_MD,0)%></td>
				<td align="right"><%=FormatNumber(TOTALMDALL,0)%></td>

				<td align="right"><%=FormatNumber(TOTALKBT_TR,0)%></td>
				<td align="right"><%=FormatNumber(TOTALKHT_TR,0)%></td>
				<td align="right"><%=FormatNumber(TOTALKHL_TR,0)%></td>
				<td align="right"><%=FormatNumber(TOTALTRALL,0)%></td>

				<td align="right"><%=FormatNumber(TOTALKBT_KU,0)%></td>
				<td align="right"><%=FormatNumber(TOTALKHT_KU,0)%></td>
				<td align="right"><%=FormatNumber(TOTALKHL_KU,0)%></td>
				<td align="right"><%=FormatNumber(TOTALKUALL,0)%></td>

				<td align="right"><%=FormatNumber(TOTALKBT,0)%></td>
				<td align="right"><%=FormatNumber(TOTALKHT,0)%></td>
				<td align="right"><%=FormatNumber(TOTALKHL,0)%></td>
				<td align="right"><%=FormatNumber(TOTALALLSEC,0)%></td>
			  
			</tr>
     </tbody>       
</table>

</form>

</div>
</body>
</html>