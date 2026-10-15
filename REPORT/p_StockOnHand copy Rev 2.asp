<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%
	Dim rsStockOnHand, strStockOnHand
	Dim v_txtStockCode
	Dim NO, cClass, nTotal

	v_txtStockCode		=	trim(request("txtStockCode"))
	v_txtGroupPlantation  = Request("txtGroupPlantation")
	v_txtPlantation       = Request("txtPlantation")
	v_txtGroupRegion      = Request("txtGroupRegion")
	v_txtRegion           = Request("txtRegion")
	v_txtcompany          = Request("txtCompany")
	v_txtfcba             = Request("txtfcba")
	v_txtDivision         = Request("txtDivision")
	v_txtBlok             = Request("txtBlok")
	v_txtPlantingDate     = Request("txtPlantingDate")
	v_txttodaydate        = Request("txttodaydate")
	v_txtYear             = Request("txtYear")
	v_txtMonth            = Request("txtmonth")
	v_txtcustomercode       = Request("txtcustomercode")
	v_txtownership         = Request("txtownership")
	v_txtSupplier         = Request("txtSupplier")  
	kurs                  = Request("kurs")  
	v_txtitemcode         = Request("txtitemcode") 
    v_txtitemname         = Request("txtitemname") 

	txtyear       = right(Request("txttodaydate"),4)
	txtmonth      = left(Request("txttodaydate"),2)
	txttodaydate  = Request("txttodaydate")


	txtComp         = Request("txtCompany")
	v_txtdate_from = trim(Request("txtdate_from"))
	v_txtdate_to   = trim(Request("txtdate_to "))

	if "" & v_txtdate_from = "" then v_txtdate_from = textDate(date)
	if "" & v_txtdate_to = "" then v_txtdate_to = textDate(date)

	nSpanTotal = 8
	v_txtComp           =   txtComp

	if request("txtYear")="" then
	 v_txtYear  = Year(Session("v_Costbook"))
	end if
	if request("txtMonth")="" then
	 v_txtMonth = Month(Session("v_Costbook"))
	end if

      strDefaultBa =  " SELECT" & _
                      "     FCCODE as FCBA" & _
                      " FROM BUSINESSUNIT BU" & _
                      " INNER JOIN USER_FCBA_TAB UFT ON (BU.FCCODE = UFT.FCBA) " & _
                      " WHERE" & _
                      "     ISDEFAULT = 'TRUE'" & _
                      "     AND FCTYPE = 'E'" & _
                      "     AND FCUSERCODE = '" & Session("CodeName") & "'"

      Set rsDefaultBa = Server.CreateObject("ADODB.Recordset")
      rsDefaultBa.Open strDefaultBa, connstr, adOpenStatic

      If Not rsDefaultBa.EOF Then
        varDefaultBa = rsDefaultBa("FCBA")
      Else
        varDefaultBa = ""
      End If
      rsDefaultBa.Close
      set rsDefaultBa = Nothing

      If (v_txtFcba = "" Or Len(Trim(v_txtFcba)) = 0) Then
        v_txtFcba = varDefaultBa
      End If  

	Dim selectFields,strFieldCons,selectCons,whereCons,GroupCons,orderCons 
    strFieldCons =  "   USER_COMPANY_API.GetGroupPlantationCode('" & Session("CodeName") & "', fcba) GroupPlantation" & _
                    " , USER_COMPANY_API.GetPlantationCode('" & Session("CodeName") & "', fcba) Plantation" & _
                    " , USER_COMPANY_API.GetGroupRegionCode('" & Session("CodeName") & "', fcba) groupregion" & _
                    " , USER_COMPANY_API.GetRegionCode('" & Session("CodeName") & "', fcba) region" & _
                    " , USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "', fcba) company" & _
                    " , BUSINESSUNIT_API.get_OrderData(fcba) NoUrut, fcba,"

             
    wherecons = " where company is not null " 

    q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where FCCOMPANYCODE = '" & v_txtCompany & "'"  
      set rsCat = connStr.Execute(q)
      if not rsCat.eof then
         CompanyName=rsCat("FCCOMPANYNAME")        
        end if 

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

    If Len(Trim(v_txtGroupPlantation)) > 0 Then
      If v_txtGroupPlantation = "--SUMMARY--" then
        selectCons  = selectCons  & " "
        selectCons2 = selectCons2  & " " 
        whereCons = whereCons & " "
        orderCons = orderCons & " "
        orderCons2 = orderCons2 & " "
      Elseif v_txtGroupPlantation = "--ALL--" then
        selectCons  = selectCons  & " "
        selectCons2 = selectCons2  & " " 
        whereCons = whereCons & " "
        orderCons = orderCons & " "
        orderCons2 = orderCons2 & " " 
        var_colspan = var_colspan + 1
      Else
        selectCons  = selectCons  & " "
        selectCons2 = selectCons2  & " " 
        whereCons = whereCons & " And GroupPlantation = '" & v_txtGroupPlantation & "'"
        orderCons = orderCons & " " 
        orderCons2 = orderCons2 & " "
        var_colspan = var_colspan + 1
      End if
    End if

    If Len(Trim(v_txtPlantation)) > 0 Then
      If v_txtPlantation = "--SUMMARY--" then
        selectCons  = selectCons  & " " 
        selectCons2 = selectCons2  & " " 
        whereCons = whereCons & " "
        orderCons = orderCons & " "
        orderCons2 = orderCons2 & " "
      Elseif v_txtPlantation = "--ALL--" then
        selectCons  = selectCons  & " , Plantation" 
        selectCons2 = selectCons2  & " , Plantation" 
        whereCons = whereCons & " "
        orderCons = orderCons & " , Plantation" 
        orderCons2 = orderCons2 & " , Plantation" 
        GroupCons = GroupCons & " , Plantation" 
        GroupCons2  =  GroupCons2 & " , Plantation"
        var_colspan = var_colspan + 1
      Else
        selectCons  = selectCons  & " , Plantation"
        selectCons2 = selectCons2  & " , Plantation" 
        whereCons = whereCons & " And Plantation = '" & v_txtPlantation & "'"
        orderCons = orderCons & " , Plantation" 
        orderCons2 = orderCons2 & " , Plantation" 
        GroupCons   =  GroupCons & " , Plantation"
        GroupCons2  =  GroupCons2 & " , Plantation"
        var_colspan = var_colspan + 1
      End if
    End if

    If Len(Trim(v_txtRegion)) > 0 Then
      If v_txtRegion = "--SUMMARY--" then
        selectCons  = selectCons  & " " 
        selectCons2 = selectCons2  & " " 
        whereCons = whereCons & " "
        orderCons = orderCons & " "
      Elseif v_txtRegion = "--ALL--" then
        selectCons  = selectCons  & " , Region"
        selectCons2 = selectCons2  & " , Region"  
        whereCons = whereCons & " "
        orderCons = orderCons & " , Region" 
        orderCons2 = orderCons2 & " , Region" 
        GroupCons   =  GroupCons & " , Region"
        GroupCons2  =  GroupCons2 & " , Region"
        var_colspan = var_colspan + 1
      Else
        selectCons  = selectCons  & " , Region" 
        selectCons2 = selectCons2  & " , Region" 
        whereCons = whereCons & " And Region = '" & v_txtRegion & "'"
        orderCons = orderCons & " , Region" 
        orderCons2 = orderCons2 & " , Region" 
        GroupCons   =  GroupCons & " , Region"
        GroupCons2  =  GroupCons2 & " , Region"
        var_colspan = var_colspan + 1
      End if
    End if

    If Len(Trim(v_txtGroupRegion)) > 0 Then
      If v_txtGroupRegion = "--SUMMARY--" then
        selectCons  = selectCons  & " "
        selectCons2 = selectCons2  & " " 
        whereCons = whereCons & " "
        orderCons = orderCons & " "
        orderCons2 = orderCons2 & " "
      Elseif v_txtGroupRegion = "--ALL--" then
        selectCons  = selectCons  & " , GroupRegion" 
        whereCons = whereCons & " "
        orderCons = orderCons & " , GroupRegion" 
        orderCons2 = orderCons2 & " , GroupRegion" 
        GroupCons   =  GroupCons & " ,  GroupRegion"
        GroupCons2  =  GroupCons2 & " ,  GroupRegion"
        var_colspan = var_colspan + 1
      Else
        selectCons  = selectCons  & " , GroupRegion"
        selectCons2  = selectCons2 & " , GroupRegion" 
        whereCons = whereCons & " And GroupRegion = '" & v_txtGroupRegion & "'"
        orderCons = orderCons & " , GroupRegion"
        orderCons2 = orderCons2 & " , GroupRegion" 
        GroupCons   =  GroupCons & " , GroupRegion"
        GroupCons2  =  GroupCons2 & " , GroupRegion"
        var_colspan = var_colspan + 1
      End if
    End if

    If Len(Trim(v_txtCompany)) > 0 Then
      If v_txtCompany = "--SUMMARY--" then
        selectCons  = selectCons  & " " 
        selectCons2  = selectCons2  & " " 
        whereCons = whereCons & " "
        orderCons = orderCons & " "
        orderCons2 = orderCons2 & " "
      Elseif v_txtCompany = "--ALL--" then
        selectCons  = selectCons  & " , company"
        selectCons2  = selectCons2  & " , company" 
        whereCons = whereCons & " "
        orderCons = orderCons & " , COMPANY_API.get_OrderData(company)"
        orderCons2 = orderCons2 & " , COMPANY_API.get_OrderData(company)" 
        GroupCons   =  GroupCons & " , company"
        GroupCons2  =  GroupCons2 & " , company"
        var_colspan = var_colspan + 1
      Else
        selectCons  = selectCons  & " , company"
        selectCons2  = selectCons2  & " , company" 
        whereCons = whereCons & " And company = '" & v_txtCompany & "'"
        orderCons = orderCons & " , COMPANY_API.get_OrderData(company)"
        orderCons2 = orderCons2 & " , COMPANY_API.get_OrderData(company)"
        GroupCons   =  GroupCons & " , company" 
        GroupCons2  =  GroupCons2 & " , company" 
        var_colspan = var_colspan + 1
      End if
    End if

    If Len(Trim(v_txtfcba)) > 0 Then
      If v_txtfcba = "--SUMMARY--" then
        selectCons  = selectCons  & " "
        selectCons2 = selectCons2  & " " 
        whereCons = whereCons & " "
        whereCons2 = whereCons2 & " "
        orderCons = orderCons & " "
      Elseif v_txtfcba = "--ALL--" then
        selectCons  = selectCons  & " , fcba"
        selectCons2 = selectCons2  & " " 
        whereCons = whereCons & " "
        whereCons2 = whereCons2 & " "
        orderCons = orderCons & " , fcba"
        GroupCons   =  GroupCons & " , fcba"
        var_colspan = var_colspan + 1
      Else
        selectCons  = selectCons  & " , fcba" 
        selectCons2 = selectCons2  & " " 
        whereCons = whereCons & " And fcba = '" & v_txtfcba & "'"
        whereCons2 = whereCons2 & " And fcba = '" & v_txtfcba & "'"
        orderCons = orderCons & " , fcba"
        GroupCons =  GroupCons & " , fcba"
        var_colspan = var_colspan + 1
      End if
    End if 

    If Len(Trim(v_txtitemcode)) > 0 Then
      If v_txtitemcode = "--SUMMARY--" then
        selectCons  = selectCons  & " "
        selectCons2 = selectCons2  & " " 
        whereCons = whereCons & " "
        whereCons2 = whereCons2 & " "
        orderCons = orderCons & " "
      Elseif v_txtitemcode = "--ALL--" then
        selectCons  = selectCons  & " "
        selectCons2 = selectCons2  & " " 
        whereCons = whereCons & " " 
        whereCons2 = whereCons2 & " "
        orderCons = orderCons & " "
        GroupCons   =  GroupCons & " " 
      Else
        selectCons  = selectCons  & " " 
        selectCons2 = selectCons2  & " " 
        whereCons = whereCons & " And stockcode like ('%" & v_txtitemcode & "%') " 
        orderCons = orderCons & " "
        GroupCons =  GroupCons & " " 
      End if
    End if

    If Len(Trim(v_txtitemname)) > 0 Then
      If v_txtitemname = "--SUMMARY--" then
        selectCons  = selectCons  & " "
        selectCons2 = selectCons2  & " " 
        whereCons = whereCons & " "
        whereCons2 = whereCons2 & " "
        orderCons = orderCons & " "
      Elseif v_txtitemname = "--ALL--" then
        selectCons  = selectCons  & " "
        selectCons2 = selectCons2  & " " 
        whereCons = whereCons & " " 
        whereCons2 = whereCons2 & " "
        orderCons = orderCons & " "
        GroupCons   =  GroupCons & " " 
      Else
        selectCons  = selectCons  & " " 
        selectCons2 = selectCons2  & " " 
        whereCons = whereCons & " And stockname like upper('%" & v_txtitemname & "%')  " 
        orderCons = orderCons & " "
        GroupCons =  GroupCons & " " 
      End if
    End if

    
	
	strStockOnHand		=	" select "& strFieldCons &"" & _
	                    "        storecode" & _
        							"        , store_api.get_fcname(storecode, fcba) storename" & _
        							"        , stockcode" & _
        							"        , stock_api.get_stockname(stockcode, fcba) stockname" & _
        							"        , stock_api.get_uom(stockcode, fcba) uom" & _
        							"        , NVL(quantity, 0) quantity" & _
        							" from stockonhand " & _
        							" where stockcode not in (" & _
        							"                 select fccode from stock" & _
        							"                 where  fcba||fccode in (" & _
        							"                                  select fcba||stockcode from nursery where fcba = '"& v_txtFcba &"' )) AND quantity > 0" 


	' if  len(trim(v_txtStockCode)) > 0 then
	' 	strStockOnHand	=	strStockOnHand & " and stockcode = '"& v_txtStockCode &"'"
	' end if 
	' 	strStockOnHand	=	strStockOnHand & " order by storecode, stockcode"

  strEPS2 = "        , storecode" & _
			"        , storename" & _
			"        , stockcode" & _
			"        , stockname" & _
			"        , uom" & _
			"        , NVL(SUM(quantity), 0) quantity"

  strEPS =    " SELECT GroupPlantation " & selectCons & strEPS2 &"  From (" & strStockOnHand & ")" & whereCons & ""  & _ 
  			  "	GROUP BY GroupPlantation ,storecode" & _
			"        , storename" & _
			"        , stockcode" & _
			"        , stockname" & _
			"        , uom" & GroupCons & "" & _
              " ORDER BY GroupPlantation, storecode, stockcode " & orderCons & ""

	 response.write strEPS
	Set rsStockOnHand = Server.CreateObject("ADODB.Recordset")
	rsStockOnHand.Open strEPS, connstr, adOpenStatic 
	
%>


<html>

<head>
	<title>ePlaS Reporting - Stock On Hand</title>
	<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
	<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
	<link rel="stylesheet" href="/css/jquery-ui.css" />
	<script src="/js/jquery-1.9.1.js"></script>
	<script src="/js/jquery-ui.js"></script>
	<link href="../plugins/tabcontent/tabcontent.css" rel="stylesheet" type="text/css" />
	<link href="../plugins/chosen/chosen.min.css" rel="stylesheet" />

	<link rel="stylesheet" href="../plugins/responsivegridsystem/col.css" media="all">
	<link rel="stylesheet" href="../plugins/responsivegridsystem/12cols.css" media="all">
	<script language="JavaScript" src='/js/x.js'></script>
	<script language="JavaScript" src='/js/xtableheaderfixed.js'></script>
	<script language="JavaScript">
	xAddEventListener(window, 'load',
	  function() {
		new xTableHeaderFixed('xthf-site', window);
	  }, false
	);
	</script>
	 <script>
	$(function() {
		$( "#tgl" ).datepicker();
	});
	</script>
	<script language="JavaScript" src="/js/date-picker.js"></script> 
	<script language="JavaScript">
	var newwin;
	function launchwin(winurl,winname,winfeatures)
	{
		//This launches a new window and then
		//focuses it if window.focus() is supported.
		newwin = window.open(winurl,winname,winfeatures);

	}
	</script>
	
	<style type="text/css">
	@media print
	{
	.printhide{display:none;}
	}
	</style>
</head>

<body>
	<div id="wrapper">
		<div id="header-wrapper">
			<div class="printhide">
				<!--#include file="../print_excel.asp"-->
				<!--#include file="menutop.asp"-->
			</div>
			<form name="frmEntry" method="post">
			<div class="printhide">
				<!--#include file="FormSelectNewSum.html"-->
				<table id="form-table" class="centered-margin">
					<tr>
						<th> HARI INI</th>
						<th> ITEM CODE</th>
            			<th> ITEM NAME</th>
					</tr>
					<tr>
						<td>
							<input type="text" name="tgl" id="tgl" value="<%=TextDate(date())%>" size="30">
						</td>
						<td>
						    <input type="text" name="txtitemcode" value="<%=v_txtitemcode%>" size="30">
						</td>   			 
            <td>
                <input type="text" name="txtitemname" value="<%=v_txtitemname%>" size="30">
            </td>                
						<td align="center"><input type="submit" value="View" /></td>
                    </tr>
				</table>
			</div>
			</form>
		</div>
		
		<table border="0" width="100%" id="product-table" class="xthf-site">
			<thead>
				<tr>
					<th colspan="<%=6 + var_colspan%>" style="background-color:#FFFFFF; border:none;">  
						<div id="header-wrapper"> 
							<div id="subheader" class="centered" tyle="width:50%; margin:auto;">
								<h2>STOCK ON HAND</h2>
								<% If Len(Trim(v_txtStockCode)) > 0 Then %>
								<h4>Stock Code : <%=v_txtStockCode%></h4>
								<br/>
								<%End If%>
							</div>
						</div>
					</th>
				</tr>
				<tr>
				  <% if v_txtGroupPlantation <> "--SUMMARY--" and v_txtGroupPlantation <> "" then %>
			            <th style="background-color:#00A300;color:#fff;" rowspan="2">GROUP PLANTATION</th>
			      <%      
			        end if %>

			      <% if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then %>
			            <th style="background-color:#00A300;color:#fff;" rowspan="2">PLANTATION</th>
			      <%      
			        end if %>
			      
			      <% if v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  "" then %>
			            <th style="background-color:#00A300;color:#fff;" rowspan="2">GROUP REGION</th>
			      <%      
			        end if %>
			      
			      <% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
			            <th style="background-color:#00A300;color:#fff;" rowspan="2">REGION</th>
			      <%
			        end if %>
			      
			      <% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then %>
			             <th style="background-color:#00A300;color:#fff;" rowspan="2">COMPANY</th>
			      <%    
			        end if %>
			      
			      <% if v_txtfcba <>  "--SUMMARY--" and v_txtfcba <>  "" then %>
			             <th style="background-color:#00A300;color:#fff;" rowspan="2">SITE</th>
			      <%    
			        end if %> 
					<th>STORE CODE</th>
					<th>STORE NAME</th>
					<th>STOCK CODE</th>
					<th>STOCK NAME</th>
					<th>UOM</th>
					<th>QUANTITY</th>
				</tr>
			</thead>
		<tbody>
<%if rsStockOnHand.eof and rsStockOnHand.eof then%>
	<tr>
		<td colspan="<%=5 + var_colspan%>" align="left">NO DATA AVAILABLE</td>
	</tr>
<%end if
		NO = 1
		nTotal = 0
		while not rsStockOnHand.eof
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
%>
			<tr class="<%=cClass%>">
				<% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <>  "" then %>
			          <% tempPlantation = rsStockOnHand("Plantation")
			             tempGroupRegion = "--ALL--"
			             tempRegion = "--SUMMARY--"
			             tempCompany ="--SUMMARY--"
			             tempfcba = "--SUMMARY--"
			             tempDivision = "--SUMMARY--"
			             tempPlantingDate = "--SUMMARY--"
			             tempBlok = "--SUMMARY--"
			          %>
			          <% if SprevGroupPlantation <> rsStockOnHand("GroupPlantation") then %>
			            <td align="center">
			              <a href="<%=ScriptName%>?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
			              <%=rsStockOnHand("GroupPlantation")%>
			              </a>
			            </td>
			            <% 
			              SprevGroupPlantation = rsStockOnHand("GroupPlantation")
			            %>  
			            
			          <% else %>
			            <td align="center"></td>          
			          <% end if %>  
			        <% end if %>

			        <% if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  "" then %>
			          <% tempPlantation = rsStockOnHand("Plantation")
			             tempGroupRegion = "--ALL--"
			             tempRegion = "--SUMMARY--"
			             tempCompany ="--SUMMARY--"
			             tempfcba = "--SUMMARY--"
			             tempDivision = "--SUMMARY--"
			             tempPlantingDate = "--SUMMARY--"
			             tempBlok = "--SUMMARY--"
			          %>
			          <% if SprevPlantation <> rsStockOnHand("Plantation") then %>
			            <td align="center">
			              <a href="<%=ScriptName%>?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
			              <%=rsStockOnHand("Plantation")%>
			              </a>
			            </td>
			            <% 
			              SprevPlantation = rsStockOnHand("plantation")
			            %>  
			            
			          <% else %>
			            <td align="center"></td>          
			          <% end if %>  
			        <% end if %>
			        
			        <% if v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  "" then %>
			            <%  if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then
			              tempPlantation = rsStockOnHand("Plantation")
			            else
			              tempPlantation = ""
			            end if
			             tempGroupRegion = rsStockOnHand("GroupRegion")
			             tempRegion = "--ALL--"
			             tempCompany ="--SUMMARY--"
			             tempfcba = "--SUMMARY--"
			             tempDivision = "--SUMMARY--"
			             tempPlantingDate = "--SUMMARY--"
			             tempBlok = "--SUMMARY--"
			          %>
			          
			          <% if SprevGroupRegion <> rsStockOnHand("GroupRegion") then %>
			            <td align="center">
			              <a href="<%=ScriptName%>?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
			              <%=rsStockOnHand("groupregion")%>
			              </a>
			            </td>
			            <% SprevGroupRegion = rsStockOnHand("GroupRegion") %> 
			          <% else %>
			            <td align="center"></td>          
			          <% end if %>  
			        <% end if %>
			        
			        <% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
			            <% if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
			                tempPlantation = rsStockOnHand("Plantation")
			             else
			                tempPlantation = ""
			             end if
			             if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
			                 tempGroupRegion = rsStockOnHand("GroupRegion")
			             else
			                 tempGroupRegion = ""
			             end if
			             
			             tempRegion = rsStockOnHand("Region") 
			             tempCompany ="--ALL--"
			             tempfcba = "--SUMMARY--"
			             tempDivision = "--SUMMARY--"
			             tempPlantingDate = "--SUMMARY--"
			             tempBlok = "--SUMMARY--"
			          %>
			          
			          <% if SprevRegion <> rsStockOnHand("Region") then %>
			            <td align="center">
			              <a href="<%=ScriptName%>?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
			              <%=rsStockOnHand("region")%>
			              </a>
			            </td>
			            <% SprevRegion = rsStockOnHand("region") %> 
			          <% else %>
			            <td align="center"></td>          
			          <% end if %>  
			        <% end if %>
			        
			        <% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then %>
			            <% 
			             if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
			                tempPlantation = rsStockOnHand("Plantation")
			             else
			                tempPlantation = ""
			             end if
			             if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
			                 tempGroupRegion = rsStockOnHand("GroupRegion")
			             else
			                 tempGroupRegion = ""
			             end if
			             if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
			                 tempRegion = rsStockOnHand("Region") 
			             else
			                 tempRegion = ""
			             end if
			             tempCompany =rsStockOnHand("Company")
			             tempfcba = "--ALL--"
			             tempDivision = "--SUMMARY--"
			             tempPlantingDate = "--SUMMARY--"
			             tempBlok = "--SUMMARY--"            
			           %>
			          
			          <% if SprevCompany <> rsStockOnHand("Company") then %>
			            <td align="center">
			              <a href="<%=ScriptName%>?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
			              <%=rsStockOnHand("company")%>
			              </a>
			            </td>
			            <% SprevCompany = rsStockOnHand("Company") %> 
			          <% else %>
			            <td align="center"></td>          
			          <% end if %>  

			        <% end if %>
			        
			        <% if v_txtFcba <>  "--SUMMARY--" and v_txtFcba <>  "" then %>
			            <% if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
			                tempPlantation = rsStockOnHand("Plantation")
			             else
			                tempPlantation = ""
			             end if
			             if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
			                 tempGroupRegion = rsStockOnHand("GroupRegion")
			             else
			                 tempGroupRegion = ""
			             end if
			             if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
			                 tempRegion = rsStockOnHand("Region") 
			             else
			                 tempRegion = ""
			             end if
			             if  v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then
			                 tempCompany =rsStockOnHand("Company")
			             else
			                 tempCompany =""
			             end if
			             tempfcba = rsStockOnHand("fcba")
			             tempDivision = "--ALL--"
			             tempPlantingDate = "--SUMMARY--"
			             tempBlok = "--SUMMARY--"            
			           %>
			          
			          <% if sprevFcba <> rsStockOnHand("fcba") then %>
			            <td align="center">
			              <a href="<%=ScriptName%>?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
			                <%=rsStockOnHand("fcba")%>
			              </a>
			            </td>
			            <% sprevFcba = rsStockOnHand("fcba") %> 
			          <% else %>
			            <td align="center"><%=rsStockOnHand("fcba")%></td>          
			          <% end if %>   
			        <% end if %>
				<td align="left"><%=rsStockOnHand("STORECODE")%></td>
				<td align="left"><%=rsStockOnHand("STORENAME")%></td>
				<td align="left"><%=rsStockOnHand("STOCKCODE")%></td>
				<td align="left"><%=rsStockOnHand("STOCKNAME")%></td>
				<td align="left"><%=rsStockOnHand("UOM")%></td>
				<td align="right"><%=formatnumber(cdbl(to_nol(rsStockOnHand("QUANTITY"))),2)%></td>				
			</tr>
<%
			NO = NO + 1
			nTotal = nTotal + Cdbl(rsStockOnHand("QUANTITY"))
			rsStockOnHand.MoveNext
		Wend
	rsStockOnHand.Close
	set rsStockOnHand = Nothing
%>
			<tr style="background-color:#F4F4A0;">
				<td align="center" colspan="<%=5 + var_colspan%>"><strong>TOTAL</strong></td>
				<td align="right"><strong><%=FbpNum2(nTotal)%></strong></td>
			</tr>
 </tbody>
</table>
   <table width="90%" class="data" style="margin:auto;">
   	   <tr> 
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	   </tr>
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
		<td align="center"><font face="Verdana" size="2"><strong><u>&nbsp;&nbsp;<%=Session("FCMANAGER")%>&nbsp;&nbsp;</u></strong></td>
		<td align="center"><font face="Verdana" size="2"><strong><u>&nbsp;&nbsp;<%=Session("KTU")%>&nbsp;&nbsp;</u></strong></td>
	   </tr>
	   	  <tr> 
		<td align="center"><font face="Verdana" size="1"><strong><i>MANAGER</i></strong></td>
		<td align="center"><font face="Verdana" size="1"><strong><i>KTU</i></strong></td>
	   </tr>
   </table>
  <script src="../js/jquery.inputmask.bundle.min.js"></script>
  <script src="../js/inputmask.binding.js"></script>
  <script src="../plugins/chosen/chosen.jquery.min.js"></script>
  <script type="text/javascript"> 
    // $(".chosen-select").empty();
    
    var config = {
      '.chosen-select'           : {},
      '.chosen-select-deselect'  : {allow_single_deselect:true},
      '.chosen-select-no-single' : {disable_search_threshold:10},
      '.chosen-select-no-results': {no_results_text:'Oops, nothing found!'},
      '.chosen-select-width'     : {width:"45%"}
    }
    for (var selector in config) {
      $(selector).chosen(config[selector]);
    }
    // $(".chosen-select").trigger("chosen:updated");

  </script>
</div>