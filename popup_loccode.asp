<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->  
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>ePlaS - Print Stock On Hand</title>
<link rel="stylesheet" href="css/stylereport2.css" type="text/css" />
<link rel="stylesheet" href="css/jquery-ui.css" />

<style type="text/css">


#btnselect {
  background-color: #4CAF50; /* Green */
  border: none;
  color: white;
  padding: 3.5px 8px;
  text-align: center;
  text-decoration: none;
  display: inline-block;
  font-size: 12px;
}


</style>    
</head>

<body>
<%

Dim sqljob,rsjob,v_txtFcba,txtSearch,txtColom,txtdate

  txtFcba = session("fcba") 
  txtSearch = request("txtSearch") 
  jobcodeTxt = request("jobcode")
  
  if request("txtdate") <>"" then
   txtdate   = request("txtdate")
  else
   txtdate   = textDate(date)
  end if
  
  q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & txtFcba & "' "
  set rsCat = connStr.Execute(q)
  if not rsCat.eof then
     fcbaname = rsCat("fcName")
     CompanyName=rsCat("FCCOMPANYNAME")
     MgrName = rsCat("FCMANAGER")
     KTUName = rsCat("FCKTU")
     CropName=rsCat("FCCROP")
     FcbaType=rsCat("fcType")

  end if

   if 	LEN(TRIM(txtSearch)) > 0 then
	      SELECT CASE TRIM(txtSearch) 
				CASE "FF"  
								txtColom = " SELECT FCCODE,TRIM(FCNAME)||'-'||TRIM(DIVISION)||'-'||TRIM(DX_BLOCK_STATUS) FCNAME FROM FIELD_HIST_TAB WHERE fcba = '" & txtFcba & "' AND DX_BLOCK_STATUS = (select JOB_FORFIELDSTATUS from JOBMASTER where fccode = '"& jobcodeTxt &"') AND Year*100+month=TO_NUMBER(TO_CHAR(TO_DATE('" & txtdate  & "','MM/DD/YYYY'),'YYYYMM')) AND ACTIVATION='Y' ORDER BY FCCODE "
				CASE "CB"  
								txtColom = " SELECT FCCODE,FCNAME FROM gcmaster WHERE fcba = '" & txtFcba & "' AND ACTIVATION='Y' AND Item='CB' ORDER BY FCCODE "					
				CASE "GC"  
								txtColom = " SELECT FCCODE,FCNAME FROM gcmaster WHERE fcba = '" & txtFcba & "' AND ACTIVATION='Y' AND Item='GC' ORDER BY FCCODE "					
				CASE "BA"  
								txtColom = " SELECT FCCODE,FCNAME FROM bank WHERE fcba = '" & txtFcba & "' AND ACTIVATION='Y' ORDER BY FCCODE "		
        CASE "MA"  
								txtColom = " SELECT FCCODE,Description FCNAME FROM Machine WHERE fcba = '" & txtFcba & "' AND ACTIVATION='Y' ORDER BY FCCODE "	
				CASE "VH"  
								txtColom = " SELECT FCCODE,FCNAME FROM Vehicle WHERE fcba = '" & txtFcba & "' AND ACTIVATION='Y' ORDER BY FCCODE "	
				CASE "ST"  
								txtColom = " SELECT FCCODE,FCNAME FROM store WHERE fcba = '" & txtFcba & "' AND ACTIVATION='Y' ORDER BY FCCODE "	
				CASE "MN"  
								txtColom = " SELECT NurseryCode FCCODE,Description FCNAME FROM nursery WHERE fcba = '" & txtFcba & "' AND Status='Y' ORDER BY NurseryCode "	
				CASE "PN"  
								txtColom = " SELECT NurseryCode FCCODE,Description FCNAME FROM nursery WHERE fcba = '" & txtFcba & "' AND Status='Y' ORDER BY NurseryCode "	
	      CASE "BM"  
								txtColom = " SELECT FCCODE,FCNAME FROM building WHERE fcba = '" & txtFcba & "' AND ACTIVATION='Y' ORDER BY FCCODE "			
				CASE "AP"  
								txtColom = " SELECT FCCODE,FCNAME FROM Supplier WHERE fcba = '" & txtFcba & "' AND ACTIVATION='Y' UNION ALL SELECT FCCODE,FCNAME FROM Transporter WHERE fcba = '" & txtFcba & "' AND ACTIVATION='Y' ORDER BY FCCODE "			
				CASE "AR"  
								txtColom = " SELECT CustomerCode FCCODE,Description FCNAME FROM customer WHERE fcba = '" & txtFcba & "' AND isactive='Y' ORDER BY CustomerCode "			
				CASE "CA"  
								txtColom = " SELECT FCCODE,FCNAME FROM Contractor WHERE fcba = '" & txtFcba & "' AND ACTIVATION='Y' ORDER BY FCCODE "			
				CASE "FA"  
								txtColom = " SELECT FCCODE,FCNAME FROM Fixedasset WHERE fcba = '" & txtFcba & "' AND isactive='1' UNION ALL  SELECT FCCODE,FCNAME FROM BUILDING WHERE fcba = '" & txtFcba & "' AND activation='Y' ORDER BY FCCODE "	
        CASE "WS"  
								txtColom = " SELECT FCCODE,FCNAME FROM Workshop WHERE fcba = '" & txtFcba & "' AND ACTIVATION='Y' ORDER BY FCCODE "									
				CASE ELSE				
	 							
				txtColom =  " SELECT '' FCCODE, '-' FCNAME " & _
	      				    " FROM DUAL"  

        END SELECT	
	else
	  txtColom =  " SELECT '' FCCODE, '-' FCNAME "
	end if

  response.write txtColom
  Set rsjob = Server.CreateObject("ADODB.Recordset")
  rsjob.Open txtColom, connstr, adOpenStatic 

%>
<div id="wrapper">
<div class="navi_to_print">

</div>
<div class="letter">
<div id="note">

<div class="page" >  
  <table border="0" width="100%" id="product-table">
    <thead>
      <tr>
        <th colspan="14" style=" background-color:#FFFFFF; color:#000000; border:none">
          <div id="header-wrapper">
             <div id="head_judul" style="float:left; text-align:left;">
              <h4>PT. <%=CompanyName%></h4>
              <h5><%=FcbaName%></h5>
            </div>
            <div id="subheader" class="centered" style="width:50%; margin:auto;"> 
              <h1 class="judul">Data Master Location Code</h1> 
            </div>  
            <div style="clear:both;"></div>
          </div>
        </th>
      </tr> 
     
      <tr>
        <th colspan="14" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></th>
      </tr>
     
        <th style="background-color: #32CD32; color: #fff;">Action</th>
        <th style="background-color: #32CD32; color: #fff;">Code</th>
        <th style="background-color: #32CD32; color: #fff;">Description</th>
        
      </tr>
    </thead>
  <tbody>
<%
   dim no 
   no =1
  while not rsjob.eof 

  %> 
    <tr class="<%=cClass%>">
      <td align="Center" ><font face="Verdana" size="1"><button id="btnselect" onclick="getSelected(<%=no%>)" >Select</button></td>
      <td align="Center" ><font face="Verdana" size="1"><%=rsjob("FCCODE")%>
        <input type=hidden id='locationcode<%=no%>' value="<%=rsjob("FCCODE")%>"></input>
      </td>
      <td align="left" ><font face="Verdana" size="1"><%=rsjob("FCNAME")%>
        <input type=hidden id='locationname<%=no%>' value="<%=rsjob("FCNAME")%>"></input>
      </td> 
      
      
    </tr>
<%      
      rsjob.MoveNext
      no = no +1
    Wend
  rsjob.Close
  set rsjob = Nothing
%> 
</tbody>  
</table> 
</body>
</html>
<script src="/js/jquery-1.9.1.js"></script>
<script>
 
  function getSelected(a) {

      var locationcode = $('#locationcode' + a).val();
      var locationname = $('#locationname' + a).val();
      window.opener.parent.document.getElementById("locationcode").value = locationcode;
      window.opener.parent.document.getElementById("locationname").innerHTML = locationname;
      window.close();
	}
</script>





