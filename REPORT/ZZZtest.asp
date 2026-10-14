<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta name="author" content="" />
<meta http-equiv="content-type" content="text/html;charset=iso-8859-1" />
<link rel="stylesheet" href="/css/stylereport.css" type="text/css" />
<title>iPla'S Reporting</title>
<script language="JavaScript" src='/js/jquery-1.8.3.js'></script>
<script type="text/javascript"> function scroll(element){   
var ele = document.getElementById(element);   
window.scrollTo(ele.offsetLeft,ele.offsetTop); } </script> 
<script>
var fixed = false;

$(document).scroll(function() {
    if( $(this).scrollTop() > 310 ) {
        if( !fixed ) {
            fixed = true;
            $('#logo-scroll').css({position:'fixed', display:'block', top:'0px'});
        }
    } else {
        if( fixed ) {
            fixed = false;
            $('#logo-scroll').css({display:'block', position:'static'});
        }
    }
});
</script>
</head>
<body>
<div id="siteInfo">
    <div class="head">
    	<h1>MENU </h1>
        <h1 style="color:#FFCC00;">LAPORAN</h1>
    </div>
</div>
<br/>
<div id="container2">
	<div class="line"></div>
<div class="nav">
<a href="#" onclick="javascript:window.history.back(-1);return false;"><img src="../IMAGES/back_03.gif" /></a>
</div>
	<!--<div class="clearfix">
		<div class="left"><a href="../login.asp" title="Back to Home"><img src="../images/latar.jpg" alt="RW Plant Reports"/></a></div>
		
		<div class='clear'></div>
	</div>
       <div class="colR">
   <div class="left"><a href="../login.asp" title="Back to Home"><img src="../images/latar.jpg" alt="RW Plant Reports"/></a></div>
   
   		<ul>
        	<li>
            </li>
        </ul>
   
   
   </div>
	<div class="colL">
    <!--End Header-->
	

	<!-- Modify by AKB, to handle the New Concept Dasboard Report of IPLAS, 17 May 2013 -->
	<%
	    Dim tempString, rsMenuString, tempstrong, rsMenuStrong
		Dim menuCode, menuParentCode, menuName, menuLink, menuHelp, moduleName,MenuTypeCode, strongCode, strongName
		Dim i,k,PUTARAN, headingrpt,subheadingrpt
					
		tempString = "  SELECT b.sequence, a.sequence sequence_line,a.fccode,a.fcname,a.menuparentcode,a.formlink,a.helpdescription,b.name,a.menutypecode" & _
					 "  FROM menu_tab a, module b" & _
					 "  WHERE a.isactive = 'TRUE' AND A.MODULE = B.CODE" & _
					 "  AND USER_ROLES_API.GetAccessUsers ('" & Session("CodeName") & "', a.fccode) = 'TRUE'" & _
				     "  AND a.menutypecode <> 'FORM'  AND  SUBSTR(a.fccode,1,3)='211'" & _
					 "  ORDER BY SEQUENCE, SEQUENCE_LINE	" 
	response.Write tempString
		Set rsMenuString = connStr.Execute(tempString)
		
		tempstrong = "  SELECT b.sequence, a.sequence sequence_line,a.fccode,a.fcname,a.menuparentcode,a.formlink,a.helpdescription,b.name,a.menutypecode" & _
					 "  FROM menu_tab a, module b" & _
					 "  WHERE a.isactive = 'TRUE' AND A.MODULE = B.CODE" & _
					 "  AND USER_ROLES_API.GetAccessUsers ('" & Session("CodeName") & "', a.fccode) = 'TRUE'" & _
				     "  AND a.menutypecode <> 'FORM'  AND MENUPARENTCODE IN('21100000')" & _
					 "  ORDER BY SEQUENCE, SEQUENCE_LINE	" 

		Set rsMenuStrong = connStr.Execute(tempstrong)
		
		headingrpt = "0"
        subheadingrpt = "0"
        PUTARAN = "0"
		i = "0" %>
  <div class="colR">
   <div ><a href="../login.asp" title="Back to Home"><img src="../images/latar.jpg" alt="RW Plant Reports"/></a></div>
      <div id="logo-scroll">
   		<ul>
        <%	While Not rsMenuStrong.EOF
			 strongCode = rsMenuStrong("fccode") 
			 strongName =  rsMenuStrong("fcname") 
			 
		%>
        	<li><a href="#<%=strongCode%>"><%=strongName%></a></li>
            
          <% rsMenuStrong.movenext 
		     wend
		   %>  
        </ul>
   
       </div>
   </div>
	<div class="colL">
        
        
	<%	While Not rsMenuString.EOF
			 menuCode = rsMenuString("fccode") 
			 menuParentCode =  rsMenuString("menuparentcode") 
			 menuName =  rsMenuString("fcname") 
			 menuLink = rsMenuString("formlink") 
			 menuHelp = rsMenuString("helpdescription")
			 moduleName = rsMenuString ("name")
			 MenuTypeCode = rsMenuString ("menutypecode") 
			 
			 IF MenuTypeCode = "REPORT" then
			 	   If cint(i)=0 then  %>
			 		    <ul>
			 		 <%
			 	   end if %>				 
			 	   <li><a href=<%=menuLink%>><%=menuName%></a></li>
			 	   <%
			 	   i= cint(i) + 1
			 	   if i= 5 then
                      PUTARAN = cint(PUTARAN) + 1 %>
					  </ul>
			          </div>
			          <div class="navlist2">	
                      <div class="guide"> &nbsp;</div>
                      <ul>			 
				      <%  
				      i = "0"
				   end if	
			 ELSE 
			      if  cint(i) > 0 then
					  for k = cint(i) to 4
					  %>
						  
				   <% next %>
                      </ul>
                      </div>
               <% end if
                  i=0
				  if  MenuTypeCode="HEADINGREPORT" and headingrpt = "0" then
			   %>
					  <div id="<%=menuCode%>" class="entry-header">
                      <h3><%=menuName%></h3></div>
					  <div class="entry-body">
					  <div class="navlist"> 
			  <%
				      headingrpt = menuCode
                      subheadingrpt = "01"
                      PUTARAN = "0"
                      i="0"
                   end if
			 
				   if MenuTypeCode="HEADINGREPORT" and headingrpt <> menuCode then
			 %>
                      </div>
                      <div class=clear></div>
                      <div id="<%=menuCode%>" class="entry-header"><h3><%=menuName%></h3></div>
					  <div class="entry-body">
		    <%			 
                      PUTARAN = "0"
                      headingrpt = menuCode 
                      subheadingrpt = "0"
                      i= "0"
                  end if
				  
				  if  MenuTypeCode="SUBHEADINGREPORT" then
                      PUTARAN = cint(PUTARAN) + 1
    
                      if ( cint(PUTARAN) mod 3 = 1 ) and cint(PUTARAN) > 3 then
		  %>			
                         </div>    
                         <div class="clear"></div>						 
						 <div class="entry-body">
		  <%
                      end if
		  %>
		  <% If subheadingrpt <> "01" then %>
					  <div class=navlist>
		  <% end if %>			  
                      <h4><%=menuName%></h4>
		  <%
                      subheadingrpt = menuCode 
                      i= "0"
                  end if 
		     END IF
		     rsMenuString.MoveNext 
	    Wend 
		
		 if  cint(i) > 0 then
			 for k = cint(i) to 4
 %>
				<li>&nbsp;</li>
 <%          next %>
             </ul>
             </div>
 <%     end if    %>	
		
	</div><!--End Content-->
	<div class='clear'></div>
    
</div><!--End Container-->

<div class='clear'></div>
</body>
</html>
