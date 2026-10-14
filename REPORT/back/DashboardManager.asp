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
</head>
<body>
<div id="container">
	<div id="header" class="clearfix">
		<div class="left"><a href="../login.asp" title="Back to Home"><img src="../images/latar.jpg" alt="RW Plant Reports"/></a></div>
		<div class="right"><h1>Reports</h1></div>
		<div class='clear'></div>
	</div><!--End Header-->
	<div class='clear'></div>
	<div id="content">
	<!-- Modify by AKB, to handle the New Concept Dasboard Report of IPLAS, 17 May 2013 -->
	<%
	    Dim tempString, rsMenuString
		Dim menuCode, menuParentCode, menuName, menuLink, menuHelp, moduleName,MenuTypeCode
		Dim i,k,PUTARAN, headingrpt,subheadingrpt
					
		tempString = "  SELECT b.sequence, a.sequence sequence_line,a.fccode,a.fcname,a.menuparentcode,a.formlink,a.helpdescription,b.name,a.menutypecode" & _
					 "  FROM menu_tab a, module b" & _
					 "  WHERE a.isactive = 'TRUE' AND A.MODULE = B.CODE" & _
					 "  AND USER_ROLES_API.GetAccessUsers ('" & Session("CodeName") & "', a.fccode) = 'TRUE'" & _
				     "  AND a.menutypecode <> 'FORM'  AND  SUBSTR(a.fccode,1,3)='212'" & _
					 "  ORDER BY SEQUENCE, SEQUENCE_LINE	" 

		Set rsMenuString = connStr.Execute(tempString)
		
		headingrpt = "0"
        subheadingrpt = "0"
        PUTARAN = "0"
		i = "0"
		
		While Not rsMenuString.EOF
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
			          <div class="navlist">				 
				      <%  
				      i = "0"
				   end if	
			 ELSE 
			      if  cint(i) > 0 then
					  for k = cint(i) to 4
					  %>
						  <li>&nbsp;</li>
				   <% next %>
                      </ul>
                      </div>
               <% end if
                  i=0
				  if  MenuTypeCode="HEADINGREPORT" and headingrpt = "0" then
			   %>
					  <div class="entry-header"><h2 class="headings"><%=menuName%></h2></div>
					  <div class="entry-body">
					  <div class="navlist"> 
			  <%
				      headingrpt = menuCode
                      subheadingrpt = "0"
                      PUTARAN = "0"
                      i="0"
                   end if
			 
				   if MenuTypeCode="HEADINGREPORT" and headingrpt <> menuCode then
			 %>
                      </div>
                      <div class=clear></div>
                      <div class="entry-header"><h2 class="headings"><%=menuName%></h2></div>
					  <div class="entry-body">
					  <div class="navlist"> 
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
                         <div class='clear'></div>						 
						 <div class=entry-body>
		  <%
                      end if
		  %>
                      <div class=navlist>
                      <h3><%=menuName%></h3>
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
