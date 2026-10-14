<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
 <%
   if  len(TRIM(Session("fcBa"))) > 0 then
   
		dim QCOPY,COPY,REPORTDASHBOARD_IMAGE
		
		QCOPY="select REPORTDASHBOARD_IMAGE from APPLICATIONIMAGE  where " & _
		 " fcBA='" & session("fcBA") & "'"
		 
		 Set COPY = Server.CreateObject("ADODB.Recordset")
		   COPY.CursorType = 3
		   COPY.CursorLocation = 3
		   COPY.Open QCOPY, connstr 
		if Not COPY.eof then   
		   
		   REPORTDASHBOARD_IMAGE = COPY("REPORTDASHBOARD_IMAGE")
		   
		end if 
			   COPY.Close
			   set COPY = Nothing
   
   end if %>     
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="author" content="" />
<meta http-equiv="content-type" content="text/html;charset=iso-8859-1" />
<link rel="stylesheet" href="/css/stylereport.css" type="text/css" />
<title>ePlaS - Electronic Plantation Solutions</title>
<script language="JavaScript" src='/js/jquery-1.8.3.js'></script>
<script type="text/javascript"> function scroll(element){   
var ele = document.getElementById(element);   
window.scrollTo(ele.offsetLeft,ele.offsetTop); } </script> 
<script>
var fixed = false;
$(document).scroll(function() {
    if( $(this).scrollTop() > 26 ) {
        if( !fixed ) {
            fixed = true;
            $('#logo-scroll').css({position:'fixed', display:'block', top:'100px'});
        }
    } else {
        if( fixed ) {
            fixed = false;
            $('#logo-scroll').css({display:'block', position:'static'});
        }
    }
});
</script>
<script type="text/javascript">
		function cancelFullScreen(el) {
            var requestMethod = el.cancelFullScreen||el.webkitCancelFullScreen||el.mozCancelFullScreen||el.exitFullscreen;
            if (requestMethod) { // cancel full screen.
                requestMethod.call(el);
            } else if (typeof window.ActiveXObject !== "undefined") { // Older IE.
                var wscript = new ActiveXObject("WScript.Shell");
                if (wscript !== null) {
                    wscript.SendKeys("{F11}");
                }
            }
        }

        function requestFullScreen(el) {
            // Supports most browsers and their versions.
            var requestMethod = el.requestFullScreen || el.webkitRequestFullScreen || el.mozRequestFullScreen || el.msRequestFullScreen;

            if (requestMethod) { // Native full screen.
                requestMethod.call(el);
            } else if (typeof window.ActiveXObject !== "undefined") { // Older IE.
                var wscript = new ActiveXObject("WScript.Shell");
                if (wscript !== null) {
                    wscript.SendKeys("{F11}");
                }
            }
            return false
        }

        function toggleFull() {
            var elem = document.body; // Make the body go full screen.
            var isInFullScreen = (document.fullScreenElement && document.fullScreenElement !== null) ||  (document.mozFullScreen || document.webkitIsFullScreen);

            if (isInFullScreen) {
                cancelFullScreen(document);
            } else {
                requestFullScreen(elem);
				 alert("Tekan 'F11' pada Keyboard Anda, untuk selalu Full Screen.");
            }
            return false;
        }
</script>
<% if Request("oldui") <> "1" then %>
<!-- Tampilan baru (css/fbp-report.css). ?oldui=1 = tampilan lama -->
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap" rel="stylesheet">
<link rel="stylesheet" href="/css/login/font-awesome.min.css" type="text/css" />
<link rel="stylesheet" href="/css/fbp-report.css?v=20261010a" type="text/css" />
<script type="text/javascript" src="/js/fbp-report.js?v=20261007"></script>
<% end if %>
</head>
<body<% if Request("oldui") <> "1" then %> class="rpt"<% end if %>>
<div id="siteInfo" style=" position:fixed; z-index:1; width:100%;">
    <div class="head" style="padding:5px;">	
    	<h1 style="margin-top:15px;">MENU <font color="#FFCC00" style="font-weight:normal;">LAPORAN</font></h1>
        <div style="color:#FFFFFF; font-size:10px;"><b><%=session("FULLNAME")%></b> <font size="2" color="#FFCC00">|</font> <b>BA : <%=Session("BA")%> ( <%=session("fcBA")%> )</b> <font size="2"  color="#FFCC00">|</font> Current Period :  <%=FormatDateTime(session("v_Costbook"),1)%>
		</div>
        
        <div id="globalNav">
		<%if Session("USERNAME")="" then%>
		<a class="mn_top" href="../logout.asp"><span>log-in</span></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<B>Please login first...</B><a class="mn_top"  href="../login.asp"><i class="home_icon">&nbsp;&nbsp;&nbsp;&nbsp;</i><span>Home</span></a></div>
		<%else%>
		<a class="mn_top" href="../logout.asp"><i class="off_icon">&nbsp;&nbsp;&nbsp;&nbsp;</i><span>log-out</span></a><a class="mn_top"  href="../login.Asp"><i class="home_icon">&nbsp;&nbsp;&nbsp;&nbsp;</i><span>Home</span></a><!--<a id="button" class="mn_top_full" href="javascript:void(0);" onclick="toggleFull();"><i class="full">&nbsp;&nbsp;&nbsp;&nbsp;</i><span>Full Screen</span></a>--></div>
		<%end if%>
       
    </div>
    
   
   </div>
<br/><br/><br/><br/><br/>
<div id="container2">
	<div class="line"></div>

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
  <div class="colR" style="height:100%;">
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
        <br/>
     <!-- <div class="nav">
        <a href="../login.asp"><img src="../IMAGES/back_03.gif" width="70"/></a>
      </div>-->
       </div>
   </div>
	<div class="colL">
        
        
	<%	
	Dim var_no : var_no = 0
	
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
			 	   if i = 10 then
				      
						  PUTARAN = cint(PUTARAN) + 1 %>
						  </ul>
						  </div>
						  <div class="navlist2 ">	
						  <div class="guide"> &nbsp;</div>
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
                      <h4><%=menuName%></h4></div>
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
                      <div id="<%=menuCode%>" class="entry-header"><h4><%=menuName%></h4></div>
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
