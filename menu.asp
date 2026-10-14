<!-- Modify by NAF, to handle the New Concept of IPLAS, 1 May 2013 -->
<!--[if gte IE 9]>
  <style type="text/css">
    .gradient {
       filter: none;
    }
  </style>
<![endif]-->
<%

Dim tempString, rsMenuString
Dim menuCode, menuParentCode, menuName, menuLink, menuHelp, moduleName


			
tempString = " SELECT sequence," & _
                     " sequence_line," & _
					 " fccode," & _
					 " fcname," & _
					 " menuparentcode," & _
					 " formlink," & _
					 " helpdescription," & _
					 " name" & _
				" FROM (SELECT 0 sequence," & _
				             " 0 sequence_line," & _
							 " a.fccode," & _
							 " a.fcname," & _
							 " a.menuparentcode," & _
							 " a.formlink," & _
							 " a.helpdescription," & _
							 " '' name" & _
						" FROM menu_tab a " & _
					   " WHERE a.isactive = 'TRUE'" & _
							 " AND USER_ROLES_API.GetAccessUsers ('" & Session("CodeName") & "', a.fccode) = 'TRUE'" & _
							 " AND a.menutypecode = 'FORM' " & _
							 " AND fccode = '0'" & _
					  " UNION" & _
					  " SELECT b.sequence," & _
					         " a.sequence sequence_line," & _
							 " a.fccode," & _
							 " a.fcname," & _
							 " a.menuparentcode," & _
							 " a.formlink," & _
							 " a.helpdescription," & _
							 " b.name" & _
						" FROM menu_tab a, module b" & _
					   " WHERE a.isactive = 'TRUE' AND A.MODULE = B.CODE" & _
							 " AND USER_ROLES_API.GetAccessUsers ('" & Session("CodeName") & "', a.fccode) = 'TRUE'" & _
							 " AND a.menutypecode = 'FORM') " & _
			" ORDER BY SEQUENCE, SEQUENCE_LINE	" 

Set rsMenuString = connStr.Execute(tempString)




%>		
	



<div class="dtree">
<div class="head">
<a href="javascript: d.openAll();">open all</a> | <a href="javascript: d.closeAll();">close all</a>
</div>
<div class="body_menu">	
	<script type="text/javascript">
		<!--
		d = new dTree('d');
		-->
	</script>	
	
	<% While Not rsMenuString.EOF %>
		     <%
			     menuCode = rsMenuString("fccode") 
				 menuParentCode =  rsMenuString("menuparentcode") 
				 menuName =  rsMenuString("fcname") 
				 menuLink = rsMenuString("formlink") 
				 menuHelp = rsMenuString("helpdescription")
				 moduleName = rsMenuString ("name")
			   
				 
			%>	

		<% if menuLink <> "" then %>
			<script type="text/javascript">
			<!--
				d.add('<%=menuCode%>', '<%=menuParentCode%>', '<%=menuName%>', 'show.asp?txtTable=<%=menuLink%>&txtModule=<%=moduleName%>', '<%=menuHelp%>')
			//-->	 
			</script>
		<% else %>
			<script type="text/javascript">
			<!--
				d.add('<%=menuCode%>', '<%=menuParentCode%>', '<%=menuName%>', '', '<%=menuHelp%>')
			//-->	 
			</script>		
		<% end if %>
	
		    <% rsMenuString.MoveNext %>
	<% Wend %>

	<script type="text/javascript">
	<!--
	document.write(d);
	//-->	 
	</script>
		
</div>	
</div>

