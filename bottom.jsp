<div id="sitebottom" style="padding-top:5px; padding-bottom:5px;">
    <div class="menu_bottom">
        <a href="#">About Us</a> | 
        <a href="#">Site Map</a> | 
        <a href="#">Privacy Policy</a> | 
        <a href="#">Contact Us</a>
		<% if LEN(trim(PROFIL("COPYRIGHT"))) > 0 then %>

		| <a href="#"><%=PROFIL("COPYRIGHT")%></a>
			
		<% ELSE %>
		
		| <a href="#">&copy;2022 FB-(Plantation System)</a>
		
		<% end if %>
   
		
    </div>
</div>