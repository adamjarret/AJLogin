<%
'AJLogin v4.1
'By: Adam Jarret
'RandomRavings.com
%>

<!--#INCLUDE file="inc_common.asp"-->
<% RestrictPageAccesToLevel(1) %>

<% PageTitle = "Member Area" %>
<!--#INCLUDE file="inc_header.asp"-->

		<div class="page-stripe">
			<p>You are logged in as <strong><%= Session("UserName") %></strong></p>
			<p>
				<a href="<%=ajlogin_scripturl%>edit_profile.asp">Manage Account Details</a> | 
				<a href="<%=ajlogin_scripturl%>login.asp?m=LogOut">Log Out</a>
			</p>
			
			<% if CInt(Session("UserLevel")) >= ajlogin_adminlevel then %>
				<p><a href="<%=ajlogin_adminurl%>">Admin</a></p>
			<% end if %>
 		</div>

<!--#INCLUDE file="inc_footer.asp"-->

<%
'AJLogin v4.1
'By: Adam Jarret
'RandomRavings.com
%>