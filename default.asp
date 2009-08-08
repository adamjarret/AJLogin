<%
'AJLogin v4.1
'By: Adam Jarret
'RandomRavings.com
%>
<!--#INCLUDE file="inc_common.asp"-->
<%
if CInt(Session("UserLevel")) = CInt(ajlogin_adminlevel) then
	Response.Redirect(ajlogin_adminurl)
else
	Response.Redirect(ajlogin_memberurl)
end if
%>
<%
'AJLogin v4.1
'By: Adam Jarret
'RandomRavings.com
%>