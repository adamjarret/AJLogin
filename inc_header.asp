<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta http-equiv="Pragma" content="no-cache" />
<title><%= ajlogin_sitename %> | AJLogin - <%= PageTitle %></title>
<link rel="stylesheet" href="<%= ajlogin_scripturl %>ajlogin.css" type="text/css"></link>
<script type="text/javascript" src="<%= ajlogin_scripturl %>js/jquery-1.3.2.min.js"></script>
<script type="text/javascript" src="<%= ajlogin_scripturl %>js/jquery.cookie.js"></script>
<script type="text/javascript" src="<%= ajlogin_scripturl %>js/jquery.validate.js"></script>
<script type="text/javascript" src="<%= ajlogin_scripturl %>js/jquery.tablesorter.min.js"></script>
<script type="text/javascript" src="<%= ajlogin_scripturl %>js/ajlogin.js"></script>
<script type="text/javascript">
	/*
	AJLogin uses the jQuery validation plugin to validate user submitted forms.
	
	For more information on how to add your own custom completion rules, see
	http://docs.jquery.com/Plugins/Validation/ 
	*/
	$(function() {
		$(".utility-form").validate();
	});
</script>
</head>

<body>
	<div class="page-header">
		<a href="<%= ajlogin_siteurl %>"><img src="<%= ajlogin_headimg %>" alt="Home" /></a>
   		<h1 class="page-title">
   			<strong>&raquo;</strong> <%= PageTitle %>
   		</h1>
   	</div>
   	<div>
