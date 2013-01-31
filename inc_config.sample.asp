<%
'AJLogin v4.1
'By: Adam Jarret
'RandomRavings.com

'===================================================================================================
'
' Make a copy of this file named inc_config.asp and configure the following values for your site.
'
'===================================================================================================

'ajlogin_dsn is the path to the database. (Server.Mappath makes the path relative to the location of the inc_common.asp file, but
'	I recommend moving the database to an offline folder for security reasons)
ajlogin_dsn = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.Mappath("ajlogin.mdb")

'ajlogin_sitename is what appears in the title of all AJLogin Pages (it can be anything you like)
ajlogin_sitename = "Your Site Name"

'ajlogin_cookiename is the name of the cookie generated by the "keep me logged in" box on the login page (it can be anything you like)
'	WARNING: Do not use non-alpha-numeric characters for this setting
ajlogin_cookiename = "YourSiteName"

'ajlogin_minpasslength is the minimum number of characters that can be used for a new member's password
ajlogin_minpasslength = 4

'ajlogin_shouldhashpasswords controls whether or not passwords are stored in the database as secure hashes 
'	It is recommended to set this option to true, unless you have a good reason for storing member passwords in plain text
'	WARNING: Changing this setting will force existing users to use the "forgot your password" link to reset their passwords
ajlogin_shouldhashpasswords = true

'ajlogin_passwordsalt should be a random string. if can be anything you like, but should be changed from the default
'	The password salt string increases the security of storing member passwords as hashes.
'	This option is ignored if ajlogin_shouldhashpasswords is set to false. 
ajlogin_passwordsalt = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"

'This is your site's URL (don't forget the trailing slash)
ajlogin_siteurl = "http://www.yoursite.com/"

'This is the full path to ajlogin on your server (don't forget the trailing slash)
ajlogin_scripturl = ajlogin_siteurl & "ajlogin/"

'ajlogin_headimg is the url of the picture you wish to use as your header (it will appear at the top of each page - the default is the RandomRavings Logo)
ajlogin_headimg = ajlogin_scripturl & "images/rr_head.gif"

'ajlogin_memberurl is the page that members are redirected to once they log in (it should be a protected page)
ajlogin_memberurl = ajlogin_scripturl & "members_area.asp"

'ajlogin_memberurl is the page that the admin is redirected to once they log in (this can be the same as ajlogin_memberurl if you like)
ajlogin_adminurl = ajlogin_scripturl & "admin.asp"

'ajlogin_adminlevel is the highest level that the admins have
ajlogin_adminlevel = 3

'whether or not to allow new members to register (hides links to registration page and prevents new accounts being created via reg.asp) 
ajlogin_enableregistration = true

'if ajlogin_requireapproval is set to true, new accounts must be approved by an admin before they are allowed to log in
ajlogin_requireapproval = true

'ajlogin_sendconfirmation sets whether or not to send a confirmation e-mail to their address before letting them log in
ajlogin_sendconfirmation = true

'ajlogin_sendwelcome sets whether or not to send a welcome e-mail to their address after registering
'	If ajlogin_sendconfirmation is true, then a confirmation e-mail is sent instead of a welcome e-mail
ajlogin_sendwelcome = false

'ajlogin_sendpassword sets whether or not to enable automated password resets via e-mail
'	If this is false the "Forgot Your Password" link will not be displayed
ajlogin_sendpassword = true

'ajlogin_sendnotification sets whether or not to send automated new memember registration notifications
ajlogin_sendnotification = true

'The address that new memember registration notifications are sent to (if enabled)
ajlogin_emailnotification = "admin@yoursite.com"

'The address that forgotton passwords and confirmation e-mails are sent from
ajlogin_emailfrom = "admin@yoursite.com"

'The reply to address for e-mails generated by AJLogin (ignored when ajlogin_usecdonts = true)
ajlogin_emailreplyto = "admin@yoursite.com"

'ajlogin_usecdonts sets whether or not to use CDONTS to send e-mail messages
'if false then CDOSYS is used. CDOSYS is the default because CDONTS is no longer supported on newer Micro$oft servers.
ajlogin_usecdonts = false

'if you are using cdosys, please set the variables below to appropriate values for your SMTP server
'if not, these values will be ignored
ajlogin_cdosys_smtpserver = "mail.yoursite.com"
ajlogin_cdosys_smtpserverport = 25 
ajlogin_cdosys_smtpauthenticate = 1 '0 = Do not authenticate, 1 = Basic (clear-text) authentication, 2 = Use NTLM authentication (Secure Password Authentication in Outlook Express)
ajlogin_cdosys_sendusername = "admin@yoursite.com" 'ignored if ajlogin_cdosys_smtpauthenticate = 0
ajlogin_cdosys_sendpassword = "password" 'ignored if ajlogin_cdosys_smtpauthenticate = 0

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'												Please do not edit settings below this line
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

AJLoginVersion="4.1"

ajlogin_pastdate = DateAdd("m", -1, Date)

ajlogin_kConfirmed = "345h34tkjh34iu245gk43lkg342kjlh524"
ajlogin_kApproved  = "45lk12h235lkh46lk24lk23lkh235klj2l"


'AJLogin v4.1
'By: Adam Jarret
'RandomRavings.com
%>