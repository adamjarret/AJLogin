
AJLogin v4.1 Readme
===================

[AJLogin](http://adamjarret.com/ajlogin) is a user management system written in classic ASP that can be very easily integrated into existing web sites. The script uses a **Microsoft Access** database to store member account details and session variables to determine if a site visitor is logged-in. Your web host must support **ASP** and **ADODB** in order to use AJLogin. In order to use any of the features involving e-mail, your web host must also support **CDONTS** or **CDOSYS**.	

The script allows visitors to your site to register for membership in order to access "members only" pages that you designate.

### Features ###

 * A browser-based control panel for site admins
 * Optional e-mail address verification for new members
 * "Keep me logged-in" functionality
 * Fully automated password reminder e-mails
 * A built in messaging system that allows for quick and painless group e-mails to the members of your site.

Please keep in mind that **protected pages must use the _.asp_ file extension** (e.g. members.asp as opposed to members.htm or members.html).

Don't be afraid to change the extension of existing _.htm_ and _.html_ files to _.asp_ for the purposes of protecting them (obviously take care to update all links that point to the page). Instructions for protecting pages can be found below.

### Configuring/Installing AJLogin ###
 1. Create a copy of the _inc_config.sample.asp_ file called _inc_config.asp_ and set appropriate preferences and site properties as instructed by the descriptive notes within the file.
 2. Visit the _install.asp_ to create your admin account and then **IMMEDIATELY DELETE THE _INSTALL.ASP_ FILE**.
 3. Upload all files to your web server. It is usually a good idea to keep all the AJLogin files (except the database) together in one folder, separate from other existing site files.
 4. Move the database to an offline folder for security purposes. Make sure that both the database and the parent folder of the database have **permissions set to allow read and write**.

### Using AJLogin ###

To protect an ASP page (that resides in the same folder as the AJLogin files) so that only members of your site can access it, paste the following code at the top of the page:

	<!--#INCLUDE file="inc_common.asp"-->
	<% RestrictPageAccesToLevel(1) %>

The value passed to the _RestrictPageAccesToLevel_ function (in this case: 1) dictates what access level a member must have in order to view the page. All new members default to level 1 and then can be promoted to higher security clearances using the admin control panel. You may use as many different access levels as you need, but the standard configuration is to use level 1 as a regular member, level 2 as a site moderator and level 3 as a site admin.

To protect an ASP page (that resides in the same folder as the AJLogin files) so that only site administrators can access it, include the following code at the top of the page:

	<!--#INCLUDE file="inc_common.asp"-->
	<% RestrictPageAccesToLevel(ajlogin_adminlevel) %>

To protect pages outside of the directory containing the AJLogin files, change the code that reads _#INCLUDE file_ to _#INCLUDE virtual_. You also must preface the included file name with the full virtual path to the file. An example that assumes the AJLogin files are in a directory called "ajlogin" in the site's root directory would look like this:

	<!--#INCLUDE virtual="/ajlogin/inc_common.asp"-->

Any questions/bugs/compliments can be sent to adam (at) adamjarret.com.