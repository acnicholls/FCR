<% @LANGUAGE = "vbscript" %>
<% Option Explicit %>
<!-- #Include file="../App_Data/inc_files/FCRmdb.inc" -->
<% 
	dim strUserName, strPassword, strSQL
	set gobjRS = server.CreateObject("ADODB.Recordset")
	'send the unser name to variable
	strUserName = Request.Form("txtUserName")
	'send the password
	strPassword = Request.Form("txtPassword")
	'call the sql string something 
	strSQL = "SELECT * from tblPass WHERE fldID = '" & strUserName & "'"
	'now arrange the variables to collect data
	call gobjRS.Open( strSQL, gobjConn, adOpenKeyset, adLockReadOnly, adcmdText )
	'if the db pointer moved, it's gotta test the password
	if gobjRS.RecordCount > 0 then
		'test the password
		if cstr(strPassword) = cstr(gobjRS.Fields("fldPass")) then 
			Session("Login") = True
			Session("UserName") = strUserName
			Session("DesiredLocation") = "All"
			Response.Redirect("home.asp")		
		end if
	end if
	Session("loginFailure") = True
	Response.Redirect("login.asp")
	gobjRS.Close
	gobjConn.close
	set gobjConn = nothing
%>
