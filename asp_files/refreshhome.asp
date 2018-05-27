<% @LANGUAGE="vbscript" %>

<% Option Explicit %>
<!-- #Include file="../App_Data/inc_files/FCRmdb.inc" -->
<% Dim strSQL

	select case request("DesiredLocation")
		case "All", "", NULL
			set gobjRS = server.CreateObject("ADODB.Recordset")
			strSQL = "SELECT * FROM tblCar"
			call gobjRS.Open(strSQL, gobjConn, adOpenStatic, adLockReadOnly, adcmdText)
		case else
			set gobjRS = server.CreateObject("ADODB.Recordset")
			strSQL = "Select * FROM tblCar WHERE Location ='" & Request("DesiredLocation") & "'"
			call gobjRS.Open(strSQL, gobjConn, adOpenStatic, adLockReadOnly, adcmdText)
	end select
	
	Response.Redirect("home.asp")
%>