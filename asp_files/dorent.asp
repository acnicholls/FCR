<% @Language=vbscript %>


<% Option Explicit %>
<!-- #Include file="../App_Data/inc_files/FCRmdb.inc" -->

<%


dim intCarID, strClientID, strSQL, fileobj, strFileName, textfileObj, strLocation
set fileobj = server.CreateObject( "Scripting.FileSystemObject" )
strFileName = "C:\inetpub\wwwroot\fcr\App_Data\logs\" & "rentlog.txt"
strClientID = request("strClientID")
intCarID = request("strCarID")

strSQL = "SELECT * FROM tblCar WHERE fldCarID = '" & intCarID & "'"
call gobjrs.open( strSQL, gobjConn, 2, adlockoptimistic)

gobjrs.fields(10) = date()
strLocation = gobjrs.fields(12)
gobjrs.fields(12) = null
gobjrs.fields(13) = cstr(strClientID)
gobjrs.update

if fileobj.FileExists( strFileName ) <> true then
	call fileobj.CreateTextFile( strFileName )
end if

	set textfileObj = fileobj.OpenTextFile( strFileName,8,true )
	
	call textfileobj.writeline("Car  " & intCarID)
	call textfileobj.writeline("Client  " & strClientID)
	call textfileobj.writeline("KM on Rent   " & gobjrs.fields(11))
	call textfileobj.writeline("Date Rented   " & date())
	call textfileobj.writeline("Rented from " & strLocation)
	call textfileobj.writeline("Operator  " & session("UserName"))
	call textfileobj.writeline("***********************************")
	
	call textfileobj.close
		

gobjrs.close
Response.Redirect("refreshhome.asp")




%>

