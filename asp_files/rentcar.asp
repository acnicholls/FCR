<%@ Language=VBScript %>
<!-- #Include file="../App_Data/inc_files/FCRmdb.inc" -->
<html>
<link rel="stylesheet" type="text/css" href="../styles/carrent.css" />
<head>
<div align='center'><img src='../images/logo.jpg'></div>
<%
		dim strSQL, strCarDesc
		strSQL = "SELECT * FROM tblCar where fldCarID='" & request("strCarID") & "'"
		call gobjRS.open( strSQL, gobjConn, adOpenStatic, adLockReadOnly, adCmdText )
		strCarDesc = gobjrs.Fields(5) & "  " & gobjrs.Fields(4) & "  " & gobjrs.Fields(1) & " " & gobjrs.Fields(2)
			 
	sub dorent()
		Response.Write("<title>Rental Form</title></head>")
		Response.Write("<body><div width='50%' class='rentreturn' align='center'><br><font class='header'>Car Rental</font><br><br>")
	    Response.Write("<font>Car Number: </font><font class='info'>" & gobjrs.Fields(0) & "</font><br><br>")
		Response.Write("<font>Car to be rented: </font><font class='info'>" & strCarDesc)
		Response.Write("</font><br><br>")
		Response.Write("<font>Current KM's: </font><font class='info'>" & gobjrs.Fields(11) & "</font><br><br>")
		Response.Write("<form action='dorent.asp?" & Request.QueryString & "' method='POST'>Client ID: ")
		Response.Write("<input name='strClientID' type='text' size=10><br>")
		Response.Write("<br><br>")
		Response.Write("<input type='submit' value='Send to Database'></form>")	
		Response.Write("<br><br></div></body></html>")
		gobjrs.Close
		set gobjrs = nothing
	end sub
	sub doreturn()
		Response.Write("<title>Return Form</title></head>")
		Response.Write("<body><div width='50%' class='rentreturn' align='center'><br><font class='header'>Car Return</font><br><br>")
	    Response.Write("<font>Car Number: </font><font class='info'>" & gobjrs.Fields(0) & "</font><br><br>")
		Response.Write("<font>Car returning: </font><font class='info'>" & strCarDesc)
		Response.Write("<br><br><font>Date Rented: </font><font class='info'>" & gobjrs.fields(10) & "</font><br><br>")
		Response.Write("<form action='doreturn.asp?" & Request.QueryString & "' method='POST'>")
		Response.Write("<font>KM's at rental: </font><font class='info'>" & gobjrs.Fields(11) & "    </font>")
		Response.Write("<font>Current KM's: </font>")
		Response.Write("<input name='intKM' type='text' size=15><br><br>")
		Response.Write("<font>Recieving location: </font>")
			Response.Write("<select name='cboRetLocation' size=1>")
					Response.Write ("<option value='' SELECTED>")
					Response.Write ("<option value='Victoria'>Victoria")
					Response.Write ("<option value='Vancouver'>Vancouver")
					Response.Write ("<option value='Nanaimo'>Nanaimo")
					Response.Write ("<option value='Hope'>Hope")
			Response.Write("</select>")
		Response.Write("<br><br>")
		Response.Write("<input type='submit' value='Send to Database'>")	
		Response.Write("</form><br><br></div></body></html>")
		gobjrs.Close
		set gobjrs = nothing
	end sub


	
	

'test the client ID number
if gobjRS.fields(13) > "" then
'full means return
	call doreturn()
else
'empty mean rent
	call dorent()
end if
%>





