<!-- #Include file="adovbs.inc" -->
<%
  dim gobjConn, gobjRS, gstrLocation 
  set gobjConn = Server.CreateObject("ADODB.Connection")
  set gobjRS = Server.CreateObject("ADODB.Recordset")	
  
   With gobjConn
        .Provider = "Microsoft.jet.oledb.4.0; persist security info=false;"
        .ConnectionString = "Data Source=c:\inetpub\wwwroot\fcr\App_Data\database\carrent.mdb"
        .Open
   End With
  
  
  
  
  
  
  
  'gobjConn.Open "DSN=FCR"
  
%>	
