<!--| STATS COUNTER - V1.2 - Copyright 2003 - Designpost (UK) -->
<!--| Unauthorized copies are not allowed |-->

<body bgcolor="#cc9966">




<%
'display recordset for login request
'
'
'
dim adocon
dim rs
dim strsql
session.lcid=2057

set adocon = server.createobject("adodb.connection")
adoCon.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("stats.mdb")

set rs = server.createobject("adodb.recordset")

strSQL = "SELECT * FROM stats"
rs.CursorType = 2
rs.LockType   = 3
rs.open strsql, adocon

'add new record

rs.addnew
rs("ipaddress") = request.servervariables("remote_host")
rs("country")   = "Server does not support this option"
rs("page")      = request.servervariables("script_name")
rs("date")      = date
rs("time")      = time
'finish update
rs.update

rs.close
set rs = nothing
set adocon = nothing
%>
