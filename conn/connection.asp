<%

currentpath=server.MapPath(".")
'''Response.Write currentpath
'''E:\Work\ASP\refund\refund\conn

datapath="E:\Work\ASP\refund\refund\database\refund.mdb"
set conn=CreateObject("Adodb.Connection")

Conn.open="PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source="&datapath& ";"

set rs=CreateObject("ADODB.Recordset")

%>
