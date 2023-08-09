<!--#include file = "conn/connection.asp"-->
<%
refundid=request("refundid")
sql="Delete from refundinfo where refundid ="&refundid
conn.execute(sql)
session("msg")="Data are  deleted from databasee"
response.redirect "listrefund.asp"
%>
