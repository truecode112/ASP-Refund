<!--#include file = "conn/connection.asp"-->
<%
sql="select * from refundinfo "
set rs=conn.execute(sql)
refundid=Request("refundid")
if (request("Submit")="Update") then
sql1="UPDATE refundinfo SET " &_
			"name='"& request("name") &"',"&_
			"address='"& request("address") &"',"&_
			"city='"& request("city") &"',"&_
			"state='"& request("state") &"',"&_
			"zip='"& request("zip") &"',"&_
			"refdate='"& request("date") &"', "&_
			"telno='"& request("telno") &"', "&_
			"reason='"& request("reason") &"', "&_
			"product='"& request("product") &"', "&_
			"refund='"& request("refund") &"', "&_
			"checkno='"& request("checkno") &"', "&_
			"customerrep='"& request("customerrep") &"' "&_
			"WHERE refundid="& refundid
conn.execute(sql1)	
session("msg")="Database are updated successfully"
response.Redirect("listrefund.asp")		
else
session("msg")="Database are not updated successfully"
response.redirect("listrefund.asp")
end if
%>

