<!--#include file = "conn/connection.asp"-->
<%
sql="select * from refundinfo"
set rs=conn.execute(sql)
if ( Request("Submit")="Add" ) then
sql1=("INSERT into refundinfo (name,address,city,state,zip,refdate,telno,reason,product,refund,checkno,customerrep) Values ('"& request("name") &"','"& request("address") &"','"& request("city") &"','"& request("state") &"','"& request("zip") &"','"& request("date") &"','"& request("telno")&"','"& request("reason") &"','"& request("product") &"','"& request("refund") &"','"& request("checkno") &"','"& request("customerrep") &"')")
conn.execute(sql1)
session("errmsg")="Database are Added Successfully"
else
session("errmsg")="Database are not Connected"
end if
%>
<style type="text/css">
<!--
.style3 {font-size: 18px}
.style8 {font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 12px; }
-->
</style>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td valign="top" >
	<!--#include file="./includes/leftmenu.asp"-->
	</td>
    <td>
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="100%">&nbsp;</td>
        </tr>
        <tr>
          <td><p>&nbsp;</p>
          <p>&nbsp;</p>
          <p><span class="style3"><%=session("errmsg")%></span></p>
          <p>&nbsp;</p>
          <p>&nbsp;</p></td>
        </tr>
      </table>        </td>
  </tr>
  <tr>
   
  </tr>
</table>
<% session("errmsg")=""%>
