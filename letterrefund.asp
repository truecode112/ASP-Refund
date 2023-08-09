<!--#include file = "conn/connection.asp"-->
<%
refundid=request("refundid")
sql="select *from refundinfo where refundid="& refundid &" "
set rs=conn.execute(sql)
today=date()
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Refund page</title>
<style type="text/css">
<!--
.style12 {font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 12; }
.style13 {font-size: 12}
.style16 {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 12px;
}
.style18 {font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 18px; }
.style20 {font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 20px; }
-->
</style>
</head>

<body>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr >
    <td >
	<table width="750" border="0" cellspacing="0" cellpadding="0" align="center">
      <tr>
        <td width="3%">&nbsp;</td>
        <td width="49%"><span class="style20"><br />Klein's Kosher Ice Cream </span></td>
        <td width="48%" colspan="4">&nbsp;</td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td colspan="4">&nbsp;</td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td><span class="style16">3614 15th Avenue</span></td>
        <td colspan="4" rowspan="4">
          <div align="center">
            <input type="image" name="imageField" src="images/logo.jpg">
            </div></td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td><span class="style16">Brooklyn NY 11218 </span></td>
        </tr>
      <tr>
        <td>&nbsp;</td>
        <td><span class="style13"></span></td>
        </tr>
      <tr>
        <td>&nbsp;</td>
        <td><span class="style16">Tel: (718) 435-0124 </span></td>
        </tr>
      <tr>
        <td>&nbsp;</td>
        <td><span class="style16">Fax: (718) 436- 1036 </span></td>
        <td colspan="4"><div align="center"><%= formatdatetime(today,1) %></div></td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td><span class="style13"></span></td>
        <td colspan="4">&nbsp;</td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td><span class="style12">Email: info@koshericecream.com </span></td>
        <td colspan="4">&nbsp;</td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td><span class="style12">http://www.koshericecream.com</span></td>
        <td colspan="4">&nbsp;</td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td colspan="4">&nbsp;</td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td><span class="style17"><%=rs("name") %></span></td>
        <td colspan="4">&nbsp;</td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td><span class="style17"><%=rs("address") %></span></td>
        <td colspan="4">&nbsp;</td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td><span class="style17"><%=rs("city") %>&nbsp;&nbsp;&nbsp;<%=rs("state") %>&nbsp;&nbsp;&nbsp;<%=rs("zip") %></span></td>
        <td colspan="4">&nbsp;</td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td colspan="4">&nbsp;</td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td colspan="5"><span class="style18">Thank you for contacting us. Customer satisfaction is our first priority and that's why calls like yours disturb us greatly. Your comments have been forwarded to our manufacturing division.<br />
            <br />
A combination of technical and human error may be the result of your situation. Please accept our sincere apology. Enclosed you will find a refund check.<br />
<br />
Once again I thank you for your time. Your call has helped us improve!.<br />
<br />
Thank you,<br />
<br />
<%=rs("customerrep")%><br />
Customer Service </span></td>
        </tr>
      <tr>
        <td>&nbsp;</td>
        <td colspan="5">&nbsp;</td>
        </tr>
      <tr>
        <td>&nbsp;</td>
        <td colspan="5"><table width="100%" border="1" cellspacing="0" cellpadding="0">
          <tr>
            <td width="14%"><strong>Date</strong></td>
            <td width="16%"><strong>Name</strong></td>
            <td width="12%"><strong>Ref ID</strong></td>
            <td width="18%"><strong>Product</strong></td>
            <td width="15%"><strong>Telephone</strong></td>
            <td width="13%"><div align="right"><strong>CheckNo</strong></div></td>
            <td width="12%"><div align="right"><strong>Refund</strong></div></td>
          </tr>
          <tr>
            <td><%=rs("refdate") %></td>
            <td><%=rs("name") %></td>
            <td><%=rs("refundid") %></td>
            <td><%=rs("product") %></td>
            <td><%=rs("telno") %></td>
            <td><div align="right"><%=rs("checkno") %></div></td>
            <td align="right"><%=formatcurrency(rs("refund")) %></td>
          </tr>
        </table></td>
        </tr>
      <tr>
        <td>&nbsp;</td>
        <td colspan="5">&nbsp;</td>
        </tr>
      <tr>
        <td>&nbsp;</td>
        <td colspan="5"><div align="right"><a href="./listrefund.asp">&lt;&lt; Back </a>&nbsp;&nbsp;|&nbsp;&nbsp;<a href="javascript: void(0)" onclick="javascript:print();">Print Page </a></div></td>
      </tr>
    </table></td>
  </tr>
</table>
</body>
</html>
