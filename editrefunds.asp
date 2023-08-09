<!--#include file = "conn/connection.asp"-->
<%
refundid=request("refundid")
sql="select *from refundinfo where refundid="& refundid &" "
set rs=conn.execute(sql)
%>
<script language="JavaScript" type="text/JavaScript">
function Validator(refund)
{
	if(document.refund.name.value == "")
	{
		alert('!!! Enter your Your name. !!!');
		document.refund.name.focus();
		return false;
	}
		if(document.refund.address.value == "")
	{
		alert('!!! Enter your Address. !!!');
		document.refund.address.focus();
		return false;
	}
		if(document.refund.city.value == "")
	{
		alert('!!! Enter your City name. !!!');
		document.refund.city.focus();
		return false;
	}
		if(document.refund.state.value == "")
	{
		alert('!!! Enter your State name. !!!');
		document.refund.state.focus();
		return false;
	}
		if(document.refund.zip.value == "")
	{
		alert('!!! Enter your Zip Code. !!!');
		document.refund.zip.focus();
		return false;
	}
		if(document.refund.telno.value == "")
	{
		alert('!!! Enter your Telephone No. !!!');
		document.refund.telno.focus();
		return false;
	}
		if(document.refund.reason.value == "")
	{
		alert('!!! Enter your Reason of Return. !!!');
		document.refund.reason.focus();
		return false;
	}
		if(document.refund.product.value == "")
	{
		alert('!!! Enter Product name. !!!');
		document.refund.product.focus();
		return false;
	}
		if(document.refund.refund.value == "")
	{
		alert('!!! Enter your Refund . !!!');
		document.refund.refund.focus();
		return false;
	}else{
		if(isNaN(document.refund.refund.value)){
			alert('!!! Enter Numbers only for Refund . !!!');
			document.refund.refund.focus();
			return false;
		}
	
	}
		
		if(document.refund.customerrep.value == "")
	{
		alert('!!! Enter your Customer Rep. !!!');
		document.refund.customerrep.focus();
		return false;
	}
	//date
	var checkstr=document.refund.date.value;
	if (checkstr=="" ){
		alert("Enter a correct date");
		document.refund.date.focus();
		return false;
	}
	else{
		if (checkstr.match(/\d{1,2}\/\d{1,2}\/\d{4}/)){
			//if ( checkstr.match(/\d{1,2}\/\d{1,2}\/\d{2,4}/) ){
			//alert("The date  is valid!!");
			//return false;
		}
		else{
			alert("The date  is not valid!!");
			return false;
		}
		
	}
	  return true;
	
}
//-->
</script>
<style type="text/css">
<!--
.style3 {font-size: 18px}
.style8 {font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 12px; }
.style9 {	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 10px;
}
-->
</style>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<form name="refund" method="post" onsubmit="return (Validator(this));"  action="updateprocess.asp?refundid=<%=rs("refundid")%>">
  <tr>
     <td valign="top"  width="19%">
	<!--#include file="./includes/leftmenu.asp"--></td>
    <td >
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
       <tr>
        <td colspan="3"><span class="style3">Update Refund  </span></td>
        </tr>
      <tr>
        <td colspan="3">&nbsp;</td>
        </tr>
      <tr>
        <td width="25%" ><span class="style8">Refund Id </span></td>
        <td width="4%" ><strong>:</strong></td>
        <td width="71%" ><input type="text" name="refundid" value="<%=rs("refundid") %>" readonly=""/></td>
      </tr>
      <tr>
        <td><span class="style8">Name</span></td>
        <td><strong>:</strong></td>
        <td><input type="text" name="name" value="<%=rs("name") %>" /></td>
      </tr>
      <tr>
        <td><span class="style8">Address</span></td>
        <td><strong>:</strong></td>
        <td><input type="text" name="address" value="<%=rs("address") %>" /></td>
      </tr>
      <tr>
        <td><span class="style8">City </span></td>
        <td><strong>:</strong></td>
        <td><input type="text" name="city" value="<%=rs("city") %>" /></td>
      </tr>
      <tr>
        <td><span class="style8">State</span></td>
        <td><strong>:</strong></td>
        <td><input type="text" name="state" value="<%=rs("state") %>" /></td>
      </tr>
      <tr>
        <td><span class="style8">Zip</span></td>
        <td><strong>:</strong></td>
        <td><input type="text" name="zip"  value="<%=rs("zip") %>"/></td>
      </tr>
      <tr>
        <td><span class="style8">Date </span></td>
        <td><strong>:</strong></td>
        <td><input type="text" name="date" value="<%=rs("refdate") %>" />
          ( <span class="style9">mm/dd/yyyy) format </span></td>
      </tr>
      <tr>
        <td><span class="style8">Telephone No. </span></td>
        <td><strong>:</strong></td>
        <td><input type="text" name="telno" value="<%=rs("telno") %>" /></td>
      </tr>
      <tr>
        <td><span class="style8">Reason Of Return </span></td>
        <td><strong>:</strong></td>
        <td><textarea name="reason"><%=rs("reason") %></textarea></td>
      </tr>
      <tr>
        <td><span class="style8">Product</span></td>
        <td><strong>:</strong></td>
        <td><input type="text" name="product" value="<%=rs("product") %>" /></td>
      </tr>
      <tr>
        <td><span class="style8">Refund</span></td>
        <td><strong>:</strong></td>
        <td><input type="text" name="refund" value="<%=rs("refund") %>" /></td>
      </tr>
      <tr>
        <td><span class="style8">Check No. </span></td>
        <td><strong>:</strong></td>
        <td><input type="text" name="checkno" value="<%=rs("checkno") %>" /></td>
      </tr>
      <tr>
        <td><span class="style8">Customer Rep </span></td>
        <td><strong>:</strong></td>
        <td><input type="text" name="customerrep" value="<%=rs("customerrep") %>" /></td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td><input type="submit" name="Submit" value="Update" /></td>
      </tr>
    </table>	</td>
  </tr>
  </form>
</table>
