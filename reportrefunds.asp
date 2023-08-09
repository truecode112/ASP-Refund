<!--#include file = "conn/connection.asp"-->
<%
Const adUseClient = 3
Const adCmdText = 1
Const adOpenStatic = 3
Const adLockReadOnly = 1

Dim iPageSize       
Dim iPageCount      
Dim iPageCurrent   
Dim strOrderBy     
Dim strSQL          
Dim objPagingConn   
Dim objPagingRS    
Dim iRecordsShown   
Dim I               
Dim vSQL1, rsSQL1, showlist
	
	sql = "Select * from refundinfo order by refundid"

			
	iPageSize = 10'recordset level 
	
	If Request.QueryString("page") = "" Then
		iPageCurrent = 1
	Else
		iPageCurrent = CInt(Request.QueryString("page"))
	End If
	
	strConn = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source="&datapath& ";"
	
	Set objPagingConn = Server.CreateObject("ADODB.Connection")
	objPagingConn.Open strConn
	
	Set objPagingRS = Server.CreateObject("ADODB.Recordset")
	objPagingRS.PageSize = iPageSize
	
	objPagingRS.CacheSize = iPageSize
	
	objPagingRS.Open sql, objPagingConn, adOpenStatic, adLockReadOnly, adCmdText
	
	iPageCount = objPagingRS.PageCount
	
	
	If iPageCurrent > iPageCount Then iPageCurrent = iPageCount 
	If iPageCurrent < 1 Then iPageCurrent = 1 
		
	If iPageCount <>0 then
		objPagingRS.AbsolutePage = iPageCurrent
		
	End if
	

%>
<style type="text/css">
<!--
.mouse_over{background:#CCCCCC}
.mouse_out{background:#ffffff}
.normal{background:#ffffff}

.style14 {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 16px;
	font-weight: bold;
}
.style19 {font-size: 14px; font-family: Verdana, Arial, Helvetica, sans-serif; font-weight: bold; }
.style21 {font-size: 14px; font-family: Verdana, Arial, Helvetica, sans-serif; font-weight: bold; color: #FFFFFF; }
.style22 {color: #FFFFFF}
-->
</style>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="72%" valign="top"><table  width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr >
        <td colspan="7"><span class="style14">Report of Refunds </span></td>
        </tr>
      <tr >
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
	  <tr>
	  <td colspan="11">
	  <table width="100%">
      <tr bgcolor="#000000">
        <td width="6%" ><span class="style21">ID</span></td>
        <td width="18%" ><span class="style21">Name</span></td>
        <td width="11%" ><span class="style21">Date</span></td>
        <td width="24%" ><span class="style21">Address</span></td>
        <td width="12%" ><div align="center" class="style22"><span class="style19">Refund</span></div></td>
        <td width="21%" ><span class="style21">Product</span></td>
        <td width="8%" ><span class="style21">Checkno</span></td>
      </tr>
	  </table>
	  </td>
	  </tr>
	  <!-- tr  class="normal"onmouseover="this.Class='mouse_over';" onmouseout="this.class='mouse_out';" -->
	   <tr  >
	  <td colspan="11" >
	  <%
	  count=0
	  if not objPagingRS.eof  then
		iRecordsShown = 0
	  Do While iRecordsShown < iPageSize And Not objPagingRS.EOF
   
	   if bgcol = "#ffffff" then
	   		'bgcol = "#c4c4c4"
			bgcol = "#ffffff"
	   else
	   		bgcol = "#ffffff"
	   end if
	   
	  %>
	   
	  <table width="100%"  cellpadding="1" cellspacing="2"  onmouseover="this.bgColor='#cccccc'" onmouseout="this.bgColor='#ffffff'" >
      <tr>
        <td width="6%"><%=objPagingRS("refundid") %> </td>
        <td width="18%"><%=objPagingRS("name") %></td>
        <td width="11%"><%=objPagingRS("refdate") %></td>
        <td width="24%"><%=objPagingRS("address") %></td>
        <td width="12%"><div align="right"><%=formatcurrency(objPagingRS("refund")) %> &nbsp;&nbsp;</div></td>
        <td width="21%"><%=objPagingRS("product") %></td>
        <td width="8%" align="right"><%=objPagingRS("checkno") %></td>
      </tr>
	   <tr>
        <td>&nbsp;</td>
        <td colspan="2"><%=objPagingRS("telno") %><br />
        Rep : <%=objPagingRS("customerrep") %></td>
        <td><%=objPagingRS("city") %> - <%=objPagingRS("state") %> - <%=objPagingRS("zip") %> </td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
	  </table>
	  
	 </td>
	  </tr>
	 
      
    </table></td>
  </tr>
  
   <%
	  iRecordsShown = iRecordsShown + 1
	  objPagingRS.movenext	 
	  loop
	  
	  else
	  
	  end if
	  %>
	  <tr >
       
        <td><div align="right"><a href="./listrefund.asp">&lt;&lt; Back </a>&nbsp;&nbsp;|&nbsp;&nbsp;<a href="javascript: void(0)" onclick="javascript:print();">Print Page </a></div></td>
      </tr>
  <tr>
    <td align="left">
      Showing Page <strong><%= iPageCurrent %></strong> of <strong><%= iPageCount %></strong> <br>
      Move to: <%
	If iPageCurrent > 1 Then
	%>
      <a href="reportrefunds.asp?page=<%= iPageCurrent - 1 %>">[Previous]</a>
      <%
	End If

	For I = 1 To iPageCount
		If I = iPageCurrent Then
			%>
      <%= I %>
      <%
		Else
			%>
      <a href="reportrefunds.asp?page=<%= I %>"><%= I %></a>
      <%
		End If
	Next 


	If iPageCurrent < iPageCount Then
		%>
      <a href="reportrefunds.asp?page=<%= iPageCurrent + 1 %>">[Next]</a>
      <%
	End If

	%>	</td>
    </tr>
</table>
<script language="javascript1.2">
<!--
	function Printpage(){
		
		alert("This will print your page!!!");
		window.document.print();
	
	}
-->
</script>