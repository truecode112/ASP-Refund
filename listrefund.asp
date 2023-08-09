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

			
	iPageSize = 20'recordset level 
	
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
	
	records = objPagingRS.recordcount
	
	iPageCount = objPagingRS.PageCount
	
	
	If iPageCurrent > iPageCount Then iPageCurrent = iPageCount 
	If iPageCurrent < 1 Then iPageCurrent = 1 
		
	If iPageCount <>0 then
		objPagingRS.AbsolutePage = iPageCurrent
		
	End if

%>
<style type="text/css">
</style>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="19%" valign="top" >
		</td>
    <td width="81%" valign="top"><table  width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr >
        <td colspan="6"><span class="style14">List of Refunds </span></td>
        </tr>
      <tr >
        <td>&nbsp;</td>
        <td colspan="5"><%=session("msg") %></td>
        </tr>
      <tr>
        <td width="7%" bgcolor="#000000"><span class="style21">S.N.</span></td>
        <td width="18%" bgcolor="#000000"><span class="style21">Name</span></td>
        <td width="21%" bgcolor="#000000"><span class="style21">Address</span></td>
        <td width="16%" bgcolor="#000000"><span class="style21">City</span></td>
        <td width="15%" bgcolor="#000000"><span class="style21">Tel No. </span></td>
        <td width="23%" bgcolor="#000000"><div align="center" class="style22"><span class="style19">Action</span></div></td>
      </tr>
	  <%
	    count= iPageSize*iPageCurrent-(iPageSize-1)
	
	  if not objPagingRS.eof  then
		iRecordsShown = 0
	 
 
	  Do While iRecordsShown < iPageSize And Not objPagingRS.EOF
	   if bgcol = "#ffffff" then
	        'bgcol = "#cccccc"
			'bgcol="#D1D1D1"
	   		'bgcol = "#c4c4c4"
			bgcol = "#DFDFDF"
	   else
	   		bgcol = "#ffffff"
	   end if
	  
	 %>
      <tr bgcolor="<%=bgcol%>" >
        <td><%=count %></td>
        <td><%=objPagingRS("name") %></td>
        <td><%=objPagingRS("address") %></td>
        <td><%=objPagingRS("city") %></td>
        <td><%=objPagingRS("telno") %></td>
        <td><div align="center"><a href="./editrefunds.asp?refundid=<%=objPagingRS("refundid")%>">[Edit]</a>&nbsp;|&nbsp;<a href="#" onclick="javascript:return suredel('<%=objPagingRS("refundid")%>');">[Delete]</a>&nbsp;|&nbsp;<a href="./letterrefund.asp?refundid=<%=objPagingRS("refundid")%>">[Letter] </a></div></td>
      </tr>
	  <%
	  count=count+1
      iRecordsShown = iRecordsShown + 1
	  objPagingRS.movenext	 
	  loop
	  
	  else
	  
	  end if
	  %>
      <tr >
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr >
         <td align="left" colspan="6">
      Showing Page <strong><%= iPageCurrent %></strong> of <strong><%= iPageCount %></strong> <br>
      Move to: <%
	If iPageCurrent > 1 Then
	%>
      <a href="listrefund.asp?page=<%= iPageCurrent - 1 %>">[Previous]</a>
      <%
	End If

	For I = 1 To iPageCount
		If I = iPageCurrent Then
			%>
      <%= I %>
      <%
		Else
			%>
      <a href="listrefund.asp?page=<%= I %>"><%= I %></a>
      <%
		End If
	Next 


	If iPageCurrent < iPageCount Then
		%>
      <a href="listrefund.asp?page=<%= iPageCurrent + 1 %>">[Next]</a>
      <%
	End If

	%>	</td>
      </tr>
    </table></td>
  </tr>
  
</table>
<% session("msg")=""  %>
<script language="javascript1.2">
	function suredel(obj){
		if(confirm("Are you sure to delete this record?")){
			document.location="deleteprocess.asp?refundid="+obj+"";
			return true;
			}

	}
</script>
