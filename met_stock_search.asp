<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim in_name
Dim rs
Dim rs_numRows

gubun = request("gubun")
view_condi=Request("view_condi")
stock_go_man = Request("stock_go_man") 

if gubun = "" then
   gubun = Request.Form("gubun")
'   view_condi = Request.Form("view_condi")
   stock_go_man = Request.Form("stock_go_man")
end if

If Request.Form("stock_go_man")  <> "" Then 
  stock_go_man = Request.Form("stock_go_man") 
End If

Set dbconn = Server.CreateObject("ADODB.connection")
Set rs = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
dbconn.open dbconnect

Sql = "SELECT * FROM emp_master where emp_no = '"&stock_go_man&"'"
Set Rs_emp = DbConn.Execute(SQL)
if not Rs_emp.eof then
    	stock_go_name = Rs_emp("emp_name")
   else
		stock_go_name = ""
end if
Rs_emp.close()

Sql = "select * from met_stock_code where (stock_go_man like '%" + stock_go_man + "%') ORDER BY stock_company,stock_level,stock_bonbu,stock_name ASC"

Rs.Open Sql, Dbconn, 1

title_line = " 창고 출입고담당자 검색 "

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>창고 출입고담당자 검색</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript" src="/java/js_window.js"></script>
		<script type="text/javascript">
			function stockesel(stock_code,stock_level,stock_name,stock_manager_code,stock_manager_name,stock_company,stock_bonbu,stock_saupbu,stock_team,gubun)
			{
				<%
				'alert(gubun);
				%>
				if(gubun =="mving")
					{ 
					opener.document.frm.condi.value = stock_name;
					opener.document.frm.stock_code.value = stock_code;
					opener.document.frm.stock_manager_code.value = stock_manager_code;
					opener.document.frm.stock_manager_name.value = stock_manager_name;
					window.close();
					}	
				
				
				<%	
				'else
				'	{ 
				'	opener.document.frm.sido.value = sido;
				'   opener.document.frm.family_gugun.value = gugun;
				'   opener.document.frm.family_dong.value = dong;
				'   opener.document.frm.family_zip.value = zip;
				'    window.close();
				'    opener.document.frm.family_addr.focus();
				'	}
				%>
			}			
			function frmcheck () {
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {
				if(document.frm.stock_go_man.value =="") {
					alert('출입고 담장자를 입력하세요');
					frm.stock_go_man.focus();
					return false;}
				{
					return true;
				}
			}
		</script>

	</head>
	<body>
		<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="met_stock_search.asp?gubun=<%=gubun%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
                        <dd>
                            <p>
							<strong>담당자 사번을 입력하세요 </strong>
								<label>
        						<input name="stock_go_man" type="text" id="stock_go_man" value="<%=stock_go_man%>" style="width:40px; text-align:left;">
								</label>
                                <label>
								<strong>담당자: </strong>
                                	<input name="stock_go_name" type="text" value="<%=stock_go_name%>" readonly="true" style="width:100px" id="stock_go_name">
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser1.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="20%" >
							<col width="15%" >
                            <col width="15%" >
							<col width="*" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">창고명</th>
								<th scope="col">창고Level</th>
								<th scope="col">창고장</th>
                                <th scope="col">소속</th>
 							</tr>
						</thead>   
						<tbody>
					<%
						    i = 0
							do until rs.eof or rs.bof
							   i = i + 1
					%>
							<tr>
								<td class="first"><a href="#" onClick="stockesel('<%=rs("stock_code")%>','<%=rs("stock_level")%>','<%=rs("stock_name")%>','<%=rs("stock_manager_code")%>','<%=rs("stock_manager_name")%>','<%=rs("stock_company")%>','<%=rs("stock_bonbu")%>','<%=rs("stock_saupbu")%>','<%=rs("stock_team")%>','<%=gubun%>');"><%=rs("stock_name")%></a>
                                </td>
								<td><%=rs("stock_level")%>&nbsp;</td>
                                <td><%=rs("stock_manager_name")%>(<%=rs("stock_manager_code")%>)&nbsp;</td>
								<td class="left"><%=rs("stock_company")%> - <%=rs("stock_bonbu")%> - <%=rs("stock_saupbu")%>&nbsp;</td>
							</tr>
					<%
								rs.movenext()
							loop
							rs.close()
							
							if i = 0 then 
					%>
                            <tr>
								<td class="first" colspan="4">내역이 없습니다</td>
							</tr>
					<%      end if   %>
						</tbody>
					</table>
				</div>
			</div>				
	</div>
                <input type="hidden" name="gubun" value="<%=gubun%>" ID="Hidden1">
                <input type="hidden" name="stock_go_man" value="<%=stock_go_man%>" ID="Hidden1">
                <input type="hidden" name="view_condi" value="<%=view_condi%>" ID="Hidden1">    
            				
	</form>
	</body>
</html>

