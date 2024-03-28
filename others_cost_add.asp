<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<!--#include virtual="/include/end_check.asp" -->
<%
u_type = request("u_type")
slip_date = request("slip_date")
slip_seq = request("slip_seq")

slip_gubun = ""
account = ""
sign_no = ""
pay_method = ""
price = 0
vat_yn = "N"
pay_yn = "N"
company = "공통"
customer = ""
emp_name = user_name
emp_no = user_id
slip_memo = ""
end_yn = "N"
cancel_yn = "N"
curr_date = mid(cstr(now()),1,10)

title_line = "비용 대행 등록"
if u_type = "U" then

	Sql="select * from general_cost where cost_reg = '1' and slip_date = '"&slip_date&"' and slip_seq = '"&slip_seq&"'"
	Set rs=DbConn.Execute(Sql)

	org_name = rs("org_name")
	account = rs("account") + "-" + rs("account_item")
	sign_no = rs("sign_no")
	pay_method = rs("pay_method")
	price = rs("price")
	company = rs("company")
	vat_yn = rs("vat_yn")
	pay_yn = rs("pay_yn")
	customer = rs("customer")
	emp_name = rs("emp_name")
	emp_no = rs("emp_no")
	slip_memo = rs("slip_memo")
	end_yn = rs("end_yn")
	cancel_yn = rs("cancel_yn")
	reg_id = rs("reg_id")
	reg_date = rs("reg_date")
	reg_user = rs("reg_user")
	mod_id = rs("mod_id")
	mod_date = rs("mod_date")
	mod_user = rs("mod_user")
	rs.close()

	title_line = "비용 대행 변경"
end if
if end_yn = "Y" then
	end_view = "마감"
  else
  	end_view = "진행"
end if
if cancel_yn = "Y" then
	cancel_view = "취소"
  else
  	cancel_view = "지급"
end if
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>비용 관리 시스템</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			$(function() {    $( "#datepicker" ).datepicker();
												$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker" ).datepicker("setDate", "<%=slip_date%>" );
			});	  
			function goAction () {
			   window.close () ;
			}
			function goBefore () {
			   history.back() ;
			}
			function frmcheck () {
				if (chkfrm()) {
					document.frm.submit ();
				}
			}			
			function chkfrm() {
				var companyChk = document.getElementById("company").value;
				//alert(companyChk);
				
				
				
				if(document.frm.slip_date.value <= document.frm.end_date.value) {
					alert('발생일자가 마감이 되어 있는 날자입니다');
					frm.slip_date.focus();
					return false;}
				if(document.frm.slip_date.value > document.frm.curr_date.value) {
					alert('발생일자가 현재일보다 클수가 없습니다.');
					frm.slip_date.focus();
					return false;}
				if(document.frm.end_yn.value =="Y") {
					alert('마감되어 수정 할 수 없습니다');
					frm.end_yn.focus();
					return false;}
				if(document.frm.slip_date.value =="") {
					alert('발생일자를 입력하세요');
					frm.slip_date.focus();
					return false;}
				if(document.frm.emp_no.value =="") {
					alert('사용자를 선택하세요');
					frm.emp_no.focus();
					return false;}
				if(document.frm.account.value =="") {
					alert('비용항목 선택하세요');
					frm.account.focus();
					return false;}
				if(document.frm.price.value =="") {
					alert('사용금액을 입력하세요');
					frm.price.focus();
					return false;}
				if(document.frm.customer.value =="") {
					alert('발생사유를 입력하세요');
					frm.customer.focus();
					return false;}
				if(document.frm.slip_memo.value =="") {
					alert('비고를 입력하세요');
					frm.slip_memo.focus();
					return false;}

				if (fnCheckSaupbu(companyChk)) {;
					{
					a=confirm('입력하시겠습니까?')
					if (a==true) {
						return true;
					}
					return false;
					}
				}
			}
			function delcheck() 
				{
				a=confirm('정말 삭제하시겠습니까?')
				if (a==true) {
					document.frm.action = "others_cost_del_ok.asp";
					document.frm.submit();
				return true;
				}
				return false;
				}
				
				
				function fnCheckSaupbu(obj) {
					var saupbu = '<%=saupbu %>';
					
					if ((obj == '기타사업부') && (saupbu != '경영지원실' )) {
						//alert(saupbu);
						//if (saupbu != '경영지원실' )	{
							alert("기타사업부는 경영지원실외 선택 할 수 없습니다.");
							return false;
						//}
					}
				}
				
        </script>
	</head>
	<body onload="update_view()">
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="others_cost_add_save.asp" method="post" name="frm">
				<div class="gView">
				  <table cellpadding="0" cellspacing="0" class="tableWrite">
				    <colgroup>
				      <col width="13%" >
				      <col width="37%" >
				      <col width="13%" >
				      <col width="*" >
			        </colgroup>
				    <tbody>
				      <tr>
				        <th class="first">발생일자</th>
				        <td class="left">
                        <input name="slip_date" type="text" id="datepicker" style="width:80px;text-align:center" value="<%=slip_date%>" readonly="true">
				          마감일자 : <%=end_date%>
				        <input name="curr_date" type="hidden" value="<%=curr_date%>">
				        <input name="slip_seq" type="hidden" value="<%=slip_seq%>">
                        </td>
				        <th>사용자</th>
				        <td class="left">
                        <select name="emp_no" id="emp_no" style="width:200px">
				          <option value="" <% if emp_no = "" then %>selected<% end if %>>선택</option>
				          <%
                                    Sql="select * from emp_master where emp_bonbu is null order by emp_no asc"
                                    rs_emp.Open Sql, Dbconn, 1
                                    do until rs_emp.eof
								  %>
				          <option value='<%=rs_emp("emp_no")%>' <%If emp_no = rs_emp("emp_no") then %>selected<% end if %>><%=rs_emp("emp_name")%>&nbsp;<%=rs_emp("emp_grade")%></option>
				          <%
                                        rs_emp.movenext()
                                    loop
                                    rs_emp.close()						
                                  %>
			            </select>
                        </td>
			          </tr>
				      <tr>
				        <th class="first">비용항목</th>
				        <td class="left">
                        <select name="account" id="account" style="width:200px">
				          <option value="" <% if account = "" then %>selected<% end if %>>선택</option>
				          <%
                                    Sql="select * from account_item where cost_yn = 'Y' or cost_yn = 'C' order by account_name, account_item asc"
                                    rs_acc.Open Sql, Dbconn, 1
                                    do until rs_acc.eof
										account_item = rs_acc("account_name") + "-" + rs_acc("account_item")
								  %>
				          <option value='<%=account_item%>' <%If account_item = account then %>selected<% end if %>><%=account_item%></option>
				          <%
                                        rs_acc.movenext()
                                    loop
                                    rs_acc.close()						
                                  %>
			            </select>
                        </td>
				        <th>사용금액</th>
				        <td class="left"><select name="pay_method" id="pay_method" style="width:80px">
				          <option value='현금' <%If pay_method = "현금" then %>selected<% end if %>>현금</option>
				          </select>
							&nbsp;
					<% if u_type = "U" then	%>
							<input name="price" type="text" id="price" style="width:100px;text-align:right" value="<%=formatnumber(price,0)%>" onKeyUp="plusComma(this);" >
					<%   else	%>
							<input name="price" type="text" id="price" style="width:100px;text-align:right" onKeyUp="plusComma(this);" >
                    <% end if	%>
                            </td>
			          </tr>
				      <tr>
				        <th class="first">고객사</th>
				        <td class="left">
				        	<%
	                	Sql = "SELECT * FROM trade WHERE use_sw = 'Y' AND mg_group = '"+mg_group+"' ORDER BY trade_name ASC"
	                  rs_trade.Open Sql, Dbconn, 1
                  %>
                  <select name="company" id="company" style="width:150px" onchange="fnCheckSaupbu(this.value)">
                  	<option value='공통' <%If company = "공통"  then %>selected<% end if %>>공통</option>
                  	<% While not rs_trade.eof %>
                    <option value='<%=rs_trade("trade_name")%>' <%If rs_trade("trade_name") = company  then %>selected<% end if %>><%=rs_trade("trade_name")%></option>
                    <% 
                    	 	rs_trade.movenext()  
                    	 Wend 
                    	 rs_trade.Close()
                    %>
                  </select>
                </td>
				        <th>상호명</th>
				        <td class="left"><input name="customer" type="text" id="customer" style="width:200px; ime-mode:active" onKeyUp="checklength(this,50);" value="<%=customer%>"></td>
			          </tr>
				      <tr>
				        <th class="first">발생사유</th>
				        <td class="left"><input name="slip_memo" type="text" id="slip_memo" style="width:200px; ime-mode:active" onKeyUp="checklength(this,50);" value="<%=slip_memo%>"></td>
				        <th>등록정보</th>
				        <td class="left"><%=reg_user%>&nbsp;<%=reg_id%>(<%=reg_date%>)</td>
			          </tr>
			        </tbody>
			      </table>
				</div>
                <br>
                <div align=center>
				<%	if end_yn <> "Y" then	%>
                    <span class="btnType01"><input type="button" value="등록" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
        		<%	end if	%>
                    <span class="btnType01"><input type="button" value="취소" onclick="javascript:goAction();"></span>
				<%	
					if u_type = "U" and user_id = reg_id then
						if end_yn = "N" or end_yn = "C" then	
				%>
                    <span class="btnType01"><input type="button" value="삭제" onclick="javascript:delcheck();" ID="Button1" NAME="Button1"></span>
        		<%	
						end if
					end if	
				%>
                </div>
                    <input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
                    <input type="hidden" name="end_yn" value="<%=end_yn%>" ID="Hidden1">
                    <input type="hidden" name="end_date" value="<%=end_date%>" ID="Hidden1">
                    <input type="hidden" name="old_date" value="<%=slip_date%>" ID="Hidden1">
                    <input type="hidden" name="cancel_yn" value="<%=cancel_yn%>" ID="Hidden1">
                    <input type="hidden" name="mod_id" value="<%=mod_id%>" ID="Hidden1">
                    <input type="hidden" name="mod_user" value="<%=mod_user%>" ID="Hidden1">
                    <input type="hidden" name="mod_date" value="<%=mod_date%>" ID="Hidden1">
				</form>
		</div>				
	</body>
</html>

