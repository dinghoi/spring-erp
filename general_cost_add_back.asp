<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/srvmg_dbcon.asp" -->
<!--#include virtual="/include/srvmg_user.asp" -->
<%
u_type = request("u_type")
slip_date = request("slip_date")
slip_seq = request("slip_seq")

slip_gubun = ""
account = ""
paper_no = ""
pay_method = ""
price = 0
vat_yn = "N"
pay_yn = "N"
customer = ""
use_man = ""
emp_no = ""
slip_memo = ""
end_yn = "N"
curr_date = mid(cstr(now()),1,10)
last_end_date = "2014-01-01"

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_acc = Server.CreateObject("ADODB.Recordset")
Set Rs_memb = Server.CreateObject("ADODB.Recordset")
Set Rs_type = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

title_line = "일반경비 등록"
if u_type = "U" then

	Sql="select * from general_cost where slip_date = '"&slip_date&"' and slip_seq = '"&slip_seq&"'"
	Set rs=DbConn.Execute(Sql)

	bonbu = rs("bonbu")
	saupbu = rs("saupbu")
	team = rs("team")
	account = rs("account") + "/" + rs("account_item")
	paper_no = rs("paper_no")
	pay_method = rs("pay_method")
	price = rs("price")
	vat_yn = rs("vat_yn")
	pay_yn = rs("pay_yn")
	customer = rs("customer")
	use_man = rs("use_man")
	emp_no = rs("emp_no")
	slip_memo = rs("slip_memo")
	end_yn = rs("end_yn")
	rs.close()

	title_line = "일반경비 변경"
end if
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>관리회계시스템</title>
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
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}			
			function chkfrm() {
				if(document.frm.slip_date.value <= document.frm.last_end_date.value) {
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
				if(document.frm.account.value =="") {
					alert('비용항목 선택하세요');
					frm.account.focus();
					return false;}
				if(document.frm.paper_no.value =="") {
					alert('전자결재번호를 입력하세요');
					frm.paper_no.focus();
					return false;}
				if(document.frm.pay_method.value =="") {
					alert('사용구분 선택하세요');
					frm.pay_method.focus();
					return false;}
				if(document.frm.price.value ==0) {
					alert('비용을 입력하세요');
					frm.price.focus();
					return false;}
				if(document.frm.customer.value =="") {
					alert('발생사유를 입력하세요');
					frm.customer.focus();
					return false;}
				if(document.frm.emp_no.value =="") {
					alert('사용자를 선택하세요');
					frm.emp_no.focus();
					return false;}
				k = 0;
				for (j=0;j<2;j++) {
					if (eval("document.frm.pay_yn[" + j + "].checked")) {
						k = k + 1
					}
				}
				if (k==0) {
					alert ("정산 여부를 체크하세요");
					return false;
				}	
				if(document.frm.slip_memo.value =="") {
					alert('비고를 입력하세요');
					frm.slip_memo.focus();
					return false;}

				{
				a=confirm('입력하시겠습니까?')
				if (a==true) {
					return true;
				}
				return false;
				}
			}
        </script>
	</head>
	<body>
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="general_cost_add_save.asp" method="post" name="frm">
				<div class="gView">
				  <table cellpadding="0" cellspacing="0" class="tableWrite">
				    <colgroup>
				      <col width="15%" >
				      <col width="35%" >
				      <col width="15%" >
				      <col width="*" >
			        </colgroup>
				    <tbody>
				      <tr>
				        <th class="first">발생일자</th>
				        <td class="left">
                        <input name="slip_date" type="text" value="<%=slip_date%>" style="width:80px;text-align:center" id="datepicker">
				          마감일 : <%=last_end_date%>
				        <input name="curr_date" type="hidden" value="<%=curr_date%>">
				        <input name="slip_seq" type="hidden" value="<%=slip_seq%>">
                        </td>
				        <th>소속</th>
				        <td class="left">
						<%=bonbu%>&nbsp;<%=saupbu%>&nbsp;<%=team%>
			            <input name="bonbu" type="hidden" value="<%=bonbu%>">
				        <input name="saupbu" type="hidden" value="<%=saupbu%>">
				        <input name="team" type="hidden" value="<%=team%>">
                        </td>
			          </tr>
				      <tr>
				        <th class="first">사용자</th>
				        <td class="left"><select name="emp_no" id="emp_no" style="width:200px">
				          <option value="" <% if emp_no = "" then %>selected<% end if %>>선택</option>
				          <%
'                            Sql="select * from memb where mg_group = '"&mg_group&"' and bonbu = '"&bonbu&"' and saupbu = '"&saupbu&"' order by user_name asc"
                            Sql="select * from memb where mg_group = '"&mg_group&"' and belong = '"&belong&"' order by user_name asc"
                            rs_memb.Open Sql, Dbconn, 1
                            do until rs_memb.eof
						  %>
				          <option value='<%=rs_memb("user_id")%>' <%If rs_memb("user_id") = emp_no then %>selected<% end if %>><%=rs_memb("user_name")%>/<%=rs_memb("reside_place")%></option>
				          <%
                                rs_memb.movenext()
                            loop
                            rs_memb.close()						
                          %>
			            </select></td>
				        <th><span class="first">비용항목</span></th>
				        <td class="left"><select name="account" id="account" style="width:150px">
		                <option value="" <% if account = "" then %>selected<% end if %>>선택</option>
				            <%
                                    Sql="select * from account_item where account_id = '비용' order by account_name, account_item asc"
                                    rs_acc.Open Sql, Dbconn, 1
                                    do until rs_acc.eof
										account_item = rs_acc("account_name") + "/" + rs_acc("account_item")
								  %>
				            <option value='<%=account_item%>' <%If account_item = account then %>selected<% end if %>><%=account_item%></option>
				            <%
                                        rs_acc.movenext()
                                    loop
                                    rs_acc.close()						
                                  %>
		                </select></td>
			          </tr>
				      <tr>
				        <th class="first">사용구분</th>
				        <td class="left">
                        <select name="pay_method" id="pay_method" style="width:150px">
				          <option value="" <% if pay_method = "" then %>selected<% end if %>>선택</option>
				          <option value='카드' <%If pay_method = "카드" then %>selected<% end if %>>카드</option>
				          <option value='현금' <%If pay_method = "현금" then %>selected<% end if %>>현금</option>
				          <option value='법인카드' <%If pay_method = "법인카드" then %>selected<% end if %>>법인카드</option>
				        </select>				        
				        </td>
				        <th>금액</th>
				        <td class="left"><input name="price" type="text" id="price" style="width:100px;text-align:right" value="<%=formatnumber(price,0)%>" onKeyUp="plusComma(this);" ></td>
			          </tr>
				      <tr>
				        <th class="first">발생사유</th>
				        <td class="left"><input name="customer" type="text" id="customer" style="width:200px; ime-mode:active" onKeyUp="checklength(this,50);" value="<%=customer%>"></td>
				        <th>전자결재NO</th>
				        <td class="left"><input name="paper_no" type="text" id="paper_no" style="width:150px; ime-mode:active" onKeyUp="checklength(this,20);" value="<%=paper_no%>"></td>
			          </tr>
				      <tr>
				        <th class="first">정산여부</th>
				        <td class="left">
                        <input type="radio" name="pay_yn" value="N" <% if pay_yn = "N" then %>checked<% end if %> style="width:30px" id="Radio3">
				          미정산
				        <input type="radio" name="pay_yn" value="Y" <% if pay_yn = "Y" then %>checked<% end if %> style="width:30px" id="Radio4">
				            정산
                        </td>
				        <th>비고</th>
				        <td class="left"><input name="slip_memo" type="text" id="slip_memo" style="width:200px; ime-mode:active" onKeyUp="checklength(this,50);" value="<%=slip_memo%>"></td>
			          </tr>
			        </tbody>
			      </table>
				</div>
                <br>
                <div align=center>
				<%	if end_yn = "N" or end_yn = "C" then	%>
                    <span class="btnType01"><input type="button" value="등록" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
        		<%	end if	%>
                    <span class="btnType01"><input type="button" value="취소" onclick="javascript:goAction();"></span>
                </div>
				<input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
				<input type="hidden" name="end_yn" value="<%=end_yn%>" ID="Hidden1">
				<input type="hidden" name="last_end_date" value="<%=last_end_date%>" ID="Hidden1">
				</form>
		</div>				
	</body>
</html>

