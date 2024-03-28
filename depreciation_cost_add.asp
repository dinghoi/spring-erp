<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<!--#include virtual="/include/end_check.asp" -->
<%

end_date = "2014-12-31"

u_type = request("u_type")
slip_date = request("slip_date")
slip_seq = request("slip_seq")

org_company = ""
account = ""
price = 0
slip_memo = ""
end_yn = "N"
curr_date = mid(cstr(now()),1,10)

title_line = "상각비 등록"
if u_type = "U" then

	Sql="select * from general_cost where slip_date = '"&slip_date&"' and slip_seq = '"&slip_seq&"'"
	Set rs=DbConn.Execute(Sql)

	org_company = rs("emp_company")
	org_name = rs("org_name")
	account = rs("account")
	price = rs("price")
	emp_name = rs("emp_name")
	emp_grade = rs("emp_grade")
	slip_memo = rs("slip_memo")
	reg_id = rs("reg_id")
	rs.close()

	title_line = "상각비 변경"
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
				if(document.frm.slip_date.value <= document.frm.end_date.value) {
					alert('비용일자가 마감이 되어 있는 날자입니다');
					frm.slip_date.focus();
					return false;}
				if(document.frm.slip_date.value > document.frm.curr_date.value) {
					alert('비용일자가 현재일보다 클수가 없습니다.');
					frm.slip_date.focus();
					return false;}
				if(document.frm.end_yn.value =="Y") {
					alert('마감되어 수정 할 수 없습니다');
					frm.end_yn.focus();
					return false;}
				if(document.frm.slip_date.value =="") {
					alert('비용일자를 입력하세요');
					frm.slip_date.focus();
					return false;}
				if(document.frm.account.value =="") {
					alert('비용구분을 선택하세요');
					frm.account.focus();
					return false;}
				if(document.frm.org_company.value =="") {
					alert('비용회사를 선택하세요');
					frm.org_company.focus();
					return false;}
				if(document.frm.price.value =="") {
					alert('금액을 입력하세요');
					frm.price.focus();
					return false;}
				if(document.frm.slip_memo.value =="") {
					alert('발행내역을 입력하세요');
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
			function delcheck() 
				{
				a=confirm('정말 삭제하시겠습니까?')
				if (a==true) {
					document.frm.action = "genneral_cost_del_ok.asp";
					document.frm.submit();
				return true;
				}
				return false;
				}
        </script>
	</head>
	<body onLoad="condi_view()">
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="depreciation_cost_add_save.asp" method="post" name="frm">
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
				        <th class="first">비용일자</th>
				        <td class="left">
                        <input name="slip_date" type="text" value="<%=slip_date%>" style="width:80px;text-align:center" id="datepicker">
				          마감일 : <%=end_date%>
				        <input name="curr_date" type="hidden" value="<%=curr_date%>">
				        <input name="slip_seq" type="hidden" value="<%=slip_seq%>">
                        </td>
				        <th>비용구분</th>
				        <td class="left">
                            <select name="account" id="account" style="width:150px">
                              <option value="" <% if account = "" then %>selected<% end if %>>선택</option>
                              <option value="대손상각비" <% if account = "대손상각비" then %>selected<% end if %>>대손상각비</option>
                              <option value="고정자산" <% if account = "고정자산" then %>selected<% end if %>>고정자산</option>
                              <option value="무형자산" <% if account = "무형자산" then %>selected<% end if %>>무형자산</option>
                            </select>
						</td>
			          </tr>
				      <tr>
				        <th class="first">비용회사</th>
				        <td class="left">
                            <select name="org_company" id="org_company" style="width:120px">
                              <option value="" <% if org_company = "" then %>selected<% end if %>>선택</option>
                              <%
																' 2019.02.22 박정신 요청 회사리스트를 빼고자 할시 org_end_date에 null 이 아닌 만료일자를 셋팅하면 리스트에 나타나지 않는다.
																Sql = "SELECT * FROM emp_org_mst WHERE ISNULL(org_end_date) AND org_level = '회사'  ORDER BY org_company ASC"
                                rs_org.Open Sql, Dbconn, 1
                                do until rs_org.eof
                                %>
                              <option value='<%=rs_org("org_name")%>' <%If org_company = rs_org("org_name") then %>selected<% end if %>><%=rs_org("org_name")%></option>
                              <%
                                    rs_org.movenext()
                                loop
                                rs_org.close()						
                                %>
                            </select>
                        </td>
				        <th>금액</th>
				        <td class="left"><% if u_type = "U" then	%>
                          <input name="price" type="text" id="price" style="width:100px;text-align:right" value="<%=formatnumber(price,0)%>"  onKeyUp="plusComma(this);" >
                          <%   else	%>
                          <input name="price" type="text" id="price" style="width:100px;text-align:right" onKeyUp="plusComma(this);" >
                        <% end if	%></td>
			          </tr>
				      <tr>
				        <th class="first">비용내역</th>
				        <td class="left"><input name="slip_memo" type="text" id="slip_memo" style="width:200px; ime-mode:active" onKeyUp="checklength(this,50);" value="<%=slip_memo%>"></td>
				        <th><span class="first">담당자</span></th>
				        <td class="left"><%=user_name%>&nbsp;<%=user_grade%></td>
			          </tr>
			        </tbody>
			      </table>
				</div>
                <br>
                <div align=center>
				<%	if end_yn = "N" then	%>
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
				<input type="hidden" name="emp_no" value="<%=emp_no%>" ID="Hidden1">
			</form>
		</div>				
	</body>
</html>

