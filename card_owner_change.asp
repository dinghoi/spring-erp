<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
u_type = request("u_type")
card_no = request("card_no")

card_no1 = ""
card_no2 = ""
card_no3 = ""
card_no4 = ""
card_type  = ""
emp_name  = ""
emp_grade = ""
card_issue = "신규"
card_limit = ""
valid_thru = ""
card_memo = ""
use_yn = "Y"

curr_date = mid(now(),1,10)
title_line = "카드 사용자 등록"
if u_type = "U" then

	Sql="select * from card_owner where card_no = '"&card_no&"'"
	Set rs=DbConn.Execute(Sql)

 	emp_no = rs("emp_no")
 	emp_name = rs("emp_name")
	card_no1 = mid(rs("card_no"),1,4)
	card_no2 = mid(rs("card_no"),6,4)
	card_no3 = mid(rs("card_no"),11,4)
	card_no4 = mid(rs("card_no"),16,4)
	card_type = rs("card_type")
	card_issue = rs("card_issue")
	card_limit = rs("card_limit")
	valid_thru = rs("valid_thru")
	create_date = rs("create_date")
	start_date = rs("start_date")
	card_memo = rs("card_memo")
	use_yn = rs("use_yn")
	reg_id = rs("reg_id")
	reg_date = mid(rs("reg_date"),1,10)
	reg_name = rs("reg_name")
	mod_id = rs("mod_id")
	mod_date = rs("mod_date")
	mod_name = rs("mod_name")
	rs.close()

	title_line = "카드 사용자 변경"
end if
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>관리 회계 시스템</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			$(function(){
				$( "#datepicker" ).datepicker();
				$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
				$( "#datepicker" ).datepicker("setDate", "<%=change_date%>" );
			});

			function goAction(){
			   window.close();
			}

			function goBefore(){
			   history.back();
			}

			function frmcheck(){
				if(chkfrm()){
					document.frm.submit();
				}
			}

			function chkfrm(){
				if(document.frm.emp_no.value ==""){
					alert('사원조회을 하세요');
					frm.emp_no.focus();
					return false;
				}

				if(document.frm.change_date.value ==""){
					alert('변경일을 입력 하세요');
					frm.change_date.focus();
					return false;
				}

				if(document.frm.start_date.value > document.frm.change_date.value){
					alert('변경일이 사용개시일보다 작을수 없습니다.');
					frm.change_date.focus();
					return false;
				}

				if(document.frm.mod_memo.value ==""){
					alert('비고를 입력하세요');
					frm.mod_memo.focus();
					return false;
				}

				{
					a=confirm('입력하시겠습니까?');

					if(a==true){
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
				<form action="/card_owner_change_save.asp" method="post" name="frm">
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
				        <th class="first">카드종류</th>
				        <td class="left"><%=card_type%></td>
				        <th>카드번호</th>
				        <td class="left"><%=card_no%></td>
			          </tr>
				      <tr>
				        <th class="first">기존사용자</th>
				        <td class="left"><%=emp_name%>&nbsp;<%=emp_grade%></td>
				        <th>변경사용자</th>
				        <td class="left">
                        <input name="emp_name" type="text" id="emp_name" style="width:60px">
			            <input name="emp_grade" type="text" id="emp_grade" style="width:60px">
                        <a href="#" onClick="pop_Window('/member/memb_search.asp','memb_search_pop','scrollbars=yes,width=600,height=400')" class="btnType03">사원조회</a>
                        </td>
			          </tr>
				      <tr>
				        <th class="first">변경일</th>
				        <td class="left"><input name="change_date" type="text" style="width:80px;text-align:center" id="datepicker"></td>
				        <th>변경사유</th>
				        <td class="left"><input name="mod_memo" type="text" id="card_memo" style="width:200px; ime-mode:active" onKeyUp="checklength(this,50);"></td>
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
                </div>
                    <input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
                    <input type="hidden" name="curr_date" value="<%=curr_date%>" ID="Hidden1">
                    <input type="hidden" name="card_no" value="<%=card_no%>" ID="Hidden1">
                    <input type="hidden" name="old_emp_no" value="<%=emp_no%>" ID="Hidden1">
                    <input type="hidden" name="emp_no" ID="Hidden1">
                    <input type="hidden" name="org_name" value="<%=org_name%>" ID="Hidden1">
                    <input type="hidden" name="start_date" value="<%=start_date%>" ID="Hidden1">
				</form>
		</div>
	</body>
</html>

