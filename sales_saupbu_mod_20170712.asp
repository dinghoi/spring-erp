<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
approve_no = request("approve_no")

Sql="select * from saupbu_sales where approve_no = '"&approve_no&"'"
Set rs=DbConn.Execute(Sql)

title_line = "매출 사업부 수정"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>영업 관리 시스템</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
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
				if(document.frm.saupbu.value =="") {
					alert('영업사업부를 선택하세요');
					frm.saupbu.focus();
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
	<body">
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="sales_saupbu_mod_save.asp" method="post" name="frm">
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
				        <th class="first">매출일자</th>
				        <td class="left"><%=rs("sales_date")%></td>
				        <th>매출회사</th>
				        <td class="left"><%=rs("sales_company")%></td>
			          </tr>
				      <tr>
				        <th class="first">영업사업부</th>
				        <td class="left">
	                        <select name="saupbu" id="saupbu" style="width:150px">
                                <option value="회사간거래" <% if rs("saupbu") ="회사간거래" then %>selected<% end if %>>회사간거래</option>
                                <option value="기타사업부" <% if rs("saupbu") ="기타사업부" then %>selected<% end if %>>기타사업부</option>
                        <%
						Sql="select saupbu from sales_org order by sort_seq asc"
						rs_org.Open Sql, Dbconn, 1
						do until rs_org.eof
                        %>
                                <option value='<%=rs_org("saupbu")%>' <%If rs("saupbu") = rs_org("saupbu") then %>selected<% end if %>><%=rs_org("saupbu")%></option>
                        <%
	                        rs_org.movenext()
                        loop
                        rs_org.close()						
                        %>
	                        </select>
                        </td>
				        <th>담당자</th>
				        <td class="left">
                        	<input name="emp_name" type="text" id="emp_name" style="width:60px" value="<%=rs("emp_name")%>" readonly="true">
                        	<input name="emp_no" type="text" id="emp_no" style="width:60px" value="<%=rs("emp_no")%>" readonly="true">
                          	<input name="emp_grade" type="hidden" id="emp_grade" style="width:60px" readonly="true">
                        <a href="#" onClick="pop_Window('emp_search.asp?gubun=<%="1"%>','emp_search_pop','scrollbars=yes,width=600,height=400')" class="btnType03">사원조회</a></td>
			          </tr>
				      <tr>
				        <th class="first">공급가액</th>
				        <td class="left"><%=formatnumber(rs("cost_amt"),0)%></td>
				        <th>세액</th>
				        <td class="left"><%=formatnumber(rs("vat_amt"),0)%></td>
			          </tr>
				      <tr>
				        <th class="first">합계금액</th>
				        <td class="left"><%=formatnumber(rs("sales_amt"),0)%></td>
				        <th>품목명</th>
				        <td class="left"><%=rs("sales_memo")%></td>
			          </tr>
			        </tbody>
			      </table>
				</div>
                <br>
                <div align=center>
                    <span class="btnType01"><input type="button" value="변경" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
                    <span class="btnType01"><input type="button" value="취소" onclick="javascript:goAction();"></span>
                </div>
                    <input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
                    <input type="hidden" name="sales_date" value="<%=rs("sales_date")%>" ID="Hidden1">
                    <input type="hidden" name="approve_no" value="<%=rs("approve_no")%>" ID="Hidden1">
				</form>
		</div>				
	</body>
</html>

