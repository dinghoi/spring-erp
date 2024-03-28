<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim from_date
Dim to_date

slip_month=Request.form("slip_month")
view_c=Request.form("view_c")
emp_name=Request.form("emp_name")

if slip_month = "" then
	slip_month = mid(cstr(now()),1,4) + mid(cstr(now()),6,2)
	view_c = "total"
	emp_name = ""
end If

from_date = mid(slip_month,1,4) + "-" + mid(slip_month,5,2) + "-01"
end_date = datevalue(from_date)
end_date = dateadd("m",1,from_date)
to_date = cstr(dateadd("d",-1,end_date))
sign_month = slip_month

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

' 포지션별
posi_sql = " and (emp_no = '"&user_id&"' or reg_id = '"&user_id&"')"

if cost_grade = "0" then
	posi_sql = ""
end if 

' 조건별 조회.........
base_sql = "select * from general_cost where (slip_gubun = '비용') and (cost_reg = '1') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"')"
order_sql = " ORDER BY slip_date ASC"

sql = base_sql + posi_sql + order_sql
Rs.Open Sql, Dbconn, 1

title_line = "비용 대행 등록관리"
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
			function getPageCode(){
				return "0 1";
			}
		</script>
		<script type="text/javascript">
			function frmcheck () {
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {
				if (document.frm.slip_month.value == "") {
					alert ("발생년월을 입력하세요.");
					return false;
				}	
				return true;
			}
			function condi_view() {

				if (eval("document.frm.view_c[0].checked")) {
					document.getElementById('emp_name_view').style.display = 'none';
				}	
				if (eval("document.frm.view_c[1].checked")) {
					document.getElementById('emp_name_view').style.display = '';
				}	
			}
		</script>

	</head>
	<body onLoad="condi_view()">
		<div id="wrap">			
			<!--#include virtual = "/include/cost_header.asp" -->
			<!--#include virtual = "/include/cost_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="others_cost_mg.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>조건 검색</dt>
                        <dd>
                            <p>
								<label>
								&nbsp;&nbsp;<strong>발생년월&nbsp;</strong>(예201401) : 
                                	<input name="slip_month" type="text" value="<%=slip_month%>" style="width:70px">
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="8%" >
							<col width="11%" >
							<col width="10%" >
							<col width="10%" >
							<col width="8%" >
							<col width="10%" >
							<col width="10%" >
							<col width="*" >
							<col width="6%" >
							<col width="6%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">발생일자</th>
								<th scope="col">소속</th>
								<th scope="col">비용구분</th>
								<th scope="col">비용항목</th>
								<th scope="col">사용금액</th>
								<th scope="col">발생사유</th>
								<th scope="col">사용자</th>
								<th scope="col">비고</th>
								<th scope="col">마감</th>
								<th scope="col">수정</th>
							</tr>
						</thead>
						<tbody>
						<%
						cost_sum = 0
						do until rs.eof
							cost_sum = cost_sum + rs("cost")
							if rs("end_yn") = "Y" then
								end_yn = "마감"
							  else
							  	end_yn = "진행"
							end if
						%>
							<tr>
								<td class="first"><%=rs("slip_date")%></td>
								<td><%=rs("org_name")%></td>
								<td><%=rs("account")%></td>
								<td><%=rs("account_item")%></td>
								<td class="right"><%=formatnumber(rs("cost"),0)%></td>
								<td><%=rs("customer")%></td>
								<td><%=rs("emp_name")%>&nbsp;<%=rs("emp_grade")%></td>
								<td><%=rs("slip_memo")%></td>
								<td><%=end_yn%></td>
								<td>
							<% if rs("end_yn") <> "Y" then %>
                                <a href="#" onClick="pop_Window('others_cost_add.asp?slip_date=<%=rs("slip_date")%>&slip_seq=<%=rs("slip_seq")%>&u_type=<%="U"%>','general_cost_add_pop','scrollbars=yes,width=800,height=250')">수정</a>
                            <%  else	%>
								마감
	                        <% end if %>
                                </td>
							</tr>
						<%
							rs.movenext()
						loop
						rs.close()
						%>
							<tr>
								<th class="first" colspan="4">합 계</th>
							  	<th class="right"><%=formatnumber(cost_sum,0)%></th>
							  	<th>&nbsp;</th>
							  	<th>&nbsp;</th>
							  	<th>&nbsp;</th>
							  	<th>&nbsp;</th>
							  	<th>&nbsp;</th>
						  	</tr>
						</tbody>
					</table>
				</div>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td width="20%">
					<div class="btnCenter">
					</div>
                    </td>                
				    <td width="50%">
                    </td>
				    <td width="30%">
					<div class="btnRight">
					<a href="#" onClick="pop_Window('others_cost_add.asp','others_cost_add_pop','scrollbars=yes,width=800,height=250')" class="btnType04">비용대행등록</a>
					</div>                  
                    </td>
			      </tr>
				  </table>
				<br>
			</form>
		</div>				
	</div>        				
	</body>
</html>

