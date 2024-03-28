<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
Dim Rs
Dim Repeat_Rows
Dim from_date
Dim to_date
Dim win_sw

slip_month=Request.form("slip_month")
view_c=Request.form("view_c")

If slip_month = "" Then
	slip_month = mid(cstr(now()),1,4) + mid(cstr(now()),6,2)
	view_c = "total"
End If

from_date = mid(slip_month,1,4) + "-" + mid(slip_month,5,2) + "-01"
end_date = datevalue(from_date)
end_date = dateadd("m",1,from_date)
to_date = cstr(dateadd("d",-1,end_date))
sign_month = slip_month

' 조건별 조회.........
' 포지션별
posi_sql = " and reg_id = '" + user_id + "'"

if position = "사업부장" or cost_grade = "2" then
	posi_sql = " and saupbu = '"&saupbu&"'"
end if
if position = "본부장" or cost_grade = "1" then
	posi_sql = " and bonbu = '"&bonbu&"'"
end if

if cost_grade = "0" then
	posi_sql = ""
end if

base_sql = "select * from general_cost where (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') and (slip_gubun ='자재' or slip_gubun ='장비')"
order_sql = " ORDER BY slip_date ASC"

sql = base_sql + posi_sql + order_sql
Rs.Open Sql, Dbconn, 1

title_line = "자재 및 장비 비용 관리"
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
				if (chkfrm()) {
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
		</script>
	</head>
	<body>
		<div id="wrap">			
			<!--#include virtual = "/include/cost_header.asp" -->
			<!--#include virtual = "/include/cost_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="etc_cost_mg.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>조건 검색</dt>
                        <dd>
                            <p>
								<label>
								&nbsp;&nbsp;<strong>구매년월&nbsp;</strong>(예201401) : 
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
							<col width="*" >
							<col width="12%" >
							<col width="7%" >
							<col width="12%" >
							<col width="7%" >
							<col width="7%" >
							<col width="7%" >
							<col width="7%" >
							<col width="10%" >
							<col width="13%" >
							<col width="5%" >
							<col width="3%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">사용조직</th>
								<th scope="col">고객사</th>
								<th scope="col">발행일자</th>
								<th scope="col">외주업체</th>
								<th scope="col">사업자번호</th>
								<th scope="col">합계</th>
								<th scope="col">공급가액</th>
								<th scope="col">부가세</th>
								<th scope="col">구매유형</th>
								<th scope="col">발행내역</th>
								<th scope="col">등록자</th>
								<th scope="col">수정</th>
							</tr>
						</thead>
						<tbody>
						<%
						price_sum = 0
						cost_sum = 0
						cost_vat_sum = 0
						do until rs.eof
							price_sum = price_sum + rs("price")
							cost_sum = cost_sum + rs("cost")
							cost_vat_sum = cost_vat_sum + rs("cost_vat")
							org_name = rs("emp_company") + "/" + rs("org_name")
							customer_no = mid(rs("customer_no"),1,3) + "-" + mid(rs("customer_no"),4,2) + "-" + mid(rs("customer_no"),6)
						%>
							<tr>
								<td class="first"><%=rs("org_name")%></td>
								<td><%=rs("company")%></td>
								<td><%=rs("slip_date")%></td>
								<td><%=rs("customer")%></td>
								<td><%=customer_no%></td>
							  	<td class="right"><%=formatnumber(rs("price"),0)%></td>
							  	<td class="right"><%=formatnumber(rs("cost"),0)%></td>
							  	<td class="right"><%=formatnumber(rs("cost_vat"),0)%></td>
								<td><%=rs("slip_gubun")%>-<%=rs("account")%></td>
								<td><%=rs("slip_memo")%></td>
								<td><%=rs("reg_user")%></td>
								<td>
							<% if rs("end_yn") = "C" or rs("end_yn") = "N" then %>
							<%   if rs("reg_id") = user_id then	%>
                                <a href="#" onClick="pop_Window('etc_cost_add.asp?slip_date=<%=rs("slip_date")%>&slip_seq=<%=rs("slip_seq")%>&u_type=<%="U"%>','etc_cost_add_pop','scrollbars=yes,width=800,height=280')">수정</a>
							<%     else	%>
								불가
                            <%	 end if	%>
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
								<th class="first" colspan="5">합 계</th>
							  	<th class="right"><%=formatnumber(price_sum,0)%></th>
							  	<th class="right"><%=formatnumber(cost_sum,0)%></th>
							  	<th class="right"><%=formatnumber(cost_vat_sum,0)%></th>
							  	<th class="right" colspan="4">&nbsp;</th>
							</tr>
						</tbody>
					</table>
				</div>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td width="20%">
					<div class="btnCenter">
                    <a href="tax_bill_cost_excel.asp?from_date=<%=from_date%>&to_date=<%=to_date%>&slip_gubun=<%="자재"%>" class="btnType04">엑셀다운로드</a>
					</div>
                    </td>                
				    <td width="50%">
                    </td>
				    <td width="30%">
					<div class="btnRight">
					<a href="#" onClick="pop_Window('etc_cost_add.asp','etc_cost_add_pop','scrollbars=yes,width=800,height=280')" class="btnType04">자재 및 장비 세금계산서 등록</a>
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

