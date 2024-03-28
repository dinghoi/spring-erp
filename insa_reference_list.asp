<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows
dim page_cnt
dim pg_cnt
Page=Request("page")
page_cnt=Request.form("page_cnt")
pg_cnt=cint(Request("pg_cnt"))
be_pg = "insa_reference_list.asp"
curr_date = datevalue(mid(cstr(now()),1,10))

if page_cnt > 0 then 
	pg_cnt = page_cnt
end if
if pg_cnt > 0 then
	page_cnt = pg_cnt
end if

if page_cnt < 10 or page_cnt > 20 then
	page_cnt = 10
end if

pgsize = page_cnt ' 화면 한 페이지 

If Page = "" Then
	Page = 1
	start_page = 1
End If
stpage = int((page - 1) * pgsize)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_emp = Server.CreateObject("ADODB.Recordset")
Set rs_sum = Server.CreateObject("ADODB.Recordset")
Set rs_etc = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")

dbconn.open DbConnect

view_sort = request("view_sort")

if view_sort = "" then
	view_sort = "DESC"
end if


order_Sql = " ORDER BY sawo_company,sawo_org_name,sawo_date,sawo_empno " + view_sort
where_sql = " WHERE sawo_target = 'Y' or sawo_target = 'N'"
'where_sql = ""

    in_pay_sum = 0 
	give_pay_sum = 0
	
    sql="select * from emp_sawo_mem " + where_sql
	Rs_sum.Open Sql, Dbconn, 1
	
	do until rs_sum.eof
	   in_pay_sum = in_pay_sum + rs_sum("sawo_in_pay")
	   give_pay_sum = give_pay_sum + rs_sum("sawo_give_pay")
	   
	   rs_sum.movenext()
	loop
    rs_sum.close()
	
	'response.write(in_pay_sum)
	'response.write(give_pay_sum)

Sql = "SELECT count(*) FROM emp_sawo_mem " + where_sql
Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

sql = "select * from emp_sawo_mem " + where_sql + order_sql + " limit "& stpage & "," &pgsize 
Rs.Open Sql, Dbconn, 1

title_line = " 신원보증 가입 현황(공사중....) "

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>인사급여 시스템</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "1 1";
			}
		</script>
		<script type="text/javascript">
			function frmcheck () {
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {
				if (document.frm.condi.value == "") {
					alert ("소속을 선택하시기 바랍니다");
					return false;
				}	
				return true;
			}
		</script>

	</head>
	<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
		<div id="wrap">			
			<!--#include virtual = "/include/insa_header.asp" -->
			<!--#include virtual = "/include/insa_sub_menu1.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_reference_mg.asp" method="post" name="frm">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="6%" >
							<col width="6%" >
							<col width="4%" >
							<col width="4%" >
                            <col width="9%" >
                            <col width="9%" >
							<col width="6%" >
							<col width="5%" >
							<col width="6%" >
							<col width="5%" >
							<col width="6%" >
                            <col width="5%" >
							<col width="6%" >
							<col width="5%" >
                            <col width="6%" >
                            <col width="4%" >
                            <col width="4%" >
                            <col width="4%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">사번</th>
								<th scope="col">성  명</th>
								<th scope="col">직급</th>
								<th scope="col">직책</th>
                                <th scope="col">회사</th>
                                <th scope="col">소속</th>
								<th scope="col">가입일</th>
								<th scope="col">가입구분</th>
								<th scope="col">탈퇴일</th>
                                <th scope="col">탈퇴구분</th>
                                <th scope="col">급여공제</th>
                                <th scope="col">납입횟수</th>
                                <th scope="col">납입금액</th>
                                <th scope="col">지급횟수</th>
                                <th scope="col">지급금액</th>
								<th colspan="3" scope="col">신원보증</th>
							</tr>
						</thead>
					<tbody>
						<%
						
						do until rs.eof
						 
		                  sawo_empno = rs("sawo_empno")
		                  sawo_emp_name = rs("sawo_emp_name")
		
                         if sawo_empno <> "" then
		                    Sql="select * from emp_master where emp_no = '"&sawo_empno&"'"
		                    Rs_emp.Open Sql, Dbconn, 1

		                   if not Rs_emp.eof then
                              emp_grade = Rs_emp("emp_grade")
		                      emp_position = Rs_emp("emp_position")
		                   end if
	                       Rs_emp.Close()
	                	 end if		
						%>
							<tr>
								<td class="first"><%=rs("sawo_empno")%></td>
                                <td>
                                <a href="insa_card00.asp?emp_no=<%=rs("sawo_empno")%>&be_pg=<%=be_pg%>&page=<%=page%>&view_sort=<%=view_sort%>&date_sw=<%=date_sw%>&page_cnt=<%=page_cnt%>"><%=rs("sawo_emp_name")%></a>
								</td>
                                <td><%=emp_grade%>&nbsp;</td>
                                <td><%=emp_position%>&nbsp;</td>
                                <td><%=rs("sawo_company")%>&nbsp;</td>
                                <td><%=rs("sawo_org_name")%>&nbsp;</td>
                                <td><%=rs("sawo_date")%>&nbsp;</td>
                                <td><%=rs("sawo_id")%>&nbsp;</td>
                                <td><%=rs("sawo_out_date")%>&nbsp;</td>
                                <td><%=rs("sawo_out")%>&nbsp;</td>
                                <% If rs("sawo_target") = "Y" then sawo_target = "공제" end if %>
                                <% If rs("sawo_target") = "N" then sawo_target = "안함" end if %>
								<td><%=sawo_target%>&nbsp;</td>
                                <td style="text-align:right">
                                <a href="#" onClick="pop_Window('insa_sawo_in_view.asp?emp_no=<%=rs("sawo_empno")%>&emp_name=<%=rs("sawo_emp_name")%>&be_pg=<%=be_pg%>&page=<%=page%>&view_sort=<%=view_sort%>&page_cnt=<%=page_cnt%>','sawo_inview','scrollbars=yes,width=800,height=400')"><%=rs("sawo_in_count")%></a>
								</td>
                                <td style="text-align:right"><%=formatnumber(clng(rs("sawo_in_pay")),0)%>&nbsp;</td>
                                <td style="text-align:right">
                                <a href="#" onClick="pop_Window('insa_sawo_give_view.asp?emp_no=<%=rs("sawo_empno")%>&emp_name=<%=rs("sawo_emp_name")%>&be_pg=<%=be_pg%>&page=<%=page%>&view_sort=<%=view_sort%>&page_cnt=<%=page_cnt%>','sawo_inview','scrollbars=yes,width=1000,height=400')"><%=rs("sawo_give_count")%></a>
                                </td>
                                <td style="text-align:right"><%=formatnumber(clng(rs("sawo_give_pay")),0)%>&nbsp;</td>
                                <td colspan="3"><a href="#" onClick="pop_Window('insa_org_reg.asp?sawo_empno=<%=rs("sawo_empno")%>&emp_name=<%=rs("sawo_emp_name")%>&u_type=<%="U"%>','insa_org_reg_pop','scrollbars=yes,width=1400,height=600')">해지</a>&nbsp;</td>
 							</tr>
						<%
							rs.movenext()
						loop
						rs.close()
						%>
                        
                        	<tr>
                              <th colspan="2">총계</th>
                              <th colspan="2"></th>
                              <th>총 납입액 :</th>
                              <th class="right"><%=formatnumber(clng(in_pay_sum),0)%></th>
                              <th colspan="2">   </th>
                              <th colspan="2">총 납입액 :</th>
                              <th colspan="2" class="right"><%=formatnumber(clng(give_pay_sum),0)%></th>
                              <th>   </th>
                              <th>잔 고 :</th>
                              <th colspan="2" class="right"><%=formatnumber(clng(in_pay_sum-give_pay_sum),0)%></th>
                              <th colspan="2"></th>
							</tr>
                    
						</tbody>
					</table>
				</div>
				<%
                intstart = (int((page-1)/10)*10) + 1
                intend = intstart + 9
                first_page = 1
                
                if intend > total_page then
                    intend = total_page
                end if
                %>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td width="15%">
					<div class="btnCenter">
                    <a href="insa_excel_sawo.asp" class="btnType04">엑셀다운로드</a>
					</div>                  
                  	</td>
				    <td>
                    <div id="paging">
                        <a href="insa_reference_mg.asp?page=<%=first_page%>&view_sort=<%=view_sort%>">[처음]</a>
                  	<% if intstart > 1 then %>
                        <a href="insa_reference_mg.asp?page=<%=intstart -1%>&view_sort=<%=view_sort%>">[이전]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
                  	<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="insa_reference_mg.asp?page=<%=i%>&view_sort=<%=view_sort%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
                  	<% if 	intend < total_page then %>
                        <a href="insa_reference_mg.asp?page=<%=intend+1%>&view_sort=<%=view_sort%>">[다음]</a> <a href="insa_reference_mg.asp?page=<%=total_page%>&view_sort=<%=view_sort%>">[마지막]</a>
                        <%	else %>
                        [다음]&nbsp;[마지막]
                      <% end if %>
                    </div>
                    </td>
                    <td>
					<div class="btnCenter">
                    <a href="insa_sawo_ask_list.asp?be_pg=<%=be_pg%>&view_condi=<%=view_condi%>&condi=<%=condi%>&ck_sw=<%="y"%>" class="btnType04">신원보증 등록</a>
					</div>                  
                    </td>                    
			      </tr>
				  </table>
			</form>
		</div>				
	</div>        				
		<input type="hidden" name="user_id">
		<input type="hidden" name="pass">
	</body>
</html>

