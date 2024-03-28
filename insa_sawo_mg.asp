<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows
dim page_cnt
dim pg_cnt

insa_grade = request.cookies("nkpmg_user")("coo_insa_grade")

Page=Request("page")
page_cnt=Request.form("page_cnt")
pg_cnt=cint(Request("pg_cnt"))
be_pg = "insa_sawo_mg.asp"
curr_date = datevalue(mid(cstr(now()),1,10))

ck_sw=Request("ck_sw")
If ck_sw = "y" Then
	view_condi=Request("view_condi")
  else
	view_condi=Request.form("view_condi")
End if

If view_condi = "" Then
	view_condi = "전체"
End If

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
Set rs_org = Server.CreateObject("ADODB.Recordset")
Set rs_sum = Server.CreateObject("ADODB.Recordset")
Set rs_etc = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")

dbconn.open DbConnect

view_sort = request("view_sort")

if view_sort = "" then
	view_sort = "ASC"
end if


order_Sql = " ORDER BY sawo_empno " + view_sort
if view_condi = "전체" then
         'where_sql = " WHERE sawo_target = 'Y' or sawo_target = 'N'"
		 where_sql = " "
   else
         where_sql = " WHERE sawo_company = '"+view_condi+"'"
end if
'where_sql = ""

    in_pay_sum = 0 
	give_pay_sum = 0
	k1_in_pay_sum = 0 
	k1_give_pay_sum = 0
	hd_in_pay_sum = 0 
	hd_give_pay_sum = 0
	kn_in_pay_sum = 0 
	kn_give_pay_sum = 0
	su_in_pay_sum = 0 
	su_give_pay_sum = 0
	ko_in_pay_sum = 0 
	ko_give_pay_sum = 0
	
	
    sql="select * from emp_sawo_mem " + where_sql
	Rs_sum.Open Sql, Dbconn, 1
	
	do until rs_sum.eof
	   in_pay_sum = in_pay_sum + rs_sum("sawo_in_pay")
	   give_pay_sum = give_pay_sum + rs_sum("sawo_give_pay")
	   if  rs_sum("sawo_company") = "케이원정보통신" then
	          k1_in_pay_sum = k1_in_pay_sum + rs_sum("sawo_in_pay")
	          k1_give_pay_sum = k1_give_pay_sum + rs_sum("sawo_give_pay")
		   elseif  rs_sum("sawo_company") = "휴디스" then
		              hd_in_pay_sum = hd_in_pay_sum + rs_sum("sawo_in_pay")
	                  hd_give_pay_sum = hd_give_pay_sum + rs_sum("sawo_give_pay")
				   elseif  rs_sum("sawo_company") = "케이네트웍스" then
		                      kn_in_pay_sum = kn_in_pay_sum + rs_sum("sawo_in_pay")
	                          kn_give_pay_sum = kn_give_pay_sum + rs_sum("sawo_give_pay")
						   elseif  rs_sum("sawo_company") = "에스유에이치" then
		                              su_in_pay_sum = su_in_pay_sum + rs_sum("sawo_in_pay")
	                                  su_give_pay_sum = su_give_pay_sum + rs_sum("sawo_give_pay")
								   elseif  rs_sum("sawo_company") = "코리아디엔씨" then
		                                      ko_in_pay_sum = ko_in_pay_sum + rs_sum("sawo_in_pay")
	                                          ko_give_pay_sum = ko_give_pay_sum + rs_sum("sawo_give_pay")
	   end if
	   
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

title_line =  view_condi + " 경조회 회원 현황 "

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>인사관리 시스템</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "8 1";
			}
		</script>
		<script type="text/javascript">
			function frmcheck () {
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {
				if (document.frm.view_condi.value == "") {
					alert ("소속을 선택하시기 바랍니다");
					return false;
				}	
				return true;
			}
		</script>

	</head>
	<body>
		<div id="wrap">			
			<!--#include virtual = "/include/insa_header.asp" -->
			<!--#include virtual = "/include/insa_sawo_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_sawo_mg.asp" method="post" name="frm">
                
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>회사 검색</dt>
                        <dd>
                            <p>
                               <strong>회사 : </strong>
                              <%
								Sql="select * from emp_org_mst where  (org_level = '회사') ORDER BY org_code ASC"
	                            rs_org.Open Sql, Dbconn, 1	
							  %>
                                <label>
								<select name="view_condi" id="view_condi" type="text" style="width:150px">
                                    <option value="전체" <%If view_condi = "전체" then %>selected<% end if %>>전체</option>
                			  <% 
								do until rs_org.eof 
			  				  %>
                					<option value='<%=rs_org("org_name")%>' <%If view_condi = rs_org("org_name") then %>selected<% end if %>><%=rs_org("org_name")%></option>
                			  <%
									rs_org.movenext()  
								loop 
								rs_org.Close()
							  %>
            					</select>
                                </label>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser1.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>                      
                                
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
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
                            <col width="3%" >
                            <col width="3%" >
                            <col width="2%" >
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
								<th colspan="3" scope="col">비고</th>
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
                                <a href="#" onClick="pop_Window('insa_card00.asp?emp_no=<%=rs("sawo_empno")%>&be_pg=<%=be_pg%>&page=<%=page%>&view_sort=<%=view_sort%>&date_sw=<%=date_sw%>&page_cnt=<%=page_cnt%>','emp_card0_pop','scrollbars=yes,width=1250,height=650')"><%=rs("sawo_emp_name")%></a>
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
                                <a href="#" onClick="pop_Window('insa_sawo_in_view.asp?emp_no=<%=rs("sawo_empno")%>&emp_name=<%=rs("sawo_emp_name")%>','sawo_inview','scrollbars=yes,width=800,height=600')"><%=rs("sawo_in_count")%></a>
								</td>
                                <td style="text-align:right"><%=formatnumber(clng(rs("sawo_in_pay")),0)%>&nbsp;</td>
                                <td style="text-align:right">
                                <a href="#" onClick="pop_Window('insa_sawo_give_view.asp?emp_no=<%=rs("sawo_empno")%>&emp_name=<%=rs("sawo_emp_name")%>','sawo_inview','scrollbars=yes,width=1000,height=600')"><%=rs("sawo_give_count")%></a>
                                </td>
                                <td style="text-align:right"><%=formatnumber(clng(rs("sawo_give_pay")),0)%>&nbsp;</td>
                                <% if user_id = "sanginlee" then %>
                                <td colspan="3"><a href="#" onClick="pop_Window('insa_sawo_giveadd.asp?sawo_empno=<%=rs("sawo_empno")%>&emp_name=<%=rs("sawo_emp_name")%>&u_type=<%=""%>','insa_sawo_giveadd_pop','scrollbars=yes,width=750,height=500')">경조금지급</a>&nbsp;</td>
                                <%    else %>
                                <td colspan="3"><%=rs("sawo_out")%>&nbsp;</td>
                                <% end if %>
 							</tr>
						<%
							rs.movenext()
						loop
						rs.close()
						%>
                        	<tr>
                              <th colspan="4"><%=view_condi%>&nbsp;총계</th>
                              <th>총 납입액 :</th>
                              <th class="right"><%=formatnumber(clng(in_pay_sum),0)%></th>
                              <th colspan="2">&nbsp;</th>
                              <th colspan="2">총 지급액 :</th>
                              <th colspan="2" class="right"><%=formatnumber(clng(give_pay_sum),0)%></th>
                              <th>&nbsp;</th>
                              <th>잔 고 :</th>
                              <th colspan="2" class="right"><%=formatnumber(clng(in_pay_sum-give_pay_sum),0)%></th>
                              <th colspan="2">&nbsp;</th>
							</tr>
                        <% if view_condi = "전체" then %>
                            <tr>
                              <td colspan="4" class="right">케이원정보통신&nbsp;&nbsp;</td>
                              <td>납입액 :</td>
                              <td class="right"><%=formatnumber(clng(k1_in_pay_sum),0)%></td>
                              <td colspan="2">&nbsp;</td>
                              <td colspan="2">지급액 :</td>
                              <td colspan="2" class="right"><%=formatnumber(clng(k1_give_pay_sum),0)%></td>
                              <td>&nbsp;</td>
                              <td>잔 고 :</td>
                              <td colspan="2" class="right"><%=formatnumber(clng(k1_in_pay_sum - k1_give_pay_sum),0)%></td>
                              <td colspan="2">&nbsp;</td>
							</tr>
                            <tr>
                              <td colspan="4" class="right">휴디스&nbsp;&nbsp;</td>
                              <td>납입액 :</td>
                              <td class="right"><%=formatnumber(clng(hd_in_pay_sum),0)%></td>
                              <td colspan="2">&nbsp;</td>
                              <td colspan="2">지급액 :</td>
                              <td colspan="2" class="right"><%=formatnumber(clng(hd_give_pay_sum),0)%></td>
                              <td>&nbsp;</td>
                              <td>잔 고 :</td>
                              <td colspan="2" class="right"><%=formatnumber(clng(hd_in_pay_sum - hd_give_pay_sum),0)%></td>
                              <td colspan="2">&nbsp;</td>
							</tr>
                            <tr>
                              <td colspan="4" class="right">케이네트웍스&nbsp;&nbsp;</td>
                              <td>납입액 :</td>
                              <td class="right"><%=formatnumber(clng(kn_in_pay_sum),0)%></td>
                              <td colspan="2">&nbsp;</td>
                              <td colspan="2">지급액 :</td>
                              <td colspan="2" class="right"><%=formatnumber(clng(kn_give_pay_sum),0)%></td>
                              <td>&nbsp;</td>
                              <td>잔 고 :</td>
                              <td colspan="2" class="right"><%=formatnumber(clng(kn_in_pay_sum - kn_give_pay_sum),0)%></td>
                              <td colspan="2">&nbsp;</td>
							</tr>
                            <tr>
                              <td colspan="4" class="right">에스유에이치&nbsp;&nbsp;</td>
                              <td>납입액 :</td>
                              <td class="right"><%=formatnumber(clng(su_in_pay_sum),0)%></td>
                              <td colspan="2">&nbsp;</td>
                              <td colspan="2">지급액 :</td>
                              <td colspan="2" class="right"><%=formatnumber(clng(su_give_pay_sum),0)%></td>
                              <td>&nbsp;</td>
                              <td>잔 고 :</td>
                              <td colspan="2" class="right"><%=formatnumber(clng(su_in_pay_sum - su_give_pay_sum),0)%></td>
                              <td colspan="2">&nbsp;</td>
							</tr>
                            <tr>
                              <td colspan="4" class="right">코리아디엔씨&nbsp;&nbsp;</td>
                              <td>납입액 :</td>
                              <td class="right"><%=formatnumber(clng(ko_in_pay_sum),0)%></td>
                              <td colspan="2">&nbsp;</td>
                              <td colspan="2">지급액 :</td>
                              <td colspan="2" class="right"><%=formatnumber(clng(ko_give_pay_sum),0)%></td>
                              <td>&nbsp;</td>
                              <td>잔 고 :</td>
                              <td colspan="2" class="right"><%=formatnumber(clng(ko_in_pay_sum - ko_give_pay_sum),0)%></td>
                              <td colspan="2">&nbsp;</td>
							</tr>
                        <% end if %>
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
                    <a href="insa_excel_sawo.asp?view_condi=<%=view_condi%>" class="btnType04">엑셀다운로드</a>
					</div>                  
                  	</td>
				    <td>
                    <div id="paging">
                        <a href="insa_sawo_mg.asp?page=<%=first_page%>&view_sort=<%=view_sort%>&view_condi=<%=view_condi%>&ck_sw=<%="y"%>">[처음]</a>
                  	<% if intstart > 1 then %>
                        <a href="insa_sawo_mg.asp?page=<%=intstart -1%>&view_sort=<%=view_sort%>&view_condi=<%=view_condi%>&ck_sw=<%="y"%>">[이전]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
                  	<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="insa_sawo_mg.asp?page=<%=i%>&view_sort=<%=view_sort%>&view_condi=<%=view_condi%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
                  	<% if 	intend < total_page then %>
                        <a href="insa_sawo_mg.asp?page=<%=intend+1%>&view_sort=<%=view_sort%>&view_condi=<%=view_condi%>&ck_sw=<%="y"%>">[다음]</a> <a href="insa_sawo_mg.asp?page=<%=total_page%>&view_sort=<%=view_sort%>&view_condi=<%=view_condi%>&ck_sw=<%="y"%>">[마지막]</a>
                        <%	else %>
                        [다음]&nbsp;[마지막]
                      <% end if %>
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

