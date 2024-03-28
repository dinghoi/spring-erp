<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows

be_pg = "insa_gun_list.asp"

Page=Request("page")

from_date=Request.form("from_date")
to_date=Request.form("to_date")

Page=Request("page")
view_condi = request("view_condi")

ck_sw=Request("ck_sw")

if ck_sw = "n" then
	view_condi = request.form("view_condi")
	from_date=Request.form("from_date")
    to_date=Request.form("to_date")
  else
	view_condi = request("view_condi")
	from_date=request("from_date")
    to_date=request("to_date")
end if

if view_condi = "" then
	view_condi = "전체"
	curr_dd = cstr(datepart("d",now))
	to_date = mid(cstr(now()),1,10)
	from_date = mid(cstr(now()-curr_dd+1),1,10)
end if

rever_yyyy = mid(cstr(from_date),1,4) '귀속년월


pgsize = 10 ' 화면 한 페이지 
If Page = "" Then
	Page = 1
	start_page = 1
End If

stpage = int((page - 1) * pgsize)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

'Sql = "SELECT * FROM k1_memb where "+condi_sql+"mg_group = '"+mg_group+"' ORDER BY user_name ASC"
'where_sql = " WHERE isNull(emp_end_date) or emp_end_date = '1900-01-01'"

if view_condi = "전체" then
       Sql = "SELECT count(*) FROM emp_year_leave WHERE year_year='" + rever_yyyy + "' and (year_empno < '900000')"	
   else
       Sql = "SELECT count(*) FROM emp_year_leave WHERE year_year='" + rever_yyyy + "' and year_company = '" + view_condi + "' and (year_empno < '900000')"	
end if
Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF


if view_condi = "전체" then
       Sql = "SELECT * FROM emp_year_leave WHERE year_year='" + rever_yyyy + "' and (year_empno < '900000') ORDER BY year_company,year_bonbu,year_saupbu,year_team,year_org_code ASC limit "& stpage & "," &pgsize 	
   else
       Sql = "SELECT * FROM emp_year_leave WHERE year_year='" + rever_yyyy + "' and year_company = '" + view_condi + "' and (year_empno < '900000') ORDER BY year_company,year_bonbu,year_saupbu,year_team,year_org_code ASC limit "& stpage & "," &pgsize 	
end if
Rs.Open Sql, Dbconn, 1

title_line = "개인별 근태 현황 "
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
				return "3 1";
			}
		</script>
		<script type="text/javascript">
			$(function() {    $( "#datepicker" ).datepicker();
												$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker" ).datepicker("setDate", "<%=from_date%>" );
			});	  
			$(function() {    $( "#datepicker1" ).datepicker();
												$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker1" ).datepicker("setDate", "<%=to_date%>" );
			});	
			function frmcheck () {
				if (formcheck(document.frm)) {
					document.frm.submit ();
				}
			}			
			function delcheck () {
				if (form_chk(document.frm_del)) {
					document.frm_del.submit ();
				}
			}			

			function form_chk(){				
				a=confirm('삭제하시겠습니까?')
				if (a==true) {
					return true;
				}
				return false;
			}//-->
		</script>

	</head>
	<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
		<div id="wrap">			
			<!--#include virtual = "/include/insa_header.asp" -->
			<!--#include virtual = "/include/insa_gun_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_gun_list.asp?ck_sw=<%="n"%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>조건 검색</dt>
                        <dd>
                            <p>
                               <strong>회사 : </strong>
                              <%
								Sql="select * from emp_org_mst where (isNull(org_end_date) or org_end_date = '1900-01-01') and (org_level = '회사') ORDER BY org_code ASC"
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
								<label>
								<strong>시작일 : </strong>
                                	<input name="from_date" type="text" value="<%=from_date%>" style="width:70px" id="datepicker">
								</label>
								<label>
								<strong>종료일 : </strong>
                                	<input name="to_date" type="text" value="<%=to_date%>" style="width:70px" id="datepicker1">
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser1.jpg" alt="검색"></a>
                                </label>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="9%" >
							<col width="*" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
                            <col width="5%" >
							<col width="5%" >
							<col width="5%" >
							<col width="5%" >
                            <col width="5%" >
                            <col width="5%" >
                            <col width="5%" >
                            <col width="5%" >
                            <col width="5%" >
                            <col width="5%" >
                            <col width="5%" >
                            <col width="5%" >
						</colgroup>
						<thead>
						    <tr>
				                <th rowspan="2" class="first" scope="col" style=" border-left:1px solid #e3e3e3;">성명</th>
                                <th rowspan="2" scope="col" style=" border-left:1px solid #e3e3e3;">소속</th>
                                <th rowspan="2" scope="col" style=" border-left:1px solid #e3e3e3;">발생<br>연차</th>
                                <th rowspan="2" scope="col" style=" border-left:1px solid #e3e3e3;">사용<br>연차</th>
                                <th rowspan="2" scope="col" style=" border-left:1px solid #e3e3e3;">잔여<br>연차</th>
                                <th colspan="5" scope="col" style=" border-bottom:1px solid #e3e3e3;">휴&nbsp;&nbsp;&nbsp;가</th>
				                <th colspan="7" scope="col" style=" border-bottom:1px solid #e3e3e3;">근&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;태</th>
			                </tr>
                            <tr>
								<th class="first" scope="col" style=" border-left:1px solid #e3e3e3;">연차</th>
								<th scope="col">반차</th>
								<th scope="col">대휴</th>
								<th scope="col">공가</th>
								<th scope="col">정기<br>휴가</th>
                                <th scope="col">시간외<br>근무</th>
                                <th scope="col">휴일<br>근무</th>
                                <th scope="col">외근</th>
                                <th scope="col">출장</th>
                                <th scope="col">조퇴</th>
                                <th scope="col">결근</th>
                                <th scope="col">기타</th>
							</tr>
						</thead>
						<tbody>
						<%
						do until rs.eof
                                yun_cnt = 0
	           			%>
							<tr>
                                <td class="first"><%=rs("year_emp_name")%>(<%=rs("year_empno")%>)&nbsp;</td>
                                <td><%=rs("year_org_name")%>&nbsp;</td>
                                <td><%=rs("year_basic_count")%>&nbsp;</td>
                                <td><%=rs("year_use_count")%>&nbsp;</td>
                                <td><%=rs("year_remain_count")%>&nbsp;</td>
                                <td><%=yun_cnt%>&nbsp;</td>
                                <td><%=yun_cnt%>&nbsp;</td>
                                <td><%=yun_cnt%>&nbsp;</td>
                                <td><%=yun_cnt%>&nbsp;</td>
                                <td><%=yun_cnt%>&nbsp;</td>
                                <td><%=yun_cnt%>&nbsp;</td>
                                <td><%=yun_cnt%>&nbsp;</td>
                                <td><%=yun_cnt%>&nbsp;</td>
                                <td><%=yun_cnt%>&nbsp;</td>
                                <td><%=yun_cnt%>&nbsp;</td>
                                <td><%=yun_cnt%>&nbsp;</td>
                                <td><%=yun_cnt%>&nbsp;</td>
							</tr>
						<%
							rs.movenext()
						loop
						rs.close()
						%>
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
				    <td>
                    <div id="paging">
                        <a href = "insa_gun_list.asp?page=<%=first_page%>&view_condi=<%=view_condi%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[처음]</a>
                  	<% if intstart > 1 then %>
                        <a href="insa_gun_list.asp?page=<%=intstart -1%>&view_condi=<%=view_condi%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[이전]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
           	<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="insa_gun_list.asp?page=<%=i%>&view_condi=<%=view_condi%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
           	<% if 	intend < total_page then %>
                        <a href="insa_gun_list.asp?page=<%=intend+1%>&view_condi=<%=view_condi%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[다음]</a> <a href="insa_gun_list.asp?page=<%=total_page%>&view_condi=<%=view_condi%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[마지막]</a>
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
	</body>
</html>

