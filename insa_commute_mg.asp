<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows

be_pg = "insa_commute_mg.asp"

Page=Request("page")
from_date=Request.form("from_date")
to_date=Request.form("to_date")

Page=Request("page")
org_company = Request.form("org_company")
org_name = Request.form("org_name")

'Response.write org_company
'Response.write org_name

ck_sw=Request("ck_sw")

	org_company = request.form("org_company")
	org_name = request.form("org_name")
	from_date=Request.form("from_date")
  to_date=Request.form("to_date")

if org_company = "" then
	org_company = "전체"
end if
if org_name = "" then
	org_name = "전체"
end if

	curr_dd = cstr(datepart("d",now))
if to_date = "" then	
	to_date = mid(cstr(now()),1,10)
end if
if from_date = "" then	
	from_date = mid(cstr(now()-curr_dd+1),1,10)
end if

pgsize = 10 ' 화면 한 페이지 
If Page = "" Then
	Page = 1
	start_page = 1
End If

stpage = int((page - 1) * pgsize)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

if org_company = "전체" then
       Sql = "select count(*) from commute where wrkt_dt between '"&from_date&"' and '"&to_date&"'"	
else
       Sql = "select count(*) from commute left outer join emp_master on commute.emp_no = emp_master.emp_no"
       Sql = Sql + " where commute.wrkt_dt between '"&from_date&"' and '"&to_date&"'"	
       
       if org_name = "전체" then
          Sql = Sql + " and emp_master.emp_company = '"&org_company&"'"	
       else
          Sql = Sql + " and emp_master.emp_company = '"&org_company&"' and emp_master.emp_saupbu = '"&org_name&"'"	
       end if
'       Sql = Sql + " and emp_master.org_company = '케이원정보통신' and emp_org_mst.org_name = '"&view_condi&"'"	
end if
'response.write sql
Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

if org_company = "전체" then
       Sql = "select * from commute where wrkt_dt between '"&from_date&"' and '"&to_date&"' ORDER BY wrkt_dt ASC, emp_no ASC, wrk_start_time ASC limit "& stpage & "," &pgsize	
   else
       Sql = "SELECT commute.* from commute left outer join emp_master on commute.emp_no = emp_master.emp_no"
       Sql = Sql + " where commute.wrkt_dt between '"&from_date&"' and '"&to_date&"'"	
       Sql = Sql + " and emp_master.emp_company = '"&org_company&"'"	
      if org_name <> "전체" then
       Sql = Sql + " and emp_master.emp_saupbu = '"&org_name&"'"	
      end if
       Sql = Sql + " ORDER BY wrkt_dt ASC, emp_no ASC, wrk_start_time ASC limit "& stpage & "," &pgsize	
end if
'Response.write Sql
Rs.Open Sql, Dbconn, 1

title_line = "조직별 출퇴근 현황 "
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
				//if (formcheck(document.frm)) {
				//	document.frm.submit ();
				//}
//				alert($(".view_org_company option:selected").val());

				document.frm.org_company.value = $(".view_org_company option:selected").val();
				document.frm.org_name.value = $(".view_org_name option:selected").val();

//				alert($(".view_org_company option:selected").val());
				document.frm.submit ();
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
		<script>
			$(function() {
			    $(".view_org_company").change(function(){
			        document.frm.org_company.value = $(".view_org_company option:selected").val();
							document.frm.org_name.value = $(".view_org_name option:selected").val();

							document.frm.submit ();
			    })
			});
			$(function() {
			    $(".view_org_name").change(function(){
			        //alert($(".view_org_name option:selected").val());
			        document.frm.org_company.value = $(".view_org_company option:selected").val();
							document.frm.org_name.value = $(".view_org_name option:selected").val();

							document.frm.submit ();
			    })
			});
		</script>

	</head>
	<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
		<div id="wrap">			
			<!--#include virtual = "/include/insa_header.asp" -->
			<!--#include virtual = "/include/insa_gun_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_commute_mg.asp?ck_sw=<%="n"%>" method="post" name="frm">
					
				<input name="org_company" type="hidden" value="<%="org_company"%>">
				<input name="org_name" type="hidden" value="<%="org_name"%>">
				
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>조건 검색</dt>
                        <dd>
                            <p>
                            	<strong>회사 </strong>
                              <%
								Sql="select distinct org_company from emp_org_mst where (isNull(org_end_date) or org_end_date = '1900-01-01') and (org_level = '회사') and org_company IN ('케이네트웍스', '케이원정보통신', '코리아디엔씨') and length(trim(org_company)) > 0  ORDER BY org_company ASC"
	                            rs_org.Open Sql, Dbconn, 1	
							  %>
                                <label>
								<select class="view_org_company" name="view_org_company id="view_org_company" type="text" style="width:150px">
                                  <option value="전체" <%If org_company = "전체" then %>selected<% end if %>>전체</option>
                			  <% 
								do until rs_org.eof 
			  				  %>
                					<option value='<%=rs_org("org_company")%>' <%If org_company = rs_org("org_company") then %>selected<% end if %>><%=rs_org("org_company")%></option>
                			  <%
									rs_org.movenext()  
								loop 
								rs_org.Close()
							  %>
            					</select>
                                </label>
                               <strong>부서 </strong>
                              <%
								'Sql="select distinct org_saupbu from emp_org_mst where (isNull(org_end_date) or org_end_date = '1900-01-01') and (org_level = '사업부') and (org_company = '"&org_company&"') and length(trim(org_saupbu)) > 0  ORDER BY org_saupbu ASC"
								Sql="select distinct org_saupbu from emp_org_mst where (org_level = '사업부') and (org_company = '"&org_company&"') and length(trim(org_saupbu)) > 0  ORDER BY org_saupbu ASC"
								
								Response.write Sq
	                            rs_org.Open Sql, Dbconn, 1	
							  %>
                                <label>
								<select class="view_org_name" name="view_org_name" id="view_org_name" type="text" style="width:150px">
                                  <option value="전체" <%If org_name = "전체" then %>selected<% end if %>>전체</option>
                			  <% 
								do until rs_org.eof 
			  				  %>
                					<option value='<%=rs_org("org_saupbu")%>' <%If org_name = rs_org("org_saupbu") then %>selected<% end if %>><%=rs_org("org_saupbu")%></option>
                			  <%
									rs_org.movenext()  
								loop 
								rs_org.Close()
							  %>
            					</select>
                                </label>
								<label>
								<strong>시작일 </strong>
                                	<input name="from_date" type="text" value="<%=from_date%>" style="width:70px" id="datepicker">
								</label>
								<label>
								<strong>종료일 </strong>
                                	<input name="to_date" type="text" value="<%=to_date%>" style="width:70px" id="datepicker1">
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser1.jpg" alt="검색"></a>
                 </label>
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
						</colgroup>
						<thead>
              <tr>
								<th class="first" scope="col">사번</th>
								<th scope="col">성  명</th>
								<th scope="col">직급</th>
								<th scope="col">직책</th>
                <th scope="col">회사</th>
                <th scope="col">소속</th>
								<th scope="col">출근일</th>
								<th scope="col">출근시간</th>
								<th scope="col">근무형태</th>					
							</tr>              
						</thead>
						<tbody>
						<%
						do until rs.eof

                         emp_no = rs("emp_no")
                         if emp_no <> "" then
		                    Sql="select * from emp_master where emp_no = '"&emp_no&"'"
		                    Rs_emp.Open Sql, Dbconn, 1

		                   if not Rs_emp.eof then
                          emp_name = Rs_emp("emp_name")
												  emp_grade = Rs_emp("emp_grade")
												  emp_job = Rs_emp("emp_job")
							            emp_position = Rs_emp("emp_position")
												  emp_org_code = Rs_emp("emp_org_code")
												  emp_org_name = Rs_emp("emp_org_name")
												  emp_company = Rs_emp("emp_company")
		                   end if
	                       Rs_emp.Close()
	                	  end if	

	           			%>
							<tr>
								<td class="first"><%=rs("emp_no")%>&nbsp;</td>
                                <td><%=emp_name%>&nbsp;</td>
                                <td><%=emp_grade%>&nbsp;</td>
                                <td><%=emp_position%>&nbsp;</td>
                                <td><%=emp_company%>&nbsp;</td>
                                <td><%=emp_org_name%>&nbsp;</td>
                                <td><%=rs("wrkt_dt")%>&nbsp;</td>
                                <td><%=rs("wrk_start_time")%>&nbsp;</td>
                                <td><%=rs("wrk_type")%>&nbsp;</td>
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
                        <a href = "insa_commute_mg.asp?page=<%=first_page%>&org_company=<%=org_company%>&org_name=<%=org_name%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[처음]</a>
                  	<% if intstart > 1 then %>
                        <a href="insa_commute_mg.asp?page=<%=intstart -1%>&org_company=<%=org_company%>&org_name=<%=org_name%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[이전]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
           	<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="insa_commute_mg.asp?page=<%=i%>&org_company=<%=org_company%>&org_name=<%=org_name%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
           	<% if 	intend < total_page then %>
                        <a href="insa_commute_mg.asp?page=<%=intend+1%>&org_company=<%=org_company%>&org_name=<%=org_name%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[다음]</a> <a href="insa_gun_mg.asp?page=<%=total_page%>&org_company=<%=org_company%>&org_name=<%=org_name%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[마지막]</a>
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

