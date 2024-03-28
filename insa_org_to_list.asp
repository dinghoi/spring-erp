<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows

be_pg = "insa_org_to_list.asp"

Page=Request("page")
in_company = request.form("in_company")

ck_sw=Request("ck_sw")

if ck_sw = "n" then
	in_company = request.form("in_company")
  else
	in_company = request("in_company")
end if

if in_company = "" then
	in_company = "케이원정보통신"
	condi_sql = " "
end if

pgsize = 10 ' 화면 한 페이지 
If Page = "" Then
	Page = 1
	start_page = 1
End If

stpage = int((page - 1) * pgsize)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

'order_Sql = " GROUP BY emp_org_code Order By emp_org_code Asc "
order_Sql = " GROUP BY emp_org_name Order By emp_org_name Asc "
where_sql = " where (isNull(emp_end_date) or emp_end_date = '1900-01-01') and (emp_no < '900000') "

if in_company = "전체" then 
      condi_sql = " "
   else	  
	  condi_sql = " and (emp_company = '"+in_company+"')"
end if

'sql = "SELECT count(*) from (SELECT emp_org_code, count(*) as emp_cnt FROM emp_master (isNull(emp_end_date) or emp_end_date = '1900-01-01') and (emp_company = '"+in_company+"') GROUP BY emp_org_code Order By emp_org_code Asc) as b_org"

Sql = "SELECT emp_org_name, count(*) as emp_cnt FROM emp_master " + where_sql + condi_sql + order_Sql
Set RsCount = Dbconn.Execute (sql)
k = 0
do until RsCount.eof
    k = k + 1
	RsCount.movenext()
loop
RsCount.close()

'tottal_record = cint(RsCount(0)) 'Result.RecordCount
tottal_record = cint(k)

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

'sql = "select * from emp_org_mst " + where_sql + condi_sql + order_sql + " limit "& stpage & "," &pgsize 
' emp_org_code 로 sum을 해야함..

   sql = "select emp_org_name, count(*) as emp_cnt from emp_master"

if in_company = "전체" then 
      sql = sql + " where (isNull(emp_end_date) or emp_end_date = '1900-01-01') and (emp_no < '900000')"
   else	  
	  sql = sql + " where (isNull(emp_end_date) or emp_end_date = '1900-01-01') and (emp_no < '900000') and (emp_company = '"+in_company+"')"
end if

   sql = sql + " GROUP BY emp_org_name Order By emp_bonbu,emp_saupbu,emp_team,emp_org_name Asc limit "& stpage & "," &pgsize 

   Rs.Open Sql, Dbconn, 1

title_line = "조직별 T.O 현황 "
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
				return "0 1";
			}
		</script>
		<script type="text/javascript">
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
			<!--#include virtual = "/include/insa_org_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_org_to_list.asp?ck_sw=<%="n"%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>조건 검색</dt>
                        <dd>
                            <p>
                                <strong>회사</strong>
							  	<%
									Sql="select * from emp_org_mst where (isNull(org_end_date) or org_end_date = '1900-01-01') and (org_level = '회사') ORDER BY org_code ASC"
                                    rs_org.Open Sql, Dbconn, 1
                                %>
        						<select name="in_company" id="in_company" type="text" style="width:150px">
                                    <option value="전체" <%If in_company = "전체" then %>selected<% end if %>>전체</option>

          					<% 
								While not rs_org.eof 
							%>
          							<option value='<%=rs_org("org_name")%>' <%If rs_org("org_name") = in_company  then %>selected<% end if %>><%=rs_org("org_name")%></option>
          					<%
									rs_org.movenext()  
								Wend 
								rs_org.Close()
							%>
        						</select>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser1.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="10%" >
							<col width="10%" >
							<col width="8%" >
							<col width="8%" >
                            <col width="15%" >
                            <col width="15%" >
							<col width="*" >
                            <col width="11%" >
						</colgroup>
						<thead>
							<tr>
								<th scope="col" class="first">조직명</th>
                                <th scope="col">조직장</th>
								<th scope="col">T.O</th>
								<th scope="col">현인원</th>
								<th colspan="3" scope="col">소&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;속</th>
                                <th scope="col">상주처 회사</th>
							</tr>
						</thead>
						<tbody>
						<%
						do until rs.eof
						   emp_cnt = cint(rs("emp_cnt"))
						   emp_org_name = rs("emp_org_name")
						   
						   org_table_org = 0
						   
						   sql = "select * from emp_org_mst where  org_company = '케이원정보통신' and org_name = '"+emp_org_name+"'"
                           Rs_org.Open Sql, Dbconn, 1 
                           if not Rs_org.eof then
                                org_name = rs_org("org_name")
							    org_emp_name = rs_org("org_emp_name")
							'    org_table_org = rs_org("org_table_org")
							'    org_company = rs_org("org_company")
							    org_bonbu = rs_org("org_bonbu")
							    org_saupbu = rs_org("org_saupbu")
							    org_team = rs_org("org_team")
								org_reside_company = rs_org("org_reside_company")
	                          else
                                org_name = emp_org_name
							    org_emp_name = ""
							'    org_table_org = 0
							'    org_company = ""
							    org_bonbu = ""
							    org_saupbu = ""
							    org_team = ""
								org_reside_company = ""
                            end if
                            Rs_org.close()						   
	           			%>
							<tr>
                                <td class="first"><%=org_name%>&nbsp;</td>
                                <td><%=org_emp_name%>&nbsp;</td>
                                <td><%=org_table_org%>&nbsp;</td>
                                <td><%=emp_cnt%>&nbsp;</td>
                                <td class="left"><%=org_bonbu%>&nbsp;</td>
                                <td class="left"><%=org_saupbu%>&nbsp;</td>
                                <td class="left"><%=org_team%>&nbsp;</td>
                                <td><%=org_reside_company%>&nbsp;</td>
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
                        <a href = "insa_org_to_list.asp?page=<%=first_page%>&in_company=<%=in_company%>&ck_sw=<%="y"%>">[처음]</a>
                  	<% if intstart > 1 then %>
                        <a href="insa_org_to_list.asp?page=<%=intstart -1%>&in_company=<%=in_company%>&ck_sw=<%="y"%>">[이전]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
           	<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="insa_org_to_list.asp?page=<%=i%>&in_company=<%=in_company%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
           	<% if 	intend < total_page then %>
                        <a href="insa_org_to_list.asp?page=<%=intend+1%>&in_company=<%=in_company%>&ck_sw=<%="y"%>">[다음]</a> <a href="insa_org_to_list.asp?page=<%=total_page%>&in_company=<%=in_company%>&ck_sw=<%="y"%>">[마지막]</a>
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

