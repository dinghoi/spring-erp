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
be_pg = "insa_org.asp"
curr_date = datevalue(mid(cstr(now()),1,10))

ck_sw=Request("ck_sw")

If ck_sw = "y" Then
	view_condi=Request("view_condi")
  else
	view_condi=Request.form("view_condi")
End if

If view_condi = "" Then
	view_condi = "케이원정보통신"
End If

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

order_Sql = " ORDER BY org_company,org_bonbu,org_saupbu,org_team,org_name ASC"
where_sql = " WHERE (isNull(org_end_date) or org_end_date = '1900-01-01') and (org_company = '"&view_condi&"')"

Sql = "SELECT count(*) FROM emp_org_mst " + where_sql
Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

sql = "select * from emp_org_mst " + where_sql + order_sql + " limit "& stpage & "," &pgsize 
Rs.Open Sql, Dbconn, 1
'response.write(sql)

title_line = " 조직 현황 "

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
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {
				if (document.frm.view_condi.value == "") {
					alert ("필드조건을 선택하시기 바랍니다");
					return false;
				}	
				return true;
			}
		</script>

	</head>
	<body>
		<div id="wrap">			
			<!--#include virtual = "/include/insa_header.asp" -->
			<!--#include virtual = "/include/insa_org_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_org.asp" method="post" name="frm">
                
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>회사 검색</dt>
                        <dd>
                            <p>
                               <strong>회사 : </strong>
                              <%
								Sql="select * from emp_org_mst where (isNull(org_end_date) or org_end_date = '1900-01-01') and (org_level = '회사') ORDER BY org_code ASC"
	                            rs_org.Open Sql, Dbconn, 1	
							  %>
                                <label>
								<select name="view_condi" id="view_condi" type="text" style="width:150px">

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
				      <col width="4%" >
				      <col width="11%" >
                      <col width="4%" >
				      <col width="6%" >
				      <col width="8%" >
                      <col width="10%" >
				      <col width="10%" >
				      <col width="10%" >
				      <col width="10%" >
				      <col width="8%" >
                      <col width="6%" >
				      <col width="8%" >
                      <col width="3%" >
			        </colgroup>
				    <thead>
				      <tr>
				        <th colspan="3" class="first" scope="col" style=" border-bottom:1px solid #e3e3e3;">조&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;직</th>
				        <th colspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">조&nbsp;&nbsp;직&nbsp;&nbsp;장</th>
                        <th colspan="4" scope="col" style=" border-bottom:1px solid #e3e3e3;">소&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;속</th>
				        <th rowspan="2" scope="col">조직생성일</th>
				        <th colspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">상위&nbsp;조직장</th>
                        <th rowspan="2" scope="col">수정</th>
			          </tr>
                      <tr>
				        <th class="first"scope="col">코드</th>
				        <th scope="col">조직명</th>
                        <th scope="col">T.O</th>
				        <th scope="col">사번</th>
				        <th scope="col">성명</th>
                        <th scope="col">회&nbsp;&nbsp;사</th>
				        <th scope="col">본&nbsp;&nbsp;부</th>
				        <th scope="col">사업부</th>
				        <th scope="col">팀</th>
				        <th scope="col">사번</th>
                        <th scope="col">성명</th>
                      </tr>
			        </thead>
				    <tbody>
                      <%
						do until rs.eof
					  %>
				      <tr>
				        <td class="first"><%=rs("org_code")%>&nbsp;</td>
                        <td><a href="#" onClick="pop_Window('insa_org_view.asp?org_code=<%=rs("org_code")%>&org_name=<%=org_name%>&u_type=<%="U"%>','insa_org_view_pop','scrollbars=yes,width=750,height=350')"><%=rs("org_name")%></a>&nbsp;</td>
                        <td><%=rs("org_table_org")%>&nbsp;</td>
                        <td><%=rs("org_empno")%>&nbsp;</td>
                        <td>
                        <a href="#" onClick="pop_Window('insa_card00.asp?emp_no=<%=rs("org_empno")%>&be_pg=<%=be_pg%>&page=<%=page%>&view_sort=<%=view_sort%>&date_sw=<%=date_sw%>&page_cnt=<%=page_cnt%>','emp_card0_pop','scrollbars=yes,width=1250,height=650')"><%=rs("org_emp_name")%></a>
						</td>
                        <td><%=rs("org_company")%>&nbsp;</td>
				        <td><%=rs("org_bonbu")%>&nbsp;</td>
                        <td><%=rs("org_saupbu")%>&nbsp;</td>
                        <td><%=rs("org_team")%>&nbsp;</td>
                        <td><%=rs("org_date")%>&nbsp;</td>
                        <td><%=rs("org_owner_empno")%>&nbsp;</td>
                        <td><%=rs("org_owner_empname")%>&nbsp;</td>
                        <td><a href="#" onClick="pop_Window('insa_org_modify.asp?org_code=<%=rs("org_code")%>&u_type=<%="U"%>','insa_org_modi_pop','scrollbars=yes,width=1250,height=400')">수정</a>&nbsp;</td>
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
				    <td width="20%">
					<div class="btnCenter">
                    <a href="insa_excel_org.asp?view_condi=<%=view_condi%>" class="btnType04">엑셀다운로드</a>
					</div>                  
                  	</td>
				    <td>
                    <div id="paging">
                        <a href="insa_org.asp?page=<%=first_page%>&view_condi=<%=view_condi%>&ck_sw=<%="y"%>">[처음]</a>
                  	<% if intstart > 1 then %>
                        <a href="insa_org.asp?page=<%=intstart -1%>&view_condi=<%=view_condi%>&ck_sw=<%="y"%>">[이전]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
   	        <% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="insa_org.asp?page=<%=i%>&view_condi=<%=view_condi%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
   	        <% if 	intend < total_page then %>
                        <a href="insa_org.asp?page=<%=intend+1%>&view_condi=<%=view_condi%>&ck_sw=<%="y"%>">[다음]</a> <a href="insa_org.asp?page=<%=total_page%>&view_condi=<%=view_condi%>&ck_sw=<%="y"%>">[마지막]</a>
                        <%	else %>
                        [다음]&nbsp;[마지막]
                      <% end if %>
                    </div>
                    </td>
				    <td width="20%">
					<div class="btnCenter">
                    <a href="#" onClick="pop_Window('insa_org_reg.asp?view_condi=<%=view_condi%>','insa_org_reg_popup','scrollbars=yes,width=1250,height=400')" class="btnType04">신규조직등록</a>
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

