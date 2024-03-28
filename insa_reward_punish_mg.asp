<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim win_sw

be_pg = "insa_reward_punish_mg.asp"

user_name = request.cookies("nkpmg_user")("coo_user_name")
user_id = request.cookies("nkpmg_user")("coo_user_id")
insa_grade = request.cookies("nkpmg_user")("coo_insa_grade")

view_condi = request("view_condi")
condi = request("condi")
Page=Request("page")

ck_sw=Request("ck_sw")

if ck_sw = "n" then
	owner_view=Request.form("owner_view")
	view_condi = request.form("view_condi")
	condi = request.form("condi")
  else
	owner_view=request("owner_view")
	view_condi = request("view_condi")
	condi = request("condi")
end if

if view_condi = "" then
	view_condi = ""
	owner_view = "C"
	condi = "전체"
	ck_sw = "n"
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
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

if view_condi <> "" then
   if condi = "전체" then
      if owner_view = "T" then  
              Sql = "SELECT count(*) FROM emp_appoint where (app_empno = '"+view_condi+"') and (app_id = '포상발령' or app_id = '징계발령')"
         else
              Sql = "SELECT count(*) FROM emp_appoint where (app_emp_name like '%"+view_condi+"%') and (app_id = '포상발령' or app_id = '징계발령')"
      end if
     else 
      if owner_view = "T" then  
              Sql = "SELECT count(*) FROM emp_appoint where app_empno = '"+view_condi+"' and app_id = '"+condi+"'"
         else
              Sql = "SELECT count(*) FROM emp_appoint where app_emp_name like '%"+view_condi+"%' and app_id = '"+condi+"'"
      end if
   end if  
   Set RsCount = Dbconn.Execute (sql)
   tottal_record = cint(RsCount(0)) 'Result.RecordCount
end if 
'Set RsCount = Dbconn.Execute (sql)

'tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

if view_condi <> "" then
   if condi = "전체" then
      if owner_view = "T" then  
              Sql = "SELECT * FROM emp_appoint where (app_empno = '"+view_condi+"') and (app_id = '포상발령' or app_id = '징계발령') ORDER BY app_empno,app_date,app_seq ASC limit "& stpage & "," &pgsize
         else
              Sql = "SELECT * FROM emp_appoint where (app_emp_name like '%"+view_condi+"%') and (app_id = '포상발령' or app_id = '징계발령') ORDER BY app_empno,app_date,app_seq ASC limit "& stpage & "," &pgsize
      end if
     else 
      if owner_view = "T" then  
              Sql = "SELECT * FROM emp_appoint where app_empno = '"+view_condi+"' and app_id = '"+condi+"' ORDER BY app_empno,app_date,app_seq ASC limit "& stpage & "," &pgsize
         else
              Sql = "SELECT * FROM emp_appoint where app_emp_name like '%"+view_condi+"%' and app_id = '"+condi+"' ORDER BY app_empno,app_date,app_seq ASC limit "& stpage & "," &pgsize
      end if
   end if   
   Rs.Open Sql, Dbconn, 1
end if
'Rs.Open Sql, Dbconn, 1


title_line = " 상벌 사항 "
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
				return "1 1";
			}
			function goAction () {
			   window.close () ;
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
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {
				if (document.frm.view_condi.value == "") {
					alert ("조건을 입력하시기 바랍니다");
					return false;
				}	
				return true;
			}
			
			function reward_punish_del(val, val2, val3) {

            if (!confirm("정말 삭제하시겠습니까 ?")) return;
            var frm = document.frm;
			document.frm.app_empno.value = val;
			document.frm.app_seq.value = val2;
			document.frm.app_emp_name.value = val3;
		
            document.frm.action = "insa_reward_punish_del.asp";
            document.frm.submit();
            }	
		</script>

	</head>
	<body>
		<div id="wrap">			
			<!--#include virtual = "/include/insa_header.asp" -->
			<!--#include virtual = "/include/insa_sub_menu1.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_reward_punish_mg.asp?ck_sw=<%="n"%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>◈조건 검색◈</dt>
                        <dd>
                            <p>
                                <label>
                            <strong>상벌 : </strong>    
                                <select name="condi" id="condi" value="<%=condi%>" style="width:100px">
                                  <option value="전체" <%If condi = "전체" then %>selected<% end if %>>전체</option>
                                  <option value="포상발령" <%If condi = "포상발령" then %>selected<% end if %>>포상발령</option>
                                  <option value="징계발령" <%If condi = "징계발령" then %>selected<% end if %>>징계발령</option>
                                </select>
                                </label>
                                <label>
                                <input name="owner_view" type="radio" value="T" <% if owner_view = "T" then %>checked<% end if %> style="width:25px">사번
                                <input name="owner_view" type="radio" value="C" <% if owner_view = "C" then %>checked<% end if %> style="width:25px">성명
                                </label>
							<strong>조건 : </strong>
								<label>
        						<input name="view_condi" type="text" id="view_condi" value="<%=view_condi%>" style="width:100px; text-align:left">
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser1.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="5%" >
                            <col width="7%" >
                            <col width="10%" >
                            <col width="6%" >
							<col width="10%" >
							<col width="13%" >
                            <col width="*" >
                            <col width="22%" >
                            <col width="4%" >
                            <col width="4%" >
                            <col width="4%" >
						</colgroup>
						<thead>
                            <tr>
                                <th>사번</th>
                                <th>성명</th>
                                <th>현소속</th>
                                <th>상벌일자</th>
                                <th>상벌유형</th>
                                <th>징계기간</th>
                                <th>상벌내용</th>
                                <th>직급/직책 및 소속</th>
                                <th>상벌</th>
                                <th>수정</th>
                                <th>비고</th>
                            </tr>
                        </thead>
						<tbody>
						<%
						if  view_condi <> "" then 
						do until rs.eof
						      app_empno = rs("app_empno")
							  Sql = "SELECT * FROM emp_master where emp_no = '"&app_empno&"'"
                              Set rs_emp = DbConn.Execute(SQL)
							  if not Rs_emp.eof then
                                   emp_company = rs_emp("emp_company")
								   emp_name = rs_emp("emp_name")
								   emp_job = rs_emp("emp_job")
                                   emp_bonbu = rs_emp("emp_bonbu")
                                   emp_saupbu = rs_emp("emp_saupbu")
                                   emp_team = rs_emp("emp_team")
                                   emp_org_code = rs_emp("emp_org_code")
                                   emp_org_name = rs_emp("emp_org_name")
							  end if
							  rs_emp.close()
							  
							  'reward_memo = replace(rs("app_reward"),chr(34),chr(39))
							  'view_reward = reward_memo
							  'if len(reward_memo) > 10 then
							  '  	view_reward = mid(reward_memo,1,10) + ".."
							  'end if
							  
							  'task_memo = replace(rs("app_comment"),chr(34),chr(39))
							  'view_memo = task_memo
							  'if len(task_memo) > 10 then
							  '  	view_memo = mid(task_memo,1,10) + ".."
							  'end if
						%>
							<tr>
                              <td><%=rs("app_empno")%>&nbsp;</td>
                              <td><%=emp_name%>(<%=emp_job%>)&nbsp;</td>
                              <td><%=emp_org_name%>(<%=emp_org_code%>)&nbsp;</td>
                              <td><%=rs("app_date")%>&nbsp;</td>
                        <% if rs("app_id") = "포상발령" then %>
						      <td class="left">(포상)<%=rs("app_id_type")%>&nbsp;</td>
                              <td class="left">&nbsp;</td> 
                              <td class="left"><%=rs("app_reward")%>&nbsp;</td> 
                        <%    elseif rs("app_id") = "징계발령" then %>
                              <td class="left">(징계)<%=rs("app_id_type")%>&nbsp;</td>
                              <td class="left"><%=rs("app_start_date")%>∼<%=rs("app_finish_date")%>&nbsp;</td>
                              <td class="left"><%=rs("app_comment")%>&nbsp;</td>
                        <% end if %>
                              <td class="left"><%=rs("app_to_grade")%>-<%=rs("app_to_position")%>(<%=rs("app_to_company")%>&nbsp;<%=rs("app_to_org")%>(<%=rs("app_to_orgcode")%>)</td>
                              
                              
                        <% if user_id = "900002" then %>      
                              <td >
                              <a href="#" onClick="pop_Window('insa_reward_punish_add.asp?app_empno=<%=rs("app_empno")%>&emp_name=<%=emp_name%>&owner_view=<%=owner_view%>&u_type=<%=""%>','insa_reward_punish_add_pop','scrollbars=yes,width=750,height=300')">등록</a></td>
							  <td><a href="#" onClick="pop_Window('insa_reward_punish_add.asp?app_empno=<%=rs("app_empno")%>&app_seq=<%=rs("app_seq")%>&emp_name=<%=emp_name%>&owner_view=<%=owner_view%>&u_type=<%="U"%>','insa_reward_punish_add_pop','scrollbars=yes,width=750,height=300')">수정</a></td>
                         <% if insa_grade = "0" then %>     
                              <td>
                              <a href="#" onClick="reward_punish_del('<%=rs("app_empno")%>', '<%=rs("app_seq")%>', '<%=emp_name%>');return false;">삭제</a></td>
                         <%     else %>
                              <td>&nbsp;</td>
                         <% end if %>
                         <% end if %>
                              <td>&nbsp;</td>
                              <td>&nbsp;</td>
                              <td>&nbsp;</td>
							</tr>
						<%
							rs.movenext()
						loop
						rs.close()
						
						end if
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
                        <a href = "insa_reward_punish_mg.asp?page=<%=first_page%>&view_condi=<%=view_condi%>&condi=<%=condi%>&owner_view=<%=owner_view%>&ck_sw=<%="y"%>">[처음]</a>
                  	<% if intstart > 1 then %>
                        <a href="insa_reward_punish_mg.asp?page=<%=intstart -1%>&view_condi=<%=view_condi%>&condi=<%=condi%>&owner_view=<%=owner_view%>&ck_sw=<%="y"%>">[이전]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
           	<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="insa_reward_punish_mg.asp?page=<%=i%>&view_condi=<%=view_condi%>&condi=<%=condi%>&owner_view=<%=owner_view%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
           	<% if 	intend < total_page then %>
                        <a href="insa_reward_punish_mg.asp?page=<%=intend+1%>&view_condi=<%=view_condi%>&condi=<%=condi%>&owner_view=<%=owner_view%>&ck_sw=<%="y"%>">[다음]</a> <a href="insa_reward_punish_mg.asp?page=<%=total_page%>&view_condi=<%=view_condi%>&condi=<%=condi%>&owner_view=<%=owner_view%>&ck_sw=<%="y"%>">[마지막]</a>
                        <%	else %>
                        [다음]&nbsp;[마지막]
                      <% end if %>
                    </div>
                    </td>
			      </tr>
				  </table>                                
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td>
					<div class="btnRight">
					<% if user_id = "900002" or user_id = "102592" then %>
                    <a href="#" onClick="pop_Window('insa_reward_punish_add.asp?family_empno=<%=view_condi%>&emp_name=<%=emp_name%>','insa_reward_punish_add_pop','scrollbars=yes,width=750,height=300')" class="btnType04">상벌사항등록</a>
                    <% end if %>
					</div>                  
                    </td>
			      </tr>
				  </table>
                  <input type="hidden" name="family_empno" value="<%=family_empno%>" ID="Hidden1">
                  <input type="hidden" name="family_seq" value="<%=family_seq%>" ID="Hidden1">
                  <input type="hidden" name="family_name" value="<%=emp_name%>" ID="Hidden1">
			</form>
		</div>				
	</div>        				
	</body>
</html>

