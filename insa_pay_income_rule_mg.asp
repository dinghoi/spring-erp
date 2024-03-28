<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows

curr_date = mid(cstr(now()),1,10)
curr_year = mid(cstr(now()),1,4)
curr_month = mid(cstr(now()),6,2)
curr_day = mid(cstr(now()),9,2)

ck_sw=Request("ck_sw")
Page=Request("page")

If ck_sw = "y" Then
	view_condi=Request("view_condi")
  else
	view_condi=Request.form("view_condi")
End if

If view_condi = "" Then
	view_condi = "5701근로소득공제"
End If

rule_code = mid(cstr(view_condi),1,4)
rule_yyyy = mid(cstr(now()),1,4) '귀속년월

pgsize = 10 ' 화면 한 페이지 

If Page = "" Then
	Page = 1
	start_page = 1
End If
stpage = int((page - 1) * pgsize)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_into = Server.CreateObject("ADODB.Recordset")
Set rs_etc = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

Sql = "select count(*) from pay_income_rule where (rule_id = '"+rule_code+"')"
Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

Sql = "select * from pay_income_rule where rule_id = '"+rule_code+"' ORDER BY rule_yyyy,rule_id,rule_cl ASC limit "& stpage & "," &pgsize 
Rs.Open Sql, Dbconn, 1

title_line = "근로소득세율 설정관리"

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>급여관리 시스템</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "4 1";
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
			<!--#include virtual = "/include/insa_pay_header.asp" -->
			<!--#include virtual = "/include/insa_pay_rule_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_pay_income_rule_mg.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>소득세율 검색</dt>
                        <dd>
                            <p>
                               <strong>세율구분 : </strong>
                              <%
								Sql="select * from emp_etc_code where emp_etc_type = '57' order by emp_etc_code asc"
					            Rs_etc.Open Sql, Dbconn, 1
							  %>
                                <label>
								<select name="view_condi" id="view_condi" type="text" style="width:150px" value="<%=view_condi%>">

                			  <% 
								do until Rs_etc.eof 
			  				  %>
                					<option value='<%=rs_etc("emp_etc_code")%><%=rs_etc("emp_etc_name")%>' <%If view_condi = (rs_etc("emp_etc_code")+rs_etc("emp_etc_name")) then %>selected<% end if %>><%=rs_etc("emp_etc_code")%><%=rs_etc("emp_etc_name")%></option>
                			  <%
									Rs_etc.movenext()  
								loop 
								Rs_etc.Close()
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
							<col width="8%" >
                            <col width="10%" >
							<col width="4%" >
							<col width="8%" >
							<col width="8%" >
							<col width="8%" >
							<col width="6%" >
                            <col width="8%" >
                            <col width="6%" >
                            <col width="*" >
                            <col width="4%" >
						</colgroup>
						<thead>
				            <tr>
				               <th rowspan="2" class="first" scope="col" style=" border-bottom:1px solid #e3e3e3;">기준<br>적용년도</th>
				               <th rowspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">세율구분</th>
                               <th rowspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">분류</th>
                               <th rowspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">총급여액<br>과세/세액</th>
                               <th colspan="3" scope="col" style=" border-bottom:1px solid #e3e3e3;">공제</th>
				               <th colspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">추가공제</th>
                               <th rowspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">비고</th>
                               <th rowspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">변경</th>
			               </tr>
                           <tr>
				              <th scope="col" style=" border-left:1px solid #e3e3e3;">기준공제액</th>
				              <th scope="col" style=" border-bottom:1px solid #e3e3e3;">초과공제액</th>
                              <th scope="col">초과<br>공제율</th>
				              <th scope="col">추가공제액</th>
				              <th scope="col">추가<br>공제율</th>
                           </tr>
						</thead>
						<tbody>
						<%
						do until rs.eof
						%>
							<tr>
								<td class="first"><%=rs("rule_yyyy")%>&nbsp;</td>
								<td ><%=view_condi%>&nbsp;</td>
                                <td ><%=rs("rule_cl")%>&nbsp;</td>
								<td class="right"><%=formatnumber(rs("rule_year_pay"),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("rule_st_deduct"),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("rule_exceed"),0)%>&nbsp;</td>
								<td class="right"><%=formatnumber(rs("rule_exceed_rate"),2)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("rule_add"),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("rule_add_rate"),2)%>&nbsp;</td>
                                <td class="left"><%=rs("rule_comment")%>&nbsp;</td>
								<td>
                                 <a href="#" onClick="pop_Window('insa_pay_income_rule_add.asp?rule_id=<%=rule_code%>&view_condi=<%=view_condi%>&rule_cl=<%=rs("rule_cl")%>&rule_yyyy=<%=rs("rule_yyyy")%>&u_type=<%="U"%>','pay_income_rule_add_popup','scrollbars=yes,width=750,height=400')">수정</a>
                                </td>
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
                        <a href="insa_pay_income_rule_mg.asp?page=<%=first_page%>&view_condi=<%=view_condi%>&ck_sw=<%="y"%>">[처음]</a>
                  	<% if intstart > 1 then %>
                        <a href="insa_pay_income_rule_mg.asp?page=<%=intstart -1%>&view_condi=<%=view_condi%>&ck_sw=<%="y"%>">[이전]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
                  	<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="insa_pay_income_rule_mg.asp?page=<%=i%>&view_condi=<%=view_condi%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
                  	<% if 	intend < total_page then %>
                        <a href="insa_pay_income_rule_mg.asp?page=<%=intend+1%>&view_condi=<%=view_condi%>&ck_sw=<%="y"%>">[다음]</a> <a href="insa_pay_income_rule_mg.asp?page=<%=total_page%>&view_condi=<%=view_condi%>&ck_sw=<%="y"%>">[마지막]</a>
                        <%	else %>
                        [다음]&nbsp;[마지막]
                      <% end if %>
                    </div>
                    </td>
				    <td width="20%">
					<div class="btnCenter">
                    <a href="#" onClick="pop_Window('insa_pay_income_rule_add.asp?rule_id=<%=rule_code%>&view_condi=<%=view_condi%>&u_type=<%=""%>','pay_income_rule_add_popup','scrollbars=yes,width=750,height=400')" class="btnType04">근로소득세율 등록</a>
					</div>                  
                    </td>
			      </tr>
				  </table>
			</form>
		</div>				
	</div>        				
	</body>
</html>

