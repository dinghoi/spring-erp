<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows
dim page_cnt
dim pg_cnt
dim month_tab(24,2)
dim quarter_tab(8,2)
dim year_tab(3,2)

Page=Request("page")
page_cnt=Request.form("page_cnt")
pg_cnt=cint(Request("pg_cnt"))

be_pg = "insa_year_imcome_agree_mg.asp"

curr_date = mid(cstr(now()),1,10)
curr_year = mid(cstr(now()),1,4)
curr_month = mid(cstr(now()),6,2)
curr_day = mid(cstr(now()),9,2)

view_condi = request("view_condi")

ck_sw=Request("ck_sw")

if ck_sw = "n" then
	view_condi = request.form("view_condi")
	agree_year=Request.form("agree_year")
  else
	view_condi = request("view_condi")
	agree_year=request("agree_year")
end if

if view_condi = "" then
	view_condi = "���̿��������"
	curr_dd = cstr(datepart("d",now))
	from_date = mid(cstr(now()-curr_dd+1),1,10)
	agree_year = mid(cstr(from_date),1,4)
end if

' �ֱ�3���⵵ ���̺�� ����
year_tab(3,1) = mid(now(),1,4)
year_tab(3,2) = cstr(year_tab(3,1)) + "��"
year_tab(2,1) = cint(mid(now(),1,4)) - 1
year_tab(2,2) = cstr(year_tab(2,1)) + "��"
year_tab(1,1) = cint(mid(now(),1,4)) - 2
year_tab(1,2) = cstr(year_tab(1,1)) + "��"

' �б� ���̺� ����
curr_mm = mid(now(),6,2)
if curr_mm > 0 and curr_mm < 4 then
	quarter_tab(8,1) = cstr(mid(now(),1,4)) + "1"
end if
if curr_mm > 3 and curr_mm < 7 then
	quarter_tab(8,1) = cstr(mid(now(),1,4)) + "2"
end if
if curr_mm > 6 and curr_mm < 10 then
	quarter_tab(8,1) = cstr(mid(now(),1,4)) + "3"
end if
if curr_mm > 9 and curr_mm < 13 then
	quarter_tab(8,1) = cstr(mid(now(),1,4)) + "4"
end if

quarter_tab(8,2) = cstr(mid(quarter_tab(8,1),1,4)) + "�� " + cstr(mid(quarter_tab(8,1),5,1)) + "/4�б�"

for i = 7 to 1 step -1
	cal_quarter = cint(quarter_tab(i+1,1)) - 1
	if cstr(mid(cal_quarter,5,1)) = "0" then
		quarter_tab(i,1) = cstr(cint(mid(cal_quarter,1,4))-1) + "4"
	  else
		quarter_tab(i,1) = cal_quarter
	end if	 
	quarter_tab(i,2) = cstr(mid(quarter_tab(i,1),1,4)) + "�� " + cstr(mid(quarter_tab(i,1),5,1)) + "/4�б�"
next

' ��� ���̺����
'cal_month = cstr(mid(dateadd("m",-1,now()),1,4)) + cstr(mid(dateadd("m",-1,now()),6,2))	
cal_month = mid(cstr(now()),1,4) + mid(cstr(now()),6,2)
month_tab(24,1) = cal_month
view_month = mid(cal_month,1,4) + "�� " + mid(cal_month,5,2) + "��"
month_tab(24,2) = view_month
for i = 1 to 23
	cal_month = cstr(int(cal_month) - 1)
	if mid(cal_month,5) = "00" then
		cal_year = cstr(int(mid(cal_month,1,4)) - 1)
		cal_month = cal_year + "12"
	end if	 
	view_month = mid(cal_month,1,4) + "�� " + mid(cal_month,5,2) + "��"
	j = 24 - i
	month_tab(j,1) = cal_month
	month_tab(j,2) = view_month
next

pgsize = 10 ' ȭ�� �� ������ 

If Page = "" Then
	Page = 1
	start_page = 1
End If
stpage = int((page - 1) * pgsize)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_emp = Server.CreateObject("ADODB.Recordset")
Set rs_org = Server.CreateObject("ADODB.Recordset")
Set rs_etc = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")

dbconn.open DbConnect

agree_id = "�����ٷΰ�༭"       

order_Sql = " ORDER BY agree_year,agree_company,agree_empno ASC"
where_sql = " WHERE (agree_id = '"+agree_id+"') and (agree_company = '"+view_condi+"') and (agree_year = '"+agree_year+"')"

Sql = "SELECT count(*) FROM emp_agree " + where_sql
Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

sql = "select * from emp_agree " + where_sql + order_sql + " limit "& stpage & "," &pgsize 
Rs.Open Sql, Dbconn, 1

title_line = " �����ٷΰ�ൿ�� ��Ȳ "

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>�λ���� �ý���</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "2 1";
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
					alert ("�Ҽ��� �����Ͻñ� �ٶ��ϴ�");
					return false;
				}	
				return true;
			}
		</script>

	</head>
	<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
		<div id="wrap">			
			<!--#include virtual = "/include/insa_header.asp" -->
			<!--#include virtual = "/include/insa_appoint_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_year_imcome_agree_mg.asp" method="post" name="frm">
                <fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>					
						<dt>�� �˻���</dt>
                        <dd>
                            <p>
                             <strong>ȸ�� : </strong>
                              <%
								Sql="select * from emp_org_mst where isNull(org_end_date) and org_level = 'ȸ��' ORDER BY org_code ASC"
	                            rs_org.Open Sql, Dbconn, 1	
							  %>
                                <label>
								<select name="view_condi" id="view_condi" type="text" style="width:130px">
                                  <option value="��ü" <%If view_condi = "��ü" then %>selected<% end if %>>��ü</option>
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
								<strong>�ͼӳ⵵ : </strong>
                                    <select name="agree_year" id="agree_year" type="text" value="<%=agree_year%>" style="width:90px">
                                    <%	for i = 3 to 1 step -1	%>
                                    <option value="<%=year_tab(i,1)%>" <%If agree_year = cstr(year_tab(i,1)) then %>selected<% end if %>><%=year_tab(i,2)%></option>
                                    <%	next	%>
                                 </select>
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser1.jpg" alt="�˻�"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="7%" >
                            <col width="7%" >
                            <col width="7%" >
							<col width="7%" >
							<col width="7%" >
                            <col width="10%" >
                            <col width="10%" >
                            <col width="10%" >
                            <col width="7%" >
							<col width="*" >
                            <col width="4%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">���</th>
								<th scope="col">����</th>
								<th scope="col">������</th>
								<th scope="col">����å</th>
                                <th scope="col">�Ի���</th>
                                <th scope="col">�ֹε�Ϲ�ȣ</th>
                                <th scope="col">ȸ��</th>
                                <th scope="col">�Ҽ�</th>
								<th scope="col">��������</th>
								<th scope="col">�ּ�</th>
                                <th scope="col">���</th>
							</tr>
						</thead>
					<tbody>
						<%
						do until rs.eof
						
						%>
							<tr>
								<td class="first"><%=rs("agree_empno")%></td>
                                <td>
                                <a href="#" onClick="pop_Window('insa_card00.asp?emp_no=<%=rs("agree_empno")%>&be_pg=<%=be_pg%>&page=<%=page%>&view_condi=<%=view_condi%>&agree_year=<%=agree_year%>&agree_id=<%=agree_id%>&page_cnt=<%=page_cnt%>','emp_card0_pop','scrollbars=yes,width=1250,height=650')"><%=rs("agree_empname")%></a>
								</td>
                                <td><%=rs("agree_grade")%>&nbsp;</td>
                                <td><%=rs("agree_position")%>&nbsp;</td>
                                <td><%=rs("agree_in_date")%>&nbsp;</td>
                                <td><%=rs("agree_person1")%>-<%=rs("agree_person2")%>&nbsp;</td>
                                <td><%=rs("agree_company")%>&nbsp;</td>
                                <td><%=rs("agree_org_name")%>(<%=rs("agree_org_code")%>)&nbsp;</td>
                                <td><%=rs("agree_date")%>&nbsp;</td>
                                <td class="left"><%=rs("agree_sido")%>&nbsp;<%=rs("agree_gugun")%>&nbsp;<%=rs("agree_dong")%>&nbsp;<%=rs("agree_addr")%></td>
                                <td class="left"><%=rs("agree_sw1")%>&nbsp;</td>
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
				    <td width="15%">
					<div class="btnCenter">
                    <a href="insa_excel_year_income_agree.asp?view_condi=<%=view_condi%>&agree_id=<%=agree_id%>&agree_year=<%=agree_year%>&agree_year=<%=agree_year%>" class="btnType04">�����ٿ�ε�</a>
					</div>                  
                  	</td>
				    <td>
                    <div id="paging">
                        <a href="insa_year_imcome_agree_mg.asp?page=<%=first_page%>&view_condi=<%=view_condi%>&agree_year=<%=agree_year%>&agree_id=<%=agree_id%>ck_sw=<%="y"%>">[ó��]</a>
                  	<% if intstart > 1 then %>
                        <a href="insa_year_imcome_agree_mg.asp?page=<%=intstart -1%>&view_condi=<%=view_condi%>&agree_year=<%=agree_year%>&agree_id=<%=agree_id%>ck_sw=<%="y"%>">[����]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
                  	<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="insa_year_imcome_agree_mg.asp?page=<%=i%>&view_condi=<%=view_condi%>&agree_year=<%=agree_year%>&agree_id=<%=agree_id%>ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
                  	<% if 	intend < total_page then %>
                        <a href="insa_year_imcome_agree_mg.asp?page=<%=intend+1%>&view_condi=<%=view_condi%>&agree_year=<%=agree_year%>&agree_id=<%=agree_id%>ck_sw=<%="y"%>">[����]</a> <a href="insa_year_imcome_agree_mg.asp?page=<%=total_page%>&view_condi=<%=view_condi%>&agree_year=<%=agree_year%>&agree_id=<%=agree_id%>ck_sw=<%="y"%>">[������]</a>
                        <%	else %>
                        [����]&nbsp;[������]
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

