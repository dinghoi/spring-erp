<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows
dim month_tab(24,2)

be_pg = "insa_sawo_in_pay_trans.asp"

curr_date = mid(cstr(now()),1,10)
curr_year = mid(cstr(now()),1,4)
curr_month = mid(cstr(now()),6,2)
curr_day = mid(cstr(now()),9,2)

Page=Request("page")
view_condi = request("view_condi")

ck_sw=Request("ck_sw")

if ck_sw = "n" then
	view_condi = request.form("view_condi")
	pmg_yymm=Request.form("pmg_yymm")
  else
	view_condi = request("view_condi")
	pmg_yymm=request("pmg_yymm")
end if

if view_condi = "" then
	view_condi = "��ü"
	curr_dd = cstr(datepart("d",now))
	from_date = mid(cstr(now()-curr_dd+1),1,10)
	pmg_yymm = cstr(mid(dateadd("m",-1,now()),1,4)) + cstr(mid(dateadd("m",-1,now()),6,2))
end if

'pmg_yymm = "201409" '�ӽ�..���α׷��׽�Ʈ

' ��� ���̺����
cal_month = cstr(mid(dateadd("m",-1,now()),1,4)) + cstr(mid(dateadd("m",-1,now()),6,2))	
'cal_month = mid(cstr(now()),1,4) + mid(cstr(now()),6,2)
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
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set Rs_year = Server.CreateObject("ADODB.Recordset")
Set Rs_give = Server.CreateObject("ADODB.Recordset")
Set Rs_dct = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

if view_condi = "��ü" then
         Sql = "select count(*) from pay_month_deduct where (de_yymm = '"+pmg_yymm+"' ) and (de_id = '1') and (de_sawo_amt > 0)"
   else
         Sql = "select count(*) from pay_month_deduct where (de_yymm = '"+pmg_yymm+"' ) and (de_id = '1') and (de_sawo_amt > 0) and (de_company = '"+view_condi+"')"
end if
Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

if view_condi = "��ü" then
         Sql = "select * from pay_month_deduct where (de_yymm = '"+pmg_yymm+"' ) and (de_id = '1') and (de_sawo_amt > 0) ORDER BY de_company,de_emp_no ASC limit "& stpage & "," &pgsize
   else
         Sql = "select * from pay_month_deduct where (de_yymm = '"+pmg_yymm+"' ) and (de_id = '1') and (de_sawo_amt > 0) and (de_company = '"+view_condi+"') ORDER BY de_company,de_emp_no ASC limit "& stpage & "," &pgsize
end if
Rs.Open Sql, Dbconn, 1

title_line = " ����ȸ ������ ���� ���� "

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
				return "8 1";
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
			
			function sawo_in_pay_transe(val, val2) {

            if (!confirm("����ȸ �����ݰ��� ���ó�� �Ͻðڽ��ϱ� ?")) return;
            var frm = document.frm;
			document.frm.pmg_yymm1.value = document.getElementById(val).value;
			document.frm.view_condi1.value = document.getElementById(val2).value;
			
            document.frm.action = "insa_sawo_in_pay_transe_save.asp";
            document.frm.submit();
            }	
		</script>

	</head>
	<body>
		<div id="wrap">			
			<!--#include virtual = "/include/insa_header.asp" -->
			<!--#include virtual = "/include/insa_sawo_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_sawo_in_pay_trans.asp?ck_sw=<%="n"%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>					
						<dt>�� �˻���</dt>
                        <dd>
                            <p>
                             <strong>ȸ�� : </strong>
                              <%
							  ' 2019.02.22 ������ ��û ȸ�縮��Ʈ�� ������ �ҽ� org_end_date�� null �� �ƴ� �������ڸ� �����ϸ� ����Ʈ�� ��Ÿ���� �ʴ´�.
							  Sql = "SELECT * FROM emp_org_mst WHERE ISNULL(org_end_date) AND org_level = 'ȸ��'  ORDER BY org_company ASC"
							  set rs_org = DBConn.Execute(Sql)
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
								<strong>�ͼӳ�� : </strong>
                                    <select name="pmg_yymm" id="pmg_yymm" type="text" value="<%=pmg_yymm%>" style="width:90px">
                                    <%	for i = 24 to 1 step -1	%>
                                    <option value="<%=month_tab(i,1)%>" <%If pmg_yymm = month_tab(i,1) then %>selected<% end if %>><%=month_tab(i,2)%></option>
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
							<col width="8%" >
							<col width="8%" >
							<col width="8%" >
							<col width="8%" >
							<col width="8%" >
                            <col width="8%" >
                            <col width="17%" >
							<col width="8%" >
							<col width="*" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">���</th>
								<th scope="col">��  ��</th>
								<th scope="col">����</th>
								<th scope="col">��å</th>
								<th scope="col">�����Ի���</th>
                                <th scope="col">�Ի���</th>
                                <th scope="col">�Ҽ�</th>
								<th scope="col">������</th>
								<th scope="col">��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��</th>
							</tr>
						</thead>
						<tbody>
						<%
						do until rs.eof
							  emp_no = rs("de_emp_no")
	           			%>
							<tr>
								<td class="first"><%=rs("de_emp_no")%>&nbsp;</td>
                                <td><%=rs("de_emp_name")%>&nbsp;</td>
                                <td><%=rs("de_grade")%>&nbsp;</td>
                                <td><%=rs("de_position")%>&nbsp;</td>
                        <%
						      Sql = "SELECT * FROM emp_master where emp_no = '"+emp_no+"'"
                              Set rs_emp = DbConn.Execute(SQL)
							  if not rs_emp.eof then
									emp_first_date = rs_emp("emp_first_date")
									emp_in_date = rs_emp("emp_in_date")
	                             else
									emp_first_date = ""
									emp_in_date = ""
                              end if
                              rs_emp.close()
                          %>
                                <td><%=emp_first_date%>&nbsp;</td>
                                <td><%=emp_in_date%>&nbsp;</td>
                                <td><%=rs("de_company")%>&nbsp;-&nbsp;<%=rs("de_org_name")%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("de_sawo_amt"),0)%>&nbsp;</td>
                                <td class="left"><%=rs("de_company")%>-<%=rs("de_bonbu")%>-<%=rs("de_saupbu")%>-<%=rs("de_team")%></td>
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
                    <a href="insa_excel_sawo_in_pay_trans.asp?view_condi=<%=view_condi%>&pmg_yymm=<%=pmg_yymm%>" class="btnType04">�����ٿ�ε�</a>
					</div>                  
                  	</td>
				    <td>
                    <div id="paging">
                        <a href = "insa_sawo_in_pay_trans.asp?page=<%=first_page%>&view_condi=<%=view_condi%>&pmg_yymm=<%=pmg_yymm%>&ck_sw=<%="y"%>">[ó��]</a>
                  	<% if intstart > 1 then %>
                        <a href="insa_sawo_in_pay_trans.asp?page=<%=intstart -1%>&view_condi=<%=view_condi%>&pmg_yymm=<%=pmg_yymm%>&ck_sw=<%="y"%>">[����]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
           	<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="insa_sawo_in_pay_trans.asp?page=<%=i%>&view_condi=<%=view_condi%>&pmg_yymm=<%=pmg_yymm%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
           	<% if 	intend < total_page then %>
                        <a href="insa_sawo_in_pay_trans.asp?page=<%=intend+1%>&view_condi=<%=view_condi%>&pmg_yymm=<%=pmg_yymm%>&ck_sw=<%="y"%>">[����]</a> <a href="insa_sawo_in_pay_trans.asp?page=<%=total_page%>&view_condi=<%=view_condi%>&pmg_yymm=<%=pmg_yymm%>&ck_sw=<%="y"%>">[������]</a>
                        <%	else %>
                        [����]&nbsp;[������]
                      <% end if %>
                    </div>
                    </td>
                    <td>
					<div class="btnRight">
                    <a href="#" onClick="sawo_in_pay_transe('pmg_yymm','view_condi');return false;" class="btnType04">����ȸ�� �̰����</a>
					</div>                  
                    </td>
			      </tr>
				  </table>
                  <input type="hidden" name="pmg_yymm1" value="<%=pmg_yymm%>" ID="Hidden1">
                  <input type="hidden" name="view_condi1" value="<%=view_condi%>" ID="Hidden1">
			</form>
		</div>				
	</div>        				
	</body>
</html>

