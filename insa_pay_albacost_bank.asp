<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows
dim month_tab(24,2)
dim quarter_tab(8,2)
dim year_tab(3,2)

be_pg = "insa_pay_albacost_bank.asp"

curr_date = mid(cstr(now()),1,10)
curr_year = mid(cstr(now()),1,4)
curr_month = mid(cstr(now()),6,2)
curr_day = mid(cstr(now()),9,2)

Page=Request("page")
view_condi = request("view_condi")

ck_sw=Request("ck_sw")

if ck_sw = "n" then
	view_condi = request.form("view_condi")
	view_bank = request.form("view_bank")
	rever_yymm=Request.form("rever_yymm")
    to_date=Request.form("to_date")
  else
	view_condi = request("view_condi")
	view_bank = request("view_bank")
	rever_yymm=request("rever_yymm")
    to_date=request("to_date") 
end if

if view_condi = "" then
	view_condi = "���̿��������"
'	view_bank = "��������"
	view_bank = "��ü"
	curr_dd = cstr(datepart("d",now))
	to_date = mid(cstr(now()),1,10)
	from_date = mid(cstr(now()-curr_dd+1),1,10)
	rever_yymm = mid(cstr(from_date),1,4) + mid(cstr(from_date),6,2)
	
	sum_alba_pay = 0
	sum_alba_trans = 0
	sum_alba_meals = 0
	sum_alba_other = 0
	sum_tax_amt1 = 0
	sum_tax_amt2 = 0
	sum_give_total = 0
	
	pay_count = 0	
	sum_curr_pay = 0	
	
end if

give_date = to_date '������

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
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_alba = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")    
dbconn.open DbConnect

if view_bank = "��ü" then
       Sql = "select count(*) from pay_alba_cost where (rever_yymm = '"+rever_yymm+"' ) and (company = '"+view_condi+"')"
   else
       Sql = "select count(*) from pay_alba_cost where (rever_yymm = '"+rever_yymm+"' ) and (company = '"+view_condi+"') and (bank_name = '"+view_bank+"')"
end if
Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

if view_bank = "��ü" then
       Sql = "select * from pay_alba_cost where (rever_yymm = '"+rever_yymm+"' ) and (company = '"+view_condi+"') ORDER BY company,draft_no ASC"
   else
       Sql = "select * from pay_alba_cost where (rever_yymm = '"+rever_yymm+"' ) and (company = '"+view_condi+"') and (bank_name = '"+view_bank+"') ORDER BY company,draft_no ASC"
end if
Rs.Open Sql, Dbconn, 1
do until rs.eof
      pay_count = pay_count + 1
				  
	  sum_alba_pay = sum_alba_pay + int(rs("alba_pay"))
	  sum_alba_trans = sum_alba_trans + int(rs("alba_trans"))
	  sum_alba_meals = sum_alba_meals + int(rs("alba_meals"))
	  sum_alba_other = sum_alba_other + int(rs("alba_other"))
	  sum_tax_amt1 = sum_tax_amt1 + int(rs("tax_amt1"))
	  sum_tax_amt2 = sum_tax_amt2 + int(rs("tax_amt2"))
      sum_give_total = sum_give_total + int(rs("alba_give_total"))
	rs.movenext()
loop
rs.close()	  
	  
if view_bank = "��ü" then
      Sql = "select * from pay_alba_cost where (rever_yymm = '"+rever_yymm+"' ) and (company = '"+view_condi+"') ORDER BY company,draft_no ASC limit "& stpage & "," &pgsize 
   else
      Sql = "select * from pay_alba_cost where (rever_yymm = '"+rever_yymm+"' ) and (company = '"+view_condi+"') and (bank_name = '"+view_bank+"') ORDER BY company,draft_no ASC limit "& stpage & "," &pgsize 
end if

Rs.Open Sql, Dbconn, 1

curr_yyyy = mid(cstr(rever_yymm),1,4)
curr_mm = mid(cstr(rever_yymm),6,2)
title_line = cstr(curr_yyyy) + "�� " + cstr(curr_mm) + "�� " + " ����ҵ� ���ະ ��ü��Ȳ"

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>�޿����� �ý���</title>
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
			<!--#include virtual = "/include/insa_pay_header.asp" -->
			<!--#include virtual = "/include/insa_pay_alba_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_pay_albacost_bank.asp?ck_sw=<%="n"%>" method="post" name="frm">
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
                                    <select name="rever_yymm" id="rever_yymm" type="text" value="<%=rever_yymm%>" style="width:90px">
                                    <%	for i = 24 to 1 step -1	%>
                                    <option value="<%=month_tab(i,1)%>" <%If rever_yymm = month_tab(i,1) then %>selected<% end if %>><%=month_tab(i,2)%></option>
                                    <%	next	%>
                                 </select>
								</label>

                            <strong>��ü���� : </strong>
                              <%
								Sql="select * from emp_etc_code where emp_etc_type = '50' order by emp_etc_name asc"
					            Rs_etc.Open Sql, Dbconn, 1
							  %>
                                <label>
								<select name="view_bank" id="view_bank" type="text" style="width:100px">
                                    <option value="��ü" <%If view_bank = "��ü" then %>selected<% end if %>>��ü</option>
                			  <% 
								do until Rs_etc.eof 
			  				  %>
                					<option value='<%=rs_etc("emp_etc_name")%>' <%If view_bank = rs_etc("emp_etc_name") then %>selected<% end if %>><%=rs_etc("emp_etc_name")%></option>
                			  <%
									Rs_etc.movenext()  
								loop 
								Rs_etc.Close()
							  %>
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
                            <col width="9%" >
                            <col width="9%" >
                            <col width="9%" >
                            <col width="12%" >
                            <col width="9%" >
                            <col width="10%" >
                            <col width="14%" >
                            <col width="10%" >
							<col width="9%" >
                            <col width="9%" >
						</colgroup>
						<thead>
							<tr>
				               <th class="first" scope="col">��Ϲ�ȣ</th>
                               <th scope="col">����</th>
                               <th scope="col">�����</th>
                               <th scope="col">�Ҽ�</th>
                               <th scope="col">�ҵ汸��</th>
				               <th scope="col">��ü����</th>
                               <th scope="col">���¹�ȣ</th>
                               <th scope="col">�����ָ�</th>
                               <th scope="col">�������޾�</th>
                               <th scope="col">�����޾�</th>
			                </tr>
						</thead>
						<tbody>
						<%
						do until rs.eof
							  draft_no = rs("draft_no")

							  curr_pay = int(rs("alba_give_total")) - (int(rs("tax_amt1")) + int(rs("tax_amt2")))
							  
							  Sql = "SELECT * FROM emp_alba_mst where draft_no = '"&draft_no&"'"
                              Set rs_alba = DbConn.Execute(SQL)
		                      if not rs_alba.eof then
		                    		draft_date = rs_alba("draft_date")
	                             else
	                    			draft_date = ""
                              end if
                              rs_alba.close()
					  
	           			%>
							<tr>
								<td class="first"><%=rs("draft_no")%>&nbsp;</td>
                                <td><%=rs("draft_man")%>&nbsp;</td>
                                <td><%=draft_date%>&nbsp;</td>
                                <td><%=rs("org_name")%>&nbsp;</td>
                                <td><%=rs("draft_tax_id")%>&nbsp;</td>
                                <td><%=rs("bank_name")%>&nbsp;</td>
                                <td><%=rs("account_no")%>&nbsp;</td>
                                <td><%=rs("account_name")%>&nbsp;</td>
                                <td class="right"><%=formatnumber(curr_pay,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(curr_pay,0)%>&nbsp;</td>
							</tr>
						<%
							rs.movenext()
						loop
						rs.close()
						
						sum_curr_pay = sum_give_total - (sum_tax_amt1 + sum_tax_amt2)
												
						%>
                          	<tr>
                                <th colspan="8" class="first">�Ѱ�&nbsp;</th>
                                <th class="right"><%=formatnumber(sum_curr_pay,0)%>&nbsp;</th>
                                <th class="right"><%=formatnumber(sum_curr_pay,0)%>&nbsp;</th>
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
                    <a href="insa_excel_pay_albacost_bank.asp?view_condi=<%=view_condi%>&rever_yymm=<%=rever_yymm%>&to_date=<%=to_date%>&view_bank=<%=view_bank%>" class="btnType04">�����ٿ�ε�</a>
					</div>                  
                  	</td>
				    <td>
                    <div id="paging">
                        <a href = "insa_pay_albacost_bank.asp?page=<%=first_page%>&view_condi=<%=view_condi%>&rever_yymm=<%=rever_yymm%>&to_date=<%=to_date%>&view_bank=<%=view_bank%>&ck_sw=<%="y"%>">[ó��]</a>
                  	<% if intstart > 1 then %>
                        <a href="insa_pay_albacost_bank.asp?page=<%=intstart -1%>&view_condi=<%=view_condi%>&rever_yymm=<%=rever_yymm%>&to_date=<%=to_date%>&view_bank=<%=view_bank%>&ck_sw=<%="y"%>">[����]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
           	<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="insa_pay_albacost_bank.asp?page=<%=i%>&view_condi=<%=view_condi%>&rever_yymm=<%=rever_yymm%>&to_date=<%=to_date%>&view_bank=<%=view_bank%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
           	<% if 	intend < total_page then %>
                        <a href="insa_pay_albacost_bank.asp?page=<%=intend+1%>&view_condi=<%=view_condi%>&rever_yymm=<%=rever_yymm%>&to_date=<%=to_date%>&view_bank=<%=view_bank%>&ck_sw=<%="y"%>">[����]</a> <a href="insa_pay_albacost_bank.asp?page=<%=total_page%>&view_condi=<%=view_condi%>&rever_yymm=<%=rever_yymm%>&to_date=<%=to_date%>&view_bank=<%=view_bank%>&ck_sw=<%="y"%>">[������]</a>
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
	</body>
</html>

