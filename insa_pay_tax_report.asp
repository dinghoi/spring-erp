<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows
dim month_tab(24,2)
dim quarter_tab(8,2)
dim year_tab(3,2)

be_pg = "insa_pay_tax_report.asp"

curr_date = mid(cstr(now()),1,10)
curr_year = mid(cstr(now()),1,4)
curr_month = mid(cstr(now()),6,2)
curr_day = mid(cstr(now()),9,2)

view_condi = request("view_condi")

ck_sw=Request("ck_sw")

if ck_sw = "n" then
	view_condi = request.form("view_condi")
	pmg_yymm=Request.form("pmg_yymm")
    in_tax_id = Request.Form("in_tax_id") 
  else
	view_condi = request("view_condi")
	pmg_yymm=request("pmg_yymm")
    in_tax_id = request("in_tax_id") 
end if

if view_condi = "" then
	view_condi = "���̿��������"
	curr_dd = cstr(datepart("d",now))
	in_tax_id = "1"
	from_date = mid(cstr(now()-curr_dd+1),1,10)
	pmg_yymm = mid(cstr(from_date),1,4) + mid(cstr(from_date),6,2)
	
	sum_give_tot = 0
    sum_income_tax = 0
    sum_wetax = 0
	sum_year_incom_tax = 0
    sum_year_wetax = 0
	sum_special_tax = 0
	sum_deduct_tot = 0
	pay_count = 0	
	sum_curr_pay = 0
	
	a02_give_tot = 0
    a02_income_tax = 0
    a02_wetax = 0
	a02_count = 0	
	
	a03_give_tot = 0
    a03_income_tax = 0
    a03_wetax = 0
	a03_count = 0	
	
	a04_give_tot = 0
    a04_income_tax = 0
    a04_wetax = 0
	a04_count = 0	
	
	a10_give_tot = 0
    a10_income_tax = 0
    a10_wetax = 0
	a10_count = 0	
	
	a21_give_tot = 0
    a21_income_tax = 0
    a21_wetax = 0
	a21_count = 0	
	
	a22_give_tot = 0
    a22_income_tax = 0
    a22_wetax = 0
	a22_count = 0	
	
	a20_give_tot = 0
    a20_income_tax = 0
    a20_wetax = 0
	a20_count = 0	
	
	sum_alba_give_total = 0
    sum_tax_amt1 = 0
    sum_tax_amt2 = 0
	sum_deduct_tot = 0
	
	a32_give_tot = 0
    a32_income_tax = 0
    a32_wetax = 0
	a32_count = 0	
	
	a30_give_tot = 0
    a30_income_tax = 0
    a30_wetax = 0
	a30_count = 0
	
	tot_give_tot = 0
    tot_income_tax = 0
    tot_wetax = 0
	tot_year_incom_tax = 0
    tot_year_wetax = 0
	tot_special_tax = 0
	tot_deduct_tot = 0
	tot_pay_count = 0	
	tot_curr_pay = 0			
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

'�ٷμҵ�
Sql = "select * from pay_month_give where (pmg_yymm = '"+pmg_yymm+"' ) and (pmg_id = '1') and (pmg_company = '"+view_condi+"') ORDER BY pmg_company,pmg_org_code,pmg_emp_no ASC"
Rs.Open Sql, Dbconn, 1
do until rs.eof
    emp_no = rs("pmg_emp_no")
    pmg_give_tot = rs("pmg_give_total")
    pay_count = pay_count + 1
				  
    sub_give_hap = int(rs("pmg_postage_pay")) + int(rs("pmg_re_pay")) + int(rs("pmg_car_pay")) + int(rs("pmg_position_pay")) + int(rs("pmg_custom_pay")) + int(rs("pmg_job_pay")) + int(rs("pmg_job_support")) + int(rs("pmg_jisa_pay")) + int(rs("pmg_long_pay")) + int(rs("pmg_disabled_pay"))
	
	sum_give_tot = sum_give_tot + int(rs("pmg_give_total"))

    Sql = "select * from pay_month_deduct where (de_yymm = '"+pmg_yymm+"' ) and (de_id = '1') and (de_emp_no = '"+emp_no+"') and (de_company = '"+view_condi+"')"
    Set Rs_dct = DbConn.Execute(SQL)
    if not Rs_dct.eof then

            de_income_tax = int(Rs_dct("de_income_tax"))
            de_wetax = int(Rs_dct("de_wetax"))
			de_year_incom_tax = int(Rs_dct("de_year_incom_tax"))
            de_year_wetax = int(Rs_dct("de_year_wetax"))
		    de_deduct_tot = int(Rs_dct("de_deduct_total"))	
	     else
            de_income_tax = 0
            de_wetax = 0
			de_year_incom_tax = 0
            de_year_wetax = 0
		    de_deduct_tot = 0
     end if
     Rs_dct.close()
	 
     sum_income_tax = sum_income_tax + de_income_tax
     sum_wetax = sum_wetax + de_wetax
	 sum_year_incom_tax = sum_year_incom_tax + de_year_incom_tax
     sum_year_wetax = sum_year_wetax + de_year_wetax
	 sum_deduct_tot = sum_deduct_tot + de_deduct_tot

	rs.movenext()
loop
rs.close()

a10_give_tot = sum_give_tot + a02_give_tot + a03_give_tot + a03_give_tot 
a10_income_tax = sum_income_tax + a02_income_tax + a03_income_tax + a04_income_tax
a10_wetax = sum_wetax + a02_wetax + a03_wetax + a04_wetax
a10_count = pay_count + a02_count + a03_count + a04_count

'�����ҵ�
a20_give_tot = a21_give_tot + a22_give_tot
a20_income_tax = a21_income_tax + a22_income_tax
a20_wetax = a21_wetax + a22_wetax
a20_count = a21_count + a22_count

'����ҵ�
Sql = "select * from pay_alba_cost where (rever_yymm = '"+pmg_yymm+"' ) and (company = '"+view_condi+"') ORDER BY company,draft_no ASC"
Rs.Open Sql, Dbconn, 1
do until rs.eof
    alba_count = alba_count + 1
				  
    sum_alba_give_total = sum_alba_give_total + int(rs("alba_give_total"))
    sum_tax_amt1 = sum_tax_amt1 + int(rs("tax_amt1"))
    sum_tax_amt2 = sum_tax_amt2 + int(rs("tax_amt2"))
	sum_deduct_tot = sum_deduct_tot + (int(rs("tax_amt1")) + int(rs("tax_amt2")) + int(rs("de_other")))
	
	rs.movenext()
loop
rs.close()

a30_give_tot = sum_alba_give_total + a32_give_tot
a30_income_tax = sum_tax_amt1 + a32_income_tax
a30_wetax = sum_tax_amt2 + a32_wetax
a30_count = alba_count + a32_count

'�Ѱ�
tot_give_tot = a10_give_tot + a20_give_tot + a30_give_tot
tot_income_tax = a10_income_tax + a20_income_tax + a30_income_tax
tot_wetax = a10_wetax + a20_wetax + a30_wetax
tot_pay_count = a10_count + a20_count + a30_count

if in_tax_id = "1" then 
   tax_id_name = "����Ű�" 
   elseif in_tax_id = "2" then 
          tax_id_name = "�б�" 
          elseif in_tax_id = "3" then 
		         tax_id_name = "����" 
end if

title_line = " ��õ¡�������Ȳ�Ű� - ������!!!!! "

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
				return "5 1";
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
	<body>
		<div id="wrap">			
			<!--#include virtual = "/include/insa_pay_header.asp" -->
			<!--#include virtual = "/include/insa_pay_tax_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_pay_tax_report.asp?ck_sw=<%="n"%>" method="post" name="frm">
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
                                    <select name="pmg_yymm" id="pmg_yymm" type="text" value="<%=pmg_yymm%>" style="width:90px">
                                    <%	for i = 24 to 1 step -1	%>
                                    <option value="<%=month_tab(i,1)%>" <%If pmg_yymm = month_tab(i,1) then %>selected<% end if %>><%=month_tab(i,2)%></option>
                                    <%	next	%>
                                 </select>
								</label>
								<label>
                                <strong>�Ű���</strong>
                                <select name="in_tax_id" id="in_tax_id" type="text" value="<%=in_tax_id%>" style="width:100px">
                                    <option value="1" <%If in_tax_id = "1" then %>selected<% end if %>>����Ű�</option>
                                    <option value="2" <%If in_tax_id = "2" then %>selected<% end if %>>��������Ű�</option>
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
							<col width="4%" >
							<col width="10%" >
							<col width="4%" >
							<col width="6%" >
							<col width="12%" >
                            <col width="12%" >
                            <col width="10%" >
                            <col width="10%" >
                            <col width="10%" >
                            <col width="10%" >
                            <col width="10%" >
							<col width="*" >
						</colgroup>
						<thead>
							<tr>
								<th rowspan="2" colspan="2" class="first" scope="col">����</th>
                                <th rowspan="2" scope="col">�ڵ�</th>
								<th colspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">�ҵ�����(�����̴�,���������)</th>
                                <th colspan="3" scope="col" style=" border-bottom:1px solid #e3e3e3;">¡������</th>
                                <th rowspan="2" scope="col">9.�������<br>ȯ�޼���</th>
                                <th rowspan="2" scope="col">10.�ҵ漼 ��<br>(���꼼����)</th>
                                <th rowspan="2" scope="col">11.�����<br>Ư����</th>
                                <th rowspan="2" scope="col">���</th>
							</tr>
                            <tr>
				                <th scope="col" style=" border-left:1px solid #e3e3e3;">4.�ο�</th> 
				                <th scope="col">5.�����޾�</th>
                                <th scope="col">6.�ҵ漼��</th>
                                <th scope="col">7.�����Ư����</th>
                                <th scope="col">8.���꼼</th>
                            </tr>
						</thead>
						<tbody>
							<tr>
								<td rowspan="5" class="first" style="background:#f8f8f8;">��<br>��<br>��<br>��</td>
                                <td>���̼���</td>
                                <td>A01</td>
                                <td class="right"><%=formatnumber(pay_count,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(sum_give_tot,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(sum_income_tax,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(sum_special_tax,0)%>&nbsp;</td>
                                <td class="right">&nbsp;</td>
                                <td class="right" style="background:#f8f8f8;">&nbsp;</td>
                                <td class="right" style="background:#f8f8f8;">&nbsp;</td>
                                <td class="right" style="background:#f8f8f8;">&nbsp;</td>
                                <td class="right">&nbsp;</td>
							</tr>
                            <tr>
                                <td style=" border-left:1px solid #e3e3e3;">�ߵ����</td>
                                <td>A02</td>
                                <td class="right"><%=formatnumber(a02_count,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(a02_give_tot,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(a02_income_tax,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(sum_special_tax,0)%>&nbsp;</td>
                                <td class="right">&nbsp;</td>
                                <td class="right" style="background:#f8f8f8;">&nbsp;</td>
                                <td class="right" style="background:#f8f8f8;">&nbsp;</td>
                                <td class="right" style="background:#f8f8f8;">&nbsp;</td>
                                <td class="right">&nbsp;</td>
							</tr>
                            <tr>
                                <td style=" border-left:1px solid #e3e3e3;">�Ͽ�ٷ�</td>
                                <td>A03</td>
                                <td class="right"><%=formatnumber(a03_count,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(a03_give_tot,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(a03_income_tax,0)%>&nbsp;</td>
                                <td class="right" style="background:#f8f8f8;">&nbsp;</td>
                                <td class="right">&nbsp;</td>
                                <td class="right" style="background:#f8f8f8;">&nbsp;</td>
                                <td class="right" style="background:#f8f8f8;">&nbsp;</td>
                                <td class="right" style="background:#f8f8f8;">&nbsp;</td>
                                <td class="right">&nbsp;</td>
							</tr>
                            <tr>
                                <td style=" border-left:1px solid #e3e3e3;">��������</td>
                                <td>A04</td>
                                <td class="right"><%=formatnumber(a04_count,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(a03_give_tot,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(a03_income_tax,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(sum_special_tax,0)%>&nbsp;</td>
                                <td class="right">&nbsp;</td>
                                <td class="right" style="background:#f8f8f8;">&nbsp;</td>
                                <td class="right" style="background:#f8f8f8;">&nbsp;</td>
                                <td class="right" style="background:#f8f8f8;">&nbsp;</td>
                                <td class="right">&nbsp;</td>
							</tr>
                            <tr>
                                <td style=" border-left:1px solid #e3e3e3;">�� �� ��</td>
                                <td>A10</td>
                                <td class="right"><%=formatnumber(a10_count,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(a10_give_tot,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(a10_income_tax,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(sum_special_tax,0)%>&nbsp;</td>
                                <td class="right">&nbsp;</td>
                                <td class="right" style="background:#f8f8f8;">&nbsp;</td>
                                <td class="right"><%=formatnumber(a10_income_tax,0)%>&nbsp;</td>
                                <td class="right" style="background:#f8f8f8;">&nbsp;</td>
                                <td class="right">&nbsp;</td>
							</tr>
                            <tr>
								<td rowspan="3" class="first" style="background:#f8f8f8;">��<br>��<br>��<br>��</td>
                                <td>���ݰ���</td>
                                <td>A21</td>
                                <td class="right"><%=formatnumber(a21_count,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(a21_give_tot,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(a21_income_tax,0)%>&nbsp;</td>
                                <td class="right" style="background:#f8f8f8;">&nbsp;</td>
                                <td class="right">&nbsp;</td>
                                <td class="right" style="background:#f8f8f8;">&nbsp;</td>
                                <td class="right" style="background:#f8f8f8;">&nbsp;</td>
                                <td class="right" style="background:#f8f8f8;">&nbsp;</td>
                                <td class="right">&nbsp;</td>
							</tr>
                            <tr>
                                <td style=" border-left:1px solid #e3e3e3;">�� ��</td>
                                <td>A22</td>
                                <td class="right"><%=formatnumber(a22_count,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(a22_give_tot,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(a22_income_tax,0)%>&nbsp;</td>
                                <td class="right" style="background:#f8f8f8;">&nbsp;</td>
                                <td class="right">&nbsp;</td>
                                <td class="right" style="background:#f8f8f8;">&nbsp;</td>
                                <td class="right" style="background:#f8f8f8;">&nbsp;</td>
                                <td class="right" style="background:#f8f8f8;">&nbsp;</td>
                                <td class="right">&nbsp;</td>
							</tr>
                            <tr>
                                <td style=" border-left:1px solid #e3e3e3;">�� �� ��</td>
                                <td>A20</td>
                                <td class="right"><%=formatnumber(a20_count,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(a20_give_tot,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(a20_income_tax,0)%>&nbsp;</td>
                                <td class="right" style="background:#f8f8f8;">&nbsp;</td>
                                <td class="right">&nbsp;</td>
                                <td class="right">&nbsp;</td>
                                <td class="right">&nbsp;</td>
                                <td class="right" style="background:#f8f8f8;">&nbsp;</td>
                                <td class="right">&nbsp;</td>
							</tr>
                            <tr>
								<td rowspan="3" class="first" style="background:#f8f8f8;">��<br>��<br>��<br>��</td>
                                <td>�ſ�¡��</td>
                                <td>A25</td>
                                <td class="right"><%=formatnumber(alba_count,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(sum_alba_give_total,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(sum_tax_amt1,0)%>&nbsp;</td>
                                <td class="right" style="background:#f8f8f8;">&nbsp;</td>
                                <td class="right">&nbsp;</td>
                                <td class="right" style="background:#f8f8f8;">&nbsp;</td>
                                <td class="right" style="background:#f8f8f8;">&nbsp;</td>
                                <td class="right" style="background:#f8f8f8;">&nbsp;</td>
                                <td class="right">&nbsp;</td>
							</tr>
                            <tr>
                                <td style=" border-left:1px solid #e3e3e3;">��������</td>
                                <td>A26</td>
                                <td class="right"><%=formatnumber(a32_count,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(a32_give_tot,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(a32_income_tax,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(sum_special_tax,0)%>&nbsp;</td>
                                <td class="right" style="background:#f8f8f8;">&nbsp;</td>
                                <td class="right" style="background:#f8f8f8;">&nbsp;</td>
                                <td class="right" style="background:#f8f8f8;">&nbsp;</td>
                                <td class="right">&nbsp;</td>
							</tr>
                            <tr>
                                <td style=" border-left:1px solid #e3e3e3;">�� �� ��</td>
                                <td>A30</td>
                                <td class="right"><%=formatnumber(a30_count,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(a30_give_tot,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(a30_income_tax,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(sum_special_tax,0)%>&nbsp;</td>
                                <td class="right">&nbsp;</td>
                                <td class="right" style="background:#f8f8f8;">&nbsp;</td>
                                <td class="right"><%=formatnumber(a30_income_tax,0)%>&nbsp;</td>
                                <td class="right" style="background:#f8f8f8;">&nbsp;</td>
                                <td class="right">&nbsp;</td>
							</tr>
                            <tr>
								<td rowspan="3" class="first" style="background:#f8f8f8;">��<br>Ÿ<br>��<br>��</td>
                                <td>���ݰ���</td>
                                <td>A41</td>
                                <td class="right"><%=formatnumber(sum_special_tax,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(sum_special_tax,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(sum_special_tax,0)%>&nbsp;</td>
                                <td class="right" style="background:#f8f8f8;">&nbsp;</td>
                                <td class="right">&nbsp;</td>
                                <td class="right" style="background:#f8f8f8;">&nbsp;</td>
                                <td class="right" style="background:#f8f8f8;">&nbsp;</td>
                                <td class="right" style="background:#f8f8f8;">&nbsp;</td>
                                <td class="right">&nbsp;</td>
							</tr>
                            <tr>
                                <td style=" border-left:1px solid #e3e3e3;">�׿�</td>
                                <td>A42</td>
                                <td class="right"><%=formatnumber(sum_special_tax,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(sum_special_tax,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(sum_special_tax,0)%>&nbsp;</td>
                                <td class="right" style="background:#f8f8f8;">&nbsp;</td>
                                <td class="right">&nbsp;</td>
                                <td class="right" style="background:#f8f8f8;">&nbsp;</td>
                                <td class="right" style="background:#f8f8f8;">&nbsp;</td>
                                <td class="right" style="background:#f8f8f8;">&nbsp;</td>
                                <td class="right">&nbsp;</td>
							</tr>
                            <tr>
                                <td style=" border-left:1px solid #e3e3e3;">������</td>
                                <td>A40</td>
                                <td class="right"><%=formatnumber(sum_special_tax,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(sum_special_tax,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(sum_special_tax,0)%>&nbsp;</td>
                                <td class="right" style="background:#f8f8f8;">&nbsp;</td>
                                <td class="right">&nbsp;</td>
                                <td class="right">&nbsp;</td>
                                <td class="right">&nbsp;</td>
                                <td class="right" style="background:#f8f8f8;">&nbsp;</td>
                                <td class="right">&nbsp;</td>
							</tr>
                            <tr>
								<td rowspan="4" class="first" style="background:#f8f8f8;">��<br>��<br>��<br>��</td>
                                <td>���ݰ���</td>
                                <td>A48</td>
                                <td class="right"><%=formatnumber(sum_special_tax,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(sum_special_tax,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(sum_special_tax,0)%>&nbsp;</td>
                                <td class="right" style="background:#f8f8f8;">&nbsp;</td>
                                <td class="right">&nbsp;</td>
                                <td class="right" style="background:#f8f8f8;">&nbsp;</td>
                                <td class="right" style="background:#f8f8f8;">&nbsp;</td>
                                <td class="right" style="background:#f8f8f8;">&nbsp;</td>
                                <td class="right">&nbsp;</td>
							</tr>
                            <tr>
                                <td style=" border-left:1px solid #e3e3e3;">��������(�ſ�)</td>
                                <td>A45</td>
                                <td class="right"><%=formatnumber(sum_special_tax,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(sum_special_tax,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(sum_special_tax,0)%>&nbsp;</td>
                                <td class="right" style="background:#f8f8f8;">&nbsp;</td>
                                <td class="right">&nbsp;</td>
                                <td class="right" style="background:#f8f8f8;">&nbsp;</td>
                                <td class="right" style="background:#f8f8f8;">&nbsp;</td>
                                <td class="right" style="background:#f8f8f8;">&nbsp;</td>
                                <td class="right">&nbsp;</td>
							</tr>
                            <tr>
                                <td style=" border-left:1px solid #e3e3e3;">��������</td>
                                <td>A46</td>
                                <td class="right"><%=formatnumber(sum_special_tax,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(sum_special_tax,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(sum_special_tax,0)%>&nbsp;</td>
                                <td class="right" style="background:#f8f8f8;">&nbsp;</td>
                                <td class="right">&nbsp;</td>
                                <td class="right" style="background:#f8f8f8;">&nbsp;</td>
                                <td class="right" style="background:#f8f8f8;">&nbsp;</td>
                                <td class="right" style="background:#f8f8f8;">&nbsp;</td>
                                <td class="right">&nbsp;</td>
							</tr>
                            <tr>
                                <td style=" border-left:1px solid #e3e3e3;">������</td>
                                <td>A47</td>
                                <td class="right"><%=formatnumber(sum_special_tax,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(sum_special_tax,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(sum_special_tax,0)%>&nbsp;</td>
                                <td class="right" style="background:#f8f8f8;">&nbsp;</td>
                                <td class="right">&nbsp;</td>
                                <td class="right">&nbsp;</td>
                                <td class="right">&nbsp;</td>
                                <td class="right" style="background:#f8f8f8;">&nbsp;</td>
                                <td class="right">&nbsp;</td>
							</tr>
                            <tr>
                                <td colspan="2">���ڼҵ�</td>
                                <td>A50</td>
                                <td class="right"><%=formatnumber(sum_special_tax,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(sum_special_tax,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(sum_special_tax,0)%>&nbsp;</td>
                                <td class="right">&nbsp;</td>
                                <td class="right">&nbsp;</td>
                                <td class="right">&nbsp;</td>
                                <td class="right">&nbsp;</td>
                                <td class="right">&nbsp;</td>
                                <td class="right">&nbsp;</td>
							</tr>
                            <tr>
                                <td colspan="2">���ҵ�</td>
                                <td>A60</td>
                                <td class="right"><%=formatnumber(sum_special_tax,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(sum_special_tax,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(sum_special_tax,0)%>&nbsp;</td>
                                <td class="right">&nbsp;</td>
                                <td class="right">&nbsp;</td>
                                <td class="right">&nbsp;</td>
                                <td class="right">&nbsp;</td>
                                <td class="right">&nbsp;</td>
                                <td class="right">&nbsp;</td>
							</tr>
                            <tr>
                                <td colspan="2">����������¡����</td>
                                <td>A69</td>
                                <td class="right"><%=formatnumber(sum_special_tax,0)%>&nbsp;</td>
                                <td class="right" style="background:#f8f8f8;">&nbsp;</td>
                                <td class="right"><%=formatnumber(sum_special_tax,0)%>&nbsp;</td>
                                <td class="right" style="background:#f8f8f8;">&nbsp;</td>
                                <td class="right">&nbsp;</td>
                                <td class="right">&nbsp;</td>
                                <td class="right">&nbsp;</td>
                                <td class="right" style="background:#f8f8f8;">&nbsp;</td>
                                <td class="right">&nbsp;</td>
							</tr>
                            <tr>
                                <td colspan="2">������ھ絵�ҵ�</td>
                                <td>A70</td>
                                <td class="right"><%=formatnumber(sum_special_tax,0)%>&nbsp;</td>
                                <td class="right">&nbsp;</td>
                                <td class="right"><%=formatnumber(sum_special_tax,0)%>&nbsp;</td>
                                <td class="right" style="background:#f8f8f8;">&nbsp;</td>
                                <td class="right">&nbsp;</td>
                                <td class="right">&nbsp;</td>
                                <td class="right">&nbsp;</td>
                                <td class="right" style="background:#f8f8f8;">&nbsp;</td>
                                <td class="right">&nbsp;</td>
							</tr>
                            <tr>
                                <td colspan="2">���ο�õ</td>
                                <td>A80</td>
                                <td class="right"><%=formatnumber(sum_special_tax,0)%>&nbsp;</td>
                                <td class="right">&nbsp;</td>
                                <td class="right"><%=formatnumber(sum_special_tax,0)%>&nbsp;</td>
                                <td class="right" style="background:#f8f8f8;">&nbsp;</td>
                                <td class="right">&nbsp;</td>
                                <td class="right">&nbsp;</td>
                                <td class="right">&nbsp;</td>
                                <td class="right" style="background:#f8f8f8;">&nbsp;</td>
                                <td class="right">&nbsp;</td>
							</tr>
                            <tr>
                                <td colspan="2">�����Ű�(����)</td>
                                <td>A90</td>
                                <td class="right" style="background:#f8f8f8;"></td>
                                <td class="right" style="background:#f8f8f8;">&nbsp;</td>
                                <td class="right">&nbsp;</td>
                                <td class="right">&nbsp;</td>
                                <td class="right">&nbsp;</td>
                                <td class="right" style="background:#f8f8f8;">&nbsp;</td>
                                <td class="right">&nbsp;</td>
                                <td class="right">&nbsp;</td>
                                <td class="right">&nbsp;</td>
							</tr>
                            <tr>
								<th colspan="2" class="first">�� �� ��</th>
                                <th>A99</th>
                                <th class="right"><%=formatnumber(tot_pay_count,0)%>&nbsp;</th>
                                <th class="right"><%=formatnumber(tot_give_tot,0)%>&nbsp;</th>
                                <th class="right"><%=formatnumber(tot_income_tax,0)%>&nbsp;</th>
                                <th class="right"><%=formatnumber(sum_special_tax,0)%>&nbsp;</th>
                                <th class="right">&nbsp;</td>
                                <th class="right">&nbsp;</td>
                                <th class="right"><%=formatnumber(tot_income_tax,0)%>&nbsp;</td>
                                <th class="right">&nbsp;</td>
                                <th class="right">&nbsp;</td>
							</tr>
 						</tbody>
					</table>
                    <table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="9%" >
							<col width="9%" >
							<col width="9%" >
							<col width="9%" >
							<col width="9%" >
                            <col width="9%" >
                            <col width="9%" >
                            <col width="9%" >
                            <col width="9%" >
                            <col width="10%" >
                            <col width="9%" >
						</colgroup>
						<thead>
							<tr>
								<th colspan="3" class="first" scope="col" style=" border-bottom:1px solid #e3e3e3;">���� ��ȯ�ޱ� ������ ���</th>
                                <th colspan="4" scope="col" style=" border-bottom:1px solid #e3e3e3;">��� �߻� ȯ�޼���</th>
								<th rowspan="2" scope="col">18.�������ȯ��<br>(14+15+16+17)</th>
                                <th rowspan="2" scope="col">19.�������<br>ȯ�޾װ�</th>
                                <th rowspan="2" scope="col">20.�����̿�<br>ȯ�޾�(18-19)</th>
                                <th rowspan="2" scope="col">21.ȯ�޽�û��</th>
							</tr>
                            <tr>
				                <th scope="col" style=" border-left:1px solid #e3e3e3;">12.������ȯ��</th> 
				                <th scope="col">13.��ȯ�޽�û</th>
                                <th scope="col">14.�ܾ�(12-13)</th>
                                <th scope="col">15.�Ϲ�ȯ��</th>
                                <th scope="col">16.��Ź���</th>
                                <th scope="col">17.������</th>
                                <th scope="col">17.�պ���</th>
                            </tr>
						</thead>
						<tbody>
							<tr>
								<td class="right"><%=formatnumber(sum_special_tax,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(sum_special_tax,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(sum_special_tax,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(sum_special_tax,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(sum_special_tax,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(sum_special_tax,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(sum_special_tax,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(sum_special_tax,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(sum_special_tax,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(sum_special_tax,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(sum_special_tax,0)%>&nbsp;</td>
							</tr>
                       </tbody>
				  </table>
				</div>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
                  	<td width="15%">
					<div class="btnCenter">
                    <a href="insa_excel_pay_tax_report.asp?view_condi=<%=view_condi%>&pmg_yymm=<%=pmg_yymm%>&to_date=<%=to_date%>" class="btnType04">�����ٿ�ε�</a>
					</div>                  
                  	</td>
                    <td>
					<div class="btnRight">
					<a href="#" onClick="pop_Window('insa_pay_tax_withholding_print.asp?view_condi=<%=view_condi%>&pmg_yymm=<%=pmg_yymm%>&in_tax_id=<%=in_tax_id%>','insa_pay_withholding_pop','scrollbars=yes,width=1060,height=900')" class="btnType04">���</a>
					</div>                  
                    </td> 
			      </tr>
				  </table>
			</form>
		</div>				
	</div>        				
	</body>
</html>

