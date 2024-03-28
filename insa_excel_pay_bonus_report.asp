<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
Dim from_date
Dim to_date
Dim win_sw
	 
view_condi=Request("view_condi")
pmg_yymm=request("pmg_yymm")
in_pmg_id=request("pmg_id")

curr_yyyy = mid(cstr(pmg_yymm),1,4)
curr_mm = mid(cstr(pmg_yymm),5,2)

if in_pmg_id = "2" then 
   pmg_id_name = "�󿩱�" 
   elseif in_pmg_id = "3" then 
          pmg_id_name = "��õ���μ�Ƽ��" 
          elseif in_pmg_id = "4" then 
		         pmg_id_name = "��������" 
end if

title_line = cstr(curr_yyyy) + "�� " + cstr(curr_mm) + "�� " + pmg_id_name + "��Ȳ(" + view_condi + ")"

savefilename = title_line + ".xls"

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// ������ ����
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set Rs_year = Server.CreateObject("ADODB.Recordset")
Set Rs_give = Server.CreateObject("ADODB.Recordset")
Set Rs_dct = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

pay_count = 0

Sql = "select * from pay_month_give where (pmg_yymm = '"+pmg_yymm+"' ) and (pmg_id = '"+in_pmg_id+"') and (pmg_company = '"+view_condi+"') ORDER BY pmg_company,pmg_org_code,pmg_emp_no ASC"
Rs.Open Sql, Dbconn, 1
do until rs.eof
    emp_no = rs("pmg_emp_no")
    pmg_give_tot = rs("pmg_give_total")
    pay_count = pay_count + 1
				  
    sum_base_pay = sum_base_pay + int(rs("pmg_base_pay"))
    sum_meals_pay = 0
    sum_give_tot = sum_give_tot + int(rs("pmg_give_total"))

    Sql = "select * from pay_month_deduct where (de_yymm = '"+pmg_yymm+"' ) and (de_id = '"+in_pmg_id+"') and (de_emp_no = '"+emp_no+"') and (de_company = '"+view_condi+"')"
    Set Rs_dct = DbConn.Execute(SQL)
    if not Rs_dct.eof then

            de_epi_amt = int(Rs_dct("de_epi_amt"))
            de_income_tax = int(Rs_dct("de_income_tax"))
            de_wetax = int(Rs_dct("de_wetax"))
		    de_deduct_tot = int(Rs_dct("de_deduct_total"))	
	     else
            de_epi_amt = 0
            de_income_tax = 0
            de_wetax = 0
		    de_deduct_tot = 0
     end if
     Rs_dct.close()

     sum_epi_amt = sum_epi_amt + de_epi_amt
     sum_income_tax = sum_income_tax + de_income_tax
     sum_wetax = sum_wetax + de_wetax
	 sum_deduct_tot = sum_deduct_tot + de_deduct_tot

	rs.movenext()
loop
rs.close()

Sql = "select * from pay_month_give where (pmg_yymm = '"+pmg_yymm+"' ) and (pmg_id = '"+in_pmg_id+"') and (pmg_company = '"+view_condi+"') ORDER BY pmg_company,pmg_org_code,pmg_emp_no ASC"

Rs.Open Sql, Dbconn, 1

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>�޿� ���� �ý���</title>
	</head>
	<body>
		<div id="wrap">			
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<div class="gView">
					<table border="1" cellpadding="0" cellspacing="0" class="tableList">
						<thead>
							<tr>
				               <th rowspan="2" class="first" scope="col" >����</th>
                               <th rowspan="2" scope="col" >�Ի���</th>
                               <th rowspan="2" scope="col" >����</th>
                               <th rowspan="2" scope="col" >�Ҽ�</th>
				               <th colspan="3" scope="col" style="background:#FFFFE6;">���� ����</th>
                               <th colspan="4" scope="col" style="background:#E0FFFF;">���� �� �������޾�</th>
                               <th rowspan="2" scope="col" >���޾�</th>
                               <th rowspan="2" scope="col" >���</th>
			                </tr>
                            <tr>
						<%
						  if in_pmg_id = "2" then %>
                                <td scope="col" >�󿩱�</td>
                        <%   elseif in_pmg_id = "3" then %>
                                <td scope="col" ">��õ��<br>�μ�Ƽ��</td>
                        <%          elseif in_pmg_id = "4" then %>
                                <td scope="col" >��������</td>
                        <% end if %>        
								<td scope="col" >&nbsp;</td>  
                                <td scope="col" >���޼Ұ�</td>
								<td scope="col" >��뺸��</td>
                                <td scope="col" >�ҵ漼</td>
								<td scope="col" >����ҵ漼</td>
                                <td scope="col" >������</td>
							</tr>
						</thead>
						<tbody>
						<%
						do until rs.eof
							  emp_no = rs("pmg_emp_no")
							  pmg_give_tot = rs("pmg_give_total")
						'	  pay_count = pay_count + 1
						  
							  Sql = "SELECT * FROM emp_master where emp_no = '"&emp_no&"'"
                              Set rs_emp = DbConn.Execute(SQL)
		                      if not rs_emp.eof then
		                    		emp_in_date = rs_emp("emp_in_date")
	                             else
	                    			emp_in_date = ""
                              end if
                              rs_emp.close()
					  %>
							<tr>
								<td class="first"><%=rs("pmg_emp_name")%>(<%=rs("pmg_emp_no")%>)</td>
                                <td><%=emp_in_date%></td>
                                <td><%=rs("pmg_grade")%></td>
                                <td><%=rs("pmg_org_name")%></td>
                                <td align="right"><%=formatnumber(rs("pmg_base_pay"),0)%></td>
                                <td align="right"><%=formatnumber(rs("pmg_meals_pay"),0)%></td>
                                <td align="right"><%=formatnumber(rs("pmg_give_total"),0)%></td>
                         <%
						      Sql = "select * from pay_month_deduct where (de_yymm = '"+pmg_yymm+"' ) and (de_id = '"+in_pmg_id+"') and (de_emp_no = '"+emp_no+"') and (de_company = '"+view_condi+"')"
                              Set Rs_dct = DbConn.Execute(SQL)
							  if not Rs_dct.eof then
                                    de_epi_amt = int(Rs_dct("de_epi_amt"))
                                    de_income_tax = int(Rs_dct("de_income_tax"))
                                    de_wetax = int(Rs_dct("de_wetax"))
		                            de_deduct_tot = int(Rs_dct("de_deduct_total"))	
	                             else
                                    de_epi_amt = 0
                                    de_income_tax = 0
                                    de_wetax = 0
		                            de_deduct_tot = 0
                              end if
                              Rs_dct.close()
							  pmg_curr_pay = pmg_give_tot - de_deduct_tot
							  
                          %>
                                <td align="right"><%=formatnumber(de_epi_amt,0)%></td>
                                <td align="right"><%=formatnumber(de_income_tax,0)%></td>
                                <td align="right"><%=formatnumber(de_wetax,0)%></td>
                                <td align="right"><%=formatnumber(de_deduct_tot,0)%></td>
                                <td align="right"><%=formatnumber(pmg_curr_pay,0)%></td>
                                <td class="right">&nbsp;</td>
                                
							</tr>
						<%
							rs.movenext()
						loop
						rs.close()
						
						sum_curr_pay = sum_give_tot - sum_deduct_tot
					
						%>
                          	<tr>
                                <th colspan="3" class="first">�Ѱ�</th>
                                <th align="right"><%=formatnumber(pay_count,0)%>&nbsp;��</th>
                                <th align="right"><%=formatnumber(sum_base_pay,0)%></th>
                                <th align="right"><%=formatnumber(sum_meals_pay,0)%></th>
                                <th align="right"><%=formatnumber(sum_give_tot,0)%></th>
                                <th align="right"><%=formatnumber(sum_epi_amt,0)%></th>
                                <th align="right"><%=formatnumber(sum_income_tax,0)%></th>
                                <th align="right"><%=formatnumber(sum_wetax,0)%></th>
                                <th align="right"><%=formatnumber(sum_deduct_tot,0)%></th>
                                <th align="right"><%=formatnumber(sum_curr_pay,0)%></th>
                                <th align="right">&nbsp;</th>
							</tr>
						</tbody>
					</table>
				</div>
		</div>				
	</div>        				
	</body>
</html>
