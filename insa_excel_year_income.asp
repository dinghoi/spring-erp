<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

Dim Rs
Dim stay_name

view_condi=Request("view_condi")   
in_empno = Request("in_empno")
inc_yyyy = Request("inc_yyyy")

curr_date = datevalue(mid(cstr(now()),1,10))

if view_condi = "" then
	view_condi = "���̿��������"
	in_empno = ""
end if

savefilename = "����������Ȳ -- "+ in_empno +""+ view_condi +"" + cstr(curr_date) + ".xls"

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// ������ ����
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_stay = Server.CreateObject("ADODB.Recordset")
Set Rs_year = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

if in_empno = "" then
   Sql = "select * from emp_master where (isNull(emp_end_date) or emp_end_date = '1900-01-01') and (emp_company = '"+view_condi+"')  and (emp_no < '900000') ORDER BY emp_company,emp_bonbu,emp_no,emp_name ASC"
   else  
   Sql = "select * from emp_master where (isNull(emp_end_date) or emp_end_date = '1900-01-01') and (emp_company = '"+view_condi+"') and (emp_no = '"+in_empno+"') ORDER BY emp_company,emp_bonbu,emp_no,emp_name ASC"
end if

Rs.Open Sql, Dbconn, 1

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
													
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<style type="text/css">
<!--
.style1 {font-size: 12px}
.style2 {
	font-size: 14px;
	font-weight: bold;
}
-->
</style>
</head>
<body>
<table  border="0" cellpadding="0" cellspacing="0">
  <tr bgcolor="#EFEFEF" class="style11">
    <td colspan="13" bgcolor="#FFFFFF"><div align="left" class="style2">&nbsp;<%=inc_yyyy%>��&nbsp;���� ������Ȳ>&nbsp;<%=view_condi%></div></td>
  </tr>
  <tr bgcolor="#EFEFEF" class="style11">
    <td><div align="center" class="style1">�ͼӳ⵵</div></td>
    <td><div align="center" class="style1">���</div></td>
    <td><div align="center" class="style1">����</div></td>
    <td><div align="center" class="style1">����</div></td>
    <td><div align="center" class="style1">��������</div></td>
    <td><div align="center" class="style1">�ֹι�ȣ</div></td>
    <td><div align="center" class="style1">����</div></td>
    <td><div align="center" class="style1">����</div></td>
    <td><div align="center" class="style1">��å</div></td>
    <td><div align="center" class="style1">�Ҽ�</div></td>
    <td><div align="center" class="style1">ȸ��</div></td>
    <td><div align="center" class="style1">����</div></td>
    <td><div align="center" class="style1">�����</div></td>
    <td><div align="center" class="style1">��</div></td>
    <td><div align="center" class="style1">����ó</div></td>
    <td><div align="center" class="style1">�����Ի���</div></td>
    <td><div align="center" class="style1">�Ի���</div></td>
    <td><div align="center" class="style1">����</div></td>
    
    <td><div align="center" class="style1">����</div></td>
    <td><div align="center" class="style1">�⺻��</div></td>
    <td><div align="center" class="style1">�Ĵ�</div></td>
    <td><div align="center" class="style1">����ٷμ���</div></td>
    <td><div align="center" class="style1">������</div></td>
    
    <td><div align="center" class="style1">������ռҵ�</div></td>
    <td><div align="center" class="style1">���ο���ǥ�ؿ���</div></td>
    <td><div align="center" class="style1">���ο��ݷ�</div></td>
    <td><div align="center" class="style1">�ǰ�����ǥ�ؿ���</div></td>
    <td><div align="center" class="style1">�ǰ������</div></td>
    
    <td><div align="center" class="style1">��뺸�����뿩��</div></td> 
    <td><div align="center" class="style1">���纸�����뿩��</div></td>
    <td><div align="center" class="style1">����纸�����뿩��</div></td>
    <td><div align="center" class="style1">�߼ұ��û�ҳ�ҵ漼���鿩��</div></td>
    <td><div align="center" class="style1">���������</div></td>
    <td><div align="center" class="style1">20������</div></td>
    <td><div align="center" class="style1">60���̻�</div></td>
    <td><div align="center" class="style1">��ο��</div></td>
    <td><div align="center" class="style1">�����</div></td>
    <td><div align="center" class="style1">�γ���</div></td>
    <td><div align="center" class="style1">�ξ簡����</div></td>
    <%' �Ʒ��κ��� �ϴ� ���Ƴ���... %>
    <% '<td><div align="center" class="style1"> %>
    <%    '<div align="left">�԰� ���γ��� </div> %>
    <%'</div></td> %>
  </tr>
    <%
		do until rs.eof 
		  emp_no = rs("emp_no")
		  Sql = "SELECT * FROM pay_year_income where incom_emp_no = '"&emp_no&"' and incom_year = '"&inc_yyyy&"'"
          Set rs_year = DbConn.Execute(SQL)
		  if not rs_year.eof then
                incom_base_pay = rs_year("incom_base_pay")
			    incom_overtime_pay = rs_year("incom_overtime_pay")
				incom_meals_pay = rs_year("incom_meals_pay")
				incom_severance_pay = rs_year("incom_severance_pay")
				incom_total_pay = rs_year("incom_total_pay")
				incom_month_amount = rs_year("incom_month_amount")
	            incom_nps_amount = rs_year("incom_nps_amount")
	            incom_nhis_amount = rs_year("incom_nhis_amount")
            	incom_family_cnt = rs_year("incom_family_cnt")
	            incom_nps = rs_year("incom_nps")
                incom_nhis = rs_year("incom_nhis")
                incom_go_yn = rs_year("incom_go_yn")
                incom_san_yn = rs_year("incom_san_yn")
                incom_long_yn = rs_year("incom_long_yn")
                incom_incom_yn = rs_year("incom_incom_yn")
                incom_wife_yn = rs_year("incom_wife_yn")
                incom_age20 = rs_year("incom_age20")
                incom_age60 = rs_year("incom_age60")
                incom_old = rs_year("incom_old")
                incom_disab = rs_year("incom_disab")
                incom_woman = rs_year("incom_woman")
	         else
                incom_base_pay = 0
			    incom_overtime_pay = 0
				incom_meals_pay = 0
				incom_severance_pay = 0
				incom_total_pay = 0
				incom_month_amount = 0
	            incom_nps_amount = 0
	            incom_nhis_amount = 0
            	incom_family_cnt = 0
	            incom_nps = 0
                incom_nhis = 0
                incom_go_yn = ""
                incom_san_yn = ""
                incom_long_yn = ""
                incom_incom_yn = ""
                incom_wife_yn = "0"
                incom_age20 = 0
                incom_age60 = 0
                incom_old = 0
                incom_disab = 0
                incom_woman = "0"
           end if
           rs_year.close()
    %>
  <tr valign="middle" class="style11">
    <td width="115"><div align="center" class="style1"><%=inc_yyyy%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("emp_no")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("emp_name")%></div></td>
    <td width="59"><div align="center" class="style1"><%=emp_sex%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("emp_type")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("emp_person1")%>-<%=rs("emp_person2")%></div></td>
    <td width="59"><div align="center" class="style1"><%=rs("emp_grade")%></div></td>
    <td width="59"><div align="center" class="style1"><%=rs("emp_job")%></div></td>
    <td width="59"><div align="center" class="style1"><%=rs("emp_position")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("emp_org_name")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("emp_company")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("emp_bonbu")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("emp_saupbu")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("emp_team")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("emp_reside_place")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("emp_first_date")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("emp_in_date")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("emp_jikmu")%></div></td>
    
    <td width="115"><div align="right" class="style1"><%=formatnumber(incom_total_pay,0)%></div></td>
    <td width="115"><div align="right" class="style1"><%=formatnumber(incom_base_pay,0)%></div></td>
    <td width="115"><div align="right" class="style1"><%=formatnumber(incom_meals_pay,0)%></div></td>
    <td width="115"><div align="right" class="style1"><%=formatnumber(incom_overtime_pay,0)%></div></td>
    <td width="115"><div align="right" class="style1"><%=formatnumber(incom_severance_pay,0)%></div></td>
    
    <td width="115"><div align="right" class="style1"><%=formatnumber(incom_month_amount,0)%></div></td>
    <td width="115"><div align="right" class="style1"><%=formatnumber(incom_nps_amount,0)%></div></td>
    <td width="115"><div align="right" class="style1"><%=formatnumber(incom_nps,0)%></div></td>
    <td width="115"><div align="right" class="style1"><%=formatnumber(incom_nhis_amount,0)%></div></td>
    <td width="115"><div align="right" class="style1"><%=formatnumber(incom_nhis,0)%></div></td>

    <td width="115"><div align="center" class="style1"><%=incom_go_yn%></div></td>
    <td width="115"><div align="center" class="style1"><%=incom_san_yn%></div></td>
    <td width="115"><div align="center" class="style1"><%=incom_long_yn%></div></td>
    <td width="115"><div align="center" class="style1"><%=incom_incom_yn%></div></td>
    <td width="115"><div align="right" class="style1"><%=incom_wife_yn%></div></td>
    <% 'response.write(rs("emp_stay_code"))
	   'response.End %>
    <%
	if incom_wife_yn = "1" then 
	      incom_family_cnt = incom_age20 + incom_age60 + 1
	   else 
          incom_family_cnt = incom_age20 + incom_age60 
    end if
	%>
    <td width="115"><div align="right" class="style1"><%=formatnumber(incom_age20,0)%></div></td>
    <td width="115"><div align="right" class="style1"><%=formatnumber(incom_age60,0)%></div></td>
    <td width="115"><div align="right" class="style1"><%=formatnumber(incom_old,0)%></div></td>
    <td width="115"><div align="right" class="style1"><%=formatnumber(incom_disab,0)%></div></td>
    <td width="115"><div align="right" class="style1"><%=incom_woman%></div></td>
    <td width="115"><div align="right" class="style1"><%=formatnumber(incom_family_cnt,0)%></div></td>
  </tr>
	<%
	Rs.MoveNext()
	loop
	%>
</table>
</body>
</html>
<%
Rs.Close()
Set Rs = Nothing
%>
