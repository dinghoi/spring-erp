<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
dim com_tab(20)
dim area_cnt(20,20)
dim sum_cnt(20)
dim area_tab
area_tab = array("����","���","�λ�","�뱸","��õ","����","����","���","����","�泲","���","����","�泲","���","����","����","����")

be_pg = "insa_area_count_org.asp"

curr_dd = cstr(datepart("d",now))
to_date = mid(cstr(now()),1,10)
from_date = mid(cstr(now()-curr_dd+1),1,10)

view_condi = request("view_condi")
condi = request("condi")  

ck_sw=Request("ck_sw")

if ck_sw = "n" then
	view_condi = request.form("view_condi")
	condi = request.form("condi")
  else
	view_condi = request("view_condi")
	condi = request("condi")  
end if

if view_condi = "" then
	view_condi = "��ü"
	condi = "��ü"
end if

'response.write(view_condi)
'response.write(company)

for i = 0 to 20
    com_tab(i) = ""
next

for i = 0 to 20
    for j = 0 to 20
	    area_cnt(i,j) = 0
    next
	sum_cnt(i) = 0
next

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_as = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

' ȸ�����̺� ȸ�� �Ǵ� ���θ�Ī ��������
if view_condi = "��ü" then
	' 2019.02.22 ������ ��û ȸ�縮��Ʈ�� ������ �ҽ� org_end_date�� null �� �ƴ� �������ڸ� �����ϸ� ����Ʈ�� ��Ÿ���� �ʴ´�.
	Sql = "SELECT * FROM emp_org_mst WHERE ISNULL(org_end_date) AND org_level = 'ȸ��'  ORDER BY org_company ASC"
   Rs_org.Open Sql, Dbconn, 1
   k = 0
   while not Rs_org.eof
	   k = k + 1
	   com_tab(k) = Rs_org("org_name")
	   Rs_org.movenext()
   Wend
 elseif condi = "��ü" then
           Sql="select * from emp_org_mst where (org_level = '����') and (org_company='"+view_condi+"') order by org_code ASC"
           Rs_org.Open Sql, Dbconn, 1
           k = 0
           while not Rs_org.eof
	             k = k + 1
	             com_tab(k) = Rs_org("org_name")
	            Rs_org.movenext()
           Wend   
		else 
		   Sql="select * from emp_org_mst where (org_level = '�����') and (org_company='"+view_condi+"') and (org_bonbu='"+condi+"') order by org_code ASC"
           Rs_org.Open Sql, Dbconn, 1
           k = 0
           while not Rs_org.eof
	             k = k + 1
	             com_tab(k) = Rs_org("org_name")
	            Rs_org.movenext()
           Wend   
end if
Rs_org.close()
k_org = k	

if view_condi = "��ü" then
   Sql = "select * from emp_master where (isNull(emp_end_date) or emp_end_date = '1900-01-01')  and (emp_no < '900000')"
   elseif condi = "��ü" then  
            Sql = "select * from emp_master where (emp_company='"+view_condi+"') and (isNull(emp_end_date) or emp_end_date = '1900-01-01')  and (emp_no < '900000')"
		  else 
		    Sql = "select * from emp_master where (emp_company='"+view_condi+"') and (emp_bonbu='"+condi+"') and (isNull(emp_end_date) or emp_end_date = '1900-01-01')  and (emp_no < '900000')"
end if
Rs.Open Sql, Dbconn, 1

do until rs.eof 
   if view_condi = "��ü" then
      com_name = rs("emp_company")
      elseif condi = "��ü" then 
                com_name = rs("emp_bonbu")
			 else
			    com_name = rs("emp_saupbu")
   end if

   k = 0                                       
   for i = 1 to k_org
       if com_tab(i) = com_name then
          k = i
	   end if
    next
	
    if k = 0 then   '�ӽ÷�... ����Ÿ�� �߸��Ǿ� �񱳰� �ȵ�
	   k = k_org + 1
	   if condi = "��ü" then 
	          com_tab(k) = view_condi
		  else
		      com_tab(k) = condi
	   end if
	 end if
	 
    j = 0
	
	select case rs("emp_sido")
		case "����"
			j = 0
		case "���"
			j = 1
		case "�λ�"
			j = 2
		case "�뱸"
			j = 3
		case "��õ"
			j = 4
		case "����"
			j = 5
		case "����"
			j = 6
		case "���"
			j = 7
		case "����"
			j = 8
		case "�泲"
			j = 9
		case "���"
			j = 10
		case "����"
			j = 11
		case "�泲"
			j = 12
		case "���"
			j = 13
		case "����"
			j = 14
		case "����"
			j = 15
		case "����"
			j = 16
	end select		
	
	area_cnt(k,j) = area_cnt(k,j) + 1
	sum_cnt(j) = sum_cnt(j) + 1
	
    rs.movenext()
loop
rs.close()

title_line = ""+ view_condi +" - ������ �ο����� "

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
				if (formcheck(document.frm)) {
					document.frm.submit ();
				}
			}
			 
			function chkfrm() {
				if(document.frm.view_condi.value =="ȸ�纰") {
					if(document.frm.condi.value =="��ü") {
						alert('ȸ�縦 �����ϼ���');
						frm.condi.focus();
						return false;}}		
				return true;
			}
		</script>

	</head>
	<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
		<div id="wrap">			
			<!--#include virtual = "/include/insa_header.asp" -->
			<!--#include virtual = "/include/insa_report_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<% '<form action="waiting.asp?pg_name=insa_grade_count.asp" method="post" name="frm"> %>
                <form action="insa_area_count_org.asp?ck_sw=<%="n"%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>					
						<dt>���� �˻�</dt>
                        <dd>
                            <p>
                               <strong>ȸ�� : </strong>
                              <%
								Sql="select * from emp_org_mst where (isNull(org_end_date) or org_end_date = '1900-01-01') and (org_level = 'ȸ��') ORDER BY org_code ASC"
	                            rs_org.Open Sql, Dbconn, 1	
							  %>
								<select name="view_condi" id="view_condi" type="text" style="width:150px">
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
								<strong>���� : </strong>
                              <%
								Sql="select * from emp_org_mst where isNull(org_end_date) and org_level = '����' and org_company = '"+view_condi+"' ORDER BY org_code ASC"
	                            rs_org.Open Sql, Dbconn, 1	
							  %>
								<select name="condi" id="condi" type="text" style="width:150px">
                                  <option value="��ü" <%If condi = "��ü" then %>selected<% end if %>>��ü</option>
                			  <% 
								do until rs_org.eof 
			  				  %>
                					<option value='<%=rs_org("org_name")%>' <%If condi = rs_org("org_name") then %>selected<% end if %>><%=rs_org("org_name")%></option>
                			  <%
									rs_org.movenext()  
								loop 
								rs_org.Close()
							  %>
            					</select>  
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser1.jpg" alt="�˻�"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="*" >
                            <% for i = 1 to 18 %>
							       <col width="5%" >
                            <% next	%>
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">��&nbsp;&nbsp;&nbsp;��</th>
                                <% 
								for i = 0 to 16 
								%>
                                <th scope="col"><%=area_tab(i)%></th>
								<%
								next
								%>
                                <th scope="col" style=" border-left:1px solid #e3e3e3;">�Ұ�</th>
							</tr>
						</thead>
						<tbody>
                        <%
                        for i = 0 to 20 
                        	if	com_tab(i) <> "" then
						%>	
                            <tr>
                                <% 
								hap_cnt = 0
								for j = 0 to 16 
								    hap_cnt = hap_cnt + area_cnt(i,j)
								next
								
								'if tot_pay = 0 then
								'      cr_pro = 0
								'   else
								'      cr_pro = (hap_pay / tot_pay) * 100
								'end if
					
								%>
                                <td><%=com_tab(i)%></td>
                                <% 
								for j = 0 to 16 
								    'ost_amt = cdbl(cost_amt) / 10000 �ݾ״��� ¥���°�
								%>
                                    <td class="right"><%=formatnumber(area_cnt(i,j),0)%></td>
								<%
								next
								%>
                                <td class="right"><%=formatnumber(hap_cnt,0)%></td> 
                             </tr>
                        <%
							end if
						next
                        %>
							<tr>
                              <th>�Ѱ�</th>
                              <% 
								hap_cnt = 0
								for j = 0 to 16 
								    hap_cnt = hap_cnt + sum_cnt(j)
								%>
                                    <th class="right"><%=formatnumber(sum_cnt(j),0)%></th>
								<%
								next
								%>
                                <th class="right"><%=formatnumber(hap_cnt,0)%></th> 
							</tr>
 						</tbody>
					</table>
				</div>
			</form>
		</div>				
	</div>        				
	</body>
</html>

