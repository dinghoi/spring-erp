<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows

Page=Request("page")
ck_sw=Request("ck_sw")

if ck_sw = "y" then
	rever_yymm=request("rever_yymm")
  else
	rever_yymm=Request.form("rever_yymm")
end if

if rever_yymm = "" then
	rever_yymm = mid(now(),1,4) + mid(now(),6,2)
end if

give_date = to_date '������

pgsize = 10 ' ȭ�� �� ������ 
If Page = "" Then
	Page = 1
	start_page = 1
End If

stpage = int((page - 1) * pgsize)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_alb = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_alco = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")   
dbconn.open DbConnect   

' �����Ǻ�
	posi_sql = " and (reg_id = '"&user_id&"') "
	
	if position = "����" then
		view_condi = "����"
	end if
	
	if position = "��Ʈ��" then
		if org_name = "��ȭ����ȣ��" then
			posi_sql = " and (org_name = '��ȭ����ȣ��' or org_name = '��ȭ��������') "
		  else
			posi_sql = " and org_name = '"&org_name&"'"
		end if
	end if
	
	if position = "����" then
		posi_sql = " and team = '"&team&"'"
	end if
	
	if position = "�������" or cost_grade = "2" then
		posi_sql = " and saupbu = '"&saupbu&"'"
	end if
	
	if position = "������" or cost_grade = "1" then 
		posi_sql = " and bonbu = '"&bonbu&"'"
	end if
	
	view_grade = position

	if cost_grade = "0" then
		posi_sql = ""
	end if

Sql = "select count(*) from pay_alba_cost where (rever_yymm = '"+rever_yymm+"' )"&posi_sql
Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF


Sql = "select * from pay_alba_cost where (rever_yymm = '"+rever_yymm+"' )"&posi_sql&" ORDER BY company,draft_no ASC limit "& stpage & "," &pgsize 

Rs.Open Sql, Dbconn, 1

title_line = " �Ƹ�����Ʈ ��� ��Ȳ "


%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>��� ���� �ý���</title>
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
				if (formcheck(document.frm)) {
					document.frm.submit ();
				}
			}			
			function delcheck () {
				if (form_chk(document.frm_del)) {
					document.frm_del.submit ();
				}
			}			

			function form_chk(){				
				a=confirm('�����Ͻðڽ��ϱ�?')
				if (a==true) {
					return true;
				}
				return false;
			}//-->
		</script>

	</head>
	<body>
		<div id="wrap">			
			<!--#include virtual = "/include/insa_pay_header.asp" -->
			<!--#include virtual = "/include/insa_pay_alba_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_pay_albacost_mg.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>					
						<dt>���� �˻�</dt>
                        <dd>
                            <p>
                                <label>
								<strong>�ͼӳ�� : </strong>
                               	<input name="rever_yymm" type="text" value="<%=rever_yymm%>" maxlength="6" onKeyUp="checkNum(this);" style="width:80px">
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser1.jpg" alt="�˻�"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="5%" >
							<col width="5%" >
							<col width="6%" >
							<col width="9%" >
                            <col width="7%" >
                            <col width="6%" >
							<col width="3%" >
                            <col width="6%" >
                            <col width="6%" >
                            <col width="6%" >
                            <col width="9%" >
							<col width="*" >
                            <col width="3%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">��Ϲ�ȣ</th>
								<th scope="col">����</th>
								<th scope="col">������</th>
								<th scope="col">�������</th>
                                <th scope="col">����</th>
                                <th scope="col">�����Ѿ�</th>
								<th scope="col">����<br>(%)</th>
                                <th scope="col">�ҵ漼</th>
                                <th scope="col">����<br>�ҵ漼</th>
                                <th scope="col">�������޾�</th>
								<th scope="col">����</th>
                                <th scope="col">���(���ڰ���/�ϼ�/�۾���/�۾�����)</th>
                                <th scope="col">�ڷ�</th>
							</tr>
						</thead>
						<tbody>
						<%
						do until rs.eof  
							  draft_no = rs("draft_no")
							  draft_tax_id = rs("draft_tax_id")
							  draft_live = ""
							  tax_percent = 3
	           			%>
							<tr>
								<td class="first"><%=rs("draft_no")%>&nbsp;</td>
                                <td><%=rs("draft_man")%>&nbsp;</td>
                                <td><%=rs("give_date")%>&nbsp;</td>
                                <td><%=rs("org_name")%>&nbsp;</td>
                                <td><%=rs("draft_tax_id")%>&nbsp;</td>
                        <%
						        give_tot = int(rs("alba_give_total"))
								tax_amt1 = int(rs("tax_amt1"))
								tax_amt2 = int(rs("tax_amt2"))
								alba_cnt = int(rs("alba_cnt"))
								alba_work = int(rs("alba_work"))
								work_comment = rs("work_comment")
								cost_company = rs("cost_company")
								curr_pay = int(rs("pay_amount"))

							  
							 ' tax_amt = give_tot * (tex_percent / 100)
							 ' tax2_amt = give_tot * (0.3 / 100)
							  'curr_pay = give_tot - tax_amt1 - tax2_amt
							 						  
							  'alba_comment = rs("sign_no") + "-" + alba_cnt + "-" + alba_work + "-" + work_comment
                              'alba_comment = replace(app_task,chr(34),chr(39))
							  'view_memo = alba_comment
							  'if len(alba_comment) > 10 then
							  '  	view_memo = mid(alba_comment,1,10) + ".."
							  'end if
                        %>
                                <td class="right"><%=formatnumber(give_tot,0)%></td>
                                <td><%=tax_percent%></td>
                                <td class="right"><%=formatnumber(tax_amt1,0)%></td>
                                <td class="right"><%=formatnumber(tax_amt2,0)%></td>
                                <td class="right"><%=formatnumber(curr_pay,0)%></td>
                                <td><%=cost_company%>&nbsp;</td>
                                <td class="left"><%=rs("sign_no")%>-<%=alba_cnt%>&nbsp;<%=alba_work%>&nbsp;<%=work_comment%></td>
                                                                
                                <td><a href="#" onClick="pop_Window('alba_cost_add.asp?draft_no=<%=rs("draft_no")%>&rever_yymm=<%=rever_yymm%>&give_date=<%=rs("give_date")%>&u_type=<%="U"%>','insa_pay_alba_add_pop','scrollbars=yes,width=800,height=520')">����</a></td>

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
                    <a href="insa_pay_albacost_mg.asp?rever_yymm=<%=rever_yymm%>" class="btnType04">�����ٿ�ε�</a>
					</div>                  
                  	</td>
				    <td>
                    <div id="paging">
                        <a href = "insa_pay_albacost_mg.asp?page=<%=first_page%>&rever_yymm=<%=rever_yymm%>&ck_sw=<%="y"%>">[ó��]</a>
                  	<% if intstart > 1 then %>
                        <a href="insa_pay_albacost_mg.asp?page=<%=intstart -1%>&rever_yymm=<%=rever_yymm%>&ck_sw=<%="y"%>">[����]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
           	<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="insa_pay_albacost_mg.asp?page=<%=i%>&rever_yymm=<%=rever_yymm%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
           	<% if 	intend < total_page then %>
                        <a href="insa_pay_albacost_mg.asp?page=<%=intend+1%>&rever_yymm=<%=rever_yymm%>&ck_sw=<%="y"%>">[����]</a> <a href="insa_pay_albacost_mg.asp?page=<%=total_page%>&rever_yymm=<%=rever_yymm%>&ck_sw=<%="y"%>">[������]</a>
                        <%	else %>
                        [����]&nbsp;[������]
                      <% end if %>
                    </div>
                    </td>
                    <td>
					<div class="btnRight">
					<a href="#" onClick="pop_Window('alba_cost_add.asp?rever_yymm=<%=rever_yymm%>','alba_cost_add_pop','scrollbars=yes,width=800,height=520')" class="btnType04">�Ƹ�����Ʈ��� �Է�</a>
					</div>  
                    </td>
			      </tr>
				  </table>
			</form>
		</div>				
	</div>        				
	</body>
</html>

