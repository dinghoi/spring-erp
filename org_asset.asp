<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Repeat_Rows
Dim field_check
Dim field_view
Dim win_sw
dim asset_cnt(1000,6)
dim asset_dept(1000)
dim asset_tot(6)
for i = 0 to 1000
	asset_dept(i) = "N"
	for j = 0 to 6
		asset_cnt(i,j) = 0
	next
next
for j = 0 to 6
	asset_tot(j) = 0
next

win_sw = "close"

ck_sw=Request("ck_sw")
Page=Request("page")

If ck_sw = "y" Then
	company=Request("company")
	field_view=Request("field_view")
	field_check=Request("field_check")
 else
	company=Request.form("company")
	field_check=Request.form("field_check")
	field_view=Request.form("field_view")
End if

If company = "" Then
	company = "01"
	field_check = "total"
end If

if asset_company <> "00" then
	company = asset_company
end if

If field_check = "total" Then
	field_view = ""
End If

pgsize = 10 ' ȭ�� �� ������ 

If Page = "" Then
	Page = 1
	start_page = 1
End If
stpage = int((page - 1) * pgsize)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_etc = Server.CreateObject("ADODB.Recordset")
Set rs_dept = Server.CreateObject("ADODB.Recordset")
Set rs_group = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

base_sql = "SELECT asset.dept_code, asset.gubun, count(*) as asset_cnt, asset_dept.high_org, asset_dept.org_first, asset_dept.org_second, asset_dept.dept_name FROM asset INNER JOIN asset_dept ON (asset.dept_code = asset_dept.dept_code) AND (asset.company = asset_dept.company) where inst_process = 'Y' and asset.company = '" + company + "'"
if field_check <> "total" then
	condi_sql = " and ( asset_dept." + field_check + " like '%" + field_view + "%' ) "
  else
  	condi_sql = " "
end if	

group_sql = " group by asset.dept_code, asset.gubun order by asset.dept_code"
sql = base_sql + condi_sql + group_sql
rs_group.Open Sql, Dbconn, 1
bi_dept = "y"
i = 0
do until rs_group.eof
	if	bi_dept = "y" then
		bi_dept = rs_group("dept_code")
	end if
	if  bi_dept <> rs_group("dept_code") then
		i = i + 1
		asset_dept(i) = bi_dept
		for j = 1 to 6
			asset_cnt(i,j) = asset_cnt(i,j) + asset_tot(j)
			asset_cnt(0,j) = asset_cnt(0,j) + asset_tot(j)
			asset_tot(j) = 0
		next		
		bi_dept = rs_group("dept_code")
	end if
	jj = cint(rs_group("gubun"))
	if jj > 4 then
		jj = 5
	end if
	asset_tot(jj) = asset_tot(jj) + cint(rs_group("asset_cnt"))
	asset_tot(6) = asset_tot(6) + cint(rs_group("asset_cnt"))	
	rs_group.movenext()
loop
if  i > 0 then
	i = i + 1
	asset_dept(i) = bi_dept
	for j = 1 to 6
		asset_cnt(i,j) = asset_cnt(i,j) + asset_tot(j)
		asset_cnt(0,j) = asset_cnt(0,j) + asset_tot(j)
	next		
end if

total_record = i 'Result.RecordCount
IF total_record mod pgsize = 0 THEN
	total_page = int(total_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((total_record / pgsize) + 1)
END IF

if company = "01" then
	title_01 = "���θ� / ����� / ������"
  else
	title_01 = "������1 / ������2 / ������3"
end if

etc_code = "75" + company
Sql="select * from etc_code where etc_code = '" + etc_code + "'"
Set rs_etc=DbConn.Execute(SQL)
if rs_etc.eof or rs_etc.bof then
	company_name = "����"
  else
	company_name = rs_etc("etc_name")
end if
rs_etc.close()						

title_line = "������ �ڻ� ��Ȳ"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>A/S ���� �ý���</title>
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
			function frmcheck () {
				if (formcheck(document.frm)) {
					document.frm.submit ();
				}
			}
			
		</script>

	</head>
	<body>
		<div id="wrap">			
			<!--#include virtual = "/include/asset_header.asp" -->
			<!--#include virtual = "/include/asset_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="org_asset.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>					
						<dt>���ǰ˻�</dt>
                        <dd>
                            <p>
                                <label>
								<strong>ȸ��</strong>
								<%
                                    if asset_company = "00" then
					
                                    %>
                                <select name="company" id="company">
                                  <% 
                                        Sql="select * from etc_code where etc_type = '75' and used_sw = 'Y' order by etc_name asc"
                                        Rs_etc.Open Sql, Dbconn, 1
                                        do until rs_etc.eof
                                            k = k + 1
                                  %>
                                  <option value='<%=mid(rs_etc("etc_code"),3,2)%>' <%If mid(rs_etc("etc_code"),3,2) = company then %>selected<% end if %>><%=rs_etc("etc_name")%></option>
                                  <%
                                            rs_etc.movenext()
                                        loop
                                        rs_etc.close()	
                                    %>
                                </select>
                                <%		else %>
                                &nbsp;<%=user_name%>
                                <input name="company" type="hidden" id="company" value="<%=company%>">
                                <%	end if %>
								</label>
                                <label>
								<strong>�ʵ�����</strong>
								  <%  if company = "01" then %>
                                        <select name="field_check" id="select5">
                                            <option value="total" <% if field_check = "total" then %>selected<% end if %>>��ü</option>
                                            <option value="high_org" <% if field_check = "high_org" then %>selected<% end if %>>��������</option>
                                            <option value="org_first" <% if field_check = "org_first" then %>selected<% end if %>>���θ�</option>
                                            <option value="org_second" <% if field_check = "org_second" then %>selected<% end if %>>�����</option>
                                            <option value="dept_name" <% if field_check = "dept_name" then %>selected<% end if %>>������</option>
                                            <option value="sido" <% if field_check = "sido" then %>selected<% end if %>>�õ�</option>
                                            <option value="tel_no" <% if field_check = "tel_no" then %>selected<% end if %>>��ȭ��ȣ</option>
                                        </select>              
                                  <%	else  %>
                                        <select name="field_check" id="select5">
                                            <option value="total" <% if field_check = "total" then %>selected<% end if %>>��ü</option>
                                            <option value="high_org" <% if field_check = "high_org" then %>selected<% end if %>>��������</option>
                                            <option value="org_first" <% if field_check = "org_first" then %>selected<% end if %>>������1</option>
                                            <option value="org_second" <% if field_check = "org_second" then %>selected<% end if %>>������1</option>
                                            <option value="dept_name" <% if field_check = "dept_name" then %>selected<% end if %>>������1</option>
                                            <option value="sido" <% if field_check = "sido" then %>selected<% end if %>>�õ�</option>
                                            <option value="tel_no" <% if field_check = "tel_no" then %>selected<% end if %>>��ȭ��ȣ</option>
                                        </select>              
                                  <%  end if  %>
								<input name="field_view" type="text" value="<%=field_view%>" style="width:150px; text-align:left" >
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser.jpg" alt="�˻�"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="10%" >
							<col width="10%" >
							<col width="*" >
							<col width="8%" >
							<col width="8%" >
							<col width="8%" >
							<col width="8%" >
							<col width="8%" >
							<col width="8%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">�Ҽ�ȸ��</th>
								<th scope="col">��������</th>
								<th scope="col"><%=title_01%></th>
								<th scope="col">����ũž</th>
								<th scope="col">�����</th>
								<th scope="col">��Ʈ��</th>
								<th scope="col">������</th>
								<th scope="col">��Ÿ</th>
								<th scope="col">�Ұ�</th>
							</tr>
						</thead>
						<tbody>
						<%
						kk = stpage + 1
						kkk = stpage + pgsize
						for k = kk to kkk
							if asset_dept(k) = "N" then
								exit for
							end if
				
							Sql="select * from asset_dept where company = '" + company + "' and dept_code = '" + asset_dept(k) + "'"
							Set rs_dept=DbConn.Execute(SQL)
							if rs_dept.eof or rs_dept.bof then
								high_org = "����"
								org_first = "����"
								org_second = "����"
								dept_name = "����"
							  else
								high_org = rs_dept("high_org")
								org_first = rs_dept("org_first")
								org_second = rs_dept("org_second")
								dept_name = rs_dept("dept_name")
							end if
							rs_dept.close()						
						%>
							<tr>
								<td class="first"><%=company_name%></td>
								<td><%=high_org%></td>
								<td><%=org_first%>&nbsp;/&nbsp;<%=org_second%>&nbsp;/&nbsp;<%=dept_name%></td>
								<td class="right"><%=formatnumber(clng(asset_cnt(k,1)),0)%></td>
								<td class="right"><%=formatnumber(clng(asset_cnt(k,2)),0)%></td>
								<td class="right"><%=formatnumber(clng(asset_cnt(k,3)),0)%></td>
								<td class="right"><%=formatnumber(clng(asset_cnt(k,4)),0)%></td>
								<td class="right"><%=formatnumber(clng(asset_cnt(k,5)),0)%></td>
								<td class="right"><a href="#" onClick="pop_Window('org_asset_view.asp?company=<%=company%>&dept_code=<%=asset_dept(k)%>','org_asset_view_popup','scrollbars=yes,width=750,height=500')"><%=formatnumber(clng(asset_cnt(k,6)),0)%></a></td>
							</tr>
						<% 
                        next
                        %>
							<tr>
								<th class="first">�ڻ��</th>
								<th>������</th>
								<th><%=total_record%></th>
								<th class="right"><%=formatnumber(clng(asset_cnt(0,1)),0)%></th>
								<th class="right"><%=formatnumber(clng(asset_cnt(0,2)),0)%></th>
								<th class="right"><%=formatnumber(clng(asset_cnt(0,3)),0)%></th>
								<th class="right"><%=formatnumber(clng(asset_cnt(0,4)),0)%></th>
								<th class="right"><%=formatnumber(clng(asset_cnt(0,5)),0)%></th>
								<th class="right"><a href="#" onClick="pop_Window('com_asset_sum.asp?company=<%=company%>&field_check=<%=field_check%>&field_view=<%=field_view%>','company_asset_sum_view_popup','scrollbars=yes,width=750,height=500')"><%=formatnumber(clng(asset_cnt(0,6)),0)%></th>
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
                    <a href = "org_asset_excel.asp?company=<%=company%>&field_check=<%=field_check%>&field_view=<%=field_view%>" class="btnType04">�����ٿ�ε�</a>
					</div>                  
                    </td>
				    <td>
                    <div id="paging">
                        <a href="org_asset.asp?page=<%=first_page%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>&company=<%=company%>">[ó��]</a>
                  	<% if intstart > 1 then %>
                        <a href="org_asset.asp?page=<%=intstart -1%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>&company=<%=company%>">[����]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
                  	<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="org_asset.asp?page=<%=i%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>&company=<%=company%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
                  	<% if 	intend < total_page then %>
                        <a href="org_asset.asp?page=<%=intend+1%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>&company=<%=company%>">[����]</a> <a href="org_asset.asp?page=<%=total_page%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>&company=<%=company%>">[������]</a>
                        <%	else %>
                        [����]&nbsp;[������]
                      <% end if %>
                    </div>
                    </td>
				    <td width="15%">
                    </td>
			      </tr>
				  </table>
			</form>
		</div>				
	</div>        				
	</body>
</html>

