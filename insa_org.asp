<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows
dim page_cnt
dim pg_cnt

Page=Request("page")
page_cnt=Request.form("page_cnt")
pg_cnt=cint(Request("pg_cnt"))
be_pg = "insa_org.asp"
curr_date = datevalue(mid(cstr(now()),1,10))

ck_sw=Request("ck_sw")

If ck_sw = "y" Then
	view_condi=Request("view_condi")
  else
	view_condi=Request.form("view_condi")
End if

If view_condi = "" Then
	view_condi = "���̿��������"
End If

pgsize = 10 ' ȭ�� �� ������ 

If Page = "" Then
	Page = 1
	start_page = 1
End If
stpage = int((page - 1) * pgsize)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")

dbconn.open DbConnect

order_Sql = " ORDER BY org_company,org_bonbu,org_saupbu,org_team,org_name ASC"
where_sql = " WHERE (isNull(org_end_date) or org_end_date = '1900-01-01') and (org_company = '"&view_condi&"')"

Sql = "SELECT count(*) FROM emp_org_mst " + where_sql
Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

sql = "select * from emp_org_mst " + where_sql + order_sql + " limit "& stpage & "," &pgsize 
Rs.Open Sql, Dbconn, 1
'response.write(sql)

title_line = " ���� ��Ȳ "

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
				return "0 1";
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
					alert ("�ʵ������� �����Ͻñ� �ٶ��ϴ�");
					return false;
				}	
				return true;
			}
		</script>

	</head>
	<body>
		<div id="wrap">			
			<!--#include virtual = "/include/insa_header.asp" -->
			<!--#include virtual = "/include/insa_org_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_org.asp" method="post" name="frm">
                
				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>					
						<dt>ȸ�� �˻�</dt>
                        <dd>
                            <p>
                               <strong>ȸ�� : </strong>
                              <%
								Sql="select * from emp_org_mst where (isNull(org_end_date) or org_end_date = '1900-01-01') and (org_level = 'ȸ��') ORDER BY org_code ASC"
	                            rs_org.Open Sql, Dbconn, 1	
							  %>
                                <label>
								<select name="view_condi" id="view_condi" type="text" style="width:150px">

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
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser1.jpg" alt="�˻�"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>                
                
				<div class="gView">
				  <table cellpadding="0" cellspacing="0" class="tableList">
				    <colgroup>
				      <col width="4%" >
				      <col width="11%" >
                      <col width="4%" >
				      <col width="6%" >
				      <col width="8%" >
                      <col width="10%" >
				      <col width="10%" >
				      <col width="10%" >
				      <col width="10%" >
				      <col width="8%" >
                      <col width="6%" >
				      <col width="8%" >
                      <col width="3%" >
			        </colgroup>
				    <thead>
				      <tr>
				        <th colspan="3" class="first" scope="col" style=" border-bottom:1px solid #e3e3e3;">��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��</th>
				        <th colspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">��&nbsp;&nbsp;��&nbsp;&nbsp;��</th>
                        <th colspan="4" scope="col" style=" border-bottom:1px solid #e3e3e3;">��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��</th>
				        <th rowspan="2" scope="col">����������</th>
				        <th colspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">����&nbsp;������</th>
                        <th rowspan="2" scope="col">����</th>
			          </tr>
                      <tr>
				        <th class="first"scope="col">�ڵ�</th>
				        <th scope="col">������</th>
                        <th scope="col">T.O</th>
				        <th scope="col">���</th>
				        <th scope="col">����</th>
                        <th scope="col">ȸ&nbsp;&nbsp;��</th>
				        <th scope="col">��&nbsp;&nbsp;��</th>
				        <th scope="col">�����</th>
				        <th scope="col">��</th>
				        <th scope="col">���</th>
                        <th scope="col">����</th>
                      </tr>
			        </thead>
				    <tbody>
                      <%
						do until rs.eof
					  %>
				      <tr>
				        <td class="first"><%=rs("org_code")%>&nbsp;</td>
                        <td><a href="#" onClick="pop_Window('insa_org_view.asp?org_code=<%=rs("org_code")%>&org_name=<%=org_name%>&u_type=<%="U"%>','insa_org_view_pop','scrollbars=yes,width=750,height=350')"><%=rs("org_name")%></a>&nbsp;</td>
                        <td><%=rs("org_table_org")%>&nbsp;</td>
                        <td><%=rs("org_empno")%>&nbsp;</td>
                        <td>
                        <a href="#" onClick="pop_Window('insa_card00.asp?emp_no=<%=rs("org_empno")%>&be_pg=<%=be_pg%>&page=<%=page%>&view_sort=<%=view_sort%>&date_sw=<%=date_sw%>&page_cnt=<%=page_cnt%>','emp_card0_pop','scrollbars=yes,width=1250,height=650')"><%=rs("org_emp_name")%></a>
						</td>
                        <td><%=rs("org_company")%>&nbsp;</td>
				        <td><%=rs("org_bonbu")%>&nbsp;</td>
                        <td><%=rs("org_saupbu")%>&nbsp;</td>
                        <td><%=rs("org_team")%>&nbsp;</td>
                        <td><%=rs("org_date")%>&nbsp;</td>
                        <td><%=rs("org_owner_empno")%>&nbsp;</td>
                        <td><%=rs("org_owner_empname")%>&nbsp;</td>
                        <td><a href="#" onClick="pop_Window('insa_org_modify.asp?org_code=<%=rs("org_code")%>&u_type=<%="U"%>','insa_org_modi_pop','scrollbars=yes,width=1250,height=400')">����</a>&nbsp;</td>
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
				    <td width="20%">
					<div class="btnCenter">
                    <a href="insa_excel_org.asp?view_condi=<%=view_condi%>" class="btnType04">�����ٿ�ε�</a>
					</div>                  
                  	</td>
				    <td>
                    <div id="paging">
                        <a href="insa_org.asp?page=<%=first_page%>&view_condi=<%=view_condi%>&ck_sw=<%="y"%>">[ó��]</a>
                  	<% if intstart > 1 then %>
                        <a href="insa_org.asp?page=<%=intstart -1%>&view_condi=<%=view_condi%>&ck_sw=<%="y"%>">[����]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
   	        <% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="insa_org.asp?page=<%=i%>&view_condi=<%=view_condi%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
   	        <% if 	intend < total_page then %>
                        <a href="insa_org.asp?page=<%=intend+1%>&view_condi=<%=view_condi%>&ck_sw=<%="y"%>">[����]</a> <a href="insa_org.asp?page=<%=total_page%>&view_condi=<%=view_condi%>&ck_sw=<%="y"%>">[������]</a>
                        <%	else %>
                        [����]&nbsp;[������]
                      <% end if %>
                    </div>
                    </td>
				    <td width="20%">
					<div class="btnCenter">
                    <a href="#" onClick="pop_Window('insa_org_reg.asp?view_condi=<%=view_condi%>','insa_org_reg_popup','scrollbars=yes,width=1250,height=400')" class="btnType04">�ű��������</a>
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

