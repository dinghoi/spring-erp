<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows
dim page_cnt
dim pg_cnt

user_name = request.cookies("nkpmg_user")("coo_user_name")
user_id = request.cookies("nkpmg_user")("coo_user_id")
org_name = request.cookies("nkpmg_user")("coo_org_name")
cost_grade = request.cookies("nkpmg_user")("coo_cost_grade")
emp_company = request.cookies("nkpmg_user")("coo_emp_company")

Page=Request("page")
page_cnt=Request.form("page_cnt")
pg_cnt=cint(Request("pg_cnt"))

view_condi = request("view_condi")
from_date=Request.form("from_date")
to_date=Request.form("to_date")

be_pg = "met_stock_out_list.asp"
curr_date = datevalue(mid(cstr(now()),1,10))

ck_sw=Request("ck_sw")

If ck_sw = "y" Then
	view_condi=Request("view_condi")
	goods_type=Request("goods_type")
	from_date=request("from_date")
    to_date=request("to_date")
  else
	view_condi=Request.form("view_condi")
	goods_type=Request.form("goods_type")
	from_date=Request.form("from_date")
    to_date=Request.form("to_date")
End if

If view_condi = "" Then
'	view_condi = emp_company
'	goods_type = "��ǰ"
    view_condi = "��ü"
	goods_type = "��ü"
	curr_dd = cstr(datepart("d",now))
	to_date = mid(cstr(now()),1,10)
	from_date = mid(cstr(now()-curr_dd+1),1,10)
	ck_sw = "n"
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
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_good = Server.CreateObject("ADODB.Recordset")
Set Rs_stock = Server.CreateObject("ADODB.Recordset")
Set Rs_trade = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")

dbconn.open DbConnect

'chulgo_id = "���"

order_Sql = " ORDER BY chulgo_date DESC"
if view_condi = "��ü" then
   if goods_type = "��ü" then
      where_sql = " WHERE (chulgo_date >= '"+from_date+"' and chulgo_date <= '"+to_date+"') and (chulgo_id = '�����ǥ' or chulgo_id = '������ǥ' or chulgo_id = '�������')"
	  else
	  where_sql = " WHERE (chulgo_goods_type = '"&goods_type&"') and (chulgo_date >= '"+from_date+"' and chulgo_date <= '"+to_date+"') and (chulgo_id = '�����ǥ' or chulgo_id = '������ǥ' or chulgo_id = '�������')"
   end if
 else
   if goods_type = "��ü" then
      where_sql = " WHERE (chulgo_stock_company = '"&view_condi&"') and (chulgo_date >= '"+from_date+"' and chulgo_date <= '"+to_date+"') and (chulgo_id = '�����ǥ' or chulgo_id = '������ǥ' or chulgo_id = '�������')"
	  else
	  where_sql = " WHERE (chulgo_goods_type = '"&goods_type&"') and (chulgo_stock_company = '"&view_condi&"') and (chulgo_date >= '"+from_date+"' and chulgo_date <= '"+to_date+"') and (chulgo_id = '�����ǥ' or chulgo_id = '������ǥ' or chulgo_id = '�������')"
   end if
end if

Sql = "SELECT count(*) FROM met_chulgo " + where_sql
Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

sql = "select * from met_chulgo " + where_sql + order_sql + " limit "& stpage & "," &pgsize
Rs.Open Sql, Dbconn, 1
'response.write(sql)

title_line = " ��� ��Ȳ "
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>��ǰ������� �ý���</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "1 1";
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
					alert ("�ʵ������� �����Ͻñ� �ٶ��ϴ�");
					return false;
				}
				return true;
			}
		</script>

	</head>
	<body>
		<div id="wrap">
			<!--#include virtual = "/include/meterials_control_header01.asp" -->
            <!--#include virtual = "/include/meterials_stock_out_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="met_stock_out_list.asp?ck_sw=<%="n"%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>
						<dt>�˻� ����</dt>
                        <dd>
                            <p>
                               <strong>���â�� : </strong>
                              <%
								Sql="select * from emp_org_mst where (isNull(org_end_date) or org_end_date = '1900-01-01') and (org_level = 'ȸ��') ORDER BY org_code ASC"
	                            rs_org.Open Sql, Dbconn, 1
							  %>
                                <label>
								<select name="view_condi" id="view_condi" type="text" style="width:120px">
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
                                <strong>�뵵���� : </strong>
                              <%
								Sql="select * from met_etc_code where etc_type = '01' order by etc_code asc"
					            Rs_etc.Open Sql, Dbconn, 1
							  %>
                                <label>
								<select name="goods_type" id="goods_type" type="text" style="width:90px">
                                    <option value="��ü" <%If goods_type = "��ü" then %>selected<% end if %>>��ü</option>
                			  <%
								do until Rs_etc.eof
			  				  %>
                					<option value='<%=rs_etc("etc_name")%>' <%If goods_type = rs_etc("etc_name") then %>selected<% end if %>><%=rs_etc("etc_name")%></option>
                			  <%
									Rs_etc.movenext()
								loop
								Rs_etc.Close()
							  %>
            					</select>
                                </label>
                               <label>
								<strong>�������(From) : </strong>
                                	<input name="from_date" type="text" value="<%=from_date%>" style="width:70px" id="datepicker">
								</label>
								<label>
								<strong> �� To : </strong>
                                	<input name="to_date" type="text" value="<%=to_date%>" style="width:70px" id="datepicker1">
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser1.jpg" alt="�˻�"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
                <h3 class="stit" style="font-size:12px;">�� ����� & �μ��� ����� �Ʒ� ��ȸ���뿡�� ������ڸ� Ŭ���Ͽ� ����Ͻø� �˴ϴ�!</h3>
				<div class="gView">
				  <table cellpadding="0" cellspacing="0" class="tableList">
				    <colgroup>
				      <col width="3%" >
                      <col width="6%" >
                      <col width="8%" >
                      <col width="6%" >
                      <col width="10%" >
				      <col width="6%" >

                      <col width="10%" >
				      <col width="6%" >
                      <col width="10%" >

                      <col width="16%" >
				      <col width="8%" >

                      <col width="*" >
                      <col width="3%" >
			        </colgroup>
				    <thead>
                      <tr>
				        <th class="first" scope="col">����</th>
                        <th scope="col">�������</th>
                        <th scope="col">����ȣ</th>
                        <th scope="col">�������</th>
                        <th scope="col">���â��</th>
                        <th scope="col">�뵵����</th>

                        <th scope="col">���ǰ��</th>
                        <th scope="col">�����</th>
                        <th scope="col">�Ҽ�</th>

                        <th scope="col">��û�׷��</th>
				        <th scope="col">��ûâ��</th>

                        <th scope="col">����</th>
                        <th scope="col">���</th>
                      </tr>
			        </thead>
				    <tbody>
                      <%
						seq = tottal_record - ( page - 1 ) * pgsize
						do until rs.eof
						   chulgo_date = rs("chulgo_date")
						   chulgo_stock = rs("chulgo_stock")
						   chulgo_seq = rs("chulgo_seq")
						   chulgo_goods_type = rs("chulgo_goods_type")

						   sql = "select * from met_chulgo_goods where (chulgo_date = '"&chulgo_date&"') and (chulgo_stock = '"&chulgo_stock&"') and (chulgo_seq = '"&chulgo_seq&"')  ORDER BY cg_goods_seq,cg_goods_code ASC"
						   Set Rs_good=DbConn.Execute(Sql)
						   if Rs_good.eof or Rs_good.bof then
								cg_goods_name = ""
							  else
							  	cg_goods_name = Rs_good("cg_goods_name")
						   end if
						   Rs_good.close()

						    task_memo = replace(rs("chulgo_memo"),chr(34),chr(39))
							view_memo = task_memo
							if len(task_memo) > 12 then
								view_memo = mid(task_memo,1,12) + ".."
							end if

					  %>
				      <tr>
				        <td class="first"><%=seq%></td>
                        <td><%=rs("chulgo_date")%>&nbsp;</td>
                        <td>
                        <a href="#" onClick="pop_Window('met_chulgo_cust_detail01.asp?chulgo_date=<%=rs("chulgo_date")%>&chulgo_stock=<%=rs("chulgo_stock")%>&chulgo_seq=<%=rs("chulgo_seq")%>&u_type=<%=""%>','met_chulgo_detail_pop','scrollbars=yes,width=930,height=650')"><%=rs("rele_no")%>&nbsp;<%=rs("rele_seq")%></a>
                        </td>
                         <td><%=rs("chulgo_id")%>&nbsp;</td>
                        <td><%=rs("chulgo_stock_name")%>&nbsp;</td>
                        <td><%=rs("chulgo_goods_type")%>&nbsp;</td>

                        <td><%=cg_goods_name%>&nbsp;��</td>
                        <td><%=rs("chulgo_emp_name")%>&nbsp;</td>
                        <td><%=rs("chulgo_org_name")%>&nbsp;</td>

                        <td><%=rs("rele_stock_company")%>(<%=rs("rele_saupbu")%>)</td>
                        <td><%=rs("rele_stock_name")%>&nbsp;</td>

                        <td class="left"><p style="cursor:pointer"><span title="<%=task_memo%>"><%=view_memo%></span></p></td>
                        <td>&nbsp;</td>
			          </tr>
				      <%
							rs.movenext()
							seq = seq -1
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
                    <a href="met_stock_out_list_excel01.asp?view_condi=<%=view_condi%>&goods_type=<%=goods_type%>&from_date=<%=from_date%>&to_date=<%=to_date%>" class="btnType04">�����ٿ�ε�</a>
					</div>
                  	</td>
				    <td>
                    <div id="paging">
                        <a href="met_stock_out_list.asp?page=<%=first_page%>&view_condi=<%=view_condi%>&goods_type=<%=goods_type%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[ó��]</a>
                  	<% if intstart > 1 then %>
                        <a href="met_stock_out_list.asp?page=<%=intstart -1%>&view_condi=<%=view_condi%>&goods_type=<%=goods_type%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[����]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
   	        <% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="met_stock_out_list.asp?page=<%=i%>&view_condi=<%=view_condi%>&goods_type=<%=goods_type%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
   	        <% if 	intend < total_page then %>
                        <a href="met_stock_out_list.asp?page=<%=intend+1%>&view_condi=<%=view_condi%>&goods_type=<%=goods_type%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[����]</a> <a href="met_stock_out_list.asp?page=<%=total_page%>&view_condi=<%=view_condi%>&goods_type=<%=goods_type%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[������]</a>
                        <%	else %>
                        [����]&nbsp;[������]
                      <% end if %>
                    </div>
                    </td>
				    <td width="20%">
					<div class="btnCenter">

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

