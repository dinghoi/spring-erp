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

stock_in_man = user_id

Page=Request("page")
page_cnt=Request.form("page_cnt")
pg_cnt=cint(Request("pg_cnt"))

condi = request("condi")
from_date=Request.form("from_date")
to_date=Request.form("to_date")

be_pg = "met_move_stin_not_enter_list.asp"
curr_date = datevalue(mid(cstr(now()),1,10))

ck_sw=Request("ck_sw")

If ck_sw = "y" Then
	condi=Request("condi")
	stock_condi=Request("stock_condi")
	from_date=request("from_date")
    to_date=request("to_date")
  else
	condi=Request.form("condi")
	stock_condi=Request.form("stock_condi")
	from_date=Request.form("from_date")
    to_date=Request.form("to_date")
End if

If condi = "" Then
	condi = "��ü"
	stock_condi = ""
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
Set Rs_reg = Server.CreateObject("ADODB.Recordset")
Set Rs_good = Server.CreateObject("ADODB.Recordset")
Set Rs_stock = Server.CreateObject("ADODB.Recordset")
Set Rs_trade = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")

dbconn.open DbConnect

chulgo_id = "â���̵�"
chulgo_type = "�κ����"  '���Ϸ�

order_Sql = " ORDER BY chulgo_date,chulgo_stock,chulgo_seq DESC"

if condi = "��ü" then  
       where_sql = " WHERE (in_stock_date = '' or in_stock_date = '0000-00-00') and (chulgo_date >= '"+from_date+"' and chulgo_date <= '"+to_date+"') "
   else 
	   where_sql = " WHERE (in_stock_date = '' or in_stock_date = '0000-00-00') and (rele_stock = '"&stock_condi&"') and (chulgo_date >= '"+from_date+"' and chulgo_date <= '"+to_date+"') "
end if
 
Sql = "SELECT count(*) FROM met_mv_go " + where_sql 
Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

sql = "select * from met_mv_go " + where_sql + order_sql + " limit "& stpage & "," &pgsize 
Rs.Open Sql, Dbconn, 1
'response.write(sql)

title_line = " â���̵� ���԰� ��Ȳ "

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
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
				return "3 1";
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
				if (document.frm.condi.value == "") {
					alert ("�ʵ������� �����Ͻñ� �ٶ��ϴ�");
					return false;
				}	
				return true;
			}
			function condi_view() {

				if (eval("document.frm.view_c[0].checked")) {
					document.getElementById('work1').style.display = 'none';
					document.getElementById('work2').style.display = 'none';
					document.getElementById('acpt1').style.display = '';
				}	
				if (eval("document.frm.view_c[1].checked")) {
					document.getElementById('work1').style.display = '';
					document.getElementById('work2').style.display = '';
					document.getElementById('acpt1').style.display = 'none';
				}	
			}
		</script>

	</head>
	<body>
		<div id="wrap">			
			<!--#include virtual = "/include/meterials_control_header.asp" -->
            <!--#include virtual = "/include/meterials_stock_move_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="met_move_stin_not_enter_list.asp?ck_sw=<%="n"%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>					
						<dt>�˻�</dt>
                        <dd>
                            <p>
                               <strong>�԰�â�� : </strong>
                                    <input name="condi" type="text" id="condi" style="width:120px" value="<%=condi%>"> 
                                    <a href="#" class="btnType03" onClick="pop_Window('met_stockin_search.asp?gubun=<%="mv_stin"%>&stock_in_man=<%=stock_in_man%>','stockin_search_pop','scrollbars=yes,width=600,height=400')">ã��</a>
                                    <input type="hidden" name="stock_condi" value="<%=stock_condi%>" ID="Hidden1">
                                    <input type="hidden" name="stock_manager_code" value="<%=stock_manager_code%>" ID="Hidden1">
                                    <input type="hidden" name="stock_manager_name" value="<%=stock_manager_name%>" ID="Hidden1">
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
                <h3 class="stit" style="font-size:12px;">�� �԰�â�� ã�⸦ �Ͻð� �˻��� �Ͻʽÿ�!</h3>
				<div class="gView">
				  <table cellpadding="0" cellspacing="0" class="tableList">
				    <colgroup>
				      <col width="3%" >
                      <col width="6%" >
                      <col width="6%" >
				      <col width="7%" >
                      <col width="6%" >
                      <col width="12%" >
				      <col width="8%" >
                      <col width="10%" >
				      <col width="7%" >
                      <col width="12%" >
                      <col width="6%" >
				      <col width="*" >
				      <col width="6%" >
			        </colgroup>
				    <thead>
                      <tr>
				        <th class="first" scope="col">����</th>
                        <th scope="col">�뵵����</th>
                        <th scope="col">�������</th>
                        <th scope="col">����ȣ</th>
                        <th scope="col">������</th>
                        <th scope="col">���â��</th>
                        <th scope="col">�����</th>
                        <th scope="col">���ǰ��</th>
                        <th scope="col">��û��(No)</th>
                        <th scope="col">��ûâ��</th>
                        <th scope="col">��û���</th>
                        <th scope="col">����</th>
                        <th scope="col">�԰�����</th>
                      </tr>
			        </thead>
				    <tbody>
                  <%
						seq = tottal_record - ( page - 1 ) * pgsize
						do until rs.eof
						   chulgo_date = rs("chulgo_date")
						   chulgo_stock = rs("chulgo_stock")
						   chulgo_seq = rs("chulgo_seq")
						   
						   rele_date = rs("rele_date")
						   rele_stock = rs("rele_stock")
						   rele_seq = rs("rele_seq")
					       
						   sql = "select * from met_mv_reg where (rele_date = '"&rele_date&"') and (rele_stock = '"&rele_stock&"') and (rele_seq = '"&rele_seq&"')"
						   Set Rs_reg=DbConn.Execute(Sql)
						   if Rs_reg.eof or Rs_reg.bof then
								rele_stock_name = ""
								rele_emp_name = ""
							  else
							  	rele_stock_name = Rs_reg("rele_stock_name")
								rele_emp_name = Rs_reg("rele_emp_name")
						   end if
						   Rs_reg.close()
						   
						   sql = "select * from met_mv_go_goods where (chulgo_date = '"&chulgo_date&"') and (chulgo_stock = '"&chulgo_stock&"') and (chulgo_seq = '"&chulgo_seq&"')  ORDER BY cg_goods_seq,cg_goods_code ASC"
						   Set Rs_good=DbConn.Execute(Sql)
						   if Rs_good.eof or Rs_good.bof then
								bg_goods_name = ""
							  else
							  	bg_goods_name = Rs_good("cg_goods_name")
						   end if
						   Rs_good.close()
						   
						   chulgo_no = mid(cstr(rs("chulgo_date")),3,2) + mid(cstr(rs("chulgo_date")),6,2) + mid(cstr(rs("chulgo_date")),9,2) 
						   rele_no = mid(cstr(rs("rele_date")),3,2) + mid(cstr(rs("rele_date")),6,2) + mid(cstr(rs("rele_date")),9,2) 

				  %>
				      <tr>
				        <td class="first"><%=seq%></td>
                        <td><%=rs("chulgo_goods_type")%>&nbsp;</td>
                        <td><%=rs("chulgo_date")%>&nbsp;</td>
                        <td>
						<a href="#" onClick="pop_Window('met_move_chulgo_detail.asp?chulgo_date=<%=rs("chulgo_date")%>&chulgo_stock=<%=rs("chulgo_stock")%>&chulgo_seq=<%=rs("chulgo_seq")%>&u_type=<%=""%>','met_move_chulgo_detail_pop','scrollbars=yes,width=930,height=650')"><%=chulgo_no%>&nbsp;<%=rs("chulgo_stock")%><%=rs("chulgo_seq")%></a>
                        </td>
                        <td><%=rs("chulgo_type")%>&nbsp;</td>
                        <td><%=rs("chulgo_stock_name")%>(<%=rs("chulgo_stock")%>)&nbsp;</td>
                        <td><%=rs("chulgo_emp_name")%>(<%=rs("chulgo_emp_no")%>)&nbsp;</td>
                        <td><%=bg_goods_name%>&nbsp;��</td>
                        <td>
						<a href="#" onClick="pop_Window('met_move_reg_detail.asp?rele_date=<%=rs("rele_date")%>&rele_stock=<%=rs("rele_stock")%>&rele_seq=<%=rs("rele_seq")%>&u_type=<%=""%>','met_move_reg_detail_pop','scrollbars=yes,width=930,height=650')"><%=rele_no%>&nbsp;<%=rs("rele_stock")%><%=rs("rele_seq")%></a>
                        </td>
                        <td><%=rele_stock_name%>&nbsp;</td>
                        <td><%=rele_emp_name%>&nbsp;</td>
                        <td class="left"><%=rs("chulgo_memo")%>&nbsp;</td>
                        <td><%=rs("in_stock_date")%>&nbsp;</td>
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
                    <a href="met_move_chulgo_excel.asp?view_condi=<%=view_condi%>&stock=<%=stock%>&goods_type=<%=goods_type%>&from_date=<%=from_date%>&to_date=<%=to_date%>" class="btnType04">�����ٿ�ε�</a>
					</div>                  
                  	</td>
				    <td>
                    <div id="paging">
                        <a href="met_move_stin_not_enter_list.asp?page=<%=first_page%>&view_condi=<%=view_condi%>&stock=<%=stock%>&goods_type=<%=goods_type%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[ó��]</a>
                  	<% if intstart > 1 then %>
                        <a href="met_move_stin_not_enter_list.asp?page=<%=intstart -1%>&view_condi=<%=view_condi%>&stock=<%=stock%>&goods_type=<%=goods_type%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[����]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
   	        <% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="met_move_stin_not_enter_list.asp?page=<%=i%>&view_condi=<%=view_condi%>&stock=<%=stock%>&goods_type=<%=goods_type%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
   	        <% if 	intend < total_page then %>
                        <a href="met_move_stin_not_enter_list.asp?page=<%=intend+1%>&view_condi=<%=view_condi%>&stock=<%=stock%>&goods_type=<%=goods_type%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[����]</a> <a href="met_move_stin_not_enter_list.asp?page=<%=total_page%>&view_condi=<%=view_condi%>&stock=<%=stock%>&goods_type=<%=goods_type%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[������]</a>
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
		<input type="hidden" name="user_id">
		<input type="hidden" name="pass">
	</body>
</html>

