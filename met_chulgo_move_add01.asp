<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
dim pummok_tab(10,20)
dim amount_tab(3,20)

user_name = request.cookies("nkpmg_user")("coo_user_name")
user_id = request.cookies("nkpmg_user")("coo_user_id")
cost_grade = request.cookies("nkpmg_user")("coo_cost_grade")

u_type = request("u_type")

view_condi=Request("view_condi")
goods_type=Request("goods_type")
rele_id=Request("rele_id")

curr_date = mid(cstr(now()),1,10)
chulgo_date = curr_date

mok_cnt = 0
pummok_cnt = 0

for i = 1 to 10
	for j = 1 to 20
		pummok_tab(i,j) = ""
	next
next
for i = 1 to 3
	for j = 1 to 20
		amount_tab(i,j) = 0
	next
next
' response.write(reg_date)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set Rs_rele = Server.CreateObject("ADODB.Recordset")
Set Rs_good = Server.CreateObject("ADODB.Recordset")
Set Rs_stock = Server.CreateObject("ADODB.Recordset")
Set rs_trade = Server.CreateObject("ADODB.Recordset")
Set Rs_max = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

if goods_type = "��ü" then
      chulgo_goods_type = "A/S����"
   else
      chulgo_goods_type = goods_type
end if

'����â�� ù ���â��� ����
      sql = "select * from met_stock_code where (stock_code = '" + user_id + "')"
	  Rs.Open Sql, Dbconn, 1
	  if not Rs.eof then
	       chulgo_stock_company = Rs("stock_company")
		   chulgo_stock_name = Rs("stock_name")
		   chulgo_stock = Rs("stock_code")
		   stock_bonbu = Rs("stock_bonbu")
		   stock_saupbu = Rs("stock_saupbu")
		   stock_team = Rs("stock_team")
		   stock_manager_code = Rs("stock_manager_code")
		   stock_manager_name = Rs("stock_manager_name")
	  end if
	  Rs.close()
	  
chulgo_id = "CE���"



title_line = " CE���(â��) ��� "

path_name = "/met_upload"

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
				return "2 1";
			}
		</script>
		<script type="text/javascript">
			$(function() {    $( "#datepicker" ).datepicker();
												$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker" ).datepicker("setDate", "<%=chulgo_date%>" );
			});	  
			$(function() {    $( "#datepicker2" ).datepicker();
												$( "#datepicker2" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker2" ).datepicker("setDate", "<%=bill_due_date%>" );
			});	  
			$(function() {    $( "#datepicker3" ).datepicker();
												$( "#datepicker3" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker3" ).datepicker("setDate", "<%=bill_issue_date%>" );
			});	  
			$(function() {    $( "#datepicker4" ).datepicker();
												$( "#datepicker4" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker4" ).datepicker("setDate", "<%=buy_collect_due_date%>" );
			});	  
			$(function() {    $( "#datepicker5" ).datepicker();
												$( "#datepicker5" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker5" ).datepicker("setDate", "<%=collect_date%>" );
			});	  
			function goAction () {
			   window.close () ;
			}
			function goBefore () {
			   history.back() ;
			}
			function frmcheck () {
				if (chkfrm1()) {
					document.frm.submit ();
				}
			}			
			function chkfrm() {
						
				{
				a=confirm('�Է��Ͻðڽ��ϱ�?')
				if (a==true) {
					return true;
				}
				return false;
				}
			}

			function chkfrm1() {
				if(document.frm.chulgo_goods_type.value == "") {
					alert('�뵵������ �����ϼ���');
					frm.chulgo_goods_type.focus();
					return false;}
				if(document.frm.chulgo_stock_name.value == "") {
					alert('���â�� �����ϼ���');
					frm.chulgo_stock_name.focus();
					return false;}
				if(document.frm.rele_company.value == "") {
					alert('����ûȸ�縦 �����ϼ���');
					frm.rele_company.focus();
					return false;}
				if(document.frm.rele_saupbu.value == "") {
					alert('����û����θ� �����ϼ���');
					frm.rele_saupbu.focus();
					return false;}
				if(document.frm.rele_stock_name.value == "") {
					alert('��ûâ��(�԰�)�� �����ϼ���');
					frm.rele_stock_name.focus();
					return false;}
				if(document.frm.chulgo_date.value == "") {
					alert('����ϸ� �Է��ϼ���');
					frm.chulgo_date.focus();
					return false;}

										
				{
				a=confirm('�Է��Ͻðڽ��ϱ�?')
				if (a==true) {
					return true;
				}
				return false;
				}
			}

			function pop_pummok() 
			{ 
				var code_ary = new Array();
//				code_ary[0] = document.frm.goods_code1.value + "/" + document.frm.goods_standard1.value;				
				code_ary[0] = document.frm.goods_code1.value + "/" + document.frm.goods_standard1.value;
				code_ary[1] = document.frm.goods_code2.value + "/" + document.frm.goods_standard2.value;
				code_ary[2] = document.frm.goods_code3.value + "/" + document.frm.goods_standard3.value;
				code_ary[3] = document.frm.goods_code4.value + "/" + document.frm.goods_standard4.value;
				code_ary[4] = document.frm.goods_code5.value + "/" + document.frm.goods_standard5.value;
				code_ary[5] = document.frm.goods_code6.value + "/" + document.frm.goods_standard6.value;
				code_ary[6] = document.frm.goods_code7.value + "/" + document.frm.goods_standard7.value;
				code_ary[7] = document.frm.goods_code8.value + "/" + document.frm.goods_standard8.value;
				code_ary[8] = document.frm.goods_code9.value + "/" + document.frm.goods_standard9.value;
				code_ary[9] = document.frm.goods_code10.value + "/" + document.frm.goods_standard10.value;
				code_ary[10] = document.frm.goods_code11.value + "/" + document.frm.goods_standard11.value;
				code_ary[11] = document.frm.goods_code12.value + "/" + document.frm.goods_standard12.value;
				code_ary[12] = document.frm.goods_code13.value + "/" + document.frm.goods_standard13.value;
				code_ary[13] = document.frm.goods_code14.value + "/" + document.frm.goods_standard14.value;
				code_ary[14] = document.frm.goods_code15.value + "/" + document.frm.goods_standard15.value;
				code_ary[15] = document.frm.goods_code16.value + "/" + document.frm.goods_standard16.value;
				code_ary[16] = document.frm.goods_code17.value + "/" + document.frm.goods_standard17.value;
				code_ary[17] = document.frm.goods_code18.value + "/" + document.frm.goods_standard18.value;
				code_ary[18] = document.frm.goods_code19.value + "/" + document.frm.goods_standard19.value;
				code_ary[19] = document.frm.goods_code20.value + "/" + document.frm.goods_standard20.value;
				goods_type = document.frm.chulgo_goods_type.value
				stock_code = document.frm.chulgo_stock.value
//				var popupW = 600;
//				var popupH = 400;
//				var left = Math.ceil((window.screen.width - popupW)/2);
//				var top = Math.ceil((window.screen.height - popupH)/2);
//				window.open('met_goods_select.asp?code_ary='+code_ary+'', '�˾�����', 'width='+ popupW +', height='+ popupH +', left='+ left +', top='+ top +', location=no, status=no, menubar=no, toolbar=no, scrollbars=yes, resizable=no');
				
				var url = "met_stock_goods_select.asp?code_ary="+code_ary+'&goods_type='+goods_type+'&stock_code='+stock_code;
//				var url = "met_goods_select.asp?code_ary="+code_ary;
				pop_Window(url,'���ǰ����','scrollbars=yes,width=960,height=600');
				
			} 
			function pop_pummok_del() 
			{ 
				var code_ary = new Array();
				var del_ary = new Array();
				code_ary[0] = document.frm.goods_code1.value
				code_ary[1] = document.frm.goods_code2.value
				code_ary[2] = document.frm.goods_code3.value
				code_ary[3] = document.frm.goods_code4.value
				code_ary[4] = document.frm.goods_code5.value
				code_ary[5] = document.frm.goods_code6.value
				code_ary[6] = document.frm.goods_code7.value
				code_ary[7] = document.frm.goods_code8.value
				code_ary[8] = document.frm.goods_code9.value
				code_ary[9] = document.frm.goods_code10.value
				code_ary[10] = document.frm.goods_code11.value
				code_ary[11] = document.frm.goods_code12.value
				code_ary[12] = document.frm.goods_code13.value
				code_ary[13] = document.frm.goods_code14.value
				code_ary[14] = document.frm.goods_code15.value
				code_ary[15] = document.frm.goods_code16.value
				code_ary[16] = document.frm.goods_code17.value
				code_ary[17] = document.frm.goods_code18.value
				code_ary[18] = document.frm.goods_code19.value
				code_ary[19] = document.frm.goods_code20.value

				if (document.frm.del_check1.checked == true) {
					del_ary[0] = 'Y'; } 					
				if (document.frm.del_check1.checked == false) {
					del_ary[0] = 'N'; } 
				if (document.frm.del_check2.checked == true) {
					del_ary[1] = 'Y'; } 					
				if (document.frm.del_check2.checked == false) {
					del_ary[1] = 'N'; } 
				if (document.frm.del_check3.checked == true) {
					del_ary[2] = 'Y'; } 					
				if (document.frm.del_check3.checked == false) {
					del_ary[2] = 'N'; } 
				if (document.frm.del_check4.checked == true) {
					del_ary[3] = 'Y'; } 					
				if (document.frm.del_check4.checked == false) {
					del_ary[3] = 'N'; } 
				if (document.frm.del_check5.checked == true) {
					del_ary[4] = 'Y'; } 					
				if (document.frm.del_check5.checked == false) {
					del_ary[4] = 'N'; } 
				if (document.frm.del_check6.checked == true) {
					del_ary[5] = 'Y'; } 					
				if (document.frm.del_check6.checked == false) {
					del_ary[5] = 'N'; } 
				if (document.frm.del_check7.checked == true) {
					del_ary[6] = 'Y'; } 					
				if (document.frm.del_check7.checked == false) {
					del_ary[6] = 'N'; } 
				if (document.frm.del_check8.checked == true) {
					del_ary[7] = 'Y'; } 					
				if (document.frm.del_check8.checked == false) {
					del_ary[7] = 'N'; } 
				if (document.frm.del_check9.checked == true) {
					del_ary[8] = 'Y'; } 					
				if (document.frm.del_check9.checked == false) {
					del_ary[8] = 'N'; } 
				if (document.frm.del_check10.checked == true) {
					del_ary[9] = 'Y'; } 					
				if (document.frm.del_check10.checked == false) {
					del_ary[9] = 'N'; } 
					
				if (document.frm.del_check11.checked == true) {
					del_ary[10] = 'Y'; } 					
				if (document.frm.del_check11.checked == false) {
					del_ary[10] = 'N'; } 
				if (document.frm.del_check12.checked == true) {
					del_ary[11] = 'Y'; } 					
				if (document.frm.del_check12.checked == false) {
					del_ary[11] = 'N'; } 
				if (document.frm.del_check13.checked == true) {
					del_ary[12] = 'Y'; } 					
				if (document.frm.del_check13.checked == false) {
					del_ary[12] = 'N'; } 
				if (document.frm.del_check14.checked == true) {
					del_ary[13] = 'Y'; } 					
				if (document.frm.del_check14.checked == false) {
					del_ary[13] = 'N'; } 
				if (document.frm.del_check15.checked == true) {
					del_ary[14] = 'Y'; } 					
				if (document.frm.del_check15.checked == false) {
					del_ary[14] = 'N'; } 
				if (document.frm.del_check16.checked == true) {
					del_ary[15] = 'Y'; } 					
				if (document.frm.del_check16.checked == false) {
					del_ary[15] = 'N'; } 
				if (document.frm.del_check17.checked == true) {
					del_ary[16] = 'Y'; } 					
				if (document.frm.del_check17.checked == false) {
					del_ary[16] = 'N'; } 
				if (document.frm.del_check18.checked == true) {
					del_ary[17] = 'Y'; } 					
				if (document.frm.del_check18.checked == false) {
					del_ary[17] = 'N'; } 
				if (document.frm.del_check19.checked == true) {
					del_ary[18] = 'Y'; } 					
				if (document.frm.del_check19.checked == false) {
					del_ary[18] = 'N'; } 
				if (document.frm.del_check20.checked == true) {
					del_ary[19] = 'Y'; } 					
				if (document.frm.del_check20.checked == false) {
					del_ary[19] = 'N'; } 
				goods_type = document.frm.chulgo_goods_type.value
				stock_code = document.frm.chulgo_stock.value	
//				var popupW = 600;
//				var popupH = 400;
//				var left = Math.ceil((window.screen.width - popupW)/2);
//				var top = Math.ceil((window.screen.height - popupH)/2);
//				window.open('met_stock_goods_del_ok.asp?code_ary='+code_ary+'&del_ary='+del_ary+'&goods_type='+goods_type+'&stock_code='+stock_code+'', '�˾�����', 'width='+ popupW +', height='+ popupH +', left='+ left +', top='+ top +', location=no, status=no, menubar=no, toolbar=no, scrollbars=no, resizable=no');
				
				var url = "met_stock_goods_del_ok.asp?code_ary="+code_ary+'&del_ary='+del_ary+'&goods_type='+goods_type+'&stock_code='+stock_code;				
				pop_Window(url,'���û�ǰ����','scrollbars=yes,width=600,height=400');
				alert("�����Ǿ����ϴ� !!!");
				NumCal();
			} 

		function NumCal(txtObj){
			var qty_ary = new Array();
			var b_qty_ary = new Array();

			for (j=1;j<21;j++) {
				qty_ary[j] = eval("document.frm.qty" + j + ".value").replace(/,/g,"");
				b_qty_ary[j] = eval("document.frm.jqty" + j + ".value").replace(/,/g,"");
				
				acpt_qty = parseInt(qty_ary[j]);
				sign_qty = parseInt(b_qty_ary[j]);

	            if (acpt_qty > sign_qty) {
					alert ("���������� �������� �����ϴ�!!");
					return false;
				}
		
			}

			if (txtObj.value.length<5) {
				txtObj.value=txtObj.value.replace(/,/g,"");
				txtObj.value=txtObj.value.replace(/\D/g,"");
			}
			var num = txtObj.value;
			if (num == "--" ||  num == "." ) num = "";
			if (num != "" ) {
				temp=new String(num);
				if(temp.length<1) return "";
							
				// ����ó��
				if(temp.substr(0,1)=="-") minus="-";
					else minus="";
							
				// �Ҽ�������ó��
				dpoint=temp.search(/\./);
						
				if(dpoint>0)
				{
				// ù��° ������ .�� �������� �ڸ��� ���������� ���� ����
				dpointVa="."+temp.substr(dpoint).replace(/\D/g,"");
				temp=temp.substr(0,dpoint);
				}else dpointVa="";
							
				// �����ܹ̿��� ����
				temp=temp.replace(/\D/g,"");
				zero=temp.search(/[1-9]/);
						
				if(zero==-1) return "";
				else if(zero!=0) temp=temp.substr(zero);
							
				if(temp.length<4) return minus+temp+dpointVa;
				buf="";
				while (true)
				{
				if(temp.length<3) { buf=temp+buf; break; }
					
				buf=","+temp.substr(temp.length-3)+buf;
				temp=temp.substr(0, temp.length-3);
				}
				if(buf.substr(0,1)==",") buf=buf.substr(1);
						
				//return minus+buf+dpointVa;
				txtObj.value = minus+buf+dpointVa;
			}else txtObj.value = "0";					
		}
		function pummok_list_view() {
				pummok_cnt = parseInt(document.frm.pummok_cnt.value);
				for (j=1;j<pummok_cnt+1;j++) {
					eval("document.getElementById('pummok_list" + j + "')").style.display = '';
				}
				NumCal();
			}
		function delcheck() 
				{
				a=confirm('���� �����Ͻðڽ��ϱ�?')
				if (a==true) {
					document.frm.method = "post";
					document.frm.enctype = "multipart/form-data";
					document.frm.action = "met_chulgo_reg_del_ok.asp";
					document.frm.submit();
				return true;
				}
				return false;
			}
		</script>

	</head>
	<body>
		<div id="container">				
			<div class="gView">
				<h3 class="insa"><%=title_line%></h3>
                <form method="post" name="frm" action="met_chulgo_move_add01_save.asp" enctype="multipart/form-data">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableWrite">
						<colgroup>
							<col width="10%" >
							<col width="15%" >
							<col width="10%" >
							<col width="15%" >
							<col width="10%" >
							<col width="15%" >
                            <col width="10%" >
							<col width="15%" >
						</colgroup>
						<tbody>
							<tr>
							  <tr>
							  <th>�������</th>
							  <td class="left"><input name="chulgo_date" type="text" value="<%=chulgo_date%>" style="width:120px;text-align:center" id="datepicker"></td>
                              <th>�뵵����</th>
							  <td class="left">
							<%
                                Sql="select * from met_etc_code where etc_type = '01' order by etc_code asc"
					            Rs_etc.Open Sql, Dbconn, 1
                            %>
                                <select name="chulgo_goods_type" id="chulgo_goods_type" style="width:120px">
                                    <option value=''>����</option> 
                            <% 
                                do until Rs_etc.eof 
                            %>
                                    <option value='<%=rs_etc("etc_name")%>' <%If chulgo_goods_type = rs_etc("etc_name") then %>selected<% end if %>><%=rs_etc("etc_name")%></option>
                            <%
                                    Rs_etc.movenext()  
                                loop 
                                Rs_etc.Close()
                            %>
                                </select>
                              </td>
                              <th>���â��</th>
							  <td colspan="3" class="left">
                              <input name="chulgo_stock_company" type="text" value="<%=chulgo_stock_company%>" readonly="true" style="width:120px">
                              
                              <input name="chulgo_stock_name" type="text" value="<%=chulgo_stock_name%>" readonly="true" style="width:120px">
                              
						      <a href="#" class="btnType03" onClick="pop_Window('meterials_stock_select.asp?gubun=<%="chulgo"%>&view_condi=<%=view_condi%>','stock_search_pop','scrollbars=yes,width=600,height=400')">ã��</a>
                              <input type="hidden" name="chulgo_stock" value="<%=chulgo_stock%>" ID="Hidden1">
                              <input type="hidden" name="stock_bonbu" value="<%=stock_bonbu%>" ID="Hidden1">
                              <input type="hidden" name="stock_saupbu" value="<%=stock_saupbu%>" ID="Hidden1">
                              <input type="hidden" name="stock_team" value="<%=stock_team%>" ID="Hidden1">
                              <input type="hidden" name="stock_manager_code" value="<%=stock_manager_code%>" ID="Hidden1">
                              <input type="hidden" name="stock_manager_name" value="<%=stock_manager_name%>" ID="Hidden1">
                              </td>
                            </tr>
							<tr>
                              <th>��ûâ��(�԰�)</th>
							  <td colspan="3" class="left">
                              <input name="rele_stock_company" type="text" value="<%=rele_stock_company%>" readonly="true" style="width:120px">
                              
                              <input name="rele_stock_name" type="text" value="<%=rele_stock_name%>" readonly="true" style="width:120px">
                              
						      <a href="#" class="btnType03" onClick="pop_Window('meterials_stock_select.asp?gubun=<%="mvreg"%>&view_condi=<%=view_condi%>','rstock_search_pop','scrollbars=yes,width=600,height=400')">ã��</a>
                              <input type="hidden" name="rele_stock" value="<%=rele_stock%>" ID="Hidden1">
                              <input type="hidden" name="rele_stock_bonbu" value="<%=rele_stock_bonbu%>" ID="Hidden1">
                              <input type="hidden" name="rele_stock_saupbu" value="<%=rele_stock_saupbu%>" ID="Hidden1">
                              <input type="hidden" name="rele_stock_team" value="<%=rele_stock_team%>" ID="Hidden1">
                              <input type="hidden" name="rele_manager_code" value="<%=rele_manager_code%>" ID="Hidden1">
                              <input type="hidden" name="rele_manager_name" value="<%=rele_manager_name%>" ID="Hidden1">
                              </td>
                              <th>��û�׷��</th>
							  <td class="left">
							<%
								' 2019.02.22 ������ ��û ȸ�縮��Ʈ�� ������ �ҽ� org_end_date�� null �� �ƴ� �������ڸ� �����ϸ� ����Ʈ�� ��Ÿ���� �ʴ´�.
								Sql = "SELECT * FROM emp_org_mst WHERE ISNULL(org_end_date) AND org_level = 'ȸ��'  ORDER BY org_company ASC"
                                rs_org.Open Sql, Dbconn, 1
                            %>
                                <select name="rele_company" id="rele_company" value="<%=rele_company%>" style="width:120px">
                                    <option value=''>����</option> 
                            <% 
                                do until rs_org.eof 
                            %>
                                    <option value='<%=rs_org("org_name")%>' <%If rs_org("org_name") = rele_company  then %>selected<% end if %>><%=rs_org("org_name")%></option>
                            <%
                                    rs_org.movenext()  
                                loop 
                                rs_org.Close()
                            %>
                                </select>
                              </td>
							  <th>��û�����</th>
							  <td class="left">
							<%
                                Sql="select org_name from emp_org_mst where org_level = '�����' group by org_name order by org_name asc"
                                rs_org.Open Sql, Dbconn, 1
                            %>
                                <select name="rele_saupbu" id="rele_saupbu" value="<%=rele_saupbu%>" style="width:120px">
                                    <option value=''>����</option> 
                            <% 
                                do until rs_org.eof 
                            %>
                                    <option value='<%=rs_org("org_name")%>' <%If rs_org("org_name") = rele_saupbu  then %>selected<% end if %>><%=rs_org("org_name")%></option>
                            <%
                                    rs_org.movenext()  
                                loop 
                                rs_org.Close()
                            %>
                                </select>
                              </td>
						    </tr>
                            <tr>
							  <th>���</th>
							  <td class="left" colspan="8" ><textarea name="chulgo_memo" rows="3" style="text-align:left; ime-mode:active" id="textarea"><%=chulgo_memo%></textarea></td>
						    </tr>
						</tbody>
					</table>
				</div>
                <br>
                <h3 class="stit" style="font-size:12px;">�� ��� ���� ���� ��</h3>
            	<table cellpadding="0" cellspacing="0" class="tableWrite">
						<colgroup>
							<col width="4%" >
							<col width="12%" >
                            <col width="12%" >
							<col width="12%" >
							<col width="*" >
                            <col width="16%" >
							<col width="10%" >
                            <col width="10%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first; left" colspan="8" scope="col" style=" border-bottom:1px solid #e3e3e3;"><a href="#" onClick="pop_pummok_del()" class="btnType03">���û���</a>&nbsp;<a href="#" onClick="pop_pummok()" class="btnType03">��� ǰ����</a></th>
							</tr>
							<tr>
								<th class="first" scope="col">����</th>
								<th scope="col">�뵵����</th>
                                <th scope="col">ǰ�񱸺�</th>
                                <th scope="col">ǰ���ڵ�</th>
								<th scope="col">ǰ���</th>
								<th scope="col">�԰�</th>
                                <th scope="col">������</th>
								<th scope="col">������</th>
							</tr>
						</thead>
						<tbody>
						<%
							for i = 1 to 20
						%>
			  				<tr id="pummok_list<%=i%>" style="display:none">
								<td class="first"><input type="checkbox" name="del_check<%=i%>" id="del_check<%=i%>" value="Y"></td>
								<td>
                                <input name="srv_type<%=i%>" type="text" id="srv_type<%=i%>" value="<%=pummok_tab(1,i)%>" style="width:90px" readonly="true">
                                </td>
                                <td>
                                <input name="goods_gubun<%=i%>" type="text" id="goods_gubun<%=i%>" value="<%=pummok_tab(2,i)%>" style="width:120px" readonly="true">
                                </td>
                                <td>
                                <input name="goods_code<%=i%>" type="text" id="goods_code<%=i%>" value="<%=pummok_tab(3,i)%>" style="width:90px" readonly="true">
                                </td>
								<td><input name="goods_name<%=i%>" type="text" id="goods_name<%=i%>" value="<%=pummok_tab(4,i)%>" style="width:250px" readonly="true">
                                </td>
								<td><input name="goods_standard<%=i%>" type="text" id="goods_standard<%=i%>" value="<%=pummok_tab(5,i)%>" style="width:160px"></td>
                                <td><input name="jqty<%=i%>" type="text" id="jqty<%=i%>" style="width:80px;text-align:right" value="<%=formatnumber(jqty,0)%>" readonly="true"></td>
								<td><input name="qty<%=i%>" type="text" id="qty<%=i%>" style="width:80px;text-align:right" value="<%=formatnumber(amount_tab(1,i),0)%>" onKeyUp="NumCal(this);">
                                <input type="hidden" name="goods_grade<%=i%>" value="<%=pummok_tab(6,i)%>" ID="Hidden1">
                                </td>
							</tr>
						<%
							next
						%>
						</tbody>
					</table>
                    <br>
					<table cellpadding="0" cellspacing="0" summary="" class="tableWrite">
						<colgroup>
							<col width="12%" >
							<col width="21%" >
							<col width="13%" >
							<col width="21%" >
							<col width="12%" >
							<col width="*" >
						</colgroup>
						<tbody>
							<tr>
							  <th>÷��</th>
                              <td colspan="5" class="left">
                              <a href="download.asp?path=<%=path_name%>&att_file=<%=chulgo_att_file%>"><%=chulgo_att_file%></a>
                              <input name="att_file" type="file" id="att_file" size="100">
                              </td>
						    </tr>
						</tbody>
					</table>                    
					<br>
				</div>
                <div align=center>
                    <span class="btnType01"><input type="button" value="����" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
                    <span class="btnType01"><input type="button" value="���" onclick="javascript:goAction();"></span>
            <% if u_type = "U" then	%>
                    <span class="btnType01"><input type="button" value="����" onclick="javascript:delcheck();"></span>
			<% end if	%>                          
                </div>
                <input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
                
                <input type="hidden" name="mok_cnt" value="<%=mok_cnt%>">
                <input type="hidden" name="pummok_cnt" value="<%=pummok_cnt%>">
                <input type="hidden" name="chulgo_id" value="<%=chulgo_id%>">
				</form>
                </div>
			</div>
		</div>		
	</body>
</html>
