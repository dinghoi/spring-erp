<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
dim pummok_tab(6,20)
dim amount_tab(3,20)

user_name = request.cookies("nkpmg_user")("coo_user_name")
user_id = request.cookies("nkpmg_user")("coo_user_id")
cost_grade = request.cookies("nkpmg_user")("coo_cost_grade")

u_type = request("u_type")

stin_in_date=Request("stin_in_date")
stin_order_no=Request("stin_order_no")
stin_order_seq=Request("stin_order_seq")

curr_date = mid(cstr(now()),1,10)

pummok_cnt = 0

path_name = "/met_upload"

for i = 1 to 6
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
Set Rs_buy = Server.CreateObject("ADODB.Recordset")
Set Rs_bg = Server.CreateObject("ADODB.Recordset")
Set Rs_good = Server.CreateObject("ADODB.Recordset")
Set Rs_max = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

if u_type = "U" then

	Sql="select * from met_stin where (stin_in_date = '"&stin_in_date&"') and (stin_order_no = '"&stin_order_no&"') and (stin_order_seq = '"&stin_order_seq&"')"
	Set rs=DbConn.Execute(Sql)

	stin_in_date = rs("stin_in_date")
	stin_order_no = rs("stin_order_no")
	stin_order_seq = rs("stin_order_seq")
	
	stin_buy_company = rs("stin_buy_company")
	buy_company = stin_buy_company
	stin_buy_saupbu = rs("stin_buy_saupbu")
	buy_saupbu = stin_buy_saupbu
		
	stin_goods_type = rs("stin_goods_type")
	buy_goods_type = stin_goods_type
    'stin_trade_no = rs("stin_trade_no")
	stin_trade_no = mid(rs("stin_trade_no"),1,3) + "-" + mid(rs("stin_trade_no"),4,2) + "-" + mid(rs("stin_trade_no"),6)
	stin_trade_name = rs("stin_trade_name")
    stin_trade_person = rs("stin_trade_person")
	stin_trade_email = rs("stin_trade_email")
	
    stin_stock_company = rs("stin_stock_company")
    stin_stock_code = rs("stin_stock_code")
    stin_stock_name = rs("stin_stock_name")
    stin_id = rs("stin_id")
    stin_type = rs("stin_type")
	
    stin_price = rs("stin_price")
    stin_cost = rs("stin_cost")
    stin_cost_vat = rs("stin_cost_vat")
	
	stin_company = rs("stin_company")
    stin_org_name = rs("stin_org_name")
    stin_emp_no = rs("stin_emp_no")
	stin_emp_name = rs("stin_emp_name")

	stin_att_file = rs("stin_att_file")
	stin_memo = rs("stin_memo")
	
	rs.close()
	
	j = 0
	Sql="select * from met_stin_goods where (stin_date = '"&stin_in_date&"') and (stin_order_no = '"&stin_order_no&"') and (stin_order_seq = '"&stin_order_seq&"')"
	Rs.Open Sql, Dbconn, 1
	do until rs.eof
		j = j + 1
		pummok_tab(1,j) = rs("stin_goods_type")
		pummok_tab(2,j) = rs("stin_goods_gubun")
		pummok_tab(3,j) = rs("stin_goods_code")
		pummok_tab(4,j) = rs("stin_goods_name")
		pummok_tab(5,j) = rs("stin_standard")

		amount_tab(1,j) = rs("stin_qty")
		amount_tab(2,j) = rs("stin_unit_cost")
		amount_tab(3,j) = rs("stin_amt")
		rs.movenext()
	loop
	pummok_cnt = j
	mok_cnt = j
    
	title_line = " �԰� ���� "
	
end if

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
				return "1 1";
			}
		</script>
		<script type="text/javascript">
			$(function() {    $( "#datepicker" ).datepicker();
												$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker" ).datepicker("setDate", "<%=stin_in_date%>" );
			});	  
			$(function() {    $( "#datepicker1" ).datepicker();
												$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker1" ).datepicker("setDate", "<%=buy_date%>" );
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
				a=confirm('�����Ͻðڽ��ϱ�?')
				if (a==true) {
					return true;
				}
				return false;
				}
			}
			
			function chkfrm1() {
				if(document.frm.buy_goods_type.value == "") {
					alert('�뵵������ �����ϼ���');
					frm.buy_goods_type.focus();
					return false;}
				if(document.frm.trade_name.value == "") {
					alert('���Űŷ�ó�� �����ϼ���');
					frm.trade_name.focus();
					return false;}
				if(document.frm.stin_stock_name.value == "") {
					alert('�԰�â�� �����ϼ���');
					frm.stin_stock_name.focus();
					return false;}
//				if(document.frm.trade_person.value == "") {
//					alert('����ó ����ڸ� �Է��ϼ���');
//					frm.trade_person.focus();
//					return false;}
//				if(document.frm.trade_email.value == "") {
//					alert('����ó ����� �̸����� �Է��ϼ���');
//					frm.trade_email.focus();
//					return false;}

					
				{
				a=confirm('�����Ͻðڽ��ϱ�?')
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
				code_ary[0] = document.frm.goods_code1.value;
				code_ary[1] = document.frm.goods_code2.value;
				code_ary[2] = document.frm.goods_code3.value;
				code_ary[3] = document.frm.goods_code4.value;
				code_ary[4] = document.frm.goods_code5.value;
				code_ary[5] = document.frm.goods_code6.value;
				code_ary[6] = document.frm.goods_code7.value;
				code_ary[7] = document.frm.goods_code8.value;
				code_ary[8] = document.frm.goods_code9.value;
				code_ary[9] = document.frm.goods_code10.value;
				code_ary[10] = document.frm.goods_code11.value;
				code_ary[11] = document.frm.goods_code12.value;
				code_ary[12] = document.frm.goods_code13.value;
				code_ary[13] = document.frm.goods_code14.value;
				code_ary[14] = document.frm.goods_code15.value;
				code_ary[15] = document.frm.goods_code16.value;
				code_ary[16] = document.frm.goods_code17.value;
				code_ary[17] = document.frm.goods_code18.value;
				code_ary[18] = document.frm.goods_code19.value;
				code_ary[19] = document.frm.goods_code20.value;
				goods_type = document.frm.buy_goods_type.value
//				var popupW = 600;
//				var popupH = 400;
//				var left = Math.ceil((window.screen.width - popupW)/2);
//				var top = Math.ceil((window.screen.height - popupH)/2);
//				window.open('met_goods_select.asp?code_ary='+code_ary+'', '�˾�����', 'width='+ popupW +', height='+ popupH +', left='+ left +', top='+ top +', location=no, status=no, menubar=no, toolbar=no, scrollbars=yes, resizable=no');
				
				var url = "met_goods_select.asp?code_ary="+code_ary+'&goods_type='+goods_type;
//				var url = "met_goods_select.asp?code_ary="+code_ary;
				pop_Window(url,'ǰ����','scrollbars=yes,width=800,height=600');
				
				
			} 
			function pop_pummok_del() 
			{ 
				var code_ary = new Array();
				var del_ary = new Array();
//				code_ary[0] = document.frm.goods_code1.value + "/" + document.frm.goods_standard1.value + "/" + document.frm.qty1.value + "/" + document.frm.buy_cost1.value + "/" + document.frm.buy_tot1.value;
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
				goods_type = document.frm.buy_goods_type.value

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
//				var popupW = 600;
//				var popupH = 400;
//				var left = Math.ceil((window.screen.width - popupW)/2);
//				var top = Math.ceil((window.screen.height - popupH)/2);
//				window.open('met_goods_del_ok.asp?code_ary='+code_ary+'&del_ary='+del_ary+'', '���û�ǰ����', 'width='+ popupW +', height='+ popupH +', left='+ left +', top='+ top +', location=no, status=no, menubar=no, toolbar=no, scrollbars=yes, resizable=no');
//				alert("�����Ǿ����ϴ� !!!");
//				NumCal();
				
				var url = "met_goods_del_ok.asp?code_ary="+code_ary+'&del_ary='+del_ary+'&goods_type='+goods_type;				
				pop_Window(url,'���û�ǰ����','scrollbars=yes,width=600,height=400');
				alert("�����Ǿ����ϴ� !!!");
				NumCal();
			} 

		function NumCal(txtObj){
			var qty_ary = new Array();
			var buy_ary = new Array();
			var buy_tot = new Array();

			for (j=1;j<21;j++) {
				qty_ary[j] = eval("document.frm.qty" + j + ".value").replace(/,/g,"");
				buy_ary[j] = eval("document.frm.buy_cost" + j + ".value").replace(/,/g,"");
				buy_tot[j] = qty_ary[j] * buy_ary[j];
				
				buy_cal = qty_ary[j] * buy_ary[j];
				buy_cal = String(buy_cal);
				num_len = buy_cal.length;
				sil_len = num_len;
				if (buy_cal.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) buy_cal = buy_cal.substr(0,num_len -3) + "," + buy_cal.substr(num_len -3,3);
				if (sil_len > 6) buy_cal = buy_cal.substr(0,num_len -6) + "," + buy_cal.substr(num_len -6,3) + "," + buy_cal.substr(num_len -2,3);
				if (sil_len > 9) buy_cal = buy_cal.substr(0,num_len -9) + "," + buy_cal.substr(num_len -9,3) + "," + buy_cal.substr(num_len -5,3) + "," + buy_cal.substr(num_len -1,3);
				eval("document.frm.buy_tot" + j + ".value = buy_cal");
				
			}

			buy_tot_cost = 0;
			for (j=1;j<21;j++) {
				buy_tot_cost = buy_tot_cost + buy_tot[j];
			}
			
			tot_cal = buy_tot_cost;
			tot_cal = String(tot_cal);
			num_len = tot_cal.length;
			sil_len = num_len;
			if (tot_cal.substr(0,1) == "-") sil_len = num_len - 1;
			if (sil_len > 3) tot_cal = tot_cal.substr(0,num_len -3) + "," + tot_cal.substr(num_len -3,3);
			if (sil_len > 6) tot_cal = tot_cal.substr(0,num_len -6) + "," + tot_cal.substr(num_len -6,3) + "," + tot_cal.substr(num_len -2,3);
			if (sil_len > 9) tot_cal = tot_cal.substr(0,num_len -9) + "," + tot_cal.substr(num_len -9,3) + "," + tot_cal.substr(num_len -5,3) + "," + tot_cal.substr(num_len -1,3);
			eval("document.frm.buy_tot_cost.value = tot_cal");
			
			buy_vat = parseInt(buy_tot_cost * (10 / 100));
			buy_hap = buy_tot_cost + buy_vat;
			
			buy_vat = String(buy_vat);
			num_len = buy_vat.length;
			sil_len = num_len;
			if (buy_vat.substr(0,1) == "-") sil_len = num_len - 1;
			if (sil_len > 3) buy_vat = buy_vat.substr(0,num_len -3) + "," + buy_vat.substr(num_len -3,3);
			if (sil_len > 6) buy_vat = buy_vat.substr(0,num_len -6) + "," + buy_vat.substr(num_len -6,3) + "," + buy_vat.substr(num_len -2,3);
			if (sil_len > 9) buy_vat = buy_vat.substr(0,num_len -9) + "," + buy_vat.substr(num_len -9,3) + "," + buy_vat.substr(num_len -5,3) + "," + buy_vat.substr(num_len -1,3);
			eval("document.frm.buy_tot_cost_vat.value = buy_vat");
			
			buy_hap = String(buy_hap);
			num_len = buy_hap.length;
			sil_len = num_len;
			if (buy_hap.substr(0,1) == "-") sil_len = num_len - 1;
			if (sil_len > 3) buy_hap = buy_hap.substr(0,num_len -3) + "," + buy_hap.substr(num_len -3,3);
			if (sil_len > 6) buy_hap = buy_hap.substr(0,num_len -6) + "," + buy_hap.substr(num_len -6,3) + "," + buy_hap.substr(num_len -2,3);
			if (sil_len > 9) buy_hap = buy_hap.substr(0,num_len -9) + "," + buy_hap.substr(num_len -9,3) + "," + buy_hap.substr(num_len -5,3) + "," + buy_hap.substr(num_len -1,3);
			eval("document.frm.buy_tot_price.value = buy_hap");

			if (txtObj.value.length<1) {
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
					document.frm.action = "met_stock_in_del_ok.asp";
					document.frm.submit();
				return true;
				}
				return false;
				}
		</script>

	</head>
	<body onload="pummok_list_view();">
		<div id="container">				
			<div class="gView">
				<h3 class="insa"><%=title_line%></h3>
                <form method="post" name="frm" action="met_stock_in_modify01_save.asp" enctype="multipart/form-data">
                    <table cellpadding="0" cellspacing="0" summary="" class="tableWrite">
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
							  <th>�뵵����</th>
							  <td class="left">
							<%
                                Sql="select * from met_etc_code where etc_type = '01' order by etc_code asc"
					            Rs_etc.Open Sql, Dbconn, 1
                            %>
                                <select name="buy_goods_type" id="buy_goods_type" style="width:120px">
                                    <option value=''>����</option> 
                            <% 
                                do until Rs_etc.eof 
                            %>
                                    <option value='<%=rs_etc("etc_name")%>' <%If buy_goods_type = rs_etc("etc_name") then %>selected<% end if %>><%=rs_etc("etc_name")%></option>
                            <%
                                    Rs_etc.movenext()  
                                loop 
                                Rs_etc.Close()
                            %>
                                </select>
                              </td>
                              <th>���ű׷��</th>
							  <td class="left">
							<%
                                Sql="select * from emp_org_mst where (isNull(org_end_date) or org_end_date = '1900-01-01') and org_level = 'ȸ��' order by org_code asc"
                                rs_org.Open Sql, Dbconn, 1
                            %>
                                <select name="buy_company" id="buy_company" style="width:120px">
                                    <option value=''>����</option> 
                            <% 
                                do until rs_org.eof 
                            %>
                                    <option value='<%=rs_org("org_name")%>' <%If rs_org("org_name") = buy_company  then %>selected<% end if %>><%=rs_org("org_name")%></option>
                            <%
                                    rs_org.movenext()  
                                loop 
                                rs_org.Close()
                            %>
                                </select>
                              </td>
							  <th>���Ż����</th>
							  <td colspan="3" class="left">
							<%
                                Sql="select org_name from emp_org_mst where org_level = '�����' group by org_name order by org_name asc"
                                rs_org.Open Sql, Dbconn, 1
                            %>
                                <select name="buy_saupbu" id="buy_saupbu" style="width:120px">
                                    <option value="����" <%If buy_saupbu = "����" then %>selected<% end if %>>����</option>
                            <% 
                                do until rs_org.eof 
                            %>
                                    <option value='<%=rs_org("org_name")%>' <%If rs_org("org_name") = buy_saupbu  then %>selected<% end if %>><%=rs_org("org_name")%></option>
                            <%
                                    rs_org.movenext()  
                                loop 
                                rs_org.Close()
                            %>
                                </select>
                              </td>
                            </tr>
                            <tr>
							  <th>�԰�����</th>
							  <td class="left"><input name="stin_in_date" type="text" value="<%=stin_in_date%>" style="width:80px;text-align:center" id="datepicker"></td>
                              <th>�԰�â��</th>
							  <td colspan="3" class="left">
                              <input name="stin_stock_name" type="text" value="<%=stin_stock_name%>" readonly="true" style="width:150px">
                              
						      <a href="#" class="btnType03" onClick="pop_Window('meterials_stock_select.asp?gubun=<%="stin"%>&view_condi=<%=view_condi%>','stock_search_pop','scrollbars=yes,width=600,height=400')">ã��</a>
                              <input type="hidden" name="stin_stock_code" value="<%=stin_stock_code%>" ID="Hidden1">
                              <input type="hidden" name="stin_stock_company" value="<%=stin_stock_company%>" ID="Hidden1">
                              <input type="hidden" name="stock_bonbu" value="<%=stock_bonbu%>" ID="Hidden1">
                              <input type="hidden" name="stock_saupbu" value="<%=stock_saupbu%>" ID="Hidden1">
                              <input type="hidden" name="stock_team" value="<%=stock_team%>" ID="Hidden1">
                              <input type="hidden" name="stock_manager_code" value="<%=stock_manager_code%>" ID="Hidden1">
                              <input type="hidden" name="stock_manager_name" value="<%=stock_manager_name%>" ID="Hidden1">
                              </td>
                              <th>�԰���</th>
							  <td class="left"><%=stin_emp_name%>&nbsp;(<%=stin_org_name%>)
                              <input type="hidden" name="emp_no" value="<%=stin_emp_no%>" ID="Hidden1">
                              <input type="hidden" name="emp_name" value="<%=stin_emp_name%>" ID="Hidden1">
                              <input type="hidden" name="emp_company" value="<%=stin_company%>" ID="Hidden1">
                              <input type="hidden" name="emp_org_name" value="<%=stin_org_name%>" ID="Hidden1">
                              </td>
						    </tr>
							<tr>
							  <th>���Űŷ�ó</th>
							  <td colspan="3" class="left"><input name="trade_name" type="text" value="<%=stin_trade_name%>" readonly="true" style="width:150px">
                              <a href="#" class="btnType03" onClick="pop_Window('insa_trade_select.asp?gubun=<%="buy"%>','trade_search_pop','scrollbars=yes,width=600,height=400')">ã��</a>
                              </td>
							  <th>����ڹ�ȣ</th>
							  <td colspan="3" class="left"><input name="trade_no" type="text" value="<%=stin_trade_no%>" readonly="true" style="width:150px">
                              <input type="hidden" name="trade_person" value="<%=stin_trade_person%>" ID="Hidden1">
                              <input type="hidden" name="trade_email" value="<%=stin_trade_email%>" ID="Hidden1">
                              </td>

						    </tr>
							<tr>
							  <th>���</th>
							  <td class="left" colspan="8" ><textarea name="buy_memo" rows="3" style="text-align:left; ime-mode:active" id="textarea"><%=stin_memo%></textarea></td>
						    </tr>
						</tbody>
					</table>
                <br>
                <h3 class="stit" style="font-size:12px;">�� �԰� ���� ���� ��</h3>
            	<table cellpadding="0" cellspacing="0" class="tableWrite">
						<colgroup>
							<col width="4%" >
							<col width="10%" >
                            <col width="*" >
                            <col width="10%" >
							<col width="16%" >
							<col width="14%" >
							<col width="11%" >
							<col width="11%" >
							<col width="11%" >
						</colgroup>
						<thead>
							<tr>
                                <th class="first; left" colspan="9" scope="col" style=" border-bottom:1px solid #e3e3e3;"><a href="#" onClick="pop_pummok_del()" class="btnType03">���û���</a>&nbsp;<a href="#" onClick="pop_pummok()" class="btnType03">ǰ���ڵ弱��</a></th>
							</tr>
							<tr>
								<th class="first" scope="col">����</th>
								<th scope="col">�뵵����</th>
                                <th scope="col">ǰ�񱸺�</th>
                                <th scope="col">ǰ���ڵ�</th>
								<th scope="col">ǰ���</th>
								<th scope="col">�԰�</th>
								<th scope="col">�԰����</th>
								<th scope="col">�԰�ܰ�</th>
								<th scope="col">�԰�ݾ�</th>
							</tr>
						</thead>
						<tbody>
						<%
							for i = 1 to 20
						%>
			  				<tr id="pummok_list<%=i%>" style="display:none">
								<td class="first"><input type="checkbox" name="del_check<%=i%>" id="del_check<%=i%>" value="Y"></td>
								<td>
                                <input name="srv_type<%=i%>" type="text" id="srv_type<%=i%>" value="<%=pummok_tab(1,i)%>" style="width:70px" readonly="true">
                                </td>
                                <td>
                                <input name="goods_gubun<%=i%>" type="text" id="goods_gubun<%=i%>" value="<%=pummok_tab(2,i)%>" style="width:120px" readonly="true">
                                </td>
                                <td>
                                <input name="goods_code<%=i%>" type="text" id="goods_code<%=i%>" value="<%=pummok_tab(3,i)%>" style="width:80px" readonly="true">
                                </td>
								<td><input name="goods_name<%=i%>" type="text" id="goods_name<%=i%>" value="<%=pummok_tab(4,i)%>" style="width:140px" readonly="true"></td>
								<td><input name="goods_standard<%=i%>" type="text" id="goods_standard<%=i%>" value="<%=pummok_tab(5,i)%>" style="width:130px"></td>
								<td><input name="qty<%=i%>" type="text" id="qty<%=i%>" value="<%=formatnumber(amount_tab(1,i),0)%>"  style="width:80px;text-align:right" onKeyUp="NumCal(this);"></td>
								<td><input name="buy_cost<%=i%>" type="text" id="buy_cost<%=i%>" value="<%=formatnumber(amount_tab(2,i),0)%>" style="width:90px;text-align:right" onKeyUp="NumCal(this);" ></td>
								<td><input name="buy_tot<%=i%>" type="text" disabled id="buy_tot<%=i%>" value="<%=formatnumber(amount_tab(3,i),0)%>"  style="width:90px;text-align:right" readonly="true"></td>
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
							  <th>�԰��Ѿ�</th>
							  <td class="left"><input name="buy_tot_price" type="text" id="buy_tot_price" style="width:150px;text-align:right" value="<%=formatnumber(buy_tot_price,0)%>" readonly="true"></td>
							  <th>�԰�ݾ�</th>
							  <td class="left"><input name="buy_tot_cost" type="text" id="buy_tot_cost" style="width:150px;text-align:right" value="<%=formatnumber(buy_tot_cost,0)%>" readonly="true"></td>
							  <th>�԰�ΰ���</th>
							  <td class="left"><input name="buy_tot_cost_vat" type="text" id="buy_tot_cost_vat" style="width:150px;text-align:right" value="<%=formatnumber(buy_tot_cost_vat,0)%>" readonly="true"></td>
						    </tr>
							<tr>
							  <th>÷��</th>
                              <td colspan="5" class="left">
                              <a href="download.asp?path=<%=path_name%>&att_file=<%=stin_att_file%>"><%=stin_att_file%></a>
                              <input name="att_file" type="file" id="att_file" size="100">
                              </td>
						    </tr>
						</tbody>
					</table>
<br>
                <div align=center>
                    <span class="btnType01"><input type="button" value="����" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
                    <span class="btnType01"><input type="button" value="���" onclick="javascript:goAction();"></span>
            <% if u_type = "U" then	%>
                    <span class="btnType01"><input type="button" value="����" onclick="javascript:delcheck();"></span>
			<% end if	%>        
                    
                </div>
                <input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
               
                <input type="hidden" name="pummok_cnt" value="<%=pummok_cnt%>">
                
                <input type="hidden" name="old_stin_in_date" value="<%=stin_in_date%>">
                <input type="hidden" name="old_stin_order_no" value="<%=stin_order_no%>">
				<input type="hidden" name="old_stin_order_seq" value="<%=stin_order_seq%>">
				<input type="hidden" name="old_stin_goods_type" value="<%=stin_goods_type%>">
				<input type="hidden" name="old_stin_att_file" value="<%=stin_att_file%>">
				</form>
                </div>
			</div>
		</div>
   	</body>
</html>
