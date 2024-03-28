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

view_condi=Request("view_condi")
goods_type=Request("goods_type")
rele_id=Request("rele_id")
rele_stock_name=Request("stock")

curr_date = mid(cstr(now()),1,10)
rele_date = curr_date
chulgo_rele_date = curr_date

rele_goods_type = goods_type
rele_emp_no = ""
rele_emp_name = ""
rele_company = ""
rele_bonbu = ""
rele_saupbu = ""
rele_team = ""
rele_org_name = ""
rele_stock_company = ""
rele_stock_name = ""

chulgo_ing = "출고의뢰"
chulgo_stock = ""
chulgo_stock_name = ""
chulgo_stock_company = ""
chulgo_date = "0000-00-00"
rele_sign_yn = "N"
rele_sign_no = ""
rele_sign_date = "0000-00-00"
rele_att_file = ""
rele_memo = ""
rele_addr = ""

mok_cnt = 0
pummok_cnt = 0

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
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set Rs_rele = Server.CreateObject("ADODB.Recordset")
Set Rs_good = Server.CreateObject("ADODB.Recordset")
Set Rs_stock = Server.CreateObject("ADODB.Recordset")
Set rs_trade = Server.CreateObject("ADODB.Recordset")
Set Rs_max = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

'rele_emp_no = "100482"
'rele_emp_name = "오세삼"
rele_emp_no = user_id
rele_emp_name = user_name

Sql = "SELECT * FROM emp_master where emp_no = '"&rele_emp_no&"'"
Set rs_emp = DbConn.Execute(SQL)
if not rs_emp.eof then
    	emp_no = rs_emp("emp_no")
		emp_name = rs_emp("emp_name")
		emp_company = rs_emp("emp_company")
		emp_bonbu = rs_emp("emp_bonbu")
		emp_saupbu = rs_emp("emp_saupbu")
		emp_team = rs_emp("emp_team")
		emp_org_code = rs_emp("emp_org_code")
		emp_org_name = rs_emp("emp_org_name")
   else
		emp_name = ""
		emp_company = ""
		emp_bonbu = ""
		emp_saupbu = ""
		emp_team = ""
		emp_org_code = ""
		emp_org_name = ""
end if
rs_emp.close()

rele_emp_no = emp_no
rele_emp_name = emp_name
rele_company = emp_company
rele_bonbu = emp_bonbu
rele_saupbu = emp_saupbu
rele_team = emp_team
rele_org_name = emp_org_name

'chulgo_stock = "1001" '테스트
'chulgo_stock_company = emp_company
'chulgo_stock_name = emp_company '테스트


title_line = " 창고이동 출고의뢰 등록 "

path_name = "/met_upload"

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>상품자재관리 시스템</title>
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
												$( "#datepicker" ).datepicker("setDate", "<%=rele_date%>" );
			});	  
			$(function() {    $( "#datepicker1" ).datepicker();
												$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker1" ).datepicker("setDate", "<%=chulgo_rele_date%>" );
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
				a=confirm('입력하시겠습니까?')
				if (a==true) {
					return true;
				}
				return false;
				}
			}

			function chkfrm1() {
				if(document.frm.rele_goods_type.value == "") {
					alert('용도구분을 선택하세요');
					frm.rele_goods_type.focus();
					return false;}
				if(document.frm.rele_stock.value == "") {
					alert('신청창고를 선택하세요');
					frm.rele_stock_name.focus();
					return false;}
				if(document.frm.chulgo_stock_name.value == "") {
					alert('출고처창고를 선택하세요');
					frm.chulgo_stock_name.focus();
					return false;}
										
				{
				a=confirm('입력하시겠습니까?')
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
				goods_type = document.frm.rele_goods_type.value
				stock_code = document.frm.chulgo_stock.value
//				var popupW = 600;
//				var popupH = 400;
//				var left = Math.ceil((window.screen.width - popupW)/2);
//				var top = Math.ceil((window.screen.height - popupH)/2);
//				window.open('met_goods_select.asp?code_ary='+code_ary+'', '팝업공지', 'width='+ popupW +', height='+ popupH +', left='+ left +', top='+ top +', location=no, status=no, menubar=no, toolbar=no, scrollbars=yes, resizable=no');
				
				var url = "met_stock_goods_select.asp?code_ary="+code_ary+'&goods_type='+goods_type+'&stock_code='+stock_code;
//				var url = "met_goods_select.asp?code_ary="+code_ary;
				pop_Window(url,'출고의뢰품목선택','scrollbars=yes,width=800,height=400');
				
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
				goods_type = document.frm.rele_goods_type.value
				stock_code = document.frm.chulgo_stock.value	
//				var popupW = 600;
//				var popupH = 400;
//				var left = Math.ceil((window.screen.width - popupW)/2);
//				var top = Math.ceil((window.screen.height - popupH)/2);
//				window.open('met_stock_goods_del_ok.asp?code_ary='+code_ary+'&del_ary='+del_ary+'&goods_type='+goods_type+'&stock_code='+stock_code+'', '팝업공지', 'width='+ popupW +', height='+ popupH +', left='+ left +', top='+ top +', location=no, status=no, menubar=no, toolbar=no, scrollbars=no, resizable=no');
				
				var url = "met_stock_goods_del_ok.asp?code_ary="+code_ary+'&del_ary='+del_ary+'&goods_type='+goods_type+'&stock_code='+stock_code;				
				pop_Window(url,'선택상품삭제','scrollbars=yes,width=600,height=400');
				alert("삭제되었습니다 !!!");
				NumCal();
			} 

		function NumCal(txtObj){
			var qty_ary = new Array();

			for (j=1;j<21;j++) {
				qty_ary[j] = eval("document.frm.qty" + j + ".value").replace(/,/g,"");
		
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
							
				// 음수처리
				if(temp.substr(0,1)=="-") minus="-";
					else minus="";
							
				// 소수점이하처리
				dpoint=temp.search(/\./);
						
				if(dpoint>0)
				{
				// 첫번째 만나는 .을 기준으로 자르고 숫자제외한 문자 삭제
				dpointVa="."+temp.substr(dpoint).replace(/\D/g,"");
				temp=temp.substr(0,dpoint);
				}else dpointVa="";
							
				// 숫자이외문자 삭제
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
				a=confirm('정말 삭제하시겠습니까?')
				if (a==true) {
					document.frm.method = "post";
					document.frm.enctype = "multipart/form-data";
					document.frm.action = "met_mode_reg_del_ok.asp";
					document.frm.submit();
				return true;
				}
				return false;
			}
		</script>

	</head>
	<body>
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
                <form method="post" name="frm" action="met_move_reg_add_save.asp" enctype="multipart/form-data">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableWrite">
						<colgroup>
							<col width="13%" >
							<col width="15%" >
							<col width="13%" >
							<col width="15%" >
							<col width="13%" >
							<col width="33%" >
						</colgroup>
						<tbody>
							<tr>
							  <th>회사</th>
							  <td class="left">
							<%
								' 2019.02.22 박정신 요청 회사리스트를 빼고자 할시 org_end_date에 null 이 아닌 만료일자를 셋팅하면 리스트에 나타나지 않는다.
								Sql = "SELECT * FROM emp_org_mst WHERE ISNULL(org_end_date) AND org_level = '회사'  ORDER BY org_company ASC"
                                rs_org.Open Sql, Dbconn, 1
                            %>
                                <select name="rele_company" id="rele_company" style="width:120px">
                                    <option value=''>선택</option> 
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
							  <th>사업부</th>
							  <td class="left">
							<%
                                Sql="select org_name from emp_org_mst where org_level = '사업부' group by org_name order by org_name asc"
                                rs_org.Open Sql, Dbconn, 1
                            %>
                                <select name="rele_saupbu" id="rele_saupbu" style="width:120px">
                                    <option value=''>선택</option> 
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
                              <th>신청담당자</th>
							  <td class="left"><%=rele_emp_name%>(<%=rele_emp_no%>)&nbsp;-&nbsp;<%=rele_org_name%>
                              <input type="hidden" name="rele_emp_no" value="<%=rele_emp_no%>" ID="rele_emp_no">
                              <input type="hidden" name="rele_emp_name" value="<%=rele_emp_name%>" ID="rele_emp_name">
                              <input type="hidden" name="rele_bonbu" value="<%=rele_bonbu%>" ID="rele_bonbu">
                              <input type="hidden" name="rele_team" value="<%=rele_team%>" ID="rele_team">
                              <input type="hidden" name="rele_org_name" value="<%=rele_org_name%>" ID="rele_org_name">
                              </td>
						    </tr>
							<tr>
							  <th>신청일자</th>
							  <td class="left"><input name="rele_date" type="text" value="<%=rele_date%>" style="width:120px;text-align:center" id="datepicker"></td>
                              <th>용도구분</th>
							  <td class="left">
							<%
                                Sql="select * from met_etc_code where etc_type = '01' order by etc_code asc"
					            Rs_etc.Open Sql, Dbconn, 1
                            %>
                                <select name="rele_goods_type" id="rele_goods_type" style="width:120px">
                                    <option value=''>선택</option> 
                            <% 
                                do until Rs_etc.eof 
                            %>
                                    <option value='<%=rs_etc("etc_name")%>' <%If rele_goods_type = rs_etc("etc_name") then %>selected<% end if %>><%=rs_etc("etc_name")%></option>
                            <%
                                    Rs_etc.movenext()  
                                loop 
                                Rs_etc.Close()
                            %>
                                </select>
                              </td>
                              <th>신청창고</th>
							  <td colspan="3" class="left">
                              <input name="rele_stock_company" type="text" value="<%=rele_stock_company%>" readonly="true" style="width:120px">
                              
                              <input name="rele_stock_name" type="text" value="<%=rele_stock_name%>" readonly="true" style="width:120px">
                              
						      <a href="#" class="btnType03" onClick="pop_Window('meterials_stock_select.asp?gubun=<%="mvreg"%>&view_condi=<%=view_condi%>','stock_search_pop','scrollbars=yes,width=600,height=400')">찾기</a>
                              <input type="hidden" name="rele_stock" value="<%=rele_stock%>" ID="Hidden1">
                              <input type="hidden" name="rele_stock_bonbu" value="<%=rele_stock_bonbu%>" ID="Hidden1">
                              <input type="hidden" name="rele_stock_saupbu" value="<%=rele_stock_saupbu%>" ID="Hidden1">
                              <input type="hidden" name="rele_stock_team" value="<%=rele_stock_team%>" ID="Hidden1">
                              <input type="hidden" name="rele_manager_code" value="<%=rele_manager_code%>" ID="Hidden1">
                              <input type="hidden" name="rele_manager_name" value="<%=rele_manager_name%>" ID="Hidden1">
                              </td>
                            </tr>
							<tr>
							  <th>출고요청일</th>
							  <td class="left"><input name="chulgo_rele_date" type="text" value="<%=chulgo_rele_date%>" style="width:120px;text-align:center" id="datepicker1"></td>
                              <th>출고처창고</th>
							  <td colspan="3" class="left">
                              <input name="chulgo_stock_company" type="text" value="<%=chulgo_stock_company%>" readonly="true" style="width:120px">
                              
                              <input name="chulgo_stock_name" type="text" value="<%=chulgo_stock_name%>" readonly="true" style="width:120px">
                              
						      <a href="#" class="btnType03" onClick="pop_Window('meterials_mstock_select.asp?gubun=<%="chulgo"%>&view_condi=<%=view_condi%>','mstock_search_pop','scrollbars=yes,width=600,height=400')">찾기</a>
                              <input type="hidden" name="chulgo_stock" value="<%=chulgo_stock%>" ID="Hidden1">
                              <input type="hidden" name="stock_bonbu" value="<%=stock_bonbu%>" ID="Hidden1">
                              <input type="hidden" name="stock_saupbu" value="<%=stock_saupbu%>" ID="Hidden1">
                              <input type="hidden" name="stock_team" value="<%=stock_team%>" ID="Hidden1">
                              &nbsp;&nbsp;출고처 창고장
                              <input name="stock_manager_name" type="text" value="<%=stock_manager_name%>" readonly="true" style="width:120px">
                              <input type="hidden" name="stock_manager_code" value="<%=stock_manager_code%>" ID="Hidden1">
                              </td>
						    </tr>
                            <tr>
							  <th>비고</th>
							  <td class="left" colspan="5" ><textarea name="chulgo_memo" rows="3" id="textarea"><%=rele_memo%></textarea></td>
						    </tr>
						</tbody>
					</table>
				</div>
                <br>
                <h3 class="stit" style="font-size:12px;">◈ 창고이동 출고의뢰 세부 내역 ◈</h3>
            	<table cellpadding="0" cellspacing="0" class="tableWrite">
						<colgroup>
							<col width="4%" >
							<col width="10%" >
                            <col width="6%" >
                            <col width="*" >
                            <col width="12%" >
							<col width="16%" >
							<col width="16%" >
                            <col width="10%" >
							<col width="10%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first; left" colspan="8" scope="col" style=" border-bottom:1px solid #e3e3e3;"><a href="#" onClick="pop_pummok_del()" class="btnType03">선택삭제</a>&nbsp;<a href="#" onClick="pop_pummok()" class="btnType03">출고의뢰 품목선택</a></th>
							</tr>
							<tr>
								<th class="first" scope="col"><input type="checkbox" name="tot_check" id="tot_check"></th>
								<th scope="col">용도구분</th>
                                <th scope="col">상태</th>
                                <th scope="col">품목구분</th>
                                <th scope="col">품목코드</th>
								<th scope="col">품목명</th>
								<th scope="col">규격</th>
                                <th scope="col">재고수량</th>
								<th scope="col">의뢰수량</th>
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
                                <input name="goods_grade<%=i%>" type="text" id="goods_grade<%=i%>" value="<%=pummok_tab(6,i)%>" style="width:40px" readonly="true">
                                </td>
                                <td>
                                <input name="goods_gubun<%=i%>" type="text" id="goods_gubun<%=i%>" value="<%=pummok_tab(2,i)%>" style="width:140px" readonly="true">
                                </td>
                                <td>
                                <input name="goods_code<%=i%>" type="text" id="goods_code<%=i%>" value="<%=pummok_tab(3,i)%>" style="width:90px" readonly="true">
                                </td>
								<td><input name="goods_name<%=i%>" type="text" id="goods_name<%=i%>" value="<%=pummok_tab(4,i)%>" style="width:140px" readonly="true"></td>
								<td><input name="goods_standard<%=i%>" type="text" id="goods_standard<%=i%>" value="<%=pummok_tab(5,i)%>" style="width:140px"></td>
                                <td><input name="jqty<%=i%>" type="text" id="jqty<%=i%>" style="width:80px;text-align:right" value="<%=formatnumber(jqty,0)%>" readonly="true"></td>
								<td><input name="qty<%=i%>" type="text" id="qty<%=i%>" style="width:80px;text-align:right" value="<%=formatnumber(amount_tab(1,i),0)%>" onKeyUp="NumCal(this);">
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
							  <th>첨부</th>
                              <td colspan="5" class="left">
                              <a href="download.asp?path=<%=path_name%>&att_file=<%=rele_att_file%>"><%=rele_att_file%></a>
                              <input name="att_file" type="file" id="att_file" size="100">
                              </td>
						    </tr>
						</tbody>
					</table>                    
					<br>
				</div>
                <div align=center>
                    <span class="btnType01"><input type="button" value="저장" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
                    <span class="btnType01"><input type="button" value="취소" onclick="javascript:goAction();"></span>
            <% if u_type = "U" then	%>
                    <span class="btnType01"><input type="button" value="삭제" onclick="javascript:delcheck();"></span>
			<% end if	%>                          
                </div>
                <input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
                <input type="hidden" name="rele_sign_yn" value="<%=rele_sign_yn%>" ID="Hidden1">
                <input type="hidden" name="old_rele_stock" value="<%=rele_stock%>">
                <input type="hidden" name="old_rele_seq" value="<%=rele_seq%>">
				<input type="hidden" name="old_rele_date" value="<%=rele_date%>">
				<input type="hidden" name="old_att_file" value="<%=rele_att_file%>">
                
                <input type="hidden" name="mok_cnt" value="<%=mok_cnt%>">
                <input type="hidden" name="pummok_cnt" value="<%=pummok_cnt%>">
				</form>
		</div>				
	</body>
</html>
