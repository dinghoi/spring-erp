<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
dim code_tab(20)
dim goods_name(20)
dim goods_type(20)
dim goods_gubun(20)
dim goods_standard(20)
dim goods_grade(20)
dim qty_tab(20)
dim j_qty_tab(20)

dim service_no(20)
dim trade_name(20)
dim trade_dept(20)
dim r_bigo(20)

' 재고조회후 곧바로 출고처리 하는 프로그램 입니다...met_culgo_cust_add01과 유사

user_name = request.cookies("nkpmg_user")("coo_user_name")
user_id = request.cookies("nkpmg_user")("coo_user_id")
cost_grade = request.cookies("nkpmg_user")("coo_cost_grade")

u_type = request("u_type")

view_condi=Request("view_condi")
stock_code=Request("stock_code")
stock_goods_type=Request("stock_goods_type")
stock_goods_code=Request("stock_goods_code")

curr_date = mid(cstr(now()),1,10)
chulgo_date = curr_date

chulgo_id = "CE출고"
chulgo_goods_type = stock_goods_type

mok_cnt = 0
pummok_cnt = 0

for i = 1 to 20
	code_tab(i) = ""
	goods_name(i) = ""
	goods_type(i) = ""
	goods_gubun(i) = ""
	goods_standard(i) = ""
	goods_grade(i) = ""
	qty_tab(i) = 0
	j_qty_tab(i) = 0
	
	service_no(i) = ""
	trade_name(i) = ""
	trade_dept(i) = ""
	r_bigo(i) = ""
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
Set Rs_jago = Server.CreateObject("ADODB.Recordset")
Set Rs_max = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

'출고창고 찾고
Sql = "SELECT * FROM met_stock_code where stock_code = '"&stock_code&"'"
Set Rs_stock = DbConn.Execute(SQL)
if not Rs_stock.eof then
       	   stock_level = Rs_stock("stock_level")
		   chulgo_stock = Rs_stock("stock_code")
		   chulgo_stock_name = Rs_stock("stock_name")
		   chulgo_stock_company = Rs_stock("stock_company")
		   stock_bonbu = Rs_stock("stock_bonbu")
		   stock_saupbu = Rs_stock("stock_saupbu")
		   stock_team = Rs_stock("stock_team")
		   stock_manager_code = Rs_stock("stock_manager_code")
		   stock_manager_name = Rs_stock("stock_manager_name")
    else
		   stock_level = ""
		   chulgo_stock = ""
		   chulgo_stock_name = ""
		   chulgo_stock_company = ""
		   stock_bonbu = ""
		   stock_saupbu = ""
		   stock_team = ""
		   stock_manager_code = ""
		   stock_manager_name = ""
end if
Rs_stock.close()

'출고품목 찾고
mok_cnt = 1
i = 1
sql="select * from met_stock_gmaster where stock_code='"&stock_code&"' and stock_goods_code='"&stock_goods_code&"' and stock_goods_type='"&stock_goods_type&"'"
set Rs_jago=dbconn.execute(sql)
if not Rs_jago.eof then
    code_tab(i) = Rs_jago("stock_goods_code")
	goods_name(i) = Rs_jago("stock_goods_name")
	goods_type(i) = Rs_jago("stock_goods_type")
	goods_gubun(i) = Rs_jago("stock_goods_gubun")
	goods_standard(i) = Rs_jago("stock_goods_standard")
	goods_grade(i) = Rs_jago("stock_goods_grade")
	j_qty_tab(i) = Rs_jago("stock_JJ_qty")
end if
Rs_jago.close()


title_line = " CE출고(창고) 등록 "

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
				a=confirm('입력하시겠습니까?')
				if (a==true) {
					return true;
				}
				return false;
				}
			}

			function chkfrm1() {
				if(document.frm.chulgo_goods_type.value == "") {
					alert('용도구분을 선택하세요');
					frm.chulgo_goods_type.focus();
					return false;}
				if(document.frm.chulgo_stock_name.value == "") {
					alert('출고창고를 선택하세요');
					frm.chulgo_stock_name.focus();
					return false;}
				if(document.frm.rele_company.value == "") {
					alert('출고요청회사를 선택하세요');
					frm.rele_company.focus();
					return false;}
				if(document.frm.rele_saupbu.value == "") {
					alert('출고요청사업부를 선택하세요');
					frm.rele_saupbu.focus();
					return false;}
				if(document.frm.rele_stock_name.value == "") {
					alert('요청창고(입고)를 선택하세요');
					frm.rele_stock_name.focus();
					return false;}
				if(document.frm.chulgo_date.value == "") {
					alert('출고일를 입력하세요');
					frm.chulgo_date.focus();
					return false;}

										
				{
				a=confirm('입력하시겠습니까?')
				if (a==true) {
					return true;
				}
				return false;
				}
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
					alert ("재고수량보다 출고수량이 많습니다!!");
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
				mok_cnt = parseInt(document.frm.mok_cnt.value);
				for (j=1;j<mok_cnt+1;j++) {
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
					document.frm.action = "met_chulgo_reg_del_ok.asp";
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
							  <th>출고일자</th>
							  <td class="left"><input name="chulgo_date" type="text" value="<%=chulgo_date%>" style="width:120px;text-align:center" id="datepicker"></td>
                              <th>용도구분</th>
							  <td class="left">
                              <input name="chulgo_goods_type" type="text" value="<%=chulgo_goods_type%>" readonly="true" style="width:120px">
                              </td>
                              <th>출고창고</th>
							  <td colspan="3" class="left">
                              <input name="chulgo_stock_company" type="text" value="<%=chulgo_stock_company%>" readonly="true" style="width:120px">
                              
                              <input name="chulgo_stock_name" type="text" value="<%=chulgo_stock_name%>" readonly="true" style="width:120px">
                              
						      <a href="#" class="btnType03" onClick="pop_Window('meterials_stock_select.asp?gubun=<%="chulgo"%>&view_condi=<%=view_condi%>','stock_search_pop','scrollbars=yes,width=600,height=400')">찾기</a>
                              <input type="hidden" name="chulgo_stock" value="<%=chulgo_stock%>" ID="Hidden1">
                              <input type="hidden" name="stock_bonbu" value="<%=stock_bonbu%>" ID="Hidden1">
                              <input type="hidden" name="stock_saupbu" value="<%=stock_saupbu%>" ID="Hidden1">
                              <input type="hidden" name="stock_team" value="<%=stock_team%>" ID="Hidden1">
                              <input type="hidden" name="stock_manager_code" value="<%=stock_manager_code%>" ID="Hidden1">
                              <input type="hidden" name="stock_manager_name" value="<%=stock_manager_name%>" ID="Hidden1">
                              </td>
                            </tr>
							<tr>
                              <th>요청창고(입고)</th>
							  <td colspan="3" class="left">
                              <input name="rele_stock_company" type="text" value="<%=rele_stock_company%>" readonly="true" style="width:120px">
                              
                              <input name="rele_stock_name" type="text" value="<%=rele_stock_name%>" readonly="true" style="width:120px">
                              
						      <a href="#" class="btnType03" onClick="pop_Window('meterials_stock_select.asp?gubun=<%="mvreg"%>&view_condi=<%=view_condi%>','rstock_search_pop','scrollbars=yes,width=600,height=400')">찾기</a>
                              <input type="hidden" name="rele_stock" value="<%=rele_stock%>" ID="Hidden1">
                              <input type="hidden" name="rele_stock_bonbu" value="<%=rele_stock_bonbu%>" ID="Hidden1">
                              <input type="hidden" name="rele_stock_saupbu" value="<%=rele_stock_saupbu%>" ID="Hidden1">
                              <input type="hidden" name="rele_stock_team" value="<%=rele_stock_team%>" ID="Hidden1">
                              <input type="hidden" name="rele_manager_code" value="<%=rele_manager_code%>" ID="Hidden1">
                              <input type="hidden" name="rele_manager_name" value="<%=rele_manager_name%>" ID="Hidden1">
                              </td>
                              <th>요청그룹사</th>
							  <td class="left">
							<%
								' 2019.02.22 박정신 요청 회사리스트를 빼고자 할시 org_end_date에 null 이 아닌 만료일자를 셋팅하면 리스트에 나타나지 않는다.
								Sql = "SELECT * FROM emp_org_mst WHERE ISNULL(org_end_date) AND org_level = '회사'  ORDER BY org_company ASC"
                                rs_org.Open Sql, Dbconn, 1
                            %>
                                <select name="rele_company" id="rele_company" value="<%=rele_company%>" style="width:120px">
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
							  <th>요청사업부</th>
							  <td class="left">
							<%
                                Sql="select org_name from emp_org_mst where org_level = '사업부' group by org_name order by org_name asc"
                                rs_org.Open Sql, Dbconn, 1
                            %>
                                <select name="rele_saupbu" id="rele_saupbu" value="<%=rele_saupbu%>" style="width:120px">
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
						    </tr>
                            <tr>
							  <th>비고</th>
							  <td class="left" colspan="8" ><textarea name="chulgo_memo" rows="3" style="text-align:left; ime-mode:active" id="textarea"><%=chulgo_memo%></textarea></td>
						    </tr>
						</tbody>
					</table>
				</div>
                <br>
                <h3 class="stit" style="font-size:12px;">◈ 출고 세부 내역 ◈</h3>
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
								<th class="first" scope="col">선택</th>
								<th scope="col">용도구분</th>
                                <th scope="col">품목구분</th>
                                <th scope="col">품목코드</th>
								<th scope="col">품목명</th>
								<th scope="col">규격</th>
                                <th scope="col">재고수량</th>
								<th scope="col">출고수량</th>
							</tr>
						</thead>
						<tbody>
						<%
							for i = 1 to 20
						%>
			  				<tr id="pummok_list<%=i%>" style="display:none">
								<td class="first"><%=i%></td>
								<td><%=goods_type(i)%>
                                <input type="hidden" name="srv_type<%=i%>" value="<%=goods_type(i)%>" id="srv_type<%=i%>">
                                </td>
                                <td><%=goods_gubun(i)%>
                                <input type="hidden" name="goods_gubun<%=i%>" value="<%=goods_gubun(i)%>" id="goods_gubun<%=i%>">
                                </td>
                                <td><%=code_tab(i)%>
                                <input type="hidden" name="goods_code<%=i%>" value="<%=code_tab(i)%>" id="goods_code<%=i%>">
                                </td>
								<td><%=goods_name(i)%>
                                <input type="hidden" name="goods_name<%=i%>" value="<%=goods_name(i)%>" id="goods_name<%=i%>">
                                <input type="hidden" name="goods_grade<%=i%>" value="<%=goods_grade(i)%>" ID="goods_grade<%=i%>">
                                </td>
								<td><%=goods_standard(i)%>
                                <input type="hidden" name="goods_standard<%=i%>" value="<%=goods_standard(i)%>" id="goods_standard<%=i%>">
                                </td>
                                <td align="right"><%=formatnumber(j_qty_tab(i),0)%>
                                <input type="hidden" name="jqty<%=i%>" value="<%=formatnumber(j_qty_tab(i),0)%>" ID="Hidden1">
                                </td>
								<td><input name="qty<%=i%>" type="text" id="qty<%=i%>" style="width:80px;text-align:right" value="<%=formatnumber(qty_tab(i),0)%>" onKeyUp="NumCal(this);">
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
                              <a href="download.asp?path=<%=path_name%>&att_file=<%=chulgo_att_file%>"><%=chulgo_att_file%></a>
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
                
                <input type="hidden" name="mok_cnt" value="<%=mok_cnt%>">
                <input type="hidden" name="pummok_cnt" value="<%=pummok_cnt%>">
                <input type="hidden" name="chulgo_id" value="<%=chulgo_id%>">
				</form>
                </div>
			</div>
		</div>		
	</body>
</html>
