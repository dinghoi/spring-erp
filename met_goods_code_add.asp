<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
u_type = request("u_type")
org_code = request("org_code")

goods_type=Request("goods_type")
goods_code=Request("goods_code")
goods_name=Request("goods_name")

curr_date = mid(cstr(now()),1,10)
goods_date = curr_date

code_last = ""

goods_level1 = ""
goods_level2 = ""
goods_seq = ""
goods_gubun = ""
goods_model = ""
goods_group = "자산"
goods_serial_no = ""
goods_name = ""
goods_standard = ""
goods_grade = ""

goods_used_sw = "Y"
goods_end_date = ""
goods_tax_id = "과세"
goods_stock_In_type = ""
goods_security_yn = "N"
goods_security_cnt = 0

part_number = ""
po_number = ""

' response.write(reg_date)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_good = Server.CreateObject("ADODB.Recordset")
Set Rs_max = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

title_line = " 품목 등록 "

if u_type = "U" then

	Sql="select * from met_goods_code where goods_code = '"&goods_code&"'"
	Set rs=DbConn.Execute(Sql)

	goods_type = rs("goods_type")
	goods_code = rs("goods_code")
	goods_level1 = rs("goods_level1")
	goods_level2 = rs("goods_level2")
	goods_seq = rs("goods_seq")
	goods_grade = rs("goods_grade")
	goods_gubun = rs("goods_gubun")
	goods_model = rs("goods_model")
    goods_group = rs("goods_group")
    goods_serial_no = rs("goods_serial_no")
	part_number = rs("part_number")
	po_number = rs("po_number")
    goods_name = rs("goods_name")
    goods_standard = rs("goods_standard")
    goods_date = rs("goods_date")
    goods_used_sw = rs("goods_used_sw")
	goods_end_date = rs("goods_end_date")
    goods_tax_id = rs("goods_tax_id")
    goods_stock_In_type = rs("goods_stock_In_type")
    goods_security_yn = rs("goods_security_yn")
    goods_security_cnt = rs("goods_security_cnt")
	goods_comment = rs("goods_comment")
	goods_bigo = rs("goods_bigo")
	reg_date = rs("reg_date")
	reg_user = rs("reg_user")
	mod_date = rs("mod_date")
	mod_user = rs("mod_user")
	if goods_end_date = "1900-01-01" then
	      goods_end_date = ""
	end if
	
	rs.close()
    
	title_line = " 품목 변경 "
	
	
	
end if

'    sql="select max(goods_code) as max_seq from met_goods_code"
'	set rs_max=dbconn.execute(sql)
	
'	if	isnull(rs_max("max_seq"))  then
'		code_last = "0000000001"
'	  else
'		max_seq = "0000000000" + cstr((int(rs_max("max_seq")) + 1))
'		code_last = right(max_seq,10)
'	end if
'    rs_max.close()
	
'if u_type = "U" then
	   code_last = goods_code
'end if
	
'goods_code = code_last

'    sql="select max(goods_seq) as max_seq from met_goods_code where goods_level1 = '"&goods_level1&"' and goods_level2 = '"&goods_level2&"'"
'	set rs_max=dbconn.execute(sql)
	
'	if	isnull(rs_max("max_seq"))  then
'		code_last = "001"
'	  else
'		max_seq = "000" + cstr((int(rs_max("max_seq")) + 1))
'		code_last = right(max_seq,3)
'	end if
'    rs_max.close()
	
'	if u_type = "U" then
'	   code_last = goods_seq
'	end if
	
'goods_seq = code_last

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
				return "6 1";
			}
		</script>
		<script type="text/javascript">
			$(function() {    $( "#datepicker" ).datepicker();
												$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker" ).datepicker("setDate", "<%=goods_date%>" );
			});	  
			$(function() {    $( "#datepicker1" ).datepicker();
												$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker1" ).datepicker("setDate", "<%=goods_end_date%>" );
			});	  
			function goAction () {
			   window.close () ;
			}
			function goBefore () {
			   history.back() ;
			}
			function frmcheck () {
				if (chkfrm()) {
					document.frm.submit ();
				}
			}			
     		function chkfrm() {
//				if(document.frm.goods_type.value =="") {
//					alert('품목구분을 선택하세요');
//					frm.goods_type.focus();
//					return false;}
				if(document.frm.goods_grade.value =="") {
					alert('상태를 선택하세요');
					frm.goods_grade.focus();
					return false;}
				if(document.frm.goods_level1.value =="") {
					alert('대분류를 선택하세요');
					frm.goods_level1.focus();
					return false;}
				if(document.frm.goods_level2.value =="") {
					alert('중분류를 선택하세요');
					frm.goods_level2.focus();
					return false;}
				if(document.frm.goods_gubun.value =="") {
					alert('품목구분을 선택하세요');
					frm.goods_gubun.focus();
					return false;}
				if(document.frm.goods_name.value =="") {
					alert('품목명을 입력하세요');
					frm.goods_name.focus();
					return false;}			
//				if(document.frm.goods_standard.value =="") {
//					alert('규격을 입력하세요');
//					frm.goods_standard.focus();
//					return false;}	
				if(document.frm.goods_date.value =="") {
					alert('품목생성일을 입력하세요');
					frm.goods_date.focus();
					return false;}	
				if(document.frm.goods_group.value =="") {
					alert('품목분류를 선택하세요');
					frm.goods_group.focus();
					return false;}	
				
				if(document.frm.goods_security_yn.value =="Y") {
					if(document.frm.goods_security_cnt.value =="") {
						alert('적정재고수량을 입력하세요');
						frm.goods_security_cnt.focus();
						return false;}}		
				if(document.frm.goods_used_sw.value =="Y") {
					if(document.frm.goods_end_date.value =="") {
						alert('중지일을 입력하세요');
						frm.goods_end_date.focus();
						return false;}}		
				{
				a=confirm('입력하시겠습니까?')
				if (a==true) {
					return true;
				}
				return false;
				}
			}	
			
			function num_chk(txtObj){
				security_cnt = parseInt(document.frm.goods_security_cnt.value.replace(/,/g,""));		

				security_cnt = String(security_cnt);
				num_len = security_cnt.length;
				sil_len = num_len;
				security_cnt = String(security_cnt);
				if (security_cnt.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) security_cnt = security_cnt.substr(0,num_len -3) + "," + security_cnt.substr(num_len -3,3);
				if (sil_len > 6) security_cnt = security_cnt.substr(0,num_len -6) + "," + security_cnt.substr(num_len -6,3) + "," + security_cnt.substr(num_len -2,3);
				document.frm.goods_security_cnt.value = security_cnt; 
			}
		</script>

	</head>
	<body>
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="met_goods_code_save.asp" method="post" name="frm">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableWrite">
						<colgroup>
						       <col width="11%" >
						       <col width="22%" >
						       <col width="11%" >
						       <col width="22%" >
						       <col width="11%" >
						       <col width="*" >
						</colgroup>
						<tbody>
                            <tr>
                                <th>품목코드</th>
                                <td class="left">
                                <input name="goods_code" type="text" id="goods_code" style="width:120px" value="<%=goods_code%>" readonly="true"></td>    
                                <th>등록일</th>
                                <td class="left">
                                <input name="goods_date" type="text" size="10" readonly="true" id="datepicker" style="width:120px;" value="<%=goods_date%>" >
              					</td>
                                <th>상태</th>
                                <td class="left">
                                <select name="goods_grade" id="goods_grade" value="<%=goods_grade%>" style="width:130px">
			            	        <option value="" <% if goods_grade = "" then %>selected<% end if %>>선택</option>
				                    <option value='신품' <%If goods_grade = "신품" then %>selected<% end if %>>신품</option>
                                    <option value='중고' <%If goods_grade = "중고" then %>selected<% end if %>>중고</option>
                                    <option value='리퍼' <%If goods_grade = "리퍼" then %>selected<% end if %>>리퍼</option>
                                </select> 
                                </td>    
                            </tr>
							<tr>
								<th>대분류</th>
                                <td class="left">
                         <% if u_type = "U" then %>
                                <input name="goods_level1" type="text" id="goods_level1" style="width:130px" value="<%=goods_level1%>" readonly="true"></td> 
                         <%    else
								Sql="select * from met_etc_code where etc_type = '02' order by group_name DESC"
								Rs_etc.Open Sql, Dbconn, 1
 						 %>
                                <select name="goods_level1" id="goods_level1" style="width:120px" value="<%=goods_level1%>">
                                <option value="" <% if goods_level1 = "" then %>selected<% end if %>>선택</option>
                         <%
								do until rs_etc.eof 
 			  			 %>
                                <option value='<%=rs_etc("etc_name")%>' <%If goods_level1 = rs_etc("etc_name") then %>selected<% end if %>><%=rs_etc("group_name")%>&nbsp;(<%=rs_etc("etc_name")%>)</option>
                 		 <%
									rs_etc.movenext() 
								loop 
								rs_etc.Close()
						 %>
            					</select>
                                </td>  
                         <% end if %>  
                                <th>중분류</th>
                                <td class="left">
                         <% if u_type = "U" then %>
                                <input name="goods_level2" type="text" id="goods_level2" style="width:120px" value="<%=goods_level2%>" readonly="true"></td> 
                         <%    else
								Sql="select * from met_etc_code where etc_type = '03' order by group_name DESC"
								Rs_etc.Open Sql, Dbconn, 1
 						 %>
                                <select name="goods_level2" id="goods_level2" style="width:120px" value="<%=goods_level2%>">
                                <option value="" <% if goods_level2 = "" then %>selected<% end if %>>선택</option>
                         <%
								do until rs_etc.eof 
 			  			 %>
                                <option value='<%=rs_etc("etc_name")%>' <%If goods_level2 = rs_etc("etc_name") then %>selected<% end if %>><%=rs_etc("group_name")%>&nbsp;(<%=rs_etc("etc_name")%>)</option>
                 		 <%
									rs_etc.movenext() 
								loop 
								rs_etc.Close()
						 %>
            					</select>
                                </td>  
                         <% end if %>         
                                <th>소분류</th>
                                <td class="left">
                                <input name="goods_seq" type="text" id="goods_seq" style="width:130px" value="<%=goods_seq%>" readonly="true"></td>
                             </tr>
                             <tr>
								<th>품목명</th>
                                <td colspan="3" class="left">
                                <input name="goods_name" type="text" id="goods_name" style="width:360px" value="<%=goods_name%>"></td>   
                                <th class="left">품목구분</th>
                                <td class="left">
                         <%
								Sql="select * from met_etc_code where etc_type = '04' order by etc_code asc"
								Rs_etc.Open Sql, Dbconn, 1
 						 %>
                                <select name="goods_gubun" id="goods_gubun" style="width:130px" value="<%=goods_gubun%>">
                                <option value="" <% if goods_gubun = "" then %>selected<% end if %>>선택</option>
                         <%
								do until rs_etc.eof 
 			  			 %>
                                <option value='<%=rs_etc("etc_name")%>' <%If goods_gubun = rs_etc("etc_name") then %>selected<% end if %>><%=rs_etc("etc_name")%></option>
                 		 <%
									rs_etc.movenext() 
								loop 
								rs_etc.Close()
						 %>
            					</select>
                                </td>        
                             </tr>
                             <tr>
								<th>규격</th>
                                <td colspan="3" class="left">
                                <input name="goods_standard" type="text" id="goods_standard" style="width:360px" value="<%=goods_standard%>"></td>
                                <th>품목분류</th>
                                <td class="left">
					            <input type="radio" name="goods_group" value="자산" <% if goods_group = "자산" then %>checked<% end if %>>자산 
              		            <input name="goods_group" type="radio" value="소모성" <% if goods_group = "소모성" then %>checked<% end if %>>소모성
                                </td>
                                  
                             </tr>
                             <tr>
                                <th>모델</th>
                                <td colspan="3" class="left">
                                <input name="goods_model" type="text" id="goods_model" style="width:360px" value="<%=goods_model%>"></td>  
                                <th>적정재고<br>관리</th>
                                <td class="left">
					            <input type="radio" name="goods_security_yn" value="Y" <% if goods_security_yn = "Y" then %>checked<% end if %>>대상 
              		            <input name="goods_security_yn" type="radio" value="N" <% if goods_security_yn = "N" then %>checked<% end if %>>비대상
                                </td>
                             </tr>
                             <tr>
                                <th>Part<br>_Number</th>
                                <td colspan="3" class="left">
                                <input name="part_number" type="text" id="part_number" style="width:360px" value="<%=part_number%>"></td>
                                
                                <th>적정재고<br>수량</th>
                                <td class="left">
                                <input name="goods_security_cnt" type="text" value="<%=formatnumber(goods_security_cnt,0)%>" style="width:130px;text-align:right" onKeyUp="num_chk(this);">
                                </td>
                             </tr>
                             <tr>
                                <th>Serial No.</th>
                                <td colspan="5" class="left">
                                <input name="goods_serial_no" type="text" id="goods_serial_no" style="width:360px" value="<%=goods_serial_no%>"></td>   
                             </tr>
                             <tr>
                                <th>상세설명</th>
                                <td class="left" colspan="5" >
                                <textarea name="goods_comment" rows="3" id="textarea"><%=goods_comment%></textarea>
                                </td>
                            </tr>
                            <tr>
                                <th>비고</th>
                                <td colspan="5" class="left">
                                <input name="goods_bigo" type="text" id="goods_bigo" style="width:360px" value="<%=goods_bigo%>"></td>
                            </tr>
                            <tr>
                                <th>사용유무</th>
                                <td colspan="2" class="left">
					            <input type="radio" name="goods_used_sw" value="Y" <% if goods_used_sw = "Y" then %>checked<% end if %>>사용 
              		            <input name="goods_used_sw" type="radio" value="N" <% if goods_used_sw = "N" then %>checked<% end if %>>중지
                                </td>
                                <th>중지일</th>
                                <td colspan="2" class="left">
                                <input name="goods_end_date" type="text" size="10" readonly="true" id="datepicker1" style="width:130px;" value="<%=goods_end_date%>" >
              					</td>
                            </tr>
                            <tr>
                                <th class="first">입력일자</th>
                                <td colspan="2" class="left"><%=reg_date%>(<%=reg_user%>)</td>
                                <th>수정일자</th>
                                <td colspan="2" class="left"><%=mod_date%>(<%=mod_user%>)</td>
                            </tr>
						</tbody>
					</table>
				</div>
                <br>
                <div align=center>
                    <span class="btnType01"><input type="button" value="저장" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
                    <span class="btnType01"><input type="button" value="취소" onclick="javascript:goAction();"></span>
                </div>
                <input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
				</form>
		</div>				
	</body>
</html>

