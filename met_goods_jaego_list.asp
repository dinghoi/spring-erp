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
cost_grade = request.cookies("nkpmg_user")("coo_cost_grade")

Page=Request("page")
page_cnt=Request.form("page_cnt")
pg_cnt=cint(Request("pg_cnt"))

be_pg = "met_goods_jaego_list.asp"
curr_date = datevalue(mid(cstr(now()),1,10))

ck_sw=Request("ck_sw")

If ck_sw = "y" Then
	field_check=Request("field_check")
	field_gubun=Request("field_gubun")
	field_grade=Request("field_grade")
	field_group=Request("field_group")
	field_name=Request("field_name")
	field_code=Request("field_code")
	field_stand=Request("field_stand")
	view_c = Request("view_c")
  else
	field_check=Request.form("field_check")
	field_gubun=Request.form("field_gubun")
	field_grade=Request.form("field_grade")
	field_group=Request.form("field_group")
	field_name=Request.form("field_name")
	field_code=Request.form("field_code")
	field_stand=Request.form("field_stand")
	view_c = Request.form("view_c")
End if

'response.write(goods_type)
'response.write(field_check)

pgsize = 10 ' 화면 한 페이지 

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
Set Rs_jae = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")

dbconn.open DbConnect

order_Sql = " ORDER BY goods_gubun,goods_name,goods_standard,goods_code ASC"

If view_c = "" Then
	ck_sw = "n"
	field_check = "total"
	view_c = "gubun"
End If

If field_check = "total" Then
       owner_sql = "select count(*) FROM met_goods_code"
	   field_check = ""
   else
       if view_c = "gubun" Then
              owner_sql = "select count(*) FROM met_goods_code  where goods_gubun like '%" + field_gubun + "%'"
       end if
	   if view_c = "grade" Then
              owner_sql = "select count(*) FROM met_goods_code  where goods_grade like '%" + field_grade + "%'"
       end if
	   if view_c = "group" Then
              owner_sql = "select count(*) FROM met_goods_code  where goods_group like '%" + field_group + "%'"
       end if
	   if view_c = "name" Then
              owner_sql = "select count(*) FROM met_goods_code  where goods_name like '%" + field_name + "%'"
       end if
	   if view_c = "code" Then
              owner_sql = "select count(*) FROM met_goods_code  where goods_code like '%" + field_code + "%'"
       end if
	   if view_c = "stand" Then
              owner_sql = "select count(*) FROM met_goods_code  where goods_standard like '%" + field_stand + "%'"
       end if
End If

sql = owner_sql  
Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

If field_check = "total" Then
       owner_sql = "select * FROM met_goods_code"
	   field_check = ""
   else
       if view_c = "gubun" Then
              owner_sql = "select * FROM met_goods_code  where goods_gubun like '%" + field_gubun + "%'"
       end if
	   if view_c = "grade" Then
              owner_sql = "select * FROM met_goods_code  where goods_grade like '%" + field_grade + "%'"
       end if
	   if view_c = "group" Then
              owner_sql = "select * FROM met_goods_code  where goods_group like '%" + field_group + "%'"
       end if
	   if view_c = "name" Then
              owner_sql = "select * FROM met_goods_code  where goods_name like '%" + field_name + "%'"
       end if
	   if view_c = "code" Then
              owner_sql = "select * FROM met_goods_code  where goods_code like '%" + field_code + "%'"
       end if
	   if view_c = "stand" Then
              owner_sql = "select * FROM met_goods_code  where goods_standard like '%" + field_stand + "%'"
       end if
End If

sql = owner_sql + order_sql + " limit "& stpage & "," &pgsize 
Rs.Open Sql, Dbconn, 1
'response.write(sql)

title_line = " 품목별 재고현황 "

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
				return "4 1";
			}
		</script>
		<script type="text/javascript">
			function frmcheck () {
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {
//				if (document.frm.goods_type.value == "") {
//					alert ("필드조건을 선택하시기 바랍니다");
//					return false;
//				}
				{
					return true;
				}
			}
			function condi_view() {

				if (eval("document.frm.view_c[0].checked")) {
					document.getElementById('gubun1').style.display = 'none';
					document.getElementById('grade1').style.display = 'none';
					document.getElementById('group1').style.display = 'none';
					document.getElementById('name1').style.display = 'none';
					document.getElementById('code1').style.display = '';
					document.getElementById('stand1').style.display = 'none';
				}	
				if (eval("document.frm.view_c[1].checked")) {
					document.getElementById('gubun1').style.display = '';
					document.getElementById('grade1').style.display = 'none';
					document.getElementById('group1').style.display = 'none';
					document.getElementById('name1').style.display = 'none';
					document.getElementById('code1').style.display = 'none';
					document.getElementById('stand1').style.display = 'none';
				}	
				if (eval("document.frm.view_c[2].checked")) {
					document.getElementById('gubun1').style.display = 'none';
					document.getElementById('grade1').style.display = 'none';
					document.getElementById('group1').style.display = 'none';
					document.getElementById('name1').style.display = '';
					document.getElementById('code1').style.display = 'none';
					document.getElementById('stand1').style.display = 'none';
				}	
				if (eval("document.frm.view_c[3].checked")) {
					document.getElementById('gubun1').style.display = 'none';
					document.getElementById('grade1').style.display = 'none';
					document.getElementById('group1').style.display = 'none';
					document.getElementById('name1').style.display = 'none';
					document.getElementById('code1').style.display = 'none';
					document.getElementById('stand1').style.display = '';
				}	
				if (eval("document.frm.view_c[4].checked")) {
					document.getElementById('gubun1').style.display = 'none';
					document.getElementById('grade1').style.display = '';
					document.getElementById('group1').style.display = 'none';
					document.getElementById('name1').style.display = 'none';
					document.getElementById('code1').style.display = 'none';
					document.getElementById('stand1').style.display = 'none';
				}	
				if (eval("document.frm.view_c[5].checked")) {
					document.getElementById('gubun1').style.display = 'none';
					document.getElementById('grade1').style.display = 'none';
					document.getElementById('group1').style.display = '';
					document.getElementById('name1').style.display = 'none';
					document.getElementById('code1').style.display = 'none';
					document.getElementById('stand1').style.display = 'none';
				}	
			}
		</script>

	</head>
	<body onLoad="condi_view()">
		<div id="wrap">			
			<!--#include virtual = "/include/meterials_control_header01.asp" -->
            <!--#include virtual = "/include/meterials_stock_jaego_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="met_goods_jaego_list.asp?ck_sw=<%="n"%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt> 품목 검색</dt>
                        <dd>
                            <p>
                                <label>
								<input type="radio" name="view_c" value="code" <% if view_c = "code" then %>checked<% end if %> style="width:25px" onClick="condi_view()">
                                품목코드
                                <input type="radio" name="view_c" value="gubun" <% if view_c = "gubun" then %>checked<% end if %> style="width:25px" onClick="condi_view()">
                                품목구분
                                <input type="radio" name="view_c" value="name" <% if view_c = "name" then %>checked<% end if %> style="width:25px" onClick="condi_view()">
                                품목명
                                <input type="radio" name="view_c" value="stand" <% if view_c = "stand" then %>checked<% end if %> style="width:25px" onClick="condi_view()">
                                규격
                                <input type="radio" name="view_c" value="grade" <% if view_c = "grade" then %>checked<% end if %> style="width:25px" onClick="condi_view()">
                                상태
                                <input type="radio" name="view_c" value="group" <% if view_c = "group" then %>checked<% end if %> style="width:25px" onClick="condi_view()">
                                품목분류
								</label>
                              <%
								Sql="select * from met_etc_code where etc_type = '04' order by etc_code asc"
					            Rs_etc.Open Sql, Dbconn, 1
							  %>
                                <label id="gubun1">
                                <strong>품목구분</strong>
                                <select name="field_gubun" id="field_gubun" type="text" style="width:100px">
                                    <option value="">선택</option>
                			  <% 
								do until Rs_etc.eof 
			  				  %>
                					<option value='<%=rs_etc("etc_name")%>' <%If field_gubun = rs_etc("etc_name") then %>selected<% end if %>><%=rs_etc("etc_name")%></option>
                			  <%
									Rs_etc.movenext()  
								loop 
								Rs_etc.Close()
							  %>
            					</select>
                                 </label>
                                 <label id="grade1">
								 <strong>상태</strong>
                                 <select name="field_grade" id="field_grade" style="width:100px">
                              		<option value="">선택</option>
								    <option value="신품" <%If field_grade = "신품" then %>selected<% end if %>>신품</option>
								    <option value="중고" <%If field_grade = "중고" then %>selected<% end if %>>중고</option>
								    <option value="리퍼" <%If field_grade = "리퍼" then %>selected<% end if %>>리퍼</option>
                                 </select>
								 </label>
                                 <label id="group1">
								 <strong>품목분류</strong>
                                 <select name="field_group" id="field_group" style="width:100px">
                              		<option value="">선택</option>
								    <option value="자산" <%If field_group = "자산" then %>selected<% end if %>>자산</option>
								    <option value="소모성" <%If field_group = "소모성" then %>selected<% end if %>>소모성</option>
                                 </select>
								 </label>
								 <label id="name1">
								 <strong>품목명</strong>
                                	<input name="field_name" type="text" value="<%=field_name%>" style="width:300px" id="field_view">
								 </label>
								 <label id="code1">
								 <strong>품목코드</strong>
                                	<input name="field_code" type="text" value="<%=field_code%>" style="width:100px" id="field_view">
								 </label>
                                 <label id="stand1">
								 <strong>규격</strong>
                                	<input name="field_stand" type="text" value="<%=field_stand%>" style="width:200px" id="field_view">
								 </label>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser1.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>                
                
				<div class="gView">
				  <table cellpadding="0" cellspacing="0" class="tableList">
				    <colgroup>
				      <col width="7%" >
				      <col width="8%" >
				      <col width="18%" >
				      <col width="16%" >
                      
                      <col width="10%" >
				      <col width="10%" >
                      <col width="*" >

                      <col width="4%" >
                      <col width="6%" >
                      <col width="4%" >
			        </colgroup>
				    <thead>
				      <tr>
				        <th class="first" scope="col">품목코드</th>
                        <th scope="col">품목구분</th>
                        <th scope="col">품목명</th>
                        <th scope="col">규격</th>
                        
                        <th scope="col">모델</th>
                        <th scope="col">Part No.</th>
                        <th scope="col">상세설명</th>
                        
                        <th scope="col">상태</th>
                        <th scope="col">현재고</th>
                        <th scope="col">상세</th>
			          </tr>
			        </thead>
				    <tbody>
                      <%
						do until rs.eof 
						   goods_code = rs("goods_code")
						   
						   if isnull(rs("goods_comment")) then 
						           goods_comment = ""
						      else
						           goods_comment = rs("goods_comment")
						   end if
						   task_memo = replace(goods_comment,chr(34),chr(39))
						   view_memo = task_memo
						   if len(task_memo) > 26 then
								view_memo = mid(task_memo,1,26) + ".."
						   end if
						   
						   h_jaego_cnt = 0
						   sql="select * from met_stock_gmaster where stock_goods_code = '"&goods_code&"'"
	                       Rs_jae.Open Sql, Dbconn, 1
						   
                           do until Rs_jae.eof
                              h_jaego_cnt = h_jaego_cnt + Rs_jae("stock_JJ_qty")
							  
							  Rs_jae.movenext()
                           loop
                           Rs_jae.close()
					  %>
				      <tr>
				        <td class="first"><%=rs("goods_code")%>&nbsp;</td>
                        <td><%=rs("goods_gubun")%>&nbsp;</td>
                        <td class="left"><%=rs("goods_name")%>&nbsp;</td>
                        <td class="left"><%=rs("goods_standard")%>&nbsp;</td>
                        
                        <td class="left"><%=rs("goods_model")%>&nbsp;</td>
                        <td class="left"><%=rs("part_number")%>&nbsp;</td>
                        <td class="left"><p style="cursor:pointer"><span title="<%=task_memo%>"><%=view_memo%></span></p></td>

                        <td><%=rs("goods_grade")%>&nbsp;</td>
                        <td class="right" style="text-align:right"><%=formatnumber(h_jaego_cnt,0)%>&nbsp;</td>
                        <td>
                        <a href="#" onClick="pop_Window('met_goods_jaego_stock_detail.asp?stock_goods_code=<%=rs("goods_code")%>&u_type=<%=""%>','met_goods_jaego_stock_detail_pop','scrollbars=yes,width=930,height=650')">조회</a>
                        </td>
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
                    <a href="met_goods_jaego_list_excel.asp?goods_type=<%=goods_type%>&view_c=<%=view_c%>&field_check=<%=field_check%>&field_gubun=<%=field_gubun%>&field_grade=<%=field_grade%>&field_group=<%=field_group%>&field_name=<%=field_name%>&field_code=<%=field_code%>&field_stand=<%=field_stand%>" class="btnType04">엑셀다운로드</a>
					</div>                  
                  	</td>
				    <td>
                    <div id="paging">
                        <a href="met_goods_jaego_list.asp?page=<%=first_page%>&goods_type=<%=goods_type%>&view_c=<%=view_c%>&field_check=<%=field_check%>&field_gubun=<%=field_gubun%>&field_grade=<%=field_grade%>&field_group=<%=field_group%>&field_name=<%=field_name%>&field_code=<%=field_code%>&field_stand=<%=field_stand%>&ck_sw=<%="y"%>">[처음]</a>
                  	<% if intstart > 1 then %>
                        <a href="met_goods_jaego_list.asp?page=<%=intstart -1%>&goods_type=<%=goods_type%>&view_c=<%=view_c%>&field_check=<%=field_check%>&field_gubun=<%=field_gubun%>&field_grade=<%=field_grade%>&field_group=<%=field_group%>&field_name=<%=field_name%>&field_code=<%=field_code%>&field_stand=<%=field_stand%>&ck_sw=<%="y"%>">[이전]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
   	        <% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="met_goods_jaego_list.asp?page=<%=i%>&goods_type=<%=goods_type%>&view_c=<%=view_c%>&field_check=<%=field_check%>&field_gubun=<%=field_gubun%>&field_grade=<%=field_grade%>&field_group=<%=field_group%>&field_name=<%=field_name%>&field_code=<%=field_code%>&field_stand=<%=field_stand%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
   	        <% if 	intend < total_page then %>
                        <a href="met_goods_jaego_list.asp?page=<%=intend+1%>&goods_type=<%=goods_type%>&view_c=<%=view_c%>&field_check=<%=field_check%>&field_gubun=<%=field_gubun%>&field_grade=<%=field_grade%>&field_group=<%=field_group%>&field_name=<%=field_name%>&field_code=<%=field_code%>&field_stand=<%=field_stand%>&ck_sw=<%="y"%>">[다음]</a> <a href="met_goods_jaego_list.asp?page=<%=total_page%>&goods_type=<%=goods_type%>&view_c=<%=view_c%>&field_check=<%=field_check%>&field_gubun=<%=field_gubun%>&field_grade=<%=field_grade%>&field_group=<%=field_group%>&field_name=<%=field_name%>&field_code=<%=field_code%>&field_stand=<%=field_stand%>&ck_sw=<%="y"%>">[마지막]</a>
                        <%	else %>
                        [다음]&nbsp;[마지막]
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
        <input type="hidden" name="field_check" value="<%=field_view%>" ID="field_check">
	</body>
</html>

