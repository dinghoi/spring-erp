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

be_pg = "met_import_serial_list.asp"
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
Set Rs_seri = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")

dbconn.open DbConnect

order_Sql = " ORDER BY goods_gubun,in_date,goods_code,serial_no,serial_seq ASC"

If view_c = "" Then
	ck_sw = "n"
	field_check = "total"
	view_c = "gubun"
End If

If field_check = "total" Then
       owner_sql = "select count(*) FROM met_goods_serial"
	   field_check = ""
   else
       if view_c = "gubun" Then
              owner_sql = "select count(*) FROM met_goods_serial  where goods_gubun like '%" + field_gubun + "%'"
       end if
	   if view_c = "pono" Then
              owner_sql = "select count(*) FROM met_goods_serial  where po_number like '%" + field_grade + "%'"
       end if
	   if view_c = "serial" Then
              owner_sql = "select count(*) FROM met_goods_serial  where serial_no like '%" + field_group + "%'"
       end if
	   if view_c = "name" Then
              owner_sql = "select count(*) FROM met_goods_serial  where goods_name like '%" + field_name + "%'"
       end if
	   if view_c = "code" Then
              owner_sql = "select count(*) FROM met_goods_serial  where goods_code like '%" + field_code + "%'"
       end if
	   if view_c = "partno" Then
              owner_sql = "select count(*) FROM met_goods_serial  where part_number like '%" + field_stand + "%'"
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
       owner_sql = "select * FROM met_goods_serial"
	   field_check = ""
   else
       if view_c = "gubun" Then
              owner_sql = "select * FROM met_goods_serial  where goods_gubun like '%" + field_gubun + "%'"
       end if
	   if view_c = "pono" Then
              owner_sql = "select * FROM met_goods_serial  where po_number like '%" + field_grade + "%'"
       end if
	   if view_c = "serial" Then
              owner_sql = "select * FROM met_goods_serial  where serial_no like '%" + field_group + "%'"
       end if
	   if view_c = "name" Then
              owner_sql = "select * FROM met_goods_serial  where goods_name like '%" + field_name + "%'"
       end if
	   if view_c = "code" Then
              owner_sql = "select * FROM met_goods_serial  where goods_code like '%" + field_code + "%'"
       end if
	   if view_c = "partno" Then
              owner_sql = "select * FROM met_goods_serial  where part_number like '%" + field_stand + "%'"
       end if
End If

sql = owner_sql + order_sql + " limit "& stpage & "," &pgsize 
Rs.Open Sql, Dbconn, 1
'response.write(sql)

title_line =  " Serial 관리대장 "

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
            <!--#include virtual = "/include/meterials_stock_nw_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="met_import_serial_list.asp?ck_sw=<%="n"%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt> 검색조건</dt>
                        <dd>
                            <p>
                                <label>
								<input type="radio" name="view_c" value="code" <% if view_c = "code" then %>checked<% end if %> style="width:25px" onClick="condi_view()">
                                품목코드
                                <input type="radio" name="view_c" value="gubun" <% if view_c = "gubun" then %>checked<% end if %> style="width:25px" onClick="condi_view()">
                                품목구분
                                <input type="radio" name="view_c" value="name" <% if view_c = "name" then %>checked<% end if %> style="width:25px" onClick="condi_view()">
                                품목명
                                <input type="radio" name="view_c" value="partno" <% if view_c = "partno" then %>checked<% end if %> style="width:25px" onClick="condi_view()">
                                Part_No
                                <input type="radio" name="view_c" value="pono" <% if view_c = "pono" then %>checked<% end if %> style="width:25px" onClick="condi_view()">
                                PO_No
                                <input type="radio" name="view_c" value="serial" <% if view_c = "serial" then %>checked<% end if %> style="width:25px" onClick="condi_view()">
                                Serial NO
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
								 <strong>PO_No</strong>
                                 <input name="field_grade" type="text" value="<%=field_grade%>" style="width:200px" id="field_view">
								 </label>
                                 <label id="group1">
								 <strong>Serial No</strong>
                                 <input name="field_group" type="text" value="<%=field_group%>" style="width:200px" id="field_view">
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
								 <strong>Part_No</strong>
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
				      <col width="6%" >
				      <col width="10%" >
				      <col width="7%" >
				      <col width="18%" >
                      <col width="10%" >
                      <col width="6%" >
                      <col width="14%" >
                      <col width="6%" >
                      <col width="16%" >
                      <col width="*" >
			        </colgroup>
				    <thead>
				      <tr>
				        <th class="first" scope="col">품목구분</th>
                        <th scope="col">Serial No</th>
                        <th scope="col">품목코드</th>
                        <th scope="col">품목명</th>
                        <th scope="col">Part_No./규격</th>
                        <th scope="col">입고일자</th>
                        <th scope="col">PO_No.</th>
                        <th scope="col">출고일자</th>
                        <th scope="col">납품처</th>
                        <th scope="col">비고</th>
			          </tr>
			        </thead>
				    <tbody>
                      <%
						do until rs.eof 

					  %>
				      <tr>
				        <td class="first"><%=rs("goods_gubun")%>&nbsp;</td>
                        <td class="left"><%=rs("serial_no")%>&nbsp;<%=rs("serial_seq")%></td>
                        <td><%=rs("goods_code")%>&nbsp;</td>
                        <td class="left"><%=rs("goods_name")%>&nbsp;</td>
                        <td class="left"><%=rs("part_number")%>&nbsp;</td>
                        <td><%=rs("in_date")%>&nbsp;</td>
                        <td class="left"><%=rs("po_number")%>&nbsp;</td>
                        <td><%=rs("chulgo_date")%>&nbsp;</td>
                        <td class="left"><%=rs("chulgo_trade")%>&nbsp;<%=rs("chulgo_trade_dept")%></td>
                        <td class="left"><%=rs("chulgo_bigo")%>&nbsp;</td>
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
                    <a href="met_import_serial_excel.asp?goods_type=<%=goods_type%>&view_c=<%=view_c%>&field_check=<%=field_check%>&field_gubun=<%=field_gubun%>&field_grade=<%=field_grade%>&field_group=<%=field_group%>&field_name=<%=field_name%>&field_code=<%=field_code%>&field_stand=<%=field_stand%>" class="btnType04">엑셀다운로드</a>
					</div>                  
                  	</td>
				    <td>
                    <div id="paging">
                        <a href="met_import_serial_list.asp?page=<%=first_page%>&goods_type=<%=goods_type%>&view_c=<%=view_c%>&field_check=<%=field_check%>&field_gubun=<%=field_gubun%>&field_grade=<%=field_grade%>&field_group=<%=field_group%>&field_name=<%=field_name%>&field_code=<%=field_code%>&field_stand=<%=field_stand%>&ck_sw=<%="y"%>">[처음]</a>
                  	<% if intstart > 1 then %>
                        <a href="met_import_serial_list.asp?page=<%=intstart -1%>&goods_type=<%=goods_type%>&view_c=<%=view_c%>&field_check=<%=field_check%>&field_gubun=<%=field_gubun%>&field_grade=<%=field_grade%>&field_group=<%=field_group%>&field_name=<%=field_name%>&field_code=<%=field_code%>&field_stand=<%=field_stand%>&ck_sw=<%="y"%>">[이전]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
   	        <% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="met_import_serial_list.asp?page=<%=i%>&goods_type=<%=goods_type%>&view_c=<%=view_c%>&field_check=<%=field_check%>&field_gubun=<%=field_gubun%>&field_grade=<%=field_grade%>&field_group=<%=field_group%>&field_name=<%=field_name%>&field_code=<%=field_code%>&field_stand=<%=field_stand%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
   	        <% if 	intend < total_page then %>
                        <a href="met_import_serial_list.asp?page=<%=intend+1%>&goods_type=<%=goods_type%>&view_c=<%=view_c%>&field_check=<%=field_check%>&field_gubun=<%=field_gubun%>&field_grade=<%=field_grade%>&field_group=<%=field_group%>&field_name=<%=field_name%>&field_code=<%=field_code%>&field_stand=<%=field_stand%>&ck_sw=<%="y"%>">[다음]</a> <a href="met_import_serial_list.asp?page=<%=total_page%>&goods_type=<%=goods_type%>&view_c=<%=view_c%>&field_check=<%=field_check%>&field_gubun=<%=field_gubun%>&field_grade=<%=field_grade%>&field_group=<%=field_group%>&field_name=<%=field_name%>&field_code=<%=field_code%>&field_stand=<%=field_stand%>&ck_sw=<%="y"%>">[마지막]</a>
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

