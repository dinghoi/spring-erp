<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

	dim abc,filenm
	Set abc = Server.CreateObject("ABCUpload4.XForm")
	abc.AbsolutePath = True
	abc.Overwrite = true
	abc.MaxUploadSize = 1024*1024*50

	tot_cnt = 0
	tot_err = 0
	tot_dept = 0
	tot_cust = 0
	tot_ddd = 0
	tot_tel = 0
	tot_sido = 0
	tot_gugun = 0
	tot_dong = 0
	tot_addr = 0
	tot_ce = 0
	tot_cnt = 0

	as_type = abc("as_type")
	company = abc("company")
	request_date = abc("request_date")
	end_date = abc("end_date")
	file_type = abc("file_type")

	if request_date = "" then
		request_date = mid(now(),1,10)
		ck_sw = "y"
	  else
	  	ck_sw = "n"
	end if
	
	Set DbConn = Server.CreateObject("ADODB.Connection")
	set cn = Server.CreateObject("ADODB.Connection")
	set rs = Server.CreateObject("ADODB.Recordset")	
	Set Rs_etc = Server.CreateObject("ADODB.Recordset")
	Set rs_com = Server.CreateObject("ADODB.Recordset")
	DbConn.Open dbconnect

	If ck_sw = "n" Then
		request_month = mid(cstr(request_date),1,4) + mid(cstr(request_date),6,2)
		sql="select max(paper_no) as max_seq from large_acpt"
		set rs=dbconn.execute(sql)
		if	isnull(rs("max_seq"))  then
			paper_no = request_month + "01"
		  else
			if mid(rs("max_seq"),1,6) = request_month then
				paper_no = cstr(int(rs("max_seq")) + 1)
			  elseif mid(rs("max_seq"),1,6) > request_month then
			  	paper_no = "error"
			  else
				paper_no = request_month + "01"
			end if
		end if
		rs.close()
	
		Set filenm = abc("att_file")(1)
		
		path = Server.MapPath ("/large_file")
		filename = filenm.safeFileName
		fileType = mid(filename,inStrRev(filename,".")+1)
		file_name = company + "_" + as_type + "_" + paper_no
		
'		save_path = path & "\" & filename
		save_path = path & "\" & file_name&"."&fileType

		if fileType = "xls" or fileType = "xlk" then
			file_type = "Y"
			filenm.save save_path
		
		
			objFile = save_path
	'		objFile = Request.form("att_file")
	'		objFile = SERVER.MapPath("att_file")
	'		objFile = SERVER.MapPath(".") & "\kwon_upload\excel_data.xls"
	'		response.write(objFile)
			
			cn.open "Driver={Microsoft Excel Driver (*.xls)};ReadOnly=1;DBQ=" & objFile & ";"
			rs.Open "select * from [1:10000]",cn,"0"
				
			rowcount=-1
			xgr = rs.getrows
			rowcount = ubound(xgr,2)
			fldcount = rs.fields.count
			tot_cnt = rowcount + 1
		  else
			objFile = "none"
			rowcount=-1
			file_type = "N"
		end if		  
	  else
		objFile = "none"
		rowcount=-1
	end if
	title_line = "대량자료 업로드"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>A/S 관리 시스템</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript" src="/java/js_window.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "1 1";
			}
			$(function() {    $( "#datepicker" ).datepicker();
											$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
											$( "#datepicker" ).datepicker("setDate", "<%=request_date%>" );
			});	  
			$(function() {    $( "#datepicker1" ).datepicker();
												$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker1" ).datepicker("setDate", "<%=end_date%>" );
			});	  
			function frmcheck () {
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {
				if (document.frm.as_type.value == "") {
					alert ("업로드유형을 선택하세요");
					return false;
				}	
				if (document.frm.company.value == "") {
					alert ("회사를 선택하세요");
					return false;
				}	
				if (document.frm.request_date.value == "") {
					alert ("개시일을 입력하세요");
					return false;
				}	
				if (document.frm.end_date.value == "") {
					alert ("마감일을 입력하세요");
					return false;
				}	
				if (document.frm.att_file.value == "") {
					alert ("업로드 엑셀 파일을 선택하세요");
					return false;
				}	
				return true;
			}
			function frm1check () {
				if (chkfrm1()) {
					document.frm1.submit ();
				}
			}
			
			function chkfrm1() {
				a=confirm('DB에 업로드 하시겠습니까?');
				if (a==true) {
					return true;
				}
				return false;
			}
		</script>

	</head>
	<body>
		<div id="wrap">			
			<!--#include virtual = "/include/header.asp" -->
			<!--#include virtual = "/include/large_sub_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="large_data_up.asp" method="post" name="frm" enctype="multipart/form-data">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>업로드내용</dt>
                        <dd>
                            <p>
								<label>
								<strong>업로드유형 : </strong>
                                <select name="as_type" id="as_type" style="width:100px">
	                				<option value="">선택</option>
                                    <option value="신규설치" <%If as_type = "신규설치" then %>selected<% end if %>>신규설치</option>
                                    <option value="신규설치공사" <%If as_type = "신규설치공사" then %>selected<% end if %>>신규설치공사</option>
                                    <option value="이전설치" <%If as_type = "이전설치" then %>selected<% end if %>>이전설치</option>
                                    <option value="이전설치공사" <%If as_type = "이전설치공사" then %>selected<% end if %>>이전설치공사</option>
                                    <option value="랜공사" <%If as_type = "랜공사" then %>selected<% end if %>>랜공사</option>
                                    <option value="이전랜공사" <%If as_type = "이전랜공사" then %>selected<% end if %>>이전랜공사</option>
                                    <option value="장비회수" <%If as_type = "장비회수" then %>selected<% end if %>>장비회수</option>
                                    <option value="예방점검" <%If as_type = "예방점검" then %>selected<% end if %>>예방점검</option>
              					</select>
								</label>
								<label>
								<strong>회사 : </strong>
								<%
                                sql="select * from trade where use_sw = 'Y' and mg_group = '" + mg_group + "' order by trade_name asc"
                                rs_com.Open Sql, Dbconn, 1
                                %>
                                <select name="company" id="company" style="width:150px">
                                  <option value="">선택</option>
                                <% 
                                do until rs_com.eof 
                                %>
          							<option value='<%=rs_com("trade_name")%>' <%If rs_com("trade_name") = company  then %>selected<% end if %>><%=rs_com("trade_name")%></option>
                                <%
                                	rs_com.movenext()  
                                loop 
                                rs_com.Close()
                                %>
                                </select>
								</label>
								<label>
								<strong>개시일 : </strong>
                                	<input name="request_date" type="text" value="<%=request_date%>" style="width:70px" id="datepicker">
								</label>
								<label>
								<strong>마감일 : </strong>
                                	<input name="end_date" type="text" value="<%=end_date%>" style="width:70px" id="datepicker1">
								</label>
								<label>
								<strong>문서번호 : </strong>
									<input name="paper_no" type="text" style="width:80px;" value="<%=paper_no%>" readonly="true" >
								</label>							
                                <label>
								<br>
                                <label>
								<strong>업로드파일 : </strong>
								<input name="att_file" type="file" id="att_file" size="113" value="<%=att_file%>" style="text-align:left"> 
								</label>
            <input name="file_type" type="hidden" id="file_type" value="<%=file_type%>">
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="3%" >
							<col width="10%" >
							<col width="6%" >
							<col width="4%" >
							<col width="4%" >
							<col width="4%" >
							<col width="4%" >
							<col width="4%" >
							<col width="4%" >
							<col width="5%" >
							<col width="10%" >
							<col width="8%" >
							<col width="*" >
							<col width="5%" >
							<col width="5%" >
							<col width="3%" >
						</colgroup>
						<thead>
							<tr>
								<th rowspan="2" class="first" scope="col">순번</th>
								<th rowspan="2" scope="col">부서</th>
								<th rowspan="2" scope="col">고객명</th>
								<th colspan="3" scope="col" style=" border-bottom:1px solid #e3e3e3;">전화번호1</th>
								<th colspan="3" scope="col" style=" border-bottom:1px solid #e3e3e3;">전화번호2</th>
								<th rowspan="2" scope="col">시도</th>
								<th rowspan="2" scope="col">구군</th>
								<th rowspan="2" scope="col">동/읍</th>
								<th rowspan="2" scope="col">번지</th>
								<th rowspan="2" scope="col">사번</th>
								<th rowspan="2" scope="col">CE명</th>
								<th rowspan="2" scope="col">건수</th>
							</tr>
							<tr>
							  <th scope="col" style=" border-left:1px solid #e3e3e3;">DDD</th>
							  <th scope="col">국</th>
							  <th scope="col">번호</th>
							  <th scope="col">DDD</th>
							  <th scope="col">국</th>
							  <th scope="col">번호</th>
				          </tr>
						</thead>
						<tbody>
						<%
						  if rowcount > -1 then
							for i=0 to rowcount
					' 부서명
							dept_sw = "Y"
							if xgr(0,i) = "" then
								dept_sw = "N"
								tot_dept = tot_dept + 1
								tot_err = tot_err + 1
							end if
					' 고객명
							cust_sw = "Y"
							if xgr(1,i) = "" then
								cust_sw = "N"
								tot_cust = tot_cust + 1
								tot_err = tot_err + 1
							end if
					' DDD
							if isnull(xgr(2,i)) then
								ddd_sw = "N"
							  else
								ddd_sw = "Y"
								sql_etc = "select * from etc_code where etc_type = '71' and etc_name = '" + xgr(2,i) +"'"
								set rs_etc=dbconn.execute(sql_etc)				
								if rs_etc.eof then
									tot_ddd = tot_ddd + 1
									tot_err = tot_err + 1
									ddd_sw = "N"
								end if
							end if
					' TEL (전화국)
							tel_sw = "Y"
							if xgr(3,i) < "100" or isnull(xgr(3,i)) then
								tel_sw = "N"
								tot_tel = tot_tel + 1
								tot_err = tot_err + 1
							end if
					' TEL NO
							tel_no_sw = "Y"
							if xgr(4,i) < "0001" or isnull(xgr(4,i)) then
								tel_no_sw = "N"
								tot_tel = tot_tel + 1
								tot_err = tot_err + 1
							end if
					' DDD
'							if isnull(xgr(5,i)) then
'								ddd1_sw = "N"
'							  else
'								ddd1_sw = "Y"
'								sql_etc = "select * from etc_code where etc_type = '71' and etc_name = '" + xgr(5,i) +"'"
'								set rs_etc=dbconn.execute(sql_etc)				
'								if rs_etc.eof then
'									tot_ddd = tot_ddd + 1
'									tot_err = tot_err + 1
'									ddd_sw = "N"
'								end if
'							end if
					' TEL (전화국)
'							tel1_sw = "Y"
'							if xgr(6,i) < "100" or isnull(xgr(6,i)) then
'								tel1_sw = "N"
'								tot_tel = tot_tel + 1
'								tot_err = tot_err + 1
'							end if
					' TEL NO
'							tel_no1_sw = "Y"
'							if xgr(7,i) < "0001" or isnull(xgr(6,i)) then
'								tel_no1_sw = "N"
'								tot_tel = tot_tel + 1
'								tot_err = tot_err + 1
'							end if
					' 시도
							sido_sw = "Y"
							sql_etc = "select * from etc_code where etc_type = '81' and etc_name = '" + xgr(8,i) +"'"
							set rs_etc=dbconn.execute(sql_etc)				
							if rs_etc.eof then
								tot_sido = tot_sido + 1
								tot_err = tot_err + 1
								sido_sw = "N"
							end if
					' 구군
							gugun_sw = "Y"
							sql_etc = "select * from ce_area where sido = '" + xgr(8,i) +"' and gugun = '" + xgr(9,i) + "'"
							set rs_etc=dbconn.execute(sql_etc)				
							if rs_etc.eof then
								tot_gugun = tot_gugun + 1
								tot_err = tot_err + 1
								gugun_sw = "N"
								mg_ce_id = ""
							  else
								mg_ce_id = rs_etc("mg_ce_id")	  
							end if
					' 동/읍
							dong_sw = "Y"
							if xgr(10,i) = "" or isnull(xgr(10,i)) then
								dong_sw = "N"
								tot_dong = tot_dong + 1
								tot_err = tot_err + 1
							end if
					' 번지
							addr_sw = "Y"
							if xgr(11,i) = "" or isnull(xgr(11,i)) then
								addr_sw = "N"
								tot_addr = tot_addr + 1
								tot_err = tot_err + 1
							end if
					' CE
							ce_sw = "Y"
							if (xgr(12,i) = "" or isnull(xgr(12,i))) and (xgr(13,i) = "" or isnull(xgr(12,i))) then
								sql_etc = "select * from memb where user_id = '" + mg_ce_id + "'"
								set rs_etc=dbconn.execute(sql_etc)				
								if rs_etc.eof then
									tot_ce = tot_ce + 1
									tot_err = tot_err + 1
									ce_sw = "N"
									mg_ce = "미등록"
								  else
									mg_ce = rs_etc("user_name")
								end if
							end if

							if xgr(12,i) <> "" then
								sql_etc = "select * from memb where user_id = '" + cstr(xgr(12,i)) + "'"
								set rs_etc=dbconn.execute(sql_etc)				
								if rs_etc.eof then
									tot_ce = tot_ce + 1
									tot_err = tot_err + 1
									ce_sw = "N"
									mg_ce_id = xgr(12,i)
									mg_ce = "미등록"
								  else
									mg_ce_id = rs_etc("user_id")
									mg_ce = rs_etc("user_name")
								end if
							end if
							
							if xgr(13,i) <> "" then
								sql_etc = "select * from memb where user_name = '" + xgr(13,i) + "'"
								set rs_etc=dbconn.execute(sql_etc)				
								if rs_etc.eof then
									tot_ce = tot_ce + 1
									tot_err = tot_err + 1
									ce_sw = "N"
									mg_ce_id = "error"
									mg_ce = xgr(13,i)
								  else
									mg_ce_id = rs_etc("user_id")
									mg_ce = rs_etc("user_name")
								end if
							end if
					' 수량
							cnt_sw = "Y"
							if xgr(14,i) = "" or xgr(14,i) = "0" or xgr(14,i) > "999" then
								cnt_sw = "N"
								tot_cnt = tot_cnt + 1
								tot_err = tot_err + 1
							end if
							%>
							<tr>
								<td class="first"><%=i+1%></td>
								<% if dept_sw = "Y" then %>
									<td><%=xgr(0,i)%></td>
                                <%   else	%>
									<td bgcolor="#FFCCFF"><%=xgr(0,i)%></td>
								<% end if 	%>                                
								<% if cust_sw = "Y" then %>
									<td><%=xgr(1,i)%></td>
                                <%   else	%>
									<td bgcolor="#FFCCFF"><%=xgr(1,i)%></td>
								<% end if 	%>                                
								<% if ddd_sw = "Y" then %>
									<td><%=xgr(2,i)%></td>
                                <%   else	%>
									<td bgcolor="#FFCCFF"><%=xgr(2,i)%></td>
								<% end if 	%>                                
								<% if tel_sw = "Y" then %>
									<td><%=xgr(3,i)%></td>
                                <%   else	%>
									<td bgcolor="#FFCCFF"><%=xgr(3,i)%></td>
								<% end if 	%>                                
								<% if tel_no_sw = "Y" then %>
									<td><%=xgr(4,i)%></td>
                                <%   else	%>
									<td bgcolor="#FFCCFF"><%=xgr(4,i)%></td>
								<% end if 	%>                                
								<% if ddd1_sw = "Y" then %>
									<td><%=xgr(5,i)%></td>
                                <%   else	%>
									<td bgcolor="#FFCCFF"><%=xgr(5,i)%></td>
								<% end if 	%>                                
								<% if tel1_sw = "Y" then %>
									<td><%=xgr(6,i)%></td>
                                <%   else	%>
									<td bgcolor="#FFCCFF"><%=xgr(6,i)%></td>
								<% end if 	%>                                
								<% if tel_no1_sw = "Y" then %>
									<td><%=xgr(7,i)%></td>
                                <%   else	%>
									<td bgcolor="#FFCCFF"><%=xgr(7,i)%></td>
								<% end if 	%>                                
								<% if sido_sw = "Y" then %>
									<td><%=xgr(8,i)%></td>
                                <%   else	%>
									<td bgcolor="#FFCCFF"><%=xgr(8,i)%></td>
								<% end if 	%>                                
								<% if gugun_sw = "Y" then %>
									<td><%=xgr(9,i)%></td>
                                <%   else	%>
									<td bgcolor="#FFCCFF"><%=xgr(9,i)%></td>
								<% end if 	%>                                
								<% if dong_sw = "Y" then %>
									<td><%=xgr(10,i)%></td>
                                <%   else	%>
									<td bgcolor="#FFCCFF"><%=xgr(10,i)%></td>
								<% end if 	%>                                
								<% if addr_sw = "Y" then %>
									<td class="left"><%=xgr(11,i)%></td>
                                <%   else	%>
									<td bgcolor="#FFCCFF" class="left"><%=xgr(11,i)%></td>
								<% end if 	%>                                
								<% if ce_sw = "Y" then %>
									<td><%=mg_ce_id%></td>
                                <%   else	%>
									<td bgcolor="#FFCCFF"><%=mg_ce_id%></td>
								<% end if 	%>                                
								<% if ce_sw = "Y" then %>
									<td><%=mg_ce%></td>
                                <%   else	%>
									<td bgcolor="#FFCCFF"><%=mg_ce%></td>
								<% end if 	%>                                
								<% if cnt_sw = "Y" then %>
									<td class="left"><%=xgr(14,i)%></td>
                                <%   else	%>
									<td bgcolor="#FFCCFF" class="left"><%=xgr(14,i)%></td>
								<% end if 	%>                                
							</tr>
						<%
							next
						end if
						%>
						</tbody>
					</table>
					<h3 class="stit">* Check 결과</h3>
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="*" >
							<col width="9%" >
							<col width="9%" >
							<col width="9%" >
							<col width="9%" >
							<col width="9%" >
							<col width="9%" >
							<col width="9%" >
							<col width="9%" >
							<col width="9%" >
							<col width="9%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">총건수</th>
								<th scope="col">총Error</th>
								<th scope="col">고객명</th>
								<th scope="col">DDD</th>
								<th scope="col">전화</th>
								<th scope="col">번호</th>
								<th scope="col">시도</th>
								<th scope="col">구군</th>
								<th scope="col">동/읍</th>
								<th scope="col">번지</th>
								<th scope="col">CE</th>
							</tr>
						</thead>
						<tbody>
							<tr>
								<td class="first"><%=tot_cnt%></td>
								<td><%=tot_err%></td>
								<td><%=tot_dept%></td>
								<td><%=tot_cust%></td>
								<td><%=tot_ddd%></td>
								<td><%=tot_tel%></td>
								<td><%=tot_sido%></td>
								<td><%=tot_gugun%></td>
								<td><%=tot_dong%></td>
								<td><%=tot_addr%></td>
								<td><%=tot_ce%></td>
							</tr>
						</tbody>
					</table>
				</div>
				</form>
				<% if tot_err = 0 and tot_cnt <> 0 then %>
				<form action="large_data_up_ok.asp" method="post" name="frm1">
					<br>
                    <div align=center>
                        <span class="btnType01"><input type="button" value="DB저장" onclick="javascript:frm1check();"NAME="Button1"></span>
                    </div>
                    <input name="objFile" type="hidden" id="objFile" value="<%=objFile%>">
                    <input name="as_type" type="hidden" id="as_type" value="<%=as_type%>">
                    <input name="request_date" type="hidden" id="request_date" value="<%=request_date%>">
                    <input name="end_date" type="hidden" id="end_date" value="<%=end_date%>">
                    <input name="paper_no" type="hidden" id="paper_no" value="<%=paper_no%>">
                    <input name="company" type="hidden" id="company" value="<%=company%>">
					<br>
				</form>
				<% end if %>
		</div>				
	</div>        				
	</body>
</html>

