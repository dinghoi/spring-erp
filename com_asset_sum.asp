<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/srvmg_dbcon.asp" -->
<!--#include virtual="/include/srvmg_user.asp" -->
<%

company = request("company")
field_check = request("field_check")
field_view = request("field_view")

condi_01 = "전체"
if field_check = "high_org" then
	condi_01 = "관리조직 = "
end if	
if field_check = "org_first" then
	condi_01 = "법인명 = "
end if	
if field_check = "org_second" then
	condi_01 = "지사명 = "
end if	
if field_check = "dept_name" then
	condi_01 = "지점명 = "
end if	
if field_check = "sido" then
	condi_01 = "시도 = "
end if	
if field_check = "tel_no = " then
	condi_01 = "전화번호"
end if	

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

etc_code = "75" + cstr(company)
sql = "select * from etc_code where etc_code = '" + etc_code + "'"
set rs_etc = dbconn.execute(sql)

base_sql = "SELECT asset.company, asset.gubun, asset.asset_name, count(*) as asset_sum FROM asset INNER JOIN asset_dept ON (asset.dept_code = asset_dept.dept_code) AND (asset.company = asset_dept.company) where asset.inst_process = 'Y' and asset.company = '" + company + "'"
if field_check <> "total" then
	condi_sql = " and ( asset_dept." + field_check + " like '%" + field_view + "%' ) "
  else
  	condi_sql = " "
end if	

group_sql = " group by asset.company, asset.gubun order by asset.company, asset.gubun"
sql = base_sql + condi_sql + group_sql
rs.Open Sql, Dbconn, 1

title_line = "조건별 자산 현황"

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>A/S 관리 시스템</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">

			function goAction () {
		  		 window.close () ;
			}

			function printWindow(){
        //		viewOff("button");   
                factory.printing.header = ""; //머리말 정의
                factory.printing.footer = ""; //꼬리말 정의
                factory.printing.portrait = false; //출력방향 설정: true - 가로, false - 세로
                factory.printing.leftMargin = 13; //외쪽 여백 설정
                factory.printing.topMargin = 10; //윗쪽 여백 설정
                factory.printing.rightMargin = 13; //오른쯕 여백 설정
                factory.printing.bottomMargin = 15; //바닦 여백 설정
        //		factory.printing.SetMarginMeasure(2); //테두리 여백 사이즈 단위를 인치로 설정
        //		factory.printing.printer = ""; //프린터 할 프린터 이름
        //		factory.printing.paperSize = "A4"; //용지선택
        //		factory.printing.pageSource = "Manusal feed"; //종이 피드 방식
        //		factory.printing.collate = true; //순서대로 출력하기
        //		factory.printing.copies = "1"; //인쇄할 매수
        //		factory.printing.SetPageRange(true,1,1); //true로 설정하고 1,3이면 1에서 3페이지 출력
        //		factory.printing.Printer(true); //출력하기
                factory.printing.Preview(); //윈도우를 통해서 출력
                factory.printing.Print(false); //윈도우를 통해서 출력
            }
        </script>

	</head>

	<style media="print"> 
    .noprint     { display: none }
    </style>

	<body>
    <object id="factory" style="display:none;" viewastext classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814" codebase="/smsx.cab#Version=7.0.0.8">
	</object>
		<div id="container">				
			<div class="gView">
			<h3 class="tit"><%=title_line%></h3>
				<form method="post" name="frm" action="">
					<table cellpadding="0" cellspacing="0" summary="" class="tableView">
						<colgroup>
							<col width="15%" >
							<col width="35%" >
							<col width="15%" >
							<col width="*" >
						</colgroup>
						<tbody>
							<tr>
							  <th>회사</th>
							  <td class="left"><%=rs_etc("etc_name")%></td>
							  <th>조건</th>
							  <td class="left"><%=condi_01%>&nbsp;<%=field_view%></td>
					      	</tr>
						</tbody>
					</table>
					<h3 class="stit">* 자산 현황</h3>
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="20%" >
							<col width="20%" >
							<col width="*" >
							<col width="15%" >
							<col width="15%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">자산구분</th>
								<th scope="col">자산코드</th>
								<th scope="col">자산명</th>
								<th scope="col">소계</th>
								<th scope="col">자산계</th>
							</tr>
						</thead>
						<tbody>
						<%
						do until rs.eof
							etc_code = "79" + cstr(rs("gubun"))
							sql = "select * from etc_code where etc_code = '" + etc_code + "'"
							set rs_gubun = dbconn.execute(sql)
							if rs_gubun.eof or rs_gubun.bof then
								gubun_name = "없음"
							  else
								gubun_name = rs_gubun("etc_name")
							end if
                        %>
							<tr>
								<td class="first"><%=gubun_name%></td>
								<td colspan="3"><table width="100%"  border="1" cellpadding="0" cellspacing="0">
						<%
                                Set rs_d = Server.CreateObject("ADODB.Recordset")
                
                                base_sql = "select asset.company, asset.gubun, asset.code_seq, asset.asset_name, count(*) as asset_cnt FROM asset INNER JOIN asset_dept ON (asset.dept_code = asset_dept.dept_code) AND (asset.company = asset_dept.company) where asset.inst_process = 'Y' and asset.gubun = '" + rs("gubun") + "' and asset.company = '" + company + "'"
                                if field_check <> "total" then
                                    condi_sql = " and ( asset_dept." + field_check + " like '%" + field_view + "%' ) "
                                  else
                                    condi_sql = " "
                                end if	
                                
                                group_sql = " group by asset.company, asset.gubun, asset.code_seq order by asset.company, asset.gubun, asset.code_seq"
                                sql = base_sql + condi_sql + group_sql
                                rs_d.Open Sql, Dbconn, 1
                
                                do until rs_d.eof
                        %>
                                    <tr>
                                      <td width="31%" class="first"><%=rs_d("company")%>-<%=rs_d("gubun")%>-<%=rs_d("code_seq")%></td>
                                      <td width="*"><%=rs_d("asset_name")%></td>
                                      <td width="23%"><%=rs_d("asset_cnt")%></td>
                                    </tr>
                        <%
                                    rs_d.movenext()
                                loop
                                rs_d.close()
                        %>
                                  </table>
                                </td>
								<td><%=rs("asset_sum")%></td>
							</tr>
						<%
                            rs.movenext()  
                        loop
                        rs.Close()
                        %>
						</tbody>
					</table>                    
					<br>
				</form>
     				<div class="noprint">
                   		<div align=center>
                    		<span class="btnType01"><input type="button" value="출력" onclick="javascript:printWindow();"></span>            
                    		<span class="btnType01"><input type="button" value="닫기" onclick="javascript:goAction();"></span>            
                    	</div>
    				</div>
				</div>
			</div>
	</body>
</html>

