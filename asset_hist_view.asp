<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

asset_no = request("asset_no")
dept_name = request("dept_name")

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set rs_hist = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

sql = "select * from asset inner join asset_dept on (asset.company = asset_dept.company) and (asset.dept_code = asset_dept.dept_code) where asset.asset_no ='" + asset_no + "'"
set rs = dbconn.execute(sql)

etc_code = "75" + cstr(mid(asset_no,1,2))
sql = "select * from etc_code where etc_code = '" + etc_code + "'"
set rs_etc = dbconn.execute(sql)

sql = "select * from asset_history inner join asset_dept on (substr(asset_history.asset_no,1,2) = asset_dept.company) and (asset_history.dept_code = asset_dept.dept_code) where asset_history.asset_no = '" + asset_no + "' order by history_seq desc"
rs_hist.Open Sql, Dbconn, 1

title_line = "자산 이전 History 조회"

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
							  <th>자산번호</th>
							  <td class="left"><%=mid(asset_no,1,2)%>-<%=mid(asset_no,3,6)%>-<%=right(asset_no,4)%>&nbsp;<%=rs_etc("etc_name")%></td>
							  <th>자산명</th>
							  <td class="left"><%=rs("asset_name")%></td>
					      	</tr>
							<tr>
							  <th>조직명</th>
							  <td class="left"><%=rs("org_first")%>&nbsp;/&nbsp;<%=rs("org_second")%>&nbsp;/&nbsp;<%=rs("dept_name")%></td>
							  <th>사용자</th>
							  <td class="left"><%=rs("user_name")%></td>
					      	</tr>
							<tr>
							  <th>최종설치일</th>
							  <td class="left"><%=rs("install_date")%></td>
							  <th>변경일</th>
							  <td class="left"><%=mid(rs("mod_date"),1,10)%></td>
					      	</tr>
						</tbody>
					</table>
					<h3 class="stit">* 이전 History</h3>
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="5%" >
							<col width="35%" >
							<col width="10%" >
							<col width="12%" >
							<col width="*" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">순번</th>
								<th scope="col">조직명</th>
								<th scope="col">사용자</th>
								<th scope="col">설치일</th>
								<th scope="col" class="left">이전사유</th>
							</tr>
						</thead>
						<tbody>
						<%
                        do until rs_hist.eof 
                        %>
							<tr>
								<td class="first"><%=rs_hist("history_seq")%></td>
								<td><%=rs_hist("org_first")%>&nbsp;/&nbsp;<%=rs_hist("org_second")%>&nbsp;/&nbsp;<%=rs_hist("dept_name")%></td>
								<td><%=rs_hist("user_name")%></td>
								<td><%=rs_hist("install_date")%></td>
								<td class="left"><%=rs_hist("trans_memo")%></td>
							</tr>
						<%
                            rs_hist.movenext()  
                        loop
                        rs_hist.Close()
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

