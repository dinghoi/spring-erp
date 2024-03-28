<!--#include virtual="/include/google_analytics.asp" -->
<div id="header">
    <h1>
		<img src="/images/logo_long_ko_main.png" alt="K-ONE" width="120" height="40"/>
	</h1>
    <h2>
		<div style="margin:6px;">
			<img src="/image/insa_title.gif" alt="인사관리 시스템" width="174" height="22"/>
		</div>
	</h2>
    <div class="login">
        <p>
        <strong><%=Request.Cookies("nkpmg_user")("coo_user_name")%>&nbsp;<%=Request.Cookies("nkpmg_user")("coo_user_grade")%></strong>님 안녕하세요.
        </p>
    </div>
    <div id="gnb">
        <ul>
            <li class="dep1"><a href="/insa/insa_report_mg.asp" target="_parent">조회 및 현황</a></li>
            <li class="dep1"><a href="/insa/insa_mg.asp" target="_parent">인사관리</a></li>
            <li class="dep1"><a href="/insa/insa_appoint_mg.asp" target="_parent">발령관리/서식</a></li>
			<li class="dep1"><a href="/insa/insa_confirm_mg.asp" target="_parent">복리후생/제증명</a></li>
            <li class="dep1"><a href="/insa/insa_car_mg.asp" target="_parent">차량관리</a></li>
            <li class="dep1"><a href="/insa/insa_org_mg.asp" target="_parent">조직 및 코드관리</a></li>
        <%If SysAdminYn = "Y" Then%>
            <li class="dep1"><a href="/insa_gun_mg.asp" target="_parent"><span style="color:red;">근태관리</span></a></li>
            <li class="dep1"><a href="/insa/insa_welfare_ask_mg.asp" target="_parent"><span style="color:red;">복리후생/제증명</span></a></li>
            <li class="dep1"><a href="/insa/insa_promotion_list.asp" target="_parent"><span style="color:red;">조회 및 현황(2)</span></a></li>

            <li class="dep1"><a href="/insa_sawo_mg.asp" target="_parent"><span style="color:red;">경조회관리</span></a></li>
        <%End If%>
        </ul>
    </div>
</div>
