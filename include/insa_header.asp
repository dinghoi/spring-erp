<!--#include virtual="/include/google_analytics.asp" -->
<div id="header">
    <h1>
		<img src="/images/logo_long_ko_main.png" alt="K-ONE" width="120" height="40"/>
	</h1>
    <h2>
		<div style="margin:6px;">
			<img src="/image/insa_title.gif" alt="�λ���� �ý���" width="174" height="22"/>
		</div>
	</h2>
    <div class="login">
        <p>
        <strong><%=Request.Cookies("nkpmg_user")("coo_user_name")%>&nbsp;<%=Request.Cookies("nkpmg_user")("coo_user_grade")%></strong>�� �ȳ��ϼ���.
        </p>
    </div>
    <div id="gnb">
        <ul>
            <li class="dep1"><a href="/insa/insa_report_mg.asp" target="_parent">��ȸ �� ��Ȳ</a></li>
            <li class="dep1"><a href="/insa/insa_mg.asp" target="_parent">�λ����</a></li>
            <li class="dep1"><a href="/insa/insa_appoint_mg.asp" target="_parent">�߷ɰ���/����</a></li>
			<li class="dep1"><a href="/insa/insa_confirm_mg.asp" target="_parent">�����Ļ�/������</a></li>
            <li class="dep1"><a href="/insa/insa_car_mg.asp" target="_parent">��������</a></li>
            <li class="dep1"><a href="/insa/insa_org_mg.asp" target="_parent">���� �� �ڵ����</a></li>
        <%If SysAdminYn = "Y" Then%>
            <li class="dep1"><a href="/insa_gun_mg.asp" target="_parent"><span style="color:red;">���°���</span></a></li>
            <li class="dep1"><a href="/insa/insa_welfare_ask_mg.asp" target="_parent"><span style="color:red;">�����Ļ�/������</span></a></li>
            <li class="dep1"><a href="/insa/insa_promotion_list.asp" target="_parent"><span style="color:red;">��ȸ �� ��Ȳ(2)</span></a></li>

            <li class="dep1"><a href="/insa_sawo_mg.asp" target="_parent"><span style="color:red;">����ȸ����</span></a></li>
        <%End If%>
        </ul>
    </div>
</div>
