<!--#include virtual="/include/google_analytics.asp" -->
<div id="header">
	<h1>
		<img src="/images/logo_long_ko_main.png" alt="K-ONE" width="120" height="40"/>
	</h1>
	<%If IntroMemberYn = "N" Then%>
	<h2>
		<div style="margin:6px;">
			<img src="/image/person_title.gif" alt="������������ �ý���" width="158" height="22"/>
		</div>
	</h2>
	<%End If%>

	<div class="login">
		<p>
		<%If IntroMemberYn = "N" Then'�ű�ȸ�� ���� ����%>
			<strong><%=request.cookies("nkpmg_user")("coo_user_name")%>&nbsp;<%=request.cookies("nkpmg_user")("coo_user_grade")%></strong>�� �ȳ��ϼ���.
		<%Else%>
			<a href="/member/logout.asp"><img src="/image/logout.gif" alt="�α׾ƿ�"/></a>
		<%End If%>
		</p>
	</div>
	<%
	Dim in_name, in_empno

	in_name = request.cookies("nkpmg_user")("coo_user_name")
	in_empno = request.cookies("nkpmg_user")("coo_user_id")
	position = request.cookies("nkpmg_user")("coo_position")
	%>
	<div id="gnb">
		<%If IntroMemberYn = "N" Then%>
		<ul>
			<li class="dep1"><a href="/person/insa_person_mg.asp" target="_parent">�λ����</a></li>
			<li class="dep1"><a href="/person/insa_individual_confirm.asp" target="_parent">�����Ļ�/������</a></li>
			<li class="dep1"><a href="/person/insa_plist_pay_mg.asp" target="_parent">�޿�����</a></li>
		 <%
		 	If SysAdminYn = "Y" Then
		 %>
		 	<li class="dep1"><a href="/insa_individual_board.asp"><span style="color:red;">�λ� �Խ���</span></a></li>
		 	<li class="dep1"><a href="/person/insa_individual_sawo.asp" target="_parent"><span style="color:red;">�����Ļ�/������</span></a></li>
		 	<li class="dep1"><a href="insa_individual_agree.asp" target="_parent"><span style="color:red;">�ٷΰ��</span></a></li>
			<li class="dep1"><a href="/insa_pay_yeartax_family.asp" target="_parent"><span style="color:red;">��������</span></a></li>
		 	<li class="dep1"><a href="/insa_individual_gun.asp" target="_parent"><span style="color:red;">����/�߷ɼ���</span></a></li>
			<li class="dep1"><a href="/insa_org_data_batch_add.asp"><span style="color:red;">1�ű����� Ÿ�迭�� ���</span></a></li>
			<li class="dep1"><a href="/insa_org_data_batch_trans1.asp"><span style="color:red;">2�뷮�������� org</span></a></li>
			<li class="dep1"><a href="/insa_org_data_batch_emp1.asp"><span style="color:red;">3�뷮 �λ�߷� emp</span></a></li>
		 	<li class="dep1"><a href="insa_pay_data_org_check.asp"><span style="color:red;">3�޿������ֱ�</span></a></li>
			<li class="dep1"><a href="insa_mst_month_trans_save.asp"><span style="color:red;">4�λ�߷ɼұ�����</span></a></li>
		 <%
		 	End If
		 %>
		</ul>
		<%Else%>
		<ul>
			<li class="dep1"><a href="/member/member_add.asp" target="_parent">ȸ�� ����</a></li>
		</ul>
		<%End If%>
	</div>
</div>