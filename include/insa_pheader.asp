<!--#include virtual="/include/google_analytics.asp" -->
<div id="header">
	<h1>
		<img src="/images/logo_long_ko_main.png" alt="K-ONE" width="120" height="40"/>
	</h1>
	<%If IntroMemberYn = "N" Then%>
	<h2>
		<div style="margin:6px;">
			<img src="/image/person_title.gif" alt="개인정보관리 시스템" width="158" height="22"/>
		</div>
	</h2>
	<%End If%>

	<div class="login">
		<p>
		<%If IntroMemberYn = "N" Then'신규회원 가입 여부%>
			<strong><%=request.cookies("nkpmg_user")("coo_user_name")%>&nbsp;<%=request.cookies("nkpmg_user")("coo_user_grade")%></strong>님 안녕하세요.
		<%Else%>
			<a href="/member/logout.asp"><img src="/image/logout.gif" alt="로그아웃"/></a>
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
			<li class="dep1"><a href="/person/insa_person_mg.asp" target="_parent">인사관리</a></li>
			<li class="dep1"><a href="/person/insa_individual_confirm.asp" target="_parent">복리후생/제증명</a></li>
			<li class="dep1"><a href="/person/insa_plist_pay_mg.asp" target="_parent">급여내역</a></li>
		 <%
		 	If SysAdminYn = "Y" Then
		 %>
		 	<li class="dep1"><a href="/insa_individual_board.asp"><span style="color:red;">인사 게시판</span></a></li>
		 	<li class="dep1"><a href="/person/insa_individual_sawo.asp" target="_parent"><span style="color:red;">복리후생/제증명</span></a></li>
		 	<li class="dep1"><a href="insa_individual_agree.asp" target="_parent"><span style="color:red;">근로계약</span></a></li>
			<li class="dep1"><a href="/insa_pay_yeartax_family.asp" target="_parent"><span style="color:red;">연말정산</span></a></li>
		 	<li class="dep1"><a href="/insa_individual_gun.asp" target="_parent"><span style="color:red;">근태/발령서식</span></a></li>
			<li class="dep1"><a href="/insa_org_data_batch_add.asp"><span style="color:red;">1신규조직 타계열사 등록</span></a></li>
			<li class="dep1"><a href="/insa_org_data_batch_trans1.asp"><span style="color:red;">2대량조직변경 org</span></a></li>
			<li class="dep1"><a href="/insa_org_data_batch_emp1.asp"><span style="color:red;">3대량 인사발령 emp</span></a></li>
		 	<li class="dep1"><a href="insa_pay_data_org_check.asp"><span style="color:red;">3급여조직넣기</span></a></li>
			<li class="dep1"><a href="insa_mst_month_trans_save.asp"><span style="color:red;">4인사발령소급정리</span></a></li>
		 <%
		 	End If
		 %>
		</ul>
		<%Else%>
		<ul>
			<li class="dep1"><a href="/member/member_add.asp" target="_parent">회원 관리</a></li>
		</ul>
		<%End If%>
	</div>
</div>