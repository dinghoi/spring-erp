<%
'====================================
'Author : ����ȣ
'Modify Date : 20201117
'Desc : ���� ��� ���� ���� �ְ�
'	���� include �� "/include/nkpmg_user.asp" ������ ���� include �Ǿ�� ��
'====================================

'====================================
' ���� ��� ����
'====================================
Dim SYSDATE, SYSDATE12, OrderByOrgName

' �ý��� ��¥(���ڿ� 8�ڸ�, YYYYMMDD )
SYSDATE = Year(Now) & Right("0" & Month(Now), 2) & Right("0" & Day(Now), 2)

' �ý��� ��¥(���ڿ� 12�ڸ�, YYYYMMDDhhmm )
SYSDATE12 = Year(Now) & Right("0" & Month(Now), 2) & Right("0" & Day(Now), 2) & Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2)

'�ý��� ������ ����
Dim SysAdminYn

Select Case user_id
	'�(����ȣ)
	Case "102592"
		SysAdminYn = "Y"
	Case Else
		SysAdminYn = "N"
End Select

'������ ���� �켱 ����
OrderByOrgName = "'���̿�', '���̳�Ʈ����', '���̽ý���', '��������ǻ��', '����������ġ', '���̴���', '�����̸�����'"
'====================================
'���� ���� ���� �޴� ���� ���� ����
'====================================
Dim IntroMemberYn, MemberPayAdmin, u_id, m_seq, m_name

u_id = f_toString(Request.cookies("nkp_member")("coo_user_id"), "")
m_seq = f_toString(Request.cookies("nkp_member")("coo_m_seq"), "")
m_name = f_toString(Request.cookies("nkp_member")("coo_m_name"), "")

IntroMemberYn = "N"'�ű� ä�� ����

If u_id = "kone" Then
	IntroMemberYn = "Y"
End If

MemberPayAdmin = "N"'������Ȳ ���� ����

If position = "�Ѱ���ǥ" Or position = "������" Or position = "�������" Or SysAdminYn = "Y" Then
	MemberPayAdmin = "Y"
End If
'====================================
'���� ���� �޴� ���� ���� ����
'====================================

'====================================
'�λ� ���� �޴� ���� ����
'====================================
Dim InsaMasterModYn	'�λ� �⺻ ���� ���� ���� ����
Dim InsaMasterDelYn	'�λ� �⺻ ���� ���� ���� ����
Dim InsaCarDelYn	'���� ���� ���� ����

'�λ� �⺻ ���� ���� ����
Select Case user_id
	'�(������, ������, ������, ������, ����ȣ, ������)
	Case "100018", "101622", "100104", "102560", "102592", "102615"
		InsaMasterModYn = "Y"
	Case Else
		InsaMasterModYn = "N"
End Select

'�λ� �⺻ ���� ���� ����
Select Case user_id
	'�(������, ������, ����ȣ, ������)
	Case "100018", "101622", "102592", "102615"
		InsaMasterDelYn = "Y"
	Case Else
		InsaMasterDelYn = "N"
End Select

'���� ���� ���� ����
Select Case user_id
	'(������, ����ȣ)
	Case "100104", "102592"
		InsaCarDelYn = "Y"
	Case Else
		InsaCarDelYn = "N"
End Select

'====================================
'�޿� ���� �޴� ���� ����
'====================================

'====================================
'��� ���� �޴� ���� ����
'====================================
Dim resideEndViewYn

'��� ���� > ����������/����� �׸� ���� ���� ����
Select Case user_id
	'Case "900001", "100359", "100952", "101100", "102305", "102306", "102592"
	Case "100359", "102592"
		resideEndViewYn = "Y"
	Case Else
		resideEndViewYn = "N"
End Select

'====================================
'���� ���� �޴� ���� ����
'====================================
Dim ProfitLossYn
Dim SubEmpReportYn
Dim SubProfitLossYn
Dim SubKdcEmpReportYn

'������Ȳ �޴� ���� ���� ����
'If CInt(coo_sales_grade) < 2 Or (coo_bonbu = "ITO �������" And coo_position = "�������") Or (coo_bonbu = "ITO �������" And coo_position = "������") Then
'	ProfitLossYn = "Y"
'Else
'	ProfitLossYn = "N"
'End If

If ProfitLossYn = "Y" Then
	'����κ� �ο� ��Ȳ ���� �޴� ���� ���� ����
	Select Case coo_user_id
		Case "90001"
			SubEmpReportYn = "Y"
		Case Else
			SubEmpReportYn = "N"
	End Select

	'����κ� �������� ���� �޴� ���� ���� ����
	Select Case coo_user_id
		'������, ������, �����, ��ȫ��, ������, �ּ���(102305, 102306)
		Case "100952", "100703", "101100", "101664", "101880", "102305", "102306", "ktrental2"
			SubProfitLossYn = "Y"
		Case Else
			SubProfitLossYn = "N"
	End Select

	'����κ� �ο� ��Ȳ(kdc) ���� �޴� ���� ���� ����
	Select Case coo_user_id
		Case "100703", "101100", "101664", "101880", "102305", "102306", "ktrental2"
			SubKdcEmpReportYn = "Y"
		Case Else
			SubKdcEmpReportYn = "N"
	End Select
End If

'����κ� ��������(����) �޴� ���� ���� ����[����ȣ_20201221]
Dim ProfitLossMonthNewYn, CompanyCostYn, CoworkYn

Select Case user_id
	'������
	Case "101880"
		ProfitLossMonthNewYn = "Y"
	Case Else
		ProfitLossMonthNewYn = "N"
End Select

'�ŷ�ó ���� View ����
Select Case user_id
	'��ǥ�̻�, ����
	Case "100001", "100740"
		CompanyCostYn = "Y"

	'�λ���, ��ǥ�̻�(���̴���)
	'Case "101245", "100262", "102627"
	''	CompanyCostYn = "Y"
	'������
	'Case "102663"
	''	CompanyCostYn = "Y"
	'����
	'Case "100020", "100125", "100173", "100180", "100186", "100187", "100244", "100703", "100953", "101246", "102271", "102665"
	''	CompanyCostYn = "Y"

	'�繫�̻�
	Case "100359"
		CompanyCostYn = "Y"
	Case Else
		CompanyCostYn = "N"
End Select

'���� View ����
Select Case user_id
	'��ǥ�̻�, ����
	Case "100001", "100740"
		CoworkYn = "Y"
	'�λ���, ������
	'Case "101245", "100262", "102663"
	'	CoworkYn = "Y"
	'����
	'Case "100020", "100125", "100173", "100180", "100186", "100187", "100244", "100703", "100953", "101246", "102271", "102665"
	'	CoworkYn = "Y"
	'�繫�̻�
	Case "100359"
		CoworkYn = "Y"
	Case Else
		CoworkYn = "N"
End Select

'����κ� �ο���Ȳ(KDC) �޴� ���� ���� ����[����ȣ_20201221]
Dim empReportKDCYn

Select Case user_id
	'������
	Case "100703"
		empReportKDCYn = "Y"
	Case Else
		empReportKDCYn = "N"
End Select

'����κ� �����Ѱ� �� ��ũ ���� ����(����)
Dim empProfitViewAll, empProfitViewSI, empProfitViewNI, empProfitGrade
empProfitViewAll = "N"
empProfitViewSI = "N"
empProfitViewNI = "N"
empProfitGrade = "N"	'������ ����

Select Case user_id
	'�����(ȸ��), �۰���(��ǥ), ������(�繫�̻�), ����ȣ(�ý��۰�����), ������(�繫), �ڼ���(�繫), �ڸ�ȣ(�繫), ������(�繫), �ձ��(��)
	Case "100001", "100740", "100359", "102592", "100842", "101672", "101902", "102825", "102498"
		'���� ����
		empProfitViewAll = "Y"
		empProfitGrade = "Y"
	'��ȸö
	Case "101245"
		'SI1, SI2 ����
		empProfitViewSI = "Y"
		empProfitGrade = "Y"
	'�̵���
	Case "102652"
		'NI, ICT ����
		empProfitViewNI = "Y"
		empProfitGrade = "Y"
End Select

'�߰� ������ ����
Select Case user_id
	'������, ����, �ֱ漺, ������, �����, ���뼮, ������, ������
	Case "102489", "100125", "100031", "100021", "101246", "100262", "100703", "102926"
		empProfitGrade = "Y"
End Select

'�ι������, �ι������(2) ����
Dim partCostView, subPartCostView

Select Case bonbu
	Case "SI1����", "SI2����", "NI����", "��������"
		partCostView = "Y"
	Case "����SI����", "����SI����", "DI����ι�"
		subPartCostView = "Y"
	Case Else
		partCostView = "N"
		subPartCostView = "N"
End Select

'====================================
'��ǰ ���� ���� �޴� ���� ����
'====================================
Dim NwInMenuYn

'N/W ����� ���� ���� ����

Select Case user_id
	'������, ������, ����ȣ
	Case "100359", "100015", "102592"
		NwInMenuYn = "Y"
	Case Else
		NwInMenuYn = "N"
End Select



'====================================
'ȸ�� ���� �޴� ���� ����
'====================================

'====================================
'�ӿ� ���� ���� �޴� ���� ����
'====================================
%>
