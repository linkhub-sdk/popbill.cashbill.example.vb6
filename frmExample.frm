VERSION 5.00
Begin VB.Form frmExample 
   Caption         =   "�˺� ���ݿ����� SDK ����"
   ClientHeight    =   11175
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16455
   LinkTopic       =   "Form1"
   ScaleHeight     =   11175
   ScaleWidth      =   16455
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.TextBox txtURL 
      Height          =   315
      Left            =   12120
      TabIndex        =   70
      Top             =   240
      Width           =   3975
   End
   Begin VB.Frame Frame7 
      Caption         =   "���ݿ����� ���� ���"
      Height          =   7185
      Left            =   240
      TabIndex        =   12
      Top             =   3720
      Width           =   15570
      Begin VB.Frame Frame9 
         Caption         =   "��ù��� ���μ��� "
         Height          =   2415
         Left            =   1800
         TabIndex        =   46
         Top             =   1440
         Width           =   3375
         Begin VB.CommandButton btnDelete_sub 
            Caption         =   "����"
            Height          =   375
            Left            =   1920
            Style           =   1  '�׷���
            TabIndex        =   49
            Top             =   1560
            Width           =   855
         End
         Begin VB.CommandButton btnCancelIssue_sub 
            BackColor       =   &H00C0C0FF&
            Caption         =   "�������"
            Height          =   375
            Left            =   600
            Style           =   1  '�׷���
            TabIndex        =   48
            Top             =   1560
            Width           =   975
         End
         Begin VB.CommandButton btnRegistIssue 
            BackColor       =   &H00C0C0FF&
            Caption         =   "��ù���"
            Height          =   430
            Left            =   600
            Style           =   1  '�׷���
            TabIndex        =   47
            Top             =   480
            Width           =   975
         End
         Begin VB.Line Line5 
            X1              =   960
            X2              =   2370
            Y1              =   1755
            Y2              =   1755
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  '�������� ����
            FillColor       =   &H00E0E0E0&
            Height          =   660
            Left            =   480
            Top             =   360
            Width           =   2415
         End
         Begin VB.Line Line4 
            X1              =   960
            X2              =   960
            Y1              =   1680
            Y2              =   840
         End
      End
      Begin VB.Frame Frame14 
         Caption         =   " �μ�/����"
         Height          =   2760
         Left            =   7680
         TabIndex        =   32
         Top             =   4125
         Width           =   5370
         Begin VB.CommandButton btnGetViewURl 
            Caption         =   "���ݿ����� ���� URL(�޴�x)"
            Height          =   375
            Left            =   210
            TabIndex        =   65
            Top             =   840
            Width           =   2625
         End
         Begin VB.CommandButton btnGetPDFURL 
            Caption         =   "PDF �ٿ�ε� URL"
            Height          =   390
            Left            =   3000
            TabIndex        =   63
            Top             =   840
            Width           =   2295
         End
         Begin VB.CommandButton btnGetEPrintUrl 
            Caption         =   "���޹޴��� �μ� �˾� URL"
            Height          =   390
            Left            =   210
            TabIndex        =   37
            Top             =   1740
            Width           =   2625
         End
         Begin VB.CommandButton btnGetPopUpURL 
            Caption         =   "���ݿ����� ���� URL"
            Height          =   390
            Left            =   210
            TabIndex        =   36
            Top             =   390
            Width           =   2625
         End
         Begin VB.CommandButton btnGetPrintURL 
            Caption         =   "���ݿ����� �μ� �˾� URL"
            Height          =   390
            Left            =   210
            TabIndex        =   35
            Top             =   1305
            Width           =   2625
         End
         Begin VB.CommandButton btnGetMassPrintURL 
            Caption         =   "�뷮 �μ� �˾� URL"
            Height          =   390
            Left            =   210
            TabIndex        =   34
            Top             =   2190
            Width           =   2625
         End
         Begin VB.CommandButton btnGetMailURL 
            Caption         =   "���ݿ����� ���ϸ�ũ URL"
            Height          =   390
            Left            =   3000
            TabIndex        =   33
            Top             =   360
            Width           =   2295
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   " ��Ÿ URL "
         Height          =   1890
         Left            =   13200
         TabIndex        =   28
         Top             =   4125
         Width           =   2265
         Begin VB.CommandButton btnGetURL_TBOX 
            Caption         =   "�ӽ� ������"
            Height          =   390
            Left            =   210
            TabIndex        =   31
            Top             =   390
            Width           =   1845
         End
         Begin VB.CommandButton btnGetURL_PBOX 
            Caption         =   "���� ������"
            Height          =   390
            Left            =   210
            TabIndex        =   30
            Top             =   825
            Width           =   1845
         End
         Begin VB.CommandButton btnGetURL_WRITE 
            Caption         =   "���� �ۼ�"
            Height          =   390
            Left            =   210
            TabIndex        =   29
            Top             =   1275
            Width           =   1845
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "�ΰ� ���"
         Height          =   2775
         Left            =   2760
         TabIndex        =   24
         Top             =   4125
         Width           =   4815
         Begin VB.CommandButton btnAssignMgtKey 
            Caption         =   "������ȣ �Ҵ�"
            Height          =   390
            Left            =   2760
            TabIndex        =   64
            Top             =   360
            Width           =   1935
         End
         Begin VB.CommandButton btnUpdateemailconfig 
            Caption         =   "�˸����� ���ۼ��� ����"
            Height          =   375
            Left            =   240
            TabIndex        =   62
            Top             =   2280
            Width           =   2415
         End
         Begin VB.CommandButton btnListemailconfig 
            Caption         =   "�˸����� ���۸�� ��ȸ"
            Height          =   375
            Left            =   240
            TabIndex        =   61
            Top             =   1800
            Width           =   2415
         End
         Begin VB.CommandButton btnSendEmail 
            Caption         =   "�̸��� ����"
            Height          =   390
            Left            =   225
            TabIndex        =   27
            Top             =   360
            Width           =   2415
         End
         Begin VB.CommandButton btnSendSMS 
            Caption         =   "���� ����"
            Height          =   390
            Left            =   225
            TabIndex        =   26
            Top             =   825
            Width           =   2415
         End
         Begin VB.CommandButton btnSendFAX 
            Caption         =   "�ѽ� ����"
            Height          =   390
            Left            =   240
            TabIndex        =   25
            Top             =   1320
            Width           =   2415
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   " ���� ���� "
         Height          =   2775
         Left            =   240
         TabIndex        =   19
         Top             =   4125
         Width           =   2265
         Begin VB.CommandButton btnSearch 
            Caption         =   "��� ��ȸ"
            Height          =   390
            Left            =   195
            TabIndex        =   50
            Top             =   1800
            Width           =   1845
         End
         Begin VB.CommandButton btnGetDetailInfo 
            Caption         =   "�� ����Ȯ��"
            Height          =   390
            Left            =   195
            TabIndex        =   23
            Top             =   1320
            Width           =   1845
         End
         Begin VB.CommandButton btnGetLogs 
            Caption         =   "���� �����̷�"
            Height          =   390
            Left            =   195
            TabIndex        =   22
            Top             =   2280
            Width           =   1845
         End
         Begin VB.CommandButton btnGetInfos 
            Caption         =   "���� �뷮 Ȯ��"
            Height          =   390
            Left            =   195
            TabIndex        =   21
            Top             =   825
            Width           =   1845
         End
         Begin VB.CommandButton btnGetInfo 
            Caption         =   "���� Ȯ��"
            Height          =   390
            Left            =   195
            TabIndex        =   20
            Top             =   390
            Width           =   1845
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "������ݿ����� ��ù��� ���μ���"
         Height          =   2385
         Left            =   5640
         TabIndex        =   17
         Top             =   1440
         Width           =   4095
         Begin VB.CommandButton btnRegistIssue_part 
            Caption         =   "�κ���� ��ù���"
            Height          =   375
            Left            =   1680
            TabIndex        =   60
            Top             =   480
            Width           =   1815
         End
         Begin VB.CommandButton btnRevokeRegistIssue 
            BackColor       =   &H00C0C0FF&
            Caption         =   "��ù���"
            Height          =   375
            Left            =   480
            MaskColor       =   &H00C0C0FF&
            TabIndex        =   52
            Top             =   480
            Width           =   975
         End
         Begin VB.CommandButton btnCancelIssue 
            BackColor       =   &H00C0C0FF&
            Caption         =   "�������"
            Height          =   375
            Left            =   480
            Style           =   1  '�׷���
            TabIndex        =   38
            Top             =   1560
            Width           =   960
         End
         Begin VB.CommandButton btnDelete 
            Caption         =   "����"
            Height          =   375
            Left            =   1920
            Style           =   1  '�׷���
            TabIndex        =   18
            Top             =   1560
            Width           =   855
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  '�������� ����
            FillColor       =   &H00E0E0E0&
            Height          =   660
            Left            =   300
            Top             =   345
            Width           =   3495
         End
         Begin VB.Line Line1 
            X1              =   960
            X2              =   960
            Y1              =   1905
            Y2              =   600
         End
         Begin VB.Line Line2 
            X1              =   1080
            X2              =   2490
            Y1              =   1760
            Y2              =   1760
         End
      End
      Begin VB.CommandButton btnCheckMgtKeyInUse 
         Caption         =   "������ȣ ��뿩�� Ȯ��"
         Height          =   375
         Left            =   6150
         TabIndex        =   16
         Top             =   255
         Width           =   2190
      End
      Begin VB.TextBox txtMgtKey 
         Height          =   330
         Left            =   3300
         TabIndex        =   15
         Top             =   285
         Width           =   2775
      End
      Begin VB.Label Label4 
         Caption         =   "����û�� �Ű�� ���ݿ������� ����ϱ� ���ؼ���'������ݿ�����'�� �����ؾ� �մϴ�."
         Height          =   375
         Left            =   5640
         TabIndex        =   53
         Top             =   960
         Width           =   4095
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "������ȣ( MgtKey) : "
         Height          =   180
         Left            =   1440
         TabIndex        =   14
         Top             =   360
         Width           =   1755
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " �˺� �⺻ API "
      Height          =   2895
      Left            =   240
      TabIndex        =   4
      Top             =   720
      Width           =   15810
      Begin VB.Frame Frame15 
         Caption         =   "��Ʈ�ʰ��� ����Ʈ"
         Height          =   2295
         Left            =   13200
         TabIndex        =   56
         Top             =   360
         Width           =   2415
         Begin VB.CommandButton btnGetPartnerURL_CHRG 
            Caption         =   "����Ʈ ���� URL"
            Height          =   410
            Left            =   120
            TabIndex        =   59
            Top             =   840
            Width           =   2175
         End
         Begin VB.CommandButton btnGetPartnerBalance 
            Caption         =   "��Ʈ�� �ܿ�����Ʈ Ȯ��"
            Height          =   410
            Left            =   120
            TabIndex        =   57
            Top             =   360
            Width           =   2175
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "�������� ����Ʈ"
         Height          =   2295
         Left            =   10920
         TabIndex        =   54
         Top             =   360
         Width           =   2175
         Begin VB.CommandButton btnGetUseHistoryURL 
            Caption         =   "����Ʈ ��볻�� URL"
            Height          =   410
            Left            =   120
            TabIndex        =   68
            Top             =   1800
            Width           =   1935
         End
         Begin VB.CommandButton btnGetPaymentURL 
            Caption         =   "����Ʈ �������� URL"
            Height          =   410
            Left            =   120
            TabIndex        =   67
            Top             =   1320
            Width           =   1935
         End
         Begin VB.CommandButton btnGetChargeURL 
            Caption         =   " ����Ʈ ���� URL"
            Height          =   410
            Left            =   120
            TabIndex        =   58
            Top             =   840
            Width           =   1935
         End
         Begin VB.CommandButton btnGetBalance 
            Caption         =   "�ܿ� ����Ʈ Ȯ��"
            Height          =   410
            Left            =   120
            TabIndex        =   55
            Top             =   360
            Width           =   1935
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   " ȸ������ ����"
         Height          =   2295
         Left            =   8880
         TabIndex        =   43
         Top             =   360
         Width           =   1935
         Begin VB.CommandButton btnUpdateCorpInfo 
            Caption         =   "ȸ������ ����"
            Height          =   410
            Left            =   120
            TabIndex        =   45
            Top             =   840
            Width           =   1695
         End
         Begin VB.CommandButton btnGetCorpInfo 
            Caption         =   "ȸ������ ��ȸ"
            Height          =   410
            Left            =   120
            TabIndex        =   44
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   " ȸ������ "
         Height          =   2295
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   1695
         Begin VB.CommandButton btnCheckID 
            Caption         =   "ID �ߺ� Ȯ��"
            Height          =   410
            Left            =   120
            TabIndex        =   39
            Top             =   840
            Width           =   1455
         End
         Begin VB.CommandButton btnCheckIsMember 
            Caption         =   "���� ���� Ȯ��"
            Height          =   410
            Left            =   120
            TabIndex        =   13
            Top             =   360
            Width           =   1455
         End
         Begin VB.CommandButton btnJoinMember 
            Caption         =   "ȸ�� ����"
            Height          =   410
            Left            =   120
            TabIndex        =   11
            Top             =   1320
            Width           =   1455
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   " ����Ʈ ���� "
         Height          =   2295
         Left            =   2040
         TabIndex        =   8
         Top             =   360
         Width           =   2295
         Begin VB.CommandButton btnGetChargeInfo 
            Caption         =   "�������� Ȯ��"
            Height          =   410
            Left            =   120
            TabIndex        =   51
            Top             =   360
            Width           =   2055
         End
         Begin VB.CommandButton btnUnitCost 
            Caption         =   "��� �ܰ� Ȯ��"
            Height          =   410
            Left            =   120
            TabIndex        =   9
            Top             =   840
            Width           =   2055
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   " ����� ���� "
         Height          =   2295
         Left            =   4440
         TabIndex        =   7
         Top             =   360
         Width           =   2055
         Begin VB.CommandButton btnGetContactInfo 
            Caption         =   "����� ���� Ȯ��"
            Height          =   410
            Left            =   120
            TabIndex        =   66
            Top             =   840
            Width           =   1815
         End
         Begin VB.CommandButton btnRegistContact 
            Caption         =   "����� �߰�"
            Height          =   410
            Left            =   120
            TabIndex        =   42
            Top             =   360
            Width           =   1815
         End
         Begin VB.CommandButton btnListContact 
            Caption         =   "����� ��� ��ȸ"
            Height          =   410
            Left            =   120
            TabIndex        =   41
            Top             =   1320
            Width           =   1815
         End
         Begin VB.CommandButton btnUpdateContact 
            Caption         =   "����� ���� ����"
            Height          =   410
            Left            =   120
            TabIndex        =   40
            Top             =   1800
            Width           =   1815
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   " �˺� �⺻ URL "
         Height          =   2295
         Left            =   6600
         TabIndex        =   5
         Top             =   360
         Width           =   2175
         Begin VB.CommandButton btnGetAccessURL 
            Caption         =   " �˺� �α��� URL"
            Height          =   410
            Left            =   120
            TabIndex        =   6
            Top             =   360
            Width           =   1935
         End
      End
   End
   Begin VB.TextBox txtUserID 
      Height          =   315
      Left            =   5880
      TabIndex        =   3
      Text            =   "testkorea"
      Top             =   210
      Width           =   1935
   End
   Begin VB.TextBox txtCorpNum 
      Height          =   315
      Left            =   2175
      TabIndex        =   1
      Text            =   "1234567890"
      Top             =   225
      Width           =   1935
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "URL : "
      Height          =   180
      Left            =   11400
      TabIndex        =   69
      Top             =   285
      Width           =   525
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "�˺�ȸ�� ���̵� : "
      Height          =   180
      Left            =   4320
      TabIndex        =   2
      Top             =   285
      Width           =   1500
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "�˺�ȸ�� ����ڹ�ȣ : "
      Height          =   180
      Left            =   240
      TabIndex        =   0
      Top             =   285
      Width           =   1860
   End
End
Attribute VB_Name = "frmExample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================================
'
' �˺� ���ݿ����� API VB 6.0 SDK Example
'
' - ������Ʈ ���� : 2022-01-17
' - ���� ������� ����ó : 1600-9854
' - ���� ������� �̸��� : code@linkhubcorp.com
' - VB6 SDK ����ȯ�� ������� �ȳ� : https://docs.popbill.com/cashbill/tutorial/vb
'
' <�׽�Ʈ �������� �غ����>
' 1) 25, 28�� ���ο� ����� ��ũ���̵�(LinkID)�� ���Ű(SecretKey)��
'    ��ũ��� ���Խ� ���Ϸ� �߱޹��� ���������� �����Ͽ� �����մϴ�.
'=========================================================================


Option Explicit

'=========================================================================
' - ��������(��ũ���̵�, ���Ű)�� ��Ʈ���� ����ȸ���� �ĺ��ϴ�
'   ������ ���Ǵ� ������ ������� �ʵ��� �����Ͻñ� �ٶ��ϴ�.
' - ����� ��ȯ���Ŀ��� ��������(��ũ���̵�, ���Ű)�� ������� �ʽ��ϴ�.
'=========================================================================

'��ũ���̵�
Private Const linkID = "TESTER"

'���Ű. ���⿡ �����Ͻñ� �ٶ��ϴ�.
Private Const SecretKey = "SwWxqU+0TErBXy/9TVjIPEnI0VTUMMSQZtJf3Ed8q3I="

'���ݿ����� ���� ��ü ����
Private CashbillService As New PBCBService

'=========================================================================
' �˺� ����Ʈ�� ���� �����Ͽ� ������ȣ�� �ο����� ���� ���ݿ������� ������ȣ�� �Ҵ��մϴ�.
' - https://docs.popbill.com/cashbill/vb/api#AssignMgtKey
'=========================================================================
Private Sub btnAssignMgtKey_Click()
    Dim Response As PBResponse
    Dim itemKey As String
    Dim mgtKey As String
    
    '���ݿ����� ������Ű, �����ȸ(Search) API�� ��ȯ�׸��� ItemKey ����
    itemKey = "020042413523200001"
            
    '�Ҵ��� ������ȣ, �ִ� 24�ڸ�, ����, ���� '-', '_'�� �����Ͽ� ����ں��� �ߺ����� �ʵ��� ����
    mgtKey = "20220101-05"
        
    Set Response = CashbillService.AssignMgtKey(txtCorpNum.Text, itemKey, mgtKey)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ����ڹ�ȣ�� ��ȸ�Ͽ� ����ȸ�� ���Կ��θ� Ȯ���մϴ�.
' - LinkID�� ���������� �����Ǿ� �ִ� ��ũ���̵� ���Դϴ�.
' - https://docs.popbill.com/cashbill/vb/api#CheckIsMember
'=========================================================================
Private Sub btnCheckIsMember_Click()
    Dim Response As PBResponse
    
    Set Response = CashbillService.CheckIsMember(txtCorpNum.Text, linkID)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ����ϰ��� �ϴ� ���̵��� �ߺ����θ� Ȯ���մϴ�.
' - https://docs.popbill.com/cashbill/vb/api#CheckID
'=========================================================================
Private Sub btnCheckID_Click()
    Dim Response As PBResponse
    
    Set Response = CashbillService.CheckID(txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ���ݿ����� PDF ������ �ٿ� ���� �� �ִ� URL�� ��ȯ�մϴ�.
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
'=========================================================================
Private Sub btnGetPDFURL_Click()
    Dim URL As String
    
    URL = CashbillService.GetPDFURL(txtCorpNum.Text, txtMgtKey.Text)
    
    If URL = "" Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + URL
    txtURL.Text = URL
End Sub


'=========================================================================
' ����ڸ� ����ȸ������ ����ó���մϴ�.
' - https://docs.popbill.com/cashbill/vb/api#JoinMember
'=========================================================================
Private Sub btnJoinMember_Click()
    Dim joinData As New PBJoinForm
    Dim Response As PBResponse
    
    '���̵�, 6���̻� 50�� �̸�
    joinData.id = "userid"
    
    '��й�ȣ, 8�� �̻� 20�� ����(����, ����, Ư������ ����)
    joinData.Password = "asdf$%^123"
    
    '��Ʈ�ʸ�ũ ���̵�
    joinData.linkID = linkID
    
    '����ڹ�ȣ, '-'����, 10�ڸ�
    joinData.CorpNum = "1234567890"
    
    '��ǥ�ڼ���, �ִ� 100��
    joinData.ceoname = "��ǥ�ڼ���"
    
    '��ȣ��, �ִ� 200��
    joinData.corpName = "ȸ����ȣ"
    
    '����� �ּ�, �ִ� 300��
    joinData.addr = "�ּ�"
    
    '����, �ִ� 100��
    joinData.bizType = "����"
    
    '����, �ִ� 100��
    joinData.bizClass = "����"

    '����� ����, �ִ� 100��
    joinData.ContactName = "����ڼ���"
    
    '����� �̸���, �ִ� 100��
    joinData.ContactEmail = "test@test.com"
    
    '����� ����ó, �ִ� 20��
    joinData.ContactTEL = "02-999-9999"
    
    '����� �޴�����ȣ, �ִ� 20��
    joinData.ContactHP = "010-1234-5678"
    
    '����� �ѽ���ȣ, �ִ� 20��
    joinData.ContactFAX = "02-999-9998"
    
    Set Response = CashbillService.JoinMember(joinData)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ���ݿ����� ����� ���ݵǴ� ����Ʈ �ܰ��� Ȯ���մϴ�.
' - https://docs.popbill.com/cashbill/vb/api#GetUnitCost
'=========================================================================
Private Sub btnUnitCost_Click()
    Dim unitCost As Double
    
    unitCost = CashbillService.GetUnitCost(txtCorpNum.Text)
    
    If unitCost < 0 Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "����ܰ� : " + CStr(unitCost)
End Sub

'=========================================================================
' �˺� ���ݿ����� API ���� ���������� Ȯ���մϴ�.
' - https://docs.popbill.com/cashbill/vb/api#GetChargeInfo
'=========================================================================
Private Sub btnGetChargeInfo_Click()
    Dim ChargeInfo As PBChargeInfo
    Dim tmp As String
    
    Set ChargeInfo = CashbillService.GetChargeInfo(txtCorpNum.Text)
     
    If ChargeInfo Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "unitCost (����ܰ�) : " + ChargeInfo.unitCost + vbCrLf
    tmp = tmp + "chargeMethod (��������) : " + ChargeInfo.chargeMethod + vbCrLf
    tmp = tmp + "rateSystem (��������) : " + ChargeInfo.rateSystem + vbCrLf
    
    MsgBox tmp
End Sub

'=========================================================================
' �˺� ����Ʈ�� �α��� ���·� ������ �� �ִ� �������� �˾� URL�� ��ȯ�մϴ�.
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
' - https://docs.popbill.com/cashbill/vb/api#GetAccessURL
'=========================================================================
Private Sub btnGetAccessURL_Click()
    Dim URL As String
     
    URL = CashbillService.GetAccessURL(txtCorpNum.Text, txtUserID.Text)
    
    If URL = "" Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + URL
    txtURL.Text = URL
End Sub

'=========================================================================
' ����ȸ�� ����ڹ�ȣ�� �����(�˺� �α��� ����)�� �߰��մϴ�.
' - https://docs.popbill.com/cashbill/vb/api#RegistContact
'=========================================================================
Private Sub btnRegistContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    '����� ���̵�, 6�� �̻� 50�� �̸�
    joinData.id = "vb6Cashbill001"
    
    '��й�ȣ, 8�� �̻� 20�� ����(����, ����, Ư������ ����)
    joinData.Password = "qwe123!@#"
    
    '����ڸ�, �ִ� 100��
    joinData.personName = "����ڸ�"
    
    '����� ����ó, �ִ� 20��
    joinData.tel = "070-1234-1234"
    
    '����� �޴�����ȣ, �ִ� 20��
    joinData.hp = "010-1234-1234"
    
    '����� �ѽ���,�ִ� 20��
    joinData.fax = "070-1234-1234"
    
    '����� �����ּ�, �ִ� 100��
    joinData.email = "test@test.com"
    
    '����� ����, 1-���� / 2-�б� / 3-ȸ��
    joinData.searchRole = 3
    
        
    Set Response = CashbillService.RegistContact(txtCorpNum.Text, joinData)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ����ȸ�� ����ڹ�ȣ�� ��ϵ� �����(�˺� �α��� ����) ������ Ȯ���մϴ�.
' https://docs.popbill.com/cashbill/vb/api#GetContactInfo
'=========================================================================
Private Sub btnGetContactInfo_Click()
    Dim tmp As String
    Dim info As PBContactInfo
    Dim ContactID As String
    
    'Ȯ���� ����� ���̵�
    ContactID = "testkorea"
    
    Set info = CashbillService.GetContactInfo(txtCorpNum.Text, ContactID, txtUserID.Text)
    
    If info Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "id(���̵�) | personName(����) | email(�̸���) | hp(�޴�����ȣ) |  fax(�ѽ���ȣ) | tel(����ó) | " _
         + "regDT(����Ͻ�) | searchRole(����� ����) | mgrYN(������ ����) | state(����) " + vbCrLf
    
   
    tmp = tmp + info.id + " | " + info.personName + " | " + info.email + " | " + info.hp + " | " + info.fax _
        + info.tel + " | " + info.regDT + " | " + CStr(info.searchRole) + " | " + CStr(info.mgrYN) + " | " + CStr(info.state) + vbCrLf
        
    MsgBox tmp
End Sub

'=========================================================================
' ����ȸ�� ����ڹ�ȣ�� ��ϵ� �����(�˺� �α��� ����) ����� Ȯ���մϴ�.
' - https://docs.popbill.com/cashbill/vb/api#ListContact
'=========================================================================
Private Sub btnListContact_Click()
    Dim resultList As Collection
    Dim tmp As String
    Dim info As PBContactInfo
    
    Set resultList = CashbillService.ListContact(txtCorpNum.Text)
     
    If resultList Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "id(���̵�) | personName(����) | email(�̸���) | hp(�޴�����ȣ) |  fax(�ѽ���ȣ) | tel(����ó) | " _
         + "regDT(����Ͻ�) | searchRole(����� ����) | mgrYN(������ ����) | state(����) " + vbCrLf
    
    For Each info In resultList
        tmp = tmp + info.id + " | " + info.personName + " | " + info.email + " | " + info.hp + " | " + info.fax _
        + info.tel + " | " + info.regDT + " | " + CStr(info.searchRole) + " | " + CStr(info.mgrYN) + " | " + CStr(info.state) + vbCrLf
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' ����ȸ�� ����ڹ�ȣ�� ��ϵ� �����(�˺� �α��� ����) ������ �����մϴ�.
' - https://docs.popbill.com/cashbill/vb/api#UpdateContact
'=========================================================================
Private Sub btnUpdateContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    '����� ���̵�
    joinData.id = "vb6Cashbill001"
    
    '����� ����, �ִ� 100��
    joinData.personName = "����ڸ�_����"
    
    '����� ����ó, �ִ� 20��
    joinData.tel = "070-1234-1234"
    
    '����� �޴�����ȣ, �ִ� 20��
    joinData.hp = "010-1234-1234"
        
    '����� �ѽ���ȣ, �ִ� 20��
    joinData.fax = "070-1234-1234"
    
    '����� �̸���, �ִ� 100��
    joinData.email = "test@test.com"

    '����� ����, 1-���� / 2-�б� / 3-ȸ��
    joinData.searchRole = 3
    
                
    Set Response = CashbillService.UpdateContact(txtCorpNum.Text, joinData, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ����ȸ���� ȸ�������� Ȯ���մϴ�.
' - https://docs.popbill.com/cashbill/vb/api#GetCorpInfo
'=========================================================================
Private Sub btnGetCorpInfo_Click()
    Dim CorpInfo As PBCorpInfo
    Dim tmp As String
    
    Set CorpInfo = CashbillService.GetCorpInfo(txtCorpNum.Text)
     
    If CorpInfo Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "ceoname (��ǥ�� ����) : " + CorpInfo.ceoname + vbCrLf
    tmp = tmp + "corpName (��ȣ) : " + CorpInfo.corpName + vbCrLf
    tmp = tmp + "addr (�ּ�) : " + CorpInfo.addr + vbCrLf
    tmp = tmp + "bizType (����) : " + CorpInfo.bizType + vbCrLf
    tmp = tmp + "bizClass (����) : " + CorpInfo.bizClass + vbCrLf
    
    MsgBox tmp
End Sub

'=========================================================================
' ����ȸ���� ȸ�������� �����մϴ�
' - https://docs.popbill.com/cashbill/vb/api#UpdateCorpInfo
'=========================================================================
Private Sub btnUpdateCorpInfo_Click()
    Dim CorpInfo As New PBCorpInfo
    Dim Response As PBResponse
    
    '��ǥ�ڸ�, �ִ� 100��
    CorpInfo.ceoname = "��ǥ��"
    
    '��ȣ, �ִ� 200��
    CorpInfo.corpName = "��ȣ"
    
    '�ּ�, �ִ� 300��
    CorpInfo.addr = "����Ư����"
    
    '����, �ִ� 100��
    CorpInfo.bizType = "����"
    
    '����, �ִ� 100��
    CorpInfo.bizClass = "����"
    
    Set Response = CashbillService.UpdateCorpInfo(txtCorpNum.Text, CorpInfo)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ����ȸ�� ����Ʈ �������� Ȯ���� ���� �������� �˾� URL�� ��ȯ�մϴ�.
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
' - https://docs.popbill.com/cashbill/vb/api#GetPaymentURL
'=========================================================================
Private Sub btnGetPaymentURL_Click()
    Dim URL As String
           
    URL = CashbillService.GetPaymentURL(txtCorpNum.Text, txtUserID.Text)
    
    If URL = "" Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + URL
    txtURL.Text = URL
End Sub

'=========================================================================
' ����ȸ�� ����Ʈ ��볻�� Ȯ���� ���� �������� �˾� URL�� ��ȯ�մϴ�.
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
' - https://docs.popbill.com/cashbill/vb/api#GetUseHistoryURL
'=========================================================================
Private Sub btnGetUseHistoryURL_Click()
    Dim URL As String
           
    URL = CashbillService.GetUseHistoryURL(txtCorpNum.Text, txtUserID.Text)
    
    If URL = "" Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + URL
    txtURL.Text = URL
End Sub

'=========================================================================
' ����ȸ���� �ܿ�����Ʈ�� Ȯ���մϴ�.
' - ���ݹ���� ��Ʈ�ʰ����� ��� ��Ʈ�� �ܿ�����Ʈ(GetPartnerBalance API)�� ���� Ȯ���Ͻñ� �ٶ��ϴ�.
' - https://docs.popbill.com/cashbill/vb/api#GetBalance
'=========================================================================
Private Sub btnGetBalance_Click()
    Dim balance As Double
    
    balance = CashbillService.GetBalance(txtCorpNum.Text)
    
    If balance < 0 Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "�ܿ�����Ʈ : " + CStr(balance)
End Sub

'=========================================================================
' ����ȸ�� ����Ʈ ������ ���� �������� �˾� URL�� ��ȯ�մϴ�.
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
' - https://docs.popbill.com/cashbill/vb/api#GetChargeURL
'=========================================================================
Private Sub btnGetChargeURL_Click()
    Dim URL As String
     
    URL = CashbillService.GetChargeURL(txtCorpNum.Text, txtUserID.Text)
    
    If URL = "" Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + URL
    txtURL.Text = URL
End Sub

'=========================================================================
' ��Ʈ���� �ܿ�����Ʈ�� Ȯ���մϴ�.
' - ���ݹ���� ���������� ��� ����ȸ�� �ܿ�����Ʈ(GetBalance API)�� �̿��Ͻñ� �ٶ��ϴ�.
' - https://docs.popbill.com/cashbill/vb/api#GetPartnerBalance
'=========================================================================
Private Sub btnGetPartnerBalance_Click()
    Dim balance As Double
    
    balance = CashbillService.GetPartnerBalance(txtCorpNum.Text)
    
    If balance < 0 Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "��Ʈ�� �ܿ�����Ʈ : " + CStr(balance)
End Sub

'=========================================================================
' ��Ʈ�� ����Ʈ ������ ���� �������� �˾� URL�� ��ȯ�մϴ�.
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
' - https://docs.popbill.com/cashbill/vb/api#GetPartnerURL
'=========================================================================
Private Sub btnGetPartnerURL_CHRG_Click()
    Dim URL As String
    
    URL = CashbillService.GetPartnerURL(txtCorpNum.Text, "CHRG")
    
    If URL = "" Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + URL
    txtURL.Text = URL
End Sub

'=========================================================================
' ��Ʈ�ʰ� ���ݿ����� ���� �������� �Ҵ��ϴ� ������ȣ ��뿩�θ� Ȯ���մϴ�.
' - �̹� ��� ���� ������ȣ�� �ߺ� ����� �Ұ��ϰ�, ���ݿ������� ������ ��쿡�� ������ȣ�� ������ �����մϴ�.
' - ������ȣ�� �ִ� 24�ڸ�, ����, ���� '-', '_'�� �����Ͽ� ����ں��� �ߺ����� �ʵ��� ����
' - https://docs.popbill.com/cashbill/vb/api#CheckMgtKeyInUse
'=========================================================================
Private Sub btnCheckMgtKeyInUse_Click()
    Dim Response As PBResponse
      
    Set Response = CashbillService.CheckMgtKeyInUse(txtCorpNum.Text, txtMgtKey.Text)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' �ۼ��� ���ݿ����� �����͸� �˺��� ����� ���ÿ� �����Ͽ� "����Ϸ�" ���·� ó���մϴ�.
' - ���ݿ����� ����û ���� ��å : https://docs.popbill.com/cashbill/ntsSendPolicy?lang=vb
' - https://docs.popbill.com/cashbill/vb/api#RegistIssue
'=========================================================================
Private Sub btnRegistIssue_Click()
    Dim Cashbill As New PBCashbill
    Dim Response As PBResponse
    Dim emailSubject As String
    
    '���ݿ����� ������ȣ, �ִ� 24�ڸ�, ����, ���� '-', '_'�� �����Ͽ� ����ں��� �ߺ����� �ʵ��� ����
    Cashbill.mgtKey = txtMgtKey.Text
    
    '[��Ұŷ��� �ʼ�] ���� ����û���ι�ȣ
    '��������(GetInfo API)�� �����׸��� ����û���ι�ȣ(confirmNum)�� Ȯ���Ͽ� ����
    Cashbill.orgConfirmNum = ""
    
    '[��Ұŷ��� �ʼ�] ���� �ŷ�����
    '��������(GetInfo API)�� �����׸��� �ŷ�����(tradeDate)�� Ȯ���Ͽ� ����
    Cashbill.orgTradeDate = ""
    
    '��������, [���ΰŷ�, ��Ұŷ�] �� ����
    Cashbill.tradeType = "���ΰŷ�"
    
    '�ŷ�����, [�ҵ������, ����������] �� ����
    Cashbill.tradeUsage = "�ҵ������"
    
    '�ŷ�����, [�Ϲ�, ��������, ���߱���] �� ����
    Cashbill.tradeOpt = "�Ϲ�"
    
    '��������, [����, �����] �� ����
    Cashbill.taxationType = "����"
    
    '�ŷ��ݾ�, ���ް��� + ����� + ����
    Cashbill.totalAmount = "11000"
    
    '���ް���
    Cashbill.supplyCost = "10000"
    
    '�ΰ���
    Cashbill.tax = "1000"
    
    '�����
    Cashbill.serviceFee = "0"
    
    '������ ����ڹ�ȣ, "-" ���� 10�ڸ�
    Cashbill.franchiseCorpNum = "1234567890"
    
    '������ ������� �ĺ���ȣ
    Cashbill.franchiseTaxRegID = ""
    
    '������ ��ȣ
    Cashbill.franchiseCorpName = "������ ��ȣ"
    
    '������ ��ǥ�� ����
    Cashbill.franchiseCEOName = "������ ��ǥ��"
    
    '������ �ּ�
    Cashbill.franchiseAddr = "������ �ּ�"
    
    '������ ��ȭ��ȣ
    Cashbill.franchiseTEL = "070-1234-1234"
        
    '�ĺ���ȣ, �ŷ����п� ���� �ۼ�
    '�ҵ������ - �ֹε��/�޴���/ī���ȣ ���簡��
    '���������� - ����ڹ�ȣ/�ֹε��/�޴���/ī���ȣ ���簡��
    Cashbill.identityNum = "0101112222"
        
    '�ֹ��ڸ�
    Cashbill.customerName = "�ֹ��ڸ�"
    
    '�ֹ���ǰ��
    Cashbill.itemName = "�ֹ���ǰ��"
    
    '�ֹ���ȣ
    Cashbill.orderNumber = "�ֹ���ȣ"
    
    '�ֹ��� �̸���
    '�˺� ����ȯ�濡�� �׽�Ʈ�ϴ� ��쿡�� �ȳ� ������ ���۵ǹǷ�,
    '���� �ŷ�ó�� �����ּҰ� ������� �ʵ��� ����
    Cashbill.email = "test@test.com"
    
    '�ֹ��� �޴���
    Cashbill.hp = "010-111-222"
    
    '���ݿ����� ���� �˸����� ���ۿ���
    Cashbill.smssendYN = False
    
    '�ȳ����� ����, �̱���� �⺻������� ����.
    emailSubject = ""
            
    Set Response = CashbillService.RegistIssue(txtCorpNum.Text, Cashbill, txtUserID.Text, emailSubject)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message + vbCrLf + "����û ���ι�ȣ : " + Response.confirmNum + vbCrLf + "�ŷ����� : " + Response.tradeDate)
End Sub

'=========================================================================
' ����û ���� ���� "����Ϸ�" ������ ���ݿ������� "�������"�ϰ� ����û �Ű� ��󿡼� �����մϴ�.
' - Delete(����)�Լ��� ȣ���Ͽ� "�������" ������ ���ݿ������� �����ϸ�, ������ȣ ������ �����մϴ�.
' - https://docs.popbill.com/cashbill/vb/api#CancelIssue
'=========================================================================
Private Sub btnCancelIssue_sub_Click()
    Dim Response As PBResponse
    Dim memo As String
    
    '�޸�
    memo = "������� �޸�"
    
    Set Response = CashbillService.CancelIssue(txtCorpNum.Text, txtMgtKey.Text, memo)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ���� ������ ������ ���ݿ������� �����մϴ�.
' - ���� ������ ����: "�ӽ�����", "�������", "���۽���"
' - ���ݿ������� �����ϸ� ���� ������ȣ(mgtKey)�� ������ �� �ֽ��ϴ�.
' - https://docs.popbill.com/cashbill/vb/api#Delete
'=========================================================================
Private Sub btnDelete_sub_Click()
    Dim Response As PBResponse
    
    Set Response = CashbillService.Delete(txtCorpNum.Text, txtMgtKey.Text)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ��� ���ݿ����� �����͸� �˺��� ����� ���ÿ� �����Ͽ� "����Ϸ�" ���·� ó���մϴ�.
' - ���ݿ����� ����û ���� ��å : https://docs.popbill.com/cashbill/ntsSendPolicy?lang=vb
' - https://docs.popbill.com/cashbill/vb/api#RevokeRegistIssue
'=========================================================================
Private Sub btnRevokeRegistIssue_Click()
    Dim Response As PBResponse
    Dim orgConfirmNum As String
    Dim orgTradeDate As String
    Dim smssendYN As Boolean
    Dim memo As String
    
    '�������ݿ����� ���ι�ȣ
    orgConfirmNum = "TB0000037"
    
    '�������ݿ����� �ŷ�����
    orgTradeDate = "20220101"
    
    '����ȳ����� ���ۿ���
    smssendYN = False
    
    '�޸�
    memo = "������ݿ����� ��ù���"
    
    Set Response = CashbillService.RevokeRegistIssue(txtCorpNum.Text, txtMgtKey.Text, orgConfirmNum, orgTradeDate, smssendYN, memo, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message + vbCrLf + "����û ���ι�ȣ : " + Response.confirmNum + vbCrLf + "�ŷ����� : " + Response.tradeDate)
End Sub

'=========================================================================
' 1���� ��� ���ݿ����� �����͸� �˺��� ����� ���ÿ� �����Ͽ� "����Ϸ�" ���·� ó���մϴ�.
' - ���ݿ����� ����û ���� ��å : https://docs.popbill.com/cashbill/ntsSendPolicy?lang=dotnet
' - https://docs.popbill.com/cashbill/vb/api#RevokeRegistIssue
'=========================================================================
Private Sub btnRegistIssue_part_Click()
    Dim Response As PBResponse
    Dim orgConfirmNum As String
    Dim orgTradeDate As String
    Dim smssendYN As Boolean
    Dim memo As String
    Dim isPartCancel As Boolean
    Dim cancelType As Integer
    Dim supplyCost As String
    Dim tax As String
    Dim serviceFee As String
    Dim totalAmount As String
    
    '�������ݿ����� ���ι�ȣ
    orgConfirmNum = "TB0000037"
    
    '�������ݿ����� �ŷ�����
    orgTradeDate = "20220101"
    
    '����ȳ����� ���ۿ���
    smssendYN = False
    
    '�޸�
    memo = "������ݿ����� ��ù���"
    
    '�κ���� ����, True-�κ����/False-��ü���
    isPartCancel = True
    
    '��һ���, 1-�ŷ����, 2-�����߱����, 3-��Ÿ
    cancelType = 1
    
    '[���] ���ް���
    supplyCost = "7000"
    
    '[���] ����
    tax = "700"
    
    '[���] �����
    serviceFee = "0"
    
    '[���] �հ�ݾ�
    totalAmount = "7700"
    
    Set Response = CashbillService.RevokeRegistIssue(txtCorpNum.Text, txtMgtKey.Text, orgConfirmNum, orgTradeDate, smssendYN, memo, txtUserID.Text, _
        isPartCancel, cancelType, supplyCost, tax, serviceFee, totalAmount)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message + vbCrLf + "����û ���ι�ȣ : " + Response.confirmNum + vbCrLf + "�ŷ����� : " + Response.tradeDate)
End Sub

'=========================================================================
' ����û ���� ���� "����Ϸ�" ������ ���ݿ������� "�������"�ϰ� ����û �Ű� ��󿡼� �����մϴ�.
' - Delete(����)�Լ��� ȣ���Ͽ� "�������" ������ ���ݿ������� �����ϸ�, ������ȣ ������ �����մϴ�.
' - https://docs.popbill.com/cashbill/vb/api#CancelIssue
'=========================================================================
Private Sub btnCancelIssue_Click()
    Dim Response As PBResponse
    Dim memo As String
    
    '�޸�
    memo = "���� ��� �޸�"
    
    Set Response = CashbillService.CancelIssue(txtCorpNum.Text, txtMgtKey.Text, memo)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ���� ������ ������ ���ݿ������� �����մϴ�.
' - ���� ������ ����: "�ӽ�����", "�������", "���۽���"
' - ���ݿ������� �����ϸ� ���� ������ȣ(mgtKey)�� ������ �� �ֽ��ϴ�.
' - https://docs.popbill.com/cashbill/vb/api#Delete
'=========================================================================
Private Sub btnDelete_Click()
    Dim Response As PBResponse
    
    Set Response = CashbillService.Delete(txtCorpNum.Text, txtMgtKey.Text)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ���ݿ����� 1���� ���� �� ��������� Ȯ���մϴ�.
' - https://docs.popbill.com/cashbill/vb/api#GetInfo
'=========================================================================
Private Sub btnGetInfo_Click()
    Dim cbInfo As PBCbInfo
    Dim tmp As String
    
    Set cbInfo = CashbillService.GetInfo(txtCorpNum.Text, txtMgtKey.Text)
     
    If cbInfo Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "itemKey (�˺���ȣ) : " + cbInfo.itemKey + vbCrLf
    tmp = tmp + "mgtKey (������ȣ) : " + cbInfo.mgtKey + vbCrLf
    tmp = tmp + "tradeDate (�ŷ�����) : " + cbInfo.tradeDate + vbCrLf
    tmp = tmp + "tradeType (��������) : " + cbInfo.tradeType + vbCrLf
    tmp = tmp + "tradeUsage (�ŷ�����) : " + cbInfo.tradeUsage + vbCrLf
    tmp = tmp + "tradeOpt (�ŷ�����) : " + cbInfo.tradeOpt + vbCrLf
    tmp = tmp + "taxationType (��������) : " + cbInfo.taxationType + vbCrLf
    tmp = tmp + "totalAmount (�ŷ��ݾ�) : " + cbInfo.totalAmount + vbCrLf
    tmp = tmp + "issueDT (�����Ͻ�) : " + cbInfo.issueDT + vbCrLf
    tmp = tmp + "regDT (����Ͻ�) : " + cbInfo.regDT + vbCrLf
    tmp = tmp + "stateMemo (���¸޸�) : " + cbInfo.stateMemo + vbCrLf
    tmp = tmp + "stateCode (�����ڵ�) : " + CStr(cbInfo.stateCode) + vbCrLf
    tmp = tmp + "stateDT (���º����Ͻ�) : " + cbInfo.stateDT + vbCrLf
    tmp = tmp + "identityNum (�ĺ���ȣ) : " + cbInfo.identityNum + vbCrLf
    tmp = tmp + "itemName (�ֹ���ǰ��) : " + cbInfo.itemName + vbCrLf
    tmp = tmp + "customerName (�ֹ��ڸ�) : " + cbInfo.customerName + vbCrLf
    tmp = tmp + "confirmNum (����û���ι�ȣ) : " + cbInfo.confirmNum + vbCrLf
    tmp = tmp + "orgConfirmNum (���� ���ݿ����� ����û���ι�ȣ) : " + cbInfo.orgConfirmNum + vbCrLf
    tmp = tmp + "orgTradeDate (���� ���ݿ����� �ŷ�����) : " + cbInfo.orgTradeDate + vbCrLf
    tmp = tmp + "ntssendDT (����û �����Ͻ�) : " + cbInfo.ntssendDT + vbCrLf
    tmp = tmp + "ntsresultDT (����û ó����� �����Ͻ�) : " + cbInfo.ntsresultDT + vbCrLf
    tmp = tmp + "ntsresultCode (����û ó����� �����ڵ�) : " + cbInfo.ntsresultCode + vbCrLf
    tmp = tmp + "ntsresultMessage (����û ó����� �޽���) : " + cbInfo.ntsresultMessage + vbCrLf
    tmp = tmp + "printYN (�μ⿩��) : " + CStr(cbInfo.printYN) + vbCrLf
    
    MsgBox tmp
End Sub

'=========================================================================
' �ټ����� ���ݿ����� ���� �� ��� ������ Ȯ���մϴ�. (1ȸ ȣ�� �� �ִ� 1,000�� Ȯ�� ����)
' - https://docs.popbill.com/cashbill/vb/api#GetInfos
'=========================================================================
Private Sub btnGetInfos_Click()
    Dim resultList As Collection
    Dim KeyList As New Collection
    Dim tmp As String
    Dim info As PBCbInfo
    
    '���ݿ����� ������ȣ�迭, �ִ� 1000��
    KeyList.Add "20220101-001"
    KeyList.Add "20220101-002"
    KeyList.Add "20220101-003"
    KeyList.Add "20220101-004"
    
    Set resultList = CashbillService.GetInfos(txtCorpNum.Text, KeyList)
     
    If resultList Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "itemKey (�˺���ȣ) | mgtKey (������ȣ) | tradeDate (�ŷ�����) | tradeType (��������) | tradeUsage (�ŷ�����) | tradeOpt (�ŷ�����) |  " + _
          "taxationType (��������) | totalAmount (�ŷ��ݾ�) | issueDT (�����Ͻ�) | regDT (����Ͻ�) | stateMemo (���¸޸�) | stateCode (�����ڵ�)  " + _
          "stateDT (���º����Ͻ�) | identityNum (�ĺ���ȣ) | itemName (�ֹ���ǰ��) | customerName (�ֹ��ڸ�) | confirmNum (����û���ι�ȣ)  " + _
          "orgConfirmNum (���� ���ݿ����� ����û���ι�ȣ) | orgTradeDate (���� ���ݿ����� �ŷ�����) | ntssendDT (����û �����Ͻ�)  " + _
          "ntsresultDT (����û ó����� �����Ͻ�) | ntsresultCode (����û ó����� �����ڵ�) | ntsresultMessage (����û ó����� �޽���) | printYN (�μ⿩��) " + vbCrLf + vbCrLf
    
    For Each info In resultList
        tmp = tmp + info.itemKey + " | " + info.mgtKey + " | " + info.tradeDate + " | " + info.tradeType + " | " + info.tradeUsage + " | " + info.tradeOpt + " | " + info.taxationType + " | "
        tmp = tmp + info.totalAmount + " | " + info.issueDT + " | " + info.regDT + " | " + info.stateMemo + " | " + CStr(info.stateCode) + " | " + info.stateDT + " | " + info.identityNum + " | "
        tmp = tmp + info.itemName + " | " + info.customerName + " | " + info.confirmNum + " | " + info.orgConfirmNum + " | " + info.orgTradeDate + " | " + info.ntssendDT + " | " + info.ntsresultDT + " | "
        tmp = tmp + info.ntsresultCode + " | " + info.ntsresultMessage + " | " + CStr(info.printYN) + vbCrLf
                
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' ���ݿ����� 1���� �������� Ȯ���մϴ�.
' - https://docs.popbill.com/cashbill/vb/api#GetDetailInfo
'=========================================================================
Private Sub btnGetDetailInfo_Click()
    Dim cbDetailInfo As PBCashbill
    Dim tmp As String
    
    Set cbDetailInfo = CashbillService.GetDetailInfo(txtCorpNum.Text, txtMgtKey.Text)
     
    If cbDetailInfo Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "mgtKey (������ȣ) : " + cbDetailInfo.mgtKey + vbCrLf
    tmp = tmp + "confirmNum (����û���ι�ȣ) : " + cbDetailInfo.confirmNum + vbCrLf
    tmp = tmp + "tradeDate (�ŷ�����) : " + cbDetailInfo.tradeDate + vbCrLf
    tmp = tmp + "tradeUsage (�ŷ�����) : " + cbDetailInfo.tradeUsage + vbCrLf
    tmp = tmp + "tradeOpt (�ŷ�����) : " + cbDetailInfo.tradeOpt + vbCrLf
    tmp = tmp + "tradeType (��������) : " + cbDetailInfo.tradeType + vbCrLf
    tmp = tmp + "taxationType (��������) : " + cbDetailInfo.taxationType + vbCrLf
    tmp = tmp + "supplyCost (���ް���) : " + cbDetailInfo.supplyCost + vbCrLf
    tmp = tmp + "tax (�ΰ���) : " + cbDetailInfo.tax + vbCrLf
    tmp = tmp + "serviceFee (�����) : " + cbDetailInfo.serviceFee + vbCrLf
    tmp = tmp + "totalAmount (�ŷ��ݾ�) : " + cbDetailInfo.totalAmount + vbCrLf
    tmp = tmp + "orgConfirmNum (�������ݿ����� ����û���ι�ȣ) : " + cbDetailInfo.orgConfirmNum + vbCrLf
    tmp = tmp + "orgTradeDate (�������ݿ����� �ŷ�����) : " + cbDetailInfo.orgTradeDate + vbCrLf
    tmp = tmp + "cancelType (��һ���) : " + CStr(cbDetailInfo.cancelType) + vbCrLf
    tmp = tmp + "franchiseCorpNum (������ ����ڹ�ȣ) : " + cbDetailInfo.franchiseCorpNum + vbCrLf
    tmp = tmp + "franchiseTaxRegID (������ ������� �ĺ���ȣ) : " + cbDetailInfo.franchiseTaxRegID + vbCrLf
    tmp = tmp + "franchiseCorpName (������ ��ȣ) : " + cbDetailInfo.franchiseCorpName + vbCrLf
    tmp = tmp + "franchiseCEOName (������ ��ǥ�� ����) : " + cbDetailInfo.franchiseCEOName + vbCrLf
    tmp = tmp + "franchiseAddr (������ �ּ�) : " + cbDetailInfo.franchiseAddr + vbCrLf
    tmp = tmp + "franchiseTEL (������ ��ȭ��ȣ) : " + cbDetailInfo.franchiseTEL + vbCrLf
    tmp = tmp + "identityNum (�ĺ���ȣ) : " + cbDetailInfo.identityNum + vbCrLf
    tmp = tmp + "customerName (�ֹ��ڸ�) : " + cbDetailInfo.customerName + vbCrLf
    tmp = tmp + "itemName (�ֹ���ǰ��) : " + cbDetailInfo.itemName + vbCrLf
    tmp = tmp + "orderNumber (�ֹ���ȣ) : " + cbDetailInfo.orderNumber + vbCrLf
    tmp = tmp + "email (�ֹ��� �̸���) : " + cbDetailInfo.email + vbCrLf
    tmp = tmp + "hp (�ֹ��� �޴���) : " + cbDetailInfo.hp + vbCrLf
    
    tmp = tmp + "smssendYN (����ȳ����� ���ۿ���) : " + CStr(cbDetailInfo.smssendYN) + vbCrLf
        
    MsgBox tmp
End Sub

'=========================================================================
' �˻����ǿ� �ش��ϴ� ���ݿ������� ��ȸ�մϴ�. (��ȸ�Ⱓ ���� : �ִ� 6����)
' - https://docs.popbill.com/cashbill/vb/api#Search
'=========================================================================
Private Sub btnSearch_Click()
    Dim cbSearchList As PBCBSearchList
    Dim DType As String
    Dim SDate As String
    Dim EDate As String
    Dim state As New Collection
    Dim tradeType As New Collection
    Dim tradeUsage As New Collection
    Dim tradeOpt As New Collection
    Dim taxationType As New Collection
    Dim Page As Integer
    Dim PerPage As Integer
    Dim Order As String
    Dim QString As String
    Dim franchiseTaxRedID As String
    
    '[�ʼ�] ��������, R-�������, T-�ŷ����� I-�����Ͻ�
    DType = "T"
    
    '[�ʼ�] ��������, ����(yyyyMMdd)
    SDate = "20220101"
    
    '[�ʼ�] ��������, ����(yyyyMMdd)
    EDate = "20220130"
    
    '�����ڵ� �迭, �̱���� ��ü ������ȸ, �����ڵ�(stateCode)�� 3�ڸ��� �迭, 2,3��° �ڸ��� ���ϵ�ī�� ����
    state.Add "3**"
    state.Add "4**"
    
    '�������� �迭, N-�Ϲ� ���ݿ�����, C-��� ���ݿ�����
    tradeType.Add "N"
    tradeType.Add "C"
        
    '�ŷ����� �迭, P-�ҵ������, C-����������
    tradeUsage.Add "P"
    tradeUsage.Add "C"
    
    '�ŷ����� �迭, N-�Ϲ�, B-��������, T-���߱���
    tradeOpt.Add "N"
    tradeOpt.Add "B"
    tradeOpt.Add "T"
    
    '�������� �迭, T-����, N-�����
    taxationType.Add "T"
    taxationType.Add "N"
                
    '������ ��ȣ, �⺻�� 1
    Page = 1
    
    '�������� ��ϰ���, �⺻�� 500
    PerPage = 30
    
    '���Ĺ��� D-��������(�⺻��), A-��������
    Order = "D"
    
    '���ݿ����� �ĺ���ȣ ��ȸ, �̱���� ��ü��ȸ
    QString = ""
    
    '������ ������� ��ȣ, �̱���� ��ü��ȸ
    '�� �ټ��� �˻��� �޸�(",")�� ����. ��) 1234,1000
    franchiseTaxRedID = ""
    
    
    Set cbSearchList = CashbillService.Search(txtCorpNum.Text, DType, SDate, EDate, state, tradeType, tradeUsage, taxationType, Page, PerPage, Order, QString, tradeOpt, franchiseTaxRedID)
     
    If cbSearchList Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = "code (�����ڵ�) : " + CStr(cbSearchList.code) + vbCrLf
    tmp = tmp + "total (�˻���� �Ǽ�) : " + CStr(cbSearchList.total) + vbCrLf
    tmp = tmp + "perPage (�������� �˻�����) : " + CStr(cbSearchList.PerPage) + vbCrLf
    tmp = tmp + "pageNum (������ ��ȣ) : " + CStr(cbSearchList.pageNum) + vbCrLf
    tmp = tmp + "pageCount (������ ����) : " + CStr(cbSearchList.pageCount) + vbCrLf
    tmp = tmp + "message (����޽���) : " + cbSearchList.message + vbCrLf + vbCrLf + vbCrLf
    
    tmp = "itemKey (�˺���ȣ) | mgtKey (������ȣ) | tradeDate (�ŷ�����) | tradeType (��������) | tradeUsage (�ŷ�����) | tradeOpt (�ŷ�����) |  " + _
         "taxationType (��������) | totalAmount (�ŷ��ݾ�) | issueDT (�����Ͻ�) | regDT (����Ͻ�) | stateMemo (���¸޸�) | stateCode (�����ڵ�)  " + _
         "stateDT (���º����Ͻ�) | identityNum (�ĺ���ȣ) | itemName (�ֹ���ǰ��) | customerName (�ֹ��ڸ�) | confirmNum (����û���ι�ȣ)  " + _
         "orgConfirmNum (���� ���ݿ����� ����û���ι�ȣ) | orgTradeDate (���� ���ݿ����� �ŷ�����) | ntssendDT (����û �����Ͻ�)  " + _
         "ntsresultDT (����û ó����� �����Ͻ�) | ntsresultCode (����û ó����� �����ڵ�) | ntsresultMessage (����û ó����� �޽���) | printYN (�μ⿩��) " + vbCrLf + vbCrLf
          
    Dim info As PBCbInfo
    
    For Each info In cbSearchList.list
        tmp = tmp + info.itemKey + " | "
        tmp = tmp + info.mgtKey + " | "
        tmp = tmp + info.tradeDate + " | "
        tmp = tmp + info.tradeType + " | "
        tmp = tmp + info.tradeUsage + " | "
        tmp = tmp + info.tradeOpt + " | "
        tmp = tmp + info.taxationType + " | "
        tmp = tmp + info.totalAmount + " | "
        tmp = tmp + info.issueDT + " | "
        tmp = tmp + info.regDT + " | "
        tmp = tmp + info.stateMemo + " | "
        tmp = tmp + CStr(info.stateCode) + " | "
        tmp = tmp + info.stateDT + " | "
        tmp = tmp + info.identityNum + " | "
        tmp = tmp + info.itemName + " | "
        tmp = tmp + info.customerName + " | "
        tmp = tmp + info.confirmNum + " | "
        tmp = tmp + info.orgConfirmNum + " | "
        tmp = tmp + info.orgTradeDate + " | "
        tmp = tmp + info.ntssendDT + " | "
        tmp = tmp + info.ntsresult + " | "
        tmp = tmp + info.ntsresultDT + " | "
        tmp = tmp + info.ntsresultCode + " | "
        tmp = tmp + info.ntsresultMessage + " | "
        tmp = tmp + CStr(info.printYN) + vbCrLf
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' ���ݿ������� ���¿� ���� �����̷��� Ȯ���մϴ�.
' - https://docs.popbill.com/cashbill/vb/api#GetLogs
'=========================================================================
Private Sub btnGetLogs_Click()
    Dim resultList As Collection
    Dim tmp As String
    Dim log As PBCbLog
    
    Set resultList = CashbillService.GetLogs(txtCorpNum.Text, txtMgtKey.Text)
     
    If resultList Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "DocLogType(�α�Ÿ��) | Log(�̷�����) | ProcType(ó������) | ProcMemo(ó���޸�) | RegDT(����Ͻ�) | IP(������)" + vbCrLf
    
    For Each log In resultList
        tmp = tmp + CStr(log.docLogType) + " | " + log.log + " | " + log.procType + " | " + log.procMemo + " | " + log.regDT + " | " + log.ip + vbCrLf
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' ���ݿ������� ���õ� �ȳ� ������ ������ �մϴ�.
' - https://docs.popbill.com/cashbill/vb/api#SendEmail
'=========================================================================
Private Sub btnSendEmail_Click()
    Dim Response As PBResponse
    Dim receiverEmail As String
    
    '���Ÿ��� �ּ�
    receiverEmail = "test@test.com"
    
    Set Response = CashbillService.SendEmail(txtCorpNum.Text, txtMgtKey.Text, receiverEmail)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ���ݿ������� ���õ� �ȳ� SMS(�ܹ�) ���ڸ� �������ϴ� �Լ���, �˺� ����Ʈ [���ڡ��ѽ�] > [����] > [���۳���] �޴����� ���۰���� Ȯ�� �� �� �ֽ��ϴ�.
' - �˸����� ���۽� ����Ʈ�� �����˴ϴ�. (���۽��н� ȯ��ó��)
' - https://docs.popbill.com/cashbill/vb/api#SendSMS
'=========================================================================
Private Sub btnSendSMS_Click()
    Dim Response As PBResponse
    Dim sendNum As String
    Dim receiveNum As String
    Dim Contents As String
    
    '�߽Ź�ȣ
    sendNum = "07043042991"
    
    '���Ź�ȣ
    receiveNum = "010-111-222"
    
    ' �޽��� ����, �ִ� 90Byte (�ѱ� 45��), ���̸� �ʰ��� ������ �����Ǿ� ���۵˴ϴ�.
    Contents = "�˸� ���� ����, �ִ� 90Byte"
    
    Set Response = CashbillService.SendSMS(txtCorpNum.Text, txtMgtKey.Text, sendNum, receiveNum, Contents)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ���ݿ������� �ѽ��� �����ϴ� �Լ���, �˺� ����Ʈ [���ڡ��ѽ�] > [�ѽ�] > [���۳���] �޴����� ���۰���� Ȯ�� �� �� �ֽ��ϴ�.
' - �ѽ� ���� ��û�� ����Ʈ�� �����˴ϴ�. (���۽��н� ȯ��ó��)
' - https://docs.popbill.com/cashbill/vb/api#SendFAX
'=========================================================================
Private Sub btnSendFAX_Click()
    Dim Response As PBResponse
    Dim sendNum As String
    Dim receiveNum As String
    
    '�߽Ź�ȣ
    sendNum = "07043042991"
    
    '���Ź�ȣ
    receiveNum = "010-111-222"
    
    Set Response = CashbillService.SendFax(txtCorpNum.Text, txtMgtKey.Text, sendNum, receiveNum)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ���ݿ����� ���� ���� �׸� ���� �߼ۼ����� Ȯ���մϴ�.
' - https://docs.popbill.com/cashbill/vb/api#ListEmailConfig
'=========================================================================
Private Sub btnListemailconfig_Click()
    Dim resultList As Collection
    Dim i As Integer
    
    Set resultList = CashbillService.ListEmailConfig(txtCorpNum.Text, txtUserID.Text)
    
    If resultList Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
 
    Dim tmp As String
    
    tmp = "������������(EmailType) | ���ۿ���(SendYN) " + vbCrLf
    
    Dim info As PBEmailConfig
    
    For i = 1 To resultList.Count
        If resultList(i).emailType = "CSH_ISSUE" Then
            tmp = tmp + "������ ���ݿ������� ���� �Ǿ����� �˷��ִ� ���� : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
        
        If resultList(i).emailType = "CSH_CANCEL" Then
            tmp = tmp + "������ ���ݿ������� ������� �Ǿ����� �˷��ִ� ���� : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
    Next
    
    MsgBox tmp

End Sub

'=========================================================================
' ���ݿ����� ���� ���� �׸� ���� �߼ۼ����� �����մϴ�.
' - https://docs.popbill.com/cashbill/vb/api#UpdateEmailConfig
'
' ������������
' CSH_ISSUE : ������ ���ݿ������� ���� �Ǿ����� �˷��ִ� ���� �Դϴ�.
' CSH_CANCEL : ������ ���ݿ������� ������� �Ǿ����� �˷��ִ� ���� �Դϴ�.
'=========================================================================
Private Sub btnUpdateemailconfig_Click()
    Dim Response As PBResponse
    Dim emailType As String
    Dim sendYN As Boolean
    
    '���� ���� ����
    emailType = "CSH_ISSUE"

    '���� ���� (True = ����, False = ������)
    sendYN = True
    
    Set Response = CashbillService.UpdateEmailConfig(txtCorpNum.Text, emailType, sendYN, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ���ݿ����� 1���� �� ���� �������� URL�� ��ȯ�մϴ�.
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
' - https://docs.popbill.com/cashbill/vb/api#GetPopUpURL
'=========================================================================
Private Sub btnGetPopUpURL_Click()
    Dim URL As String
    
    URL = CashbillService.GetPopUpURL(txtCorpNum.Text, txtMgtKey.Text)
    
    If URL = "" Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + URL
    txtURL.Text = URL
End Sub

'=========================================================================
' ���ݿ����� 1���� �� ���� ������(����Ʈ ���, ���� �޴� �� ��ư ����)�� URL�� ��ȯ�մϴ�.
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
' - https://docs.popbill.com/cashbill/vb/api#GetViewURL
'=========================================================================
Private Sub btnGetViewURl_Click()
    Dim URL As String
    
    URL = CashbillService.GetViewURL(txtCorpNum.Text, txtMgtKey.Text)
    
    If URL = "" Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + URL
    txtURL.Text = URL
End Sub

'=========================================================================
' ���ݿ����� 1���� �μ��ϱ� ���� �������� �˾� URL�� ��ȯ�մϴ�.
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
' - https://docs.popbill.com/cashbill/vb/api#GetPrintURL
'=========================================================================
Private Sub btnGetPrintURL_Click()
    Dim URL As String
    
    URL = CashbillService.GetPrintURL(txtCorpNum.Text, txtMgtKey.Text)
    
    If URL = "" Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + URL
    txtURL.Text = URL
End Sub

'=========================================================================
' 1���� ���ݿ����� �μ� �˾� URL�� ��ȯ�մϴ�. (���޹޴��ڿ�)
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
'=========================================================================
Private Sub btnGetEPrintUrl_Click()
    Dim URL As String
    
    URL = CashbillService.GetEPrintURL(txtCorpNum.Text, txtMgtKey.Text)
    
    If URL = "" Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + URL
    txtURL.Text = URL
End Sub

'=========================================================================
' �ټ����� ���ݿ������� �μ��ϱ� ���� �������� �˾� URL�� ��ȯ�մϴ�. (�ִ� 100��)
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
' - https://docs.popbill.com/cashbill/vb/api#GetMassPrintURL
'=========================================================================
Private Sub btnGetMassPrintURL_Click()
    Dim URL As String
    Dim KeyList As New Collection
    
    '������ȣ �迭, �ִ� 100��
    KeyList.Add "20220101-01"
    KeyList.Add "20220101-02"
    KeyList.Add "20220101-03"
    KeyList.Add "20220101-04"
    
    URL = CashbillService.GetMassPrintURL(txtCorpNum.Text, KeyList)
     
    If URL = "" Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + URL
    txtURL.Text = URL
End Sub

'=========================================================================
' �����ڰ� �����ϴ� ���ݿ����� �ȳ� ������ �ϴܿ� ��ư URL �ּҸ� ��ȯ�մϴ�.
' - �Լ� ȣ��� ��ȯ ���� URL���� ��ȿ�ð��� �����ϴ�.
' - https://docs.popbill.com/cashbill/vb/api#GetMailURL
'=========================================================================
Private Sub btnGetMailURL_Click()
    Dim URL As String
    
    URL = CashbillService.GetMailURL(txtCorpNum.Text, txtMgtKey.Text)
    
    If URL = "" Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + URL
    txtURL.Text = URL
End Sub

'=========================================================================
' �˺� ���ݿ����� �ӽù����� �˾� URL�� ��ȯ�մϴ�.
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
' - https://docs.popbill.com/cashbill/vb/api#GetURL
'=========================================================================
Private Sub btnGetURL_TBOX_Click()
    Dim URL As String
    
    URL = CashbillService.GetURL(txtCorpNum.Text, txtUserID.Text, "TBOX")
    
    If URL = "" Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + URL
    txtURL.Text = URL
End Sub

'=========================================================================
' �˺� ���ݿ����� ���๮���� �˾� URL�� ��ȯ�մϴ�.
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
' - https://docs.popbill.com/cashbill/vb/api#GetURL
'=========================================================================
Private Sub btnGetURL_PBOX_Click()
    Dim URL As String
    
    URL = CashbillService.GetURL(txtCorpNum.Text, txtUserID.Text, "PBOX")
    
    If URL = "" Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + URL
    txtURL.Text = URL
End Sub

'=========================================================================
' �˺� ���ݿ����� �ۼ� �˾� URL�� ��ȯ�մϴ�.
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
' - https://docs.popbill.com/cashbill/vb/api#GetURL
'=========================================================================
Private Sub btnGetURL_WRITE_Click()
    Dim URL As String
    
    URL = CashbillService.GetURL(txtCorpNum.Text, txtUserID.Text, "WRITE")
    
    If URL = "" Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + URL
    txtURL.Text = URL
End Sub

Private Sub Form_Load()
    CashbillService.Initialize linkID, SecretKey
    
    '����ȯ�漳����, True-���߿� False-�����
    CashbillService.IsTest = True
    
    '������ū IP���ѱ�� ��뿩��, True-���, False-�̻��, �⺻��(True)
    CashbillService.IPRestrictOnOff = True
    
    '���ýý��� �ð� ��뿩�� True-���, False-�̻��, �⺻��(False)
    CashbillService.UseLocalTimeYN = False
End Sub

