VERSION 5.00
Begin VB.Form frmExample 
   Caption         =   "�˺� ���ݿ����� SDK ����"
   ClientHeight    =   10860
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16170
   LinkTopic       =   "Form1"
   ScaleHeight     =   10860
   ScaleWidth      =   16170
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.Frame Frame7 
      Caption         =   "���ݿ����� ���� ���"
      Height          =   7185
      Left            =   240
      TabIndex        =   12
      Top             =   3480
      Width           =   12330
      Begin VB.Frame Frame9 
         Caption         =   "��ù��� ���μ��� "
         Height          =   2295
         Left            =   1800
         TabIndex        =   46
         Top             =   1440
         Width           =   3135
         Begin VB.CommandButton btnDelete_ 
            Caption         =   "����"
            Height          =   375
            Left            =   1755
            Style           =   1  '�׷���
            TabIndex        =   49
            Top             =   1560
            Width           =   855
         End
         Begin VB.CommandButton btnCanceIssue_ 
            BackColor       =   &H00C0C0FF&
            Caption         =   "�������"
            Height          =   375
            Left            =   480
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
            Left            =   360
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
         Caption         =   " ���� ���� "
         Height          =   2760
         Left            =   5880
         TabIndex        =   32
         Top             =   4125
         Width           =   3210
         Begin VB.CommandButton btnGetEPrintUrl 
            Caption         =   "���޹޴��� �μ� �˾� URL"
            Height          =   390
            Left            =   210
            TabIndex        =   37
            Top             =   1260
            Width           =   2745
         End
         Begin VB.CommandButton btnGetPopUpURL 
            Caption         =   "���� ���� ���� �˾� URL"
            Height          =   390
            Left            =   210
            TabIndex        =   36
            Top             =   390
            Width           =   2745
         End
         Begin VB.CommandButton btnGetPrintURL 
            Caption         =   "�μ� �˾� URL"
            Height          =   390
            Left            =   210
            TabIndex        =   35
            Top             =   825
            Width           =   2745
         End
         Begin VB.CommandButton btnGetMassPrintURL 
            Caption         =   "�ٷ� �μ� �˾� URL"
            Height          =   390
            Left            =   210
            TabIndex        =   34
            Top             =   1710
            Width           =   2745
         End
         Begin VB.CommandButton btnGetMailURL 
            Caption         =   "�̸���(���޹޴���) ��ũ URL"
            Height          =   390
            Left            =   210
            TabIndex        =   33
            Top             =   2160
            Width           =   2745
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   " ��Ÿ URL "
         Height          =   1890
         Left            =   9240
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
         Begin VB.CommandButton btnGetURL_SBOX 
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
         Caption         =   " �ΰ� ����"
         Height          =   2775
         Left            =   2760
         TabIndex        =   24
         Top             =   4125
         Width           =   2895
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
            Caption         =   "���� �����ȸ"
            Height          =   390
            Left            =   195
            TabIndex        =   50
            Top             =   2160
            Width           =   1845
         End
         Begin VB.CommandButton btnGetDetailInfo 
            Caption         =   "���� �� ����"
            Height          =   390
            Left            =   195
            TabIndex        =   23
            Top             =   1710
            Width           =   1845
         End
         Begin VB.CommandButton btnGetLogs 
            Caption         =   "���� �̷�"
            Height          =   390
            Left            =   195
            TabIndex        =   22
            Top             =   1260
            Width           =   1845
         End
         Begin VB.CommandButton btnGetInfos 
            Caption         =   "���� ����(�뷮)"
            Height          =   390
            Left            =   210
            TabIndex        =   21
            Top             =   825
            Width           =   1845
         End
         Begin VB.CommandButton btnGetInfo 
            Caption         =   "���� ����"
            Height          =   390
            Left            =   210
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
         Begin VB.CommandButton btnRevokeRegistIssue_part 
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
      Height          =   2535
      Left            =   240
      TabIndex        =   4
      Top             =   720
      Width           =   15690
      Begin VB.Frame Frame15 
         Caption         =   "��Ʈ�ʰ��� ����Ʈ"
         Height          =   1935
         Left            =   13080
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
         Height          =   1935
         Left            =   10920
         TabIndex        =   54
         Top             =   360
         Width           =   2055
         Begin VB.CommandButton btnPopbillURL_CHRG 
            Caption         =   " ����Ʈ ���� URL"
            Height          =   410
            Left            =   120
            TabIndex        =   58
            Top             =   840
            Width           =   1815
         End
         Begin VB.CommandButton btnGetBalance 
            Caption         =   "�ܿ� ����Ʈ Ȯ��"
            Height          =   410
            Left            =   120
            TabIndex        =   55
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   " ȸ������ ����"
         Height          =   1935
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
         Begin VB.CommandButton btnListCorpInfo 
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
         Height          =   1935
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
         Height          =   1935
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
         Height          =   1935
         Left            =   4440
         TabIndex        =   7
         Top             =   360
         Width           =   2055
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
            Top             =   840
            Width           =   1815
         End
         Begin VB.CommandButton btnUpdateContact 
            Caption         =   "����� ���� ����"
            Height          =   410
            Left            =   120
            TabIndex        =   40
            Top             =   1320
            Width           =   1815
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   " �˺� �⺻ URL "
         Height          =   1935
         Left            =   6600
         TabIndex        =   5
         Top             =   360
         Width           =   2175
         Begin VB.CommandButton btnGetPopbillURL_LOGIN 
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
' - VB6 SDK ����ȯ�� ������� �ȳ� :
' - ������Ʈ ���� : 2017-08-30
' - ���� ������� ����ó : 1600-9854 / 070-4304-2991
' - ���� ������� �̸��� : code@linkhub.co.kr
'
' <�׽�Ʈ �������� �غ����>
' 1) 27, 30�� ���ο� ����� ��ũ���̵�(LinkID)�� ���Ű(SecretKey)��
'    ��ũ��� ���Խ� ���Ϸ� �߱޹��� ���������� �����Ͽ� �����մϴ�.
' 2) �˺� ���߿� ����Ʈ(test.popbill.com)�� ����ȸ������ �����մϴ�.
'=========================================================================


Option Explicit


'=========================================================================
' - ��������(��ũ���̵�, ���Ű)�� ��Ʈ���� ����ȸ���� �ĺ��ϴ�
'   ������ ���Ǵ� ������ ������� �ʵ��� �����Ͻñ� �ٶ��ϴ�.
' - ����� ��ȯ���Ŀ��� ��������(��ũ���̵�, ���Ű)�� ������� �ʽ��ϴ�.
'=========================================================================

'��ũ���̵�
Private Const LinkID = "TESTER"

'���Ű. ���⿡ �����Ͻñ� �ٶ��ϴ�.
Private Const SecretKey = "SwWxqU+0TErBXy/9TVjIPEnI0VTUMMSQZtJf3Ed8q3I="

'���ݿ����� ���� ��ü ����
Private CashbillService As New PBCBService


'=========================================================================
' [����Ϸ�] ������ ���ݿ������� [�������] �մϴ�.
' - ������Ҵ� ����û ���������� �����մϴ�.
' - ������ҵ� ���ݿ������� ����û�� ���۵��� �ʽ��ϴ�.
'=========================================================================

Private Sub btnCanceIssue__Click()
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
' [����Ϸ�] ������ ���ݿ������� [�������] �մϴ�.
' - ������Ҵ� ����û ���������� �����մϴ�.
' - ������ҵ� ���ݿ������� ����û�� ���۵��� �ʽ��ϴ�.
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

Private Sub btnCancelIssue_rev_Click()
    Dim Response As PBResponse
    Dim memo As String
    
    '�޸�
    memo = "���� ��� �޸�"
    
    Set Response = CashbillService.CancelIssue(txtCorpNum.Text, txtMgtKey.Text, memo)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(CashbillService.LastErrCode) + "] " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
End Sub

'=========================================================================
' �˺� ȸ�����̵� �ߺ����θ� Ȯ���մϴ�.
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
' �ش� ������� ��Ʈ�� ����ȸ�� ���Կ��θ� Ȯ���մϴ�.
' - LinkID�� ���������� �����Ǿ� �ִ� ��ũ���̵� ���Դϴ�.
'=========================================================================

Private Sub btnCheckIsMember_Click()
    Dim Response As PBResponse
    
    Set Response = CashbillService.CheckIsMember(txtCorpNum.Text, LinkID)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ���ݿ����� ������ȣ �ߺ����θ� Ȯ���մϴ�.
' - ������ȣ�� 1~24�ڸ��� ����, ���� '-', '_' �������� ������ �� �ֽ��ϴ�.
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
' 1���� ���ݿ������� �����մϴ�.
' - ���ݿ������� �����ϸ� ���� ����������ȣ(mgtKey)�� ������ �� �ֽ��ϴ�.
' - ���������� ���� ���� : [�ӽ�����], [�������]
'=========================================================================

Private Sub btnDelete__Click()
    Dim Response As PBResponse
      
    Set Response = CashbillService.Delete(txtCorpNum.Text, txtMgtKey.Text)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' 1���� ���ݿ������� �����մϴ�.
' - ���ݿ������� �����ϸ� ���� ����������ȣ(mgtKey)�� ������ �� �ֽ��ϴ�.
' - ���������� ���� ���� : [�ӽ�����], [�������]
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
' ����ȸ���� �ܿ�����Ʈ�� Ȯ���մϴ�.
' - ���ݹ���� ��Ʈ�ʰ����� ��� ��Ʈ�� �ܿ�����Ʈ(GetPartnerBalance API)
'   �� ���� Ȯ���Ͻñ� �ٶ��ϴ�.
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
' ����ȸ���� ���ݿ����� API ���� ���������� Ȯ���մϴ�.
'=========================================================================

Private Sub btnGetChargeInfo_Click()
    Dim ChargeInfo As PBChargeInfo
    
    Set ChargeInfo = CashbillService.GetChargeInfo(txtCorpNum.Text)
     
    If ChargeInfo Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = tmp + "unitCost (����ܰ�) : " + ChargeInfo.unitCost + vbCrLf
    tmp = tmp + "chargeMethod (��������) : " + ChargeInfo.chargeMethod + vbCrLf
    tmp = tmp + "rateSystem (��������) : " + ChargeInfo.rateSystem + vbCrLf
    
    MsgBox tmp
End Sub

'=========================================================================
' ���ݿ����� 1���� �������� ��ȸ�մϴ�.
' - �����׸� ���� �ڼ��� ������ "[���ݿ����� API �����Ŵ���] > 4.1.
'   ���ݿ����� ����" �� �����Ͻñ� �ٶ��ϴ�.
'=========================================================================

Private Sub btnGetDetailInfo_Click()

    Dim cbDetailInfo As PBCashbill
   
    
    Set cbDetailInfo = CashbillService.GetDetailInfo(txtCorpNum.Text, txtMgtKey.Text)
     
    If cbDetailInfo Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = tmp + "mgtKey (������ȣ) : " + cbDetailInfo.mgtKey + vbCrLf
    tmp = tmp + "confirmNum (����û���ι�ȣ) : " + cbDetailInfo.confirmNum + vbCrLf
    tmp = tmp + "tradeDate (�ŷ�����) : " + cbDetailInfo.tradeDate + vbCrLf
    tmp = tmp + "tradeUsage (�ŷ�����) : " + cbDetailInfo.tradeUsage + vbCrLf
    tmp = tmp + "tradeType (���ݿ����� ����) : " + cbDetailInfo.tradeType + vbCrLf
    tmp = tmp + "taxationType (��������) : " + cbDetailInfo.taxationType + vbCrLf
    tmp = tmp + "supplyCost (���ް���) : " + cbDetailInfo.supplyCost + vbCrLf
    tmp = tmp + "tax (����) : " + cbDetailInfo.tax + vbCrLf
    tmp = tmp + "serviceFee (�����) : " + cbDetailInfo.serviceFee + vbCrLf
    tmp = tmp + "totalAmount (�ŷ��ݾ�) : " + cbDetailInfo.totalAmount + vbCrLf
    
    tmp = tmp + "franchiseCorpNum (������ ����ڹ�ȣ) : " + cbDetailInfo.franchiseCorpNum + vbCrLf
    tmp = tmp + "franchiseCorpName (������ ��ȣ) : " + cbDetailInfo.franchiseCorpName + vbCrLf
    tmp = tmp + "franchiseCEOName (������ ��ǥ�ڸ�) : " + cbDetailInfo.franchiseCEOName + vbCrLf
    tmp = tmp + "franchiseAddr (������ �ּ�) : " + cbDetailInfo.franchiseAddr + vbCrLf
    tmp = tmp + "franchiseTEL (������ ����ó) : " + cbDetailInfo.franchiseTEL + vbCrLf
    
    tmp = tmp + "identityNum (�ŷ�ó �ĺ���ȣ) : " + cbDetailInfo.identityNum + vbCrLf
    tmp = tmp + "customerName (����) : " + cbDetailInfo.customerName + vbCrLf
    tmp = tmp + "itemName (��ǰ��) : " + cbDetailInfo.itemName + vbCrLf
    tmp = tmp + "orderNumber (�ֹ���ȣ) : " + cbDetailInfo.orderNumber + vbCrLf
    tmp = tmp + "email (�� �̸���) : " + cbDetailInfo.email + vbCrLf
    tmp = tmp + "hp (�� �޴�����ȣ) : " + cbDetailInfo.hp + vbCrLf
    tmp = tmp + "smssendYN (�˸����� ���ۿ���) : " + CStr(cbDetailInfo.smssendYN) + vbCrLf
    
    tmp = tmp + "orgConfirmNum (�������ݿ����� ����û���ι�ȣ) : " + cbDetailInfo.orgConfirmNum + vbCrLf
    tmp = tmp + "orgTradeDate (�������ݿ����� �ŷ�����) : " + cbDetailInfo.orgTradeDate + vbCrLf
    tmp = tmp + "cancelType (��һ���) : " + CStr(cbDetailInfo.cancelType) + vbCrLf
    
    MsgBox tmp
    
End Sub

'=========================================================================
' ���ݿ����� �μ�(���޹޴���) URL�� ��ȯ�մϴ�.
' - URL ������å�� ���� ��ȯ�� URL�� 30���� ��ȿ�ð��� �����ϴ�.
'=========================================================================

Private Sub btnGetEPrintUrl_Click()
    Dim url As String
    
    url = CashbillService.GetEPrintURL(txtCorpNum.Text, txtMgtKey.Text)
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' 1���� ���ݿ����� ����/��� ������ Ȯ���մϴ�.
' - �����׸� ���� �ڼ��� ������ "[���ݿ����� API �����Ŵ���] > 4.2.
'   ���ݿ����� �������� ����"�� �����Ͻñ� �ٶ��ϴ�.
'=========================================================================

Private Sub btnGetInfo_Click()
    Dim cbInfo As PBCbInfo
 
    Set cbInfo = CashbillService.GetInfo(txtCorpNum.Text, txtMgtKey.Text)
     
    If cbInfo Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = tmp + "itemKey (������Ű) : " + cbInfo.itemKey + vbCrLf
    tmp = tmp + "mgtKey (����������ȣ) : " + cbInfo.mgtKey + vbCrLf
    tmp = tmp + "tradeDate (�ŷ�����) : " + cbInfo.tradeDate + vbCrLf
    tmp = tmp + "issueDT (�����Ͻ�) : " + cbInfo.issueDT + vbCrLf
    tmp = tmp + "regDT (����Ͻ�) : " + cbInfo.regDT + vbCrLf
    tmp = tmp + "taxationType (��������) : " + cbInfo.taxationType + vbCrLf
    tmp = tmp + "totalAmount (�ŷ��ݾ�) : " + cbInfo.totalAmount + vbCrLf
    tmp = tmp + "tradeUsage (�ŷ��뵵) : " + cbInfo.tradeUsage + vbCrLf
    tmp = tmp + "tradeType (���ݿ����� ����) : " + cbInfo.tradeType + vbCrLf
    tmp = tmp + "stateCode (�����ڵ�) : " + CStr(cbInfo.stateCode) + vbCrLf
    tmp = tmp + "stateDT (���º����Ͻ�) : " + cbInfo.stateDT + vbCrLf
    
    tmp = tmp + "identityNum (�ŷ�ó �ĺ���ȣ) : " + cbInfo.identityNum + vbCrLf
    tmp = tmp + "itemName (��ǰ��) : " + cbInfo.itemName + vbCrLf
    tmp = tmp + "customerName (����) : " + cbInfo.customerName + vbCrLf
    
    tmp = tmp + "confirmNum (����û���ι�ȣ) : " + cbInfo.confirmNum + vbCrLf
    tmp = tmp + "ntssendDT (����û �����Ͻ�) : " + cbInfo.ntssendDT + vbCrLf
    tmp = tmp + "ntsresultDT (����û ó����� �����Ͻ�) : " + cbInfo.ntsresultDT + vbCrLf
    tmp = tmp + "ntsresultCode (����û ó����� �����ڵ�) : " + cbInfo.ntsresultCode + vbCrLf
    tmp = tmp + "ntsresultMessage (����û ó����� �޽���) : " + cbInfo.ntsresultMessage + vbCrLf
    tmp = tmp + "orgConfirmNum (���� ���ݿ����� ����û ���ι�ȣ) : " + cbInfo.orgConfirmNum + vbCrLf
    tmp = tmp + "orgTradeDate (���� ���ݿ����� �ŷ�����) : " + cbInfo.orgTradeDate + vbCrLf
    
    tmp = tmp + "printYN (�μ⿩��) : " + CStr(cbInfo.printYN) + vbCrLf
   
    MsgBox tmp
    
    
End Sub

'=========================================================================
' �ټ����� ���ݿ����� ����/��� ������ Ȯ���մϴ�.
' - �����׸� ���� �ڼ��� ������ "[���ݿ����� API �����Ŵ���] > 4.2.
'   ���ݿ����� �������� ����"�� �����Ͻñ� �ٶ��ϴ�.
'=========================================================================

Private Sub btnGetInfos_Click()
    Dim resultList As Collection
    Dim KeyList As New Collection
    
    '���ݿ����� ������ȣ �迭, �ִ� 1000��
    KeyList.Add "20161011-01"
    KeyList.Add "20161011-02"
    KeyList.Add "20161011-03"
    KeyList.Add "20161011-04"
    
    Set resultList = CashbillService.GetInfos(txtCorpNum.Text, KeyList)
     
    If resultList Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = "ItemKey | StateCode | TaxType | WriteDate | RegDT" + vbCrLf
    
    Dim info As PBCbInfo
    
    For Each info In resultList
        tmp = tmp + info.itemKey + " | " + CStr(info.stateCode) + " | " + info.taxationType + " | " + info.tradeDate + " | " + info.confirmNum + vbCrLf
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' ���ݿ����� ���� �����̷��� Ȯ���մϴ�.
' - ���� �����̷� Ȯ��(GetLogs API) �����׸� ���� �ڼ��� ������
'   "[���ݿ����� API �����Ŵ���] > 3.4.4 ���� �����̷� Ȯ��"
'   �� �����Ͻñ� �ٶ��ϴ�.
'=========================================================================

Private Sub btnGetLogs_Click()
    Dim resultList As Collection
    
    Set resultList = CashbillService.GetLogs(txtCorpNum.Text, txtMgtKey.Text)
     
    If resultList Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = "DocLogType | Log | ProcType | ProcMemo | RegDT | IP" + vbCrLf
    
    Dim log As PBCbLog
    
    For Each log In resultList
        tmp = tmp + CStr(log.docLogType) + " | " + log.log + " | " + log.procType + " | " + log.procMemo + " | " + log.regDT + " | " + log.ip + vbCrLf
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' ���޹޴��� ���ϸ�ũ URL�� ��ȯ�մϴ�.
' - ���ϸ�ũ URL�� ��ȿ�ð��� �������� �ʽ��ϴ�.
'=========================================================================

Private Sub btnGetMailURL_Click()
    Dim url As String
    
    url = CashbillService.GetMailURL(txtCorpNum.Text, txtMgtKey.Text)
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' �ټ����� ���ݿ����� �μ��˾� URL�� ��ȯ�մϴ�.
' ������å���� ���� ��ȯ�� URL�� ��ȿ�ð��� 30���Դϴ�.
'=========================================================================

Private Sub btnGetMassPrintURL_Click()
    Dim url As String
    Dim KeyList As New Collection
    
    KeyList.Add "20161011-01"
    KeyList.Add "20161011-02"
    KeyList.Add "20161011-03"
    KeyList.Add "20161011-04"
    
    url = CashbillService.GetMassPrintURL(txtCorpNum.Text, KeyList)
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    
End Sub

'=========================================================================
' ��Ʈ���� �ܿ�����Ʈ�� Ȯ���մϴ�.
' - ���ݹ���� ���������� ��� ����ȸ�� �ܿ�����Ʈ(GetBalance API)��
'   �̿��Ͻñ� �ٶ��ϴ�.
'=========================================================================

Private Sub btnGetPartnerBalance_Click()
    Dim balance As Double
    
    balance = CashbillService.GetPartnerBalance(txtCorpNum.Text)
    
    If balance < 0 Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "�ܿ�����Ʈ : " + CStr(balance)
    
End Sub

'=========================================================================
' ��Ʈ�� ����Ʈ���� URL�� ��ȯ�մϴ�.
' - URL ������å�� ���� ��ȯ�� URL�� 30���� ��ȿ�ð��� �����ϴ�.
'=========================================================================

Private Sub btnGetPartnerURL_CHRG_Click()
    Dim url As String
    
    url = CashbillService.GetPartnerURL(txtCorpNum.Text, "CHRG")
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' �˺�(www.popbill.com)�� �α��ε� �˺� URL�� ��ȯ�մϴ�.
' - ������å�� ���� ��ȯ�� URL�� 30���� ��ȿ�ð��� �����ϴ�.
'=========================================================================

Private Sub btnGetPopbillURL_LOGIN_Click()
    Dim url As String
    
    url = CashbillService.GetPopbillURL(txtCorpNum.Text, txtUserID.Text, "LOGIN")
    
    If url = "" Then
         MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' 1���� ���ݿ����� ���� �˾� URL�� ��ȯ�մϴ�.
' - ������å���� ���� ��ȯ�� URL�� ��ȿ�ð��� 30���Դϴ�.
'=========================================================================

Private Sub btnGetPopUpURL_Click()
    Dim url As String
    
    url = CashbillService.GetPopUpURL(txtCorpNum.Text, txtMgtKey.Text)
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' 1���� ���ݿ����� �μ��˾� URL�� ��ȯ�մϴ�.
' - ������å���� ���� ��ȯ�� URL�� ��ȿ�ð��� 30���Դϴ�.
'=========================================================================

Private Sub btnGetPrintURL_Click()
    Dim url As String
    
    url = CashbillService.GetPrintURL(txtCorpNum.Text, txtMgtKey.Text)
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' �˺� > ���ݿ����� > ���๮���� �˾� URL�� ��ȯ�մϴ�.
' - ������å���� ���� ��ȯ�� URL�� ��ȿ�ð��� 30���Դϴ�.
'=========================================================================

Private Sub btnGetURL_SBOX_Click()
    Dim url As String
    
    url = CashbillService.GetURL(txtCorpNum.Text, txtUserID.Text, "PBOX")
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' �˺� > ���ݿ����� > �ӽ�(����)������ �˾� URL�� ��ȯ�մϴ�.
' - ������å���� ���� ��ȯ�� URL�� ��ȿ�ð��� 30���Դϴ�.
'=========================================================================

Private Sub btnGetURL_TBOX_Click()
    Dim url As String
    
    url = CashbillService.GetURL(txtCorpNum.Text, txtUserID.Text, "TBOX")
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' �˺� > ���ݿ����� > ���ݿ����� �ۼ� �˾� URL�� ��ȯ�մϴ�.
' - ������å���� ���� ��ȯ�� URL�� ��ȿ�ð��� 30���Դϴ�.
'=========================================================================

Private Sub btnGetURL_WRITE_Click()
    Dim url As String
    
    url = CashbillService.GetURL(txtCorpNum.Text, txtUserID.Text, "WRITE")
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' 1���� �ӽ����� ���ݿ������� ����ó���մϴ�.
' - ������ ���� ���� 5�� ������ ����� ���ݿ������� ������ ���� 2�ÿ� ����û
'   ���۰���� Ȯ���� �� �ֽ��ϴ�.
' - ���ݿ����� ����û ���� ��å�� ���� ������ "[���ݿ����� API �����Ŵ���]
'   > 1.4. ����û ������å"�� �����Ͻñ� �ٶ��ϴ�.
'=========================================================================

Private Sub btnIssue_Click()
    Dim Response As PBResponse
    Dim memo As String
    
    '�޸�
    memo = "����޸�"
    
    Set Response = CashbillService.Issue(txtCorpNum.Text, txtMgtKey.Text, memo)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
    
End Sub


Private Sub btnJoinMember_Click()
    Dim joinData As New PBJoinForm
    Dim Response As PBResponse
    
    '��ũ ���̵�
    joinData.LinkID = LinkID
    
    '����ڹ�ȣ, '-'����, 10�ڸ�
    joinData.CorpNum = "6748500389"
    
    '��ǥ�ڼ���, �ִ� 30��
    joinData.ceoname = "��ǥ�ڼ���"
    
    '��ȣ��, �ִ� 70��
    joinData.corpName = "ȸ����ȣ"
    
    '�ּ�, �ִ� 300��
    joinData.addr = "�ּ�"
    
    '����, �ִ� 40��
    joinData.bizType = "����"
    
    '����, �ִ� 40��
    joinData.bizClass = "����"
    
    '���̵�, 6���̻� 20�� �̸�
    joinData.id = "testkorea_1011"
    
    '��й�ȣ, 6���̻� 20�� �̸�
    joinData.pwd = "pwd_must_be_long_enough"
    
    '����ڸ�, �ִ� 30��
    joinData.ContactName = "����ڼ���"
    
    '����� ����ó, �ִ� 20��
    joinData.ContactTEL = "02-999-9999"
    
    '����� �޴�����ȣ, �ִ� 20��
    joinData.ContactHP = "010-1234-5678"
    
    '����� �ѽ���ȣ, �ִ� 20��
    joinData.ContactFAX = "02-999-9998"
    
    '����� ����, �ִ� 70��
    joinData.ContactEmail = "test@test.com"
    
    Set Response = CashbillService.JoinMember(joinData)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
    
    
End Sub

'=========================================================================
' ����ȸ���� ����� ����� Ȯ���մϴ�.
'=========================================================================

Private Sub btnListContact_Click()
    Dim resultList As Collection
        
    Set resultList = CashbillService.ListContact(txtCorpNum.Text)
     
    If resultList Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = "id | email | hp | personName | searchAllAllowYN | tel | fax | mgrYN | regDT | state" + vbCrLf
    
    Dim info As PBContactInfo
    
    For Each info In resultList
        tmp = tmp + info.id + " | " + info.email + " | " + info.hp + " | " + info.personName + " | " + CStr(info.searchAllAllowYN) _
                + info.tel + " | " + info.fax + " | " + CStr(info.mgrYN) + " | " + info.regDT + " | " + CStr(info.state) + vbCrLf
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' ����ȸ���� ȸ�������� Ȯ���մϴ�.
'=========================================================================

Private Sub btnListCorpInfo_Click()
    Dim CorpInfo As PBCorpInfo
    
    Set CorpInfo = CashbillService.GetCorpInfo(txtCorpNum.Text)
     
    If CorpInfo Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = tmp + "ceoname(��ǥ�ڼ���) : " + CorpInfo.ceoname + vbCrLf
    tmp = tmp + "corpName(ȸ���) : " + CorpInfo.corpName + vbCrLf
    tmp = tmp + "addr(�ּ�) : " + CorpInfo.addr + vbCrLf
    tmp = tmp + "bizType(����) : " + CorpInfo.bizType + vbCrLf
    tmp = tmp + "bizClass(����) : " + CorpInfo.bizClass + vbCrLf
    
    MsgBox tmp
End Sub
'=========================================================================
' ���ݿ����� ���� �������� �׸� ���� ���ۿ��θ� ������� ��ȯ�մϴ�
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
' ���ݿ����� ���� �������� �׸� ���� ���ۿ��θ� �����մϴ�.
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
' ����ȸ�� ����Ʈ ���� URL�� ��ȯ�մϴ�.
' - URL ������å�� ���� ��ȯ�� URL�� 30���� ��ȿ�ð��� �����ϴ�.
'=========================================================================

Private Sub btnPopbillURL_CHRG_Click()
    Dim url As String
    
    url = CashbillService.GetPopbillURL(txtCorpNum.Text, txtUserID.Text, "CHRG")
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' ����ȸ���� ����ڸ� �űԷ� ����մϴ�.
'=========================================================================

Private Sub btnRegistContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    '����� ���̵�, 6�� �̻� 20�� �̸�
    joinData.id = "testkorea_20161010"
    
    '��й�ȣ, 6�� �̻� 20�� �̸�
    joinData.pwd = "test@test.com"
    
    '����ڸ�, �ִ� 30��
    joinData.personName = "����ڸ�"
    
    '����� ����ó
    joinData.tel = "070-1234-1234"
    
    '����� �޴�����ȣ
    joinData.hp = "010-1234-1234"
    
    '����� �����ּ�
    joinData.email = "test@test.com"
    
    '����� �ѽ���ȣ
    joinData.fax = "070-1234-1234"
    
    'ȸ����ȸ ���ѿ���, true-ȸ����ȸ / false-������ȸ
    joinData.searchAllAllowYN = True
    
    '������ ���ѿ���
    joinData.mgrYN = False
        
    Set Response = CashbillService.RegistContact(txtCorpNum.Text, joinData)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' 1���� ���ݿ������� �ӽ����� �մϴ�.
' - [�ӽ�����] ������ ���ݿ������� ����(Issue API)�� ȣ���ؾ߸� ����û��
'   ���۵˴ϴ�.
' - ������ ���� ���� 5�� ������ ����� ���ݿ������� ������ ���� 2�ÿ� ����û
'   ���۰���� Ȯ���� �� �ֽ��ϴ�.
' - ���ݿ����� ����û ���� ��å�� ���� ������ "[���ݿ����� API �����Ŵ���]
'   > 1.4. ����û ������å"�� �����Ͻñ� �ٶ��ϴ�.
' - ������ݿ����� �ۼ���� �ȳ� - http://blog.linkhub.co.kr/702
'=========================================================================

Private Sub btnRegister_Click()
    Dim Cashbill As New PBCashbill
    
    '���ݿ����� ������ȣ, 1~24�ڸ� ����,������������ ����ں��� �ߺ����� �ʵ��� ����
    Cashbill.mgtKey = txtMgtKey.Text
    
    '���ݿ����� ����, [���ΰŷ�, ��Ұŷ�] �� ����
    Cashbill.tradeType = "���ΰŷ�"
    
    '[��Ұŷ��� �ʼ�] ���� ����û���ι�ȣ
    '��������(GetInfo API)�� �����׸��� ����û���ι�ȣ(confirmNum)�� Ȯ���Ͽ� ����
    Cashbill.orgConfirmNum = ""
    
    '[��Ұŷ��� �ʼ�] ���� ���ݿ����� �ŷ�����
    '��������(GetInfo API)�� �����׸��� �ŷ�����(tradeDate)�� Ȯ���Ͽ� ����
    Cashbill.orgTradeDate = ""
    
    '������ ����ڹ�ȣ, "-" ���� 10�ڸ�
    Cashbill.franchiseCorpNum = "1234567890"
    
    '������ ��ȣ��
    Cashbill.franchiseCorpName = "������ ��ȣ"
    
    '������ ��ǥ�� ����
    Cashbill.franchiseCEOName = "������ ��ǥ��"
    
    '������ �ּ�
    Cashbill.franchiseAddr = "������ �ּ�"
    
    '������ ����ó
    Cashbill.franchiseTEL = "070-1234-1234"
    
    '�ŷ�����, [�ҵ������, ����������] �� ����
    Cashbill.tradeUsage = "�ҵ������"
    
    '�ŷ�ó �ĺ���ȣ, �ŷ������� ���� �ۼ�
    '�ҵ������ - �ֹε��/�޴���/ī���ȣ ���簡��
    '���������� - ����ڹ�ȣ/�ֹε��/�޴���/ī���ȣ ���簡��
    Cashbill.identityNum = "0101112222"
    
    '��������, [����, �����] �� ����
    Cashbill.taxationType = "����"
    
    '���ް���
    Cashbill.supplyCost = "10000"
    
    '�����
    Cashbill.serviceFee = "0"
    
    '����
    Cashbill.tax = "1000"
    
    '�հ�ݾ�, ���ް��� + ����� + ����
    Cashbill.totalAmount = "11000"
    
    '�ֹ�����
    Cashbill.customerName = "����"
    
    '��ǰ��
    Cashbill.itemName = "��ǰ��"
    
    '�ֹ���ȣ
    Cashbill.orderNumber = "�ֹ���ȣ"
    
    '���̸���
    Cashbill.email = "test@test.com"
    
    '���޴�����ȣ
    Cashbill.hp = "010-111-222"
    
    '���ݿ����� ���� �˸����� ���ۿ���
    Cashbill.smssendYN = False
    
    Dim Response As PBResponse
    
    Set Response = CashbillService.Register(txtCorpNum.Text, Cashbill)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
    

End Sub

'=========================================================================
' 1���� ���ݿ������� ��ù����մϴ�.
' - ������ ���� ���� 5�� ������ ����� ���ݿ������� ������ ���� 2�ÿ� ����û
'   ���۰���� Ȯ���� �� �ֽ��ϴ�.
' - ���ݿ����� ����û ���� ��å�� ���� ������ "[���ݿ����� API �����Ŵ���]
'   > 1.4. ����û ������å"�� �����Ͻñ� �ٶ��ϴ�.
' - ������ݿ����� �ۼ���� �ȳ� - http://blog.linkhub.co.kr/702
'=========================================================================

Private Sub btnRegistIssue_Click()
    Dim Cashbill As New PBCashbill
    
    '���ݿ����� ������ȣ, 1~24�ڸ� ����,������������ ����ں��� �ߺ����� �ʵ��� ����
    Cashbill.mgtKey = txtMgtKey.Text
    
    '���ݿ����� ����, [���ΰŷ�, ��Ұŷ�] �� ����
    Cashbill.tradeType = "���ΰŷ�"
    
    '[��Ұŷ��� �ʼ�] ���� ����û���ι�ȣ
    '��������(GetInfo API)�� �����׸��� ����û���ι�ȣ(confirmNum)�� Ȯ���Ͽ� ����
    Cashbill.orgConfirmNum = ""
    
    '[��Ұŷ��� �ʼ�] ���� ���ݿ����� �ŷ�����
    '��������(GetInfo API)�� �����׸��� �ŷ�����(tradeDate)�� Ȯ���Ͽ� ����
    Cashbill.orgTradeDate = ""
    
    '������ ����ڹ�ȣ, "-" ���� 10�ڸ�
    Cashbill.franchiseCorpNum = txtCorpNum.Text
    
    '������ ��ȣ��
    Cashbill.franchiseCorpName = "������ ��ȣ"
    
    '������ ��ǥ�� ����
    Cashbill.franchiseCEOName = "������ ��ǥ��"
    
    '������ �ּ�
    Cashbill.franchiseAddr = "������ �ּ�"
    
    '������ ����ó
    Cashbill.franchiseTEL = "070-1234-1234"
    
    '�ŷ�����, [�ҵ������, ����������] �� ����
    Cashbill.tradeUsage = "�ҵ������"
    
    '�ŷ�ó �ĺ���ȣ, �ŷ������� ���� �ۼ�
    '�ҵ������ - �ֹε��/�޴���/ī���ȣ ���簡��
    '���������� - ����ڹ�ȣ/�ֹε��/�޴���/ī���ȣ ���簡��
    Cashbill.identityNum = "0101112222"
    
    '��������, [����, �����] �� ����
    Cashbill.taxationType = "����"
    
    '���ް���
    Cashbill.supplyCost = "10000"
    
    '�����
    Cashbill.serviceFee = "0"
    
    '����
    Cashbill.tax = "1000"
    
    '�հ�ݾ�, ���ް��� + ����� + ����
    Cashbill.totalAmount = "11000"
    
    '�ֹ�����
    Cashbill.customerName = "����"
    
    '��ǰ��
    Cashbill.itemName = "��ǰ��"
    
    '�ֹ���ȣ
    Cashbill.orderNumber = "�ֹ���ȣ"
    
    '���̸���
    Cashbill.email = "test@test.com"
    
    '���޴�����ȣ
    Cashbill.hp = "010-111-222"
    
    '���ݿ����� ���� �˸����� ���ۿ���
    Cashbill.smssendYN = False
        
    Dim Response As PBResponse
    
    Set Response = CashbillService.RegistIssue(txtCorpNum.Text, Cashbill)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
    
End Sub

'=========================================================================
' 1���� ������ݿ������� ��ù����մϴ�.
' - ������ ���� ���� 5�� ������ ����� ���ݿ������� ������ ���� 2�ÿ� ����û
'   ���۰���� Ȯ���� �� �ֽ��ϴ�.
' - ���ݿ����� ����û ���� ��å�� ���� ������ "[���ݿ����� API �����Ŵ���]
'   > 1.4. ����û ������å"�� �����Ͻñ� �ٶ��ϴ�.
' - ������ݿ����� �ۼ���� �ȳ� - http://blog.linkhub.co.kr/702
'=========================================================================

Private Sub btnRevokeRegistIssue_Click()
    Dim Response As PBResponse
    Dim orgConfirmNum As String
    Dim orgTradeDate As String
    
    
    '�������ݿ����� ���ι�ȣ
    orgConfirmNum = "820116333"
    
    '�������ݿ����� �ŷ�����
    orgTradeDate = "20170711"
    
    Set Response = CashbillService.RevokeRegistIssue(txtCorpNum.Text, txtMgtKey.Text, orgConfirmNum, orgTradeDate)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' 1���� (�κ� ������ݿ������� ��ù����մϴ�.
' - ������ ���� ���� 5�� ������ ����� ���ݿ������� ������ ���� 2�ÿ� ����û
'   ���۰���� Ȯ���� �� �ֽ��ϴ�.
' - ���ݿ����� ����û ���� ��å�� ���� ������ "[���ݿ����� API �����Ŵ���]
'   > 1.4. ����û ������å"�� �����Ͻñ� �ٶ��ϴ�.
' - ������ݿ����� �ۼ���� �ȳ� - http://blog.linkhub.co.kr/702
'=========================================================================

Private Sub btnRevokeRegistIssue_part_Click()
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
    orgConfirmNum = "820116333"
    
    '�������ݿ����� �ŷ�����
    orgTradeDate = "20170711"
    
    '�ȳ����� ���ۿ���
    smssendYN = False
    
    '�޸�
    memo = "��ù��� �޸�"
    
    '�κ���ҿ���, True-�κ����, False-��ü���
    isPartCancel = True
    
    '��һ���(Integer), 1-�ŷ����, 2-�����߱����, 3-��Ÿ
    cancelType = 1
    
    '[���] ���ް���
    supplyCost = "3000"
    
    '[���] ����
    tax = "300"
    
    '[���] �����
    serviceFee = "0"
    
    '[���] �հ�ݾ�
    totalAmount = "3300"
    
    Set Response = CashbillService.RevokeRegistIssue(txtCorpNum.Text, txtMgtKey.Text, orgConfirmNum, orgTradeDate, smssendYN, memo, txtUserID.Text, _
        isPartCancel, cancelType, supplyCost, tax, serviceFee, totalAmount)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' �˻������� ����Ͽ� ���ݿ����� ����� ��ȸ�մϴ�.
' - �����׸� ���� �ڼ��� ������ "[���ݿ����� API �����Ŵ���] >
'   4.2. ���ݿ����� �������� ����" �� �����Ͻñ� �ٶ��ϴ�.
'=========================================================================

Private Sub btnSearch_Click()
    Dim cbSearchList As PBCBSearchList
    Dim DType As String
    Dim SDate As String
    Dim EDate As String
    Dim state As New Collection
    Dim tradeType As New Collection
    Dim tradeUsage As New Collection
    Dim taxationType As New Collection
    Dim Page As Integer
    Dim PerPage As Integer
    Dim QString As String
    Dim Order As String
    
    '[�ʼ�] ��������, R-�������, T-�ŷ����� I-��������
    DType = "T"
    
    '[�ʼ�] ��������, ����(yyyyMMdd)
    SDate = "20160901"
    
    '[�ʼ�] ��������, ����(yyyyMMdd)
    EDate = "20161031"
    
    '���ۻ����ڵ� �迭, �̱���� ��ü��ȸ, 2,3��° �ڸ� ���ϵ�ī��(*) ����
    '[����] ���ݿ����� API �����Ŵ��� "5.1. ���ݿ����� �����ڵ�"
    state.Add "2**"
    state.Add "3**"
    state.Add "4**"
    
    '���ݿ����� ���� �迭, N-�Ϲ� ���ݿ�����, C-��� ���ݿ�����
    tradeType.Add "N"
    tradeType.Add "C"
    
    '�ŷ����� �迭, P-�ҵ����, C-��������
    tradeUsage.Add "P"
    tradeUsage.Add "C"
    
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
    
    Set cbSearchList = CashbillService.Search(txtCorpNum.Text, DType, SDate, EDate, state, tradeType, _
                                tradeUsage, taxationType, Page, PerPage, Order, QString)
     
    If cbSearchList Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = "code (�����ڵ�)  : " + CStr(cbSearchList.code) + vbCrLf
    tmp = tmp + "message (����޽���) : " + cbSearchList.message + vbCrLf + vbCrLf + vbCrLf
    tmp = tmp + "total (�� �˻���� �Ǽ�) : " + CStr(cbSearchList.total) + vbCrLf
    tmp = tmp + "perPage (�������� �˻�����) : " + CStr(cbSearchList.PerPage) + vbCrLf
    tmp = tmp + "pageNum (������ ��ȣ) : " + CStr(cbSearchList.pageNum) + vbCrLf
    tmp = tmp + "pageCount (������ ����) : " + CStr(cbSearchList.pageCount) + vbCrLf
    
    tmp = tmp + "ItemKey | MgtKey | TradeDate | TradeUsage | IssueDT | CustomerName | ItemName | IdentityNum | TaxationType | TotalAmount | tradeType | StateCode | TxationType | TradeDate | confirmNum " + vbCrLf
    
    Dim info As PBCbInfo
    
    For Each info In cbSearchList.list
        tmp = tmp + info.itemKey + " | "
        tmp = tmp + info.mgtKey + " | "
        tmp = tmp + info.tradeDate + " | "
        tmp = tmp + info.tradeUsage + " | "
        tmp = tmp + info.issueDT + " | "
        tmp = tmp + info.customerName + " | "
        tmp = tmp + info.itemName + " | "
        tmp = tmp + info.identityNum + " | "
        tmp = tmp + info.taxationType + " | "
        tmp = tmp + info.totalAmount + " | "
        tmp = tmp + info.tradeType + " | "
        tmp = tmp + CStr(info.stateCode) + " | "
        tmp = tmp + info.taxationType + " | "
        tmp = tmp + info.tradeDate + " | "
        tmp = tmp + info.confirmNum + vbCrLf
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' ���� �ȳ������� �������մϴ�.
'=========================================================================

Private Sub btnSendEmail_Click()
    Dim Response As PBResponse
    Dim receiveEmail As String
    
    '���Ÿ����ּ�
    receiveEmail = "test@test.com"
    
    Set Response = CashbillService.SendEmail(txtCorpNum.Text, txtMgtKey.Text, receiveEmail)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ���ݿ������� �ѽ������մϴ�.
' - �ѽ� ���� ��û�� ����Ʈ�� �����˴ϴ�. (���۽��н� ȯ��ó��)
' - ���۳��� Ȯ���� "�˺� �α���" > [���� �ѽ�] > [�ѽ�] > [���۳���]
'   �޴����� ���۰���� Ȯ���� �� �ֽ��ϴ�.
'=========================================================================

Private Sub btnSendFAX_Click()
    Dim Response As PBResponse
    Dim senderNum As String
    Dim receiveNum As String
    
    '�߽Ź�ȣ
    senderNum = "07043042991"
    
    '���Ź�ȣ
    receiveNum = "010-111-222"
    
    Set Response = CashbillService.SendFax(txtCorpNum.Text, txtMgtKey.Text, senderNum, receiveNum)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' �˸����ڸ� �����մϴ�. (�ܹ�/SMS- �ѱ� �ִ� 45��)
' - �˸����� ���۽� ����Ʈ�� �����˴ϴ�. (���۽��н� ȯ��ó��)
' - ���۳��� Ȯ���� "�˺� �α���" > [���� �ѽ�] > [���۳���] �ǿ���
'   ���۰���� Ȯ���� �� �ֽ��ϴ�.
'=========================================================================

Private Sub btnSendSMS_Click()
    Dim Response As PBResponse
    Dim senderNum As String
    Dim receiveNum As String
    Dim Contents As String
    
    '�߽Ź�ȣ
    senderNum = "07075103710"
    
    '���Ź�ȣ
    receiveNum = "010-111-222"
    
    '���ڸ޽��� ����, 90Byte�� �ʰ��� ������ �����Ǿ� ���۵�
    Contents = "�˸� ���� ����, �ִ� 90Byte"
      
    Set Response = CashbillService.SendSMS(txtCorpNum.Text, txtMgtKey.Text, senderNum, receiveNum, Contents)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ���ݿ����� ����ܰ��� Ȯ���մϴ�.
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
' 1���� ���ݿ������� �����մϴ�.
' - [�ӽ�����] ������ ���ݿ������� ������ �� �ֽ��ϴ�.
' - ����û�� �Ű�� ���ݿ������� ������ �� ������, ��� ���ݿ������� �����Ͽ�
'   ���ó�� �� �� �ֽ��ϴ�.
' - ������ݿ����� �ۼ���� �ȳ� - http://blog.linkhub.co.kr/702
'=========================================================================

Private Sub btnUpdate_Click()
    Dim Cashbill As New PBCashbill
    Dim Response As PBResponse
    
    '���ݿ����� ������ȣ, 1~24�ڸ� ����,������������ ����ں��� �ߺ����� �ʵ��� ����
    Cashbill.mgtKey = txtMgtKey.Text
    
    '���ݿ����� ����, [���ΰŷ�, ��Ұŷ�] �� ����
    Cashbill.tradeType = "���ΰŷ�"
    
    '������ ����ڹ�ȣ, "-" ���� 10�ڸ�
    Cashbill.franchiseCorpNum = "1234567890"
    
    '������ ��ȣ��
    Cashbill.franchiseCorpName = "������ ��ȣ_����"
    
    '������ ��ǥ�� ����
    Cashbill.franchiseCEOName = "������ ��ǥ��_����"
    
    '������ �ּ�
    Cashbill.franchiseAddr = "������ �ּ�"
    
    '������ ����ó
    Cashbill.franchiseTEL = "070-1234-1234"
    
    '�ŷ�����, [�ҵ������, ����������] �� ����
    Cashbill.tradeUsage = "�ҵ������"
    
    '�ŷ�ó �ĺ���ȣ, �ŷ������� ���� �ۼ�
    '�ҵ������ - �ֹε��/�޴���/ī���ȣ ���簡��
    '���������� - ����ڹ�ȣ/�ֹε��/�޴���/ī���ȣ ���簡��
    Cashbill.identityNum = "01041680206"
    
    '��������, [����, �����] �� ����
    Cashbill.taxationType = "����"
    
    '���ް���
    Cashbill.supplyCost = "10000"
    
    '�����
    Cashbill.serviceFee = "0"
    
    '����
    Cashbill.tax = "1000"
    
    '�հ�ݾ�, ���ް��� + ����� + ����
    Cashbill.totalAmount = "11000"
    
    '�ֹ�����
    Cashbill.customerName = "����"
    
    '��ǰ��
    Cashbill.itemName = "��ǰ��"
    
    '�ֹ���ȣ
    Cashbill.orderNumber = "�ֹ���ȣ"
    
    '���̸���
    Cashbill.email = "test@test.com"
    
    '���޴�����ȣ
    Cashbill.hp = "010-111-222"
    
    '���ݿ����� ���� �˸����� ���ۿ���
    Cashbill.smssendYN = False
    
    Set Response = CashbillService.Update(txtCorpNum.Text, txtMgtKey.Text, Cashbill)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ����ȸ���� ����� ������ �����մϴ�.
'=========================================================================

Private Sub btnUpdateContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    '����� ���̵�
    joinData.id = txtUserID.Text
    
    '����ڸ�
    joinData.personName = "����ڸ�_����"
    
    '����ó
    joinData.tel = "070-1234-1234"
    
    '�޴�����ȣ
    joinData.hp = "010-1234-1234"
    
    '�̸��� �ּ�
    joinData.email = "test@test.com"
    
    '�ѽ���ȣ
    joinData.fax = "070-1234-1234"
    
    '��ü��ȸ����, Ture-ȸ����ȸ, False-������
    joinData.searchAllAllowYN = True
    
    '������ ���ѿ���
    joinData.mgrYN = False

                
    Set Response = CashbillService.UpdateContact(txtCorpNum.Text, joinData, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ����ȸ���� ȸ�������� �����մϴ�
'=========================================================================

Private Sub btnUpdateCorpInfo_Click()
    Dim CorpInfo As New PBCorpInfo
    Dim Response As PBResponse
    
    '��ǥ�ڼ���, �ִ� 30��
    CorpInfo.ceoname = "��ǥ��"
    
    '��ȣ, �ִ� 70��
    CorpInfo.corpName = "��ȣ"
    
    ' �ּ�, �ִ� 300��
    CorpInfo.addr = "����Ư����"
    
    '����, �ִ� 40��
    CorpInfo.bizType = "����"
    
    '����, �ִ� 40��
    CorpInfo.bizClass = "����"
    
    Set Response = CashbillService.UpdateCorpInfo(txtCorpNum.Text, CorpInfo)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(CashbillService.LastErrCode) + vbCrLf + "����޽��� : " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub




Private Sub Form_Load()
    CashbillService.Initialize LinkID, SecretKey
    
    '����ȯ�� ������ True-�׽�Ʈ��, False-�����
    CashbillService.IsTest = True
End Sub

