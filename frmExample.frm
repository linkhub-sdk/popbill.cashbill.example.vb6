VERSION 5.00
Begin VB.Form frmExample 
   Caption         =   "�˺� ���ݿ����� SDK ����"
   ClientHeight    =   8925
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9570
   LinkTopic       =   "Form1"
   ScaleHeight     =   8925
   ScaleWidth      =   9570
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.Frame Frame7 
      Caption         =   "���ݿ����� ���� ���"
      Height          =   6105
      Left            =   120
      TabIndex        =   15
      Top             =   2730
      Width           =   9330
      Begin VB.Frame Frame14 
         Caption         =   " ���� ���� "
         Height          =   2520
         Left            =   3900
         TabIndex        =   38
         Top             =   3405
         Width           =   2970
         Begin VB.CommandButton btnGetEPrintUrl 
            Caption         =   "���޹޴��� �μ� �˾� URL"
            Height          =   390
            Left            =   90
            TabIndex        =   43
            Top             =   1140
            Width           =   2745
         End
         Begin VB.CommandButton btnGetPopUpURL 
            Caption         =   "���� ���� ���� �˾� URL"
            Height          =   390
            Left            =   90
            TabIndex        =   42
            Top             =   270
            Width           =   2745
         End
         Begin VB.CommandButton btnGetPrintURL 
            Caption         =   "�μ� �˾� URL"
            Height          =   390
            Left            =   90
            TabIndex        =   41
            Top             =   705
            Width           =   2745
         End
         Begin VB.CommandButton btnGetMassPrintURL 
            Caption         =   "�ٷ� �μ� �˾� URL"
            Height          =   390
            Left            =   90
            TabIndex        =   40
            Top             =   1590
            Width           =   2745
         End
         Begin VB.CommandButton btnGetMailURL 
            Caption         =   "�̸���(���޹޴���) ��ũ URL"
            Height          =   390
            Left            =   90
            TabIndex        =   39
            Top             =   2040
            Width           =   2745
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   " ��Ÿ URL "
         Height          =   1665
         Left            =   6960
         TabIndex        =   34
         Top             =   3405
         Width           =   2025
         Begin VB.CommandButton btnGetURL_TBOX 
            Caption         =   "�ӽ� ������"
            Height          =   390
            Left            =   90
            TabIndex        =   37
            Top             =   270
            Width           =   1845
         End
         Begin VB.CommandButton btnGetURL_SBOX 
            Caption         =   "���� ������"
            Height          =   390
            Left            =   90
            TabIndex        =   36
            Top             =   705
            Width           =   1845
         End
         Begin VB.CommandButton btnGetURL_WRITE 
            Caption         =   "���� �ۼ�"
            Height          =   390
            Left            =   90
            TabIndex        =   35
            Top             =   1155
            Width           =   1845
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   " �ΰ� ����"
         Height          =   2055
         Left            =   2235
         TabIndex        =   30
         Top             =   3405
         Width           =   1575
         Begin VB.CommandButton btnSendEmail 
            Caption         =   "�̸��� ����"
            Height          =   390
            Left            =   105
            TabIndex        =   33
            Top             =   270
            Width           =   1335
         End
         Begin VB.CommandButton btnSendSMS 
            Caption         =   "���� ����"
            Height          =   390
            Left            =   105
            TabIndex        =   32
            Top             =   705
            Width           =   1335
         End
         Begin VB.CommandButton btnSendFAX 
            Caption         =   "�ѽ� ����"
            Height          =   390
            Left            =   90
            TabIndex        =   31
            Top             =   1170
            Width           =   1335
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   " ���� ���� "
         Height          =   2055
         Left            =   105
         TabIndex        =   25
         Top             =   3405
         Width           =   2025
         Begin VB.CommandButton btnGetDetailInfo 
            Caption         =   "���� �� ����"
            Height          =   390
            Left            =   75
            TabIndex        =   29
            Top             =   1590
            Width           =   1845
         End
         Begin VB.CommandButton btnGetLogs 
            Caption         =   "���� �̷�"
            Height          =   390
            Left            =   75
            TabIndex        =   28
            Top             =   1140
            Width           =   1845
         End
         Begin VB.CommandButton btnGetInfos 
            Caption         =   "���� ����(�뷮)"
            Height          =   390
            Left            =   90
            TabIndex        =   27
            Top             =   705
            Width           =   1845
         End
         Begin VB.CommandButton btnGetInfo 
            Caption         =   "���� ����"
            Height          =   390
            Left            =   90
            TabIndex        =   26
            Top             =   270
            Width           =   1845
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "���ݿ����� ó�� ���μ��� "
         Height          =   2385
         Left            =   2640
         TabIndex        =   20
         Top             =   825
         Width           =   3735
         Begin VB.CommandButton btnIssue 
            BackColor       =   &H00C0C0FF&
            Caption         =   "����"
            Height          =   525
            Left            =   1230
            Style           =   1  '�׷���
            TabIndex        =   45
            Top             =   1140
            Width           =   1020
         End
         Begin VB.CommandButton btnCancelIssue 
            BackColor       =   &H00C0C0FF&
            Caption         =   "�������"
            Height          =   375
            Left            =   1335
            Style           =   1  '�׷���
            TabIndex        =   44
            Top             =   1830
            Width           =   855
         End
         Begin VB.CommandButton btnRegister 
            BackColor       =   &H00C0C0FF&
            Caption         =   "���"
            Height          =   375
            Left            =   1305
            Style           =   1  '�׷���
            TabIndex        =   23
            Top             =   480
            Width           =   855
         End
         Begin VB.CommandButton btnUpdate 
            BackColor       =   &H00C0C0FF&
            Caption         =   "����"
            Height          =   375
            Left            =   2265
            Style           =   1  '�׷���
            TabIndex        =   22
            Top             =   480
            Width           =   855
         End
         Begin VB.CommandButton btnDelete 
            Caption         =   "����"
            Height          =   375
            Left            =   2490
            Style           =   1  '�׷���
            TabIndex        =   21
            Top             =   1830
            Width           =   855
         End
         Begin VB.Line Line1 
            X1              =   1740
            X2              =   1740
            Y1              =   2100
            Y2              =   795
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  '����
            Caption         =   "�ӽ�����"
            Height          =   180
            Left            =   465
            TabIndex        =   24
            Top             =   555
            Width           =   720
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  '�������� ����
            FillColor       =   &H00E0E0E0&
            Height          =   660
            Left            =   300
            Top             =   345
            Width           =   3135
         End
         Begin VB.Line Line3 
            X1              =   2925
            X2              =   2925
            Y1              =   2010
            Y2              =   615
         End
         Begin VB.Line Line2 
            X1              =   1710
            X2              =   3120
            Y1              =   2010
            Y2              =   2010
         End
      End
      Begin VB.CommandButton btnCheckMgtKeyInUse 
         Caption         =   "������ȣ ��뿩�� Ȯ��"
         Height          =   375
         Left            =   5310
         TabIndex        =   19
         Top             =   255
         Width           =   2190
      End
      Begin VB.TextBox txtMgtKey 
         Height          =   330
         Left            =   2460
         TabIndex        =   18
         Top             =   285
         Width           =   2775
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "����������ȣ( MgtKey) : "
         Height          =   180
         Left            =   360
         TabIndex        =   17
         Top             =   360
         Width           =   2115
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " �˺� �⺻ API "
      Height          =   2055
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   9330
      Begin VB.Frame Frame2 
         Caption         =   " ȸ������"
         Height          =   1575
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   1935
         Begin VB.CommandButton btnCheckIsMember 
            Caption         =   "���� ���� Ȯ��"
            Height          =   495
            Left            =   240
            TabIndex        =   16
            Top             =   360
            Width           =   1455
         End
         Begin VB.CommandButton btnJoinMember 
            Caption         =   "ȸ�� ����"
            Height          =   495
            Left            =   240
            TabIndex        =   14
            Top             =   960
            Width           =   1455
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   " ����Ʈ ����"
         Height          =   1575
         Left            =   2160
         TabIndex        =   10
         Top             =   360
         Width           =   2055
         Begin VB.CommandButton btnGetBalance 
            Caption         =   "�ܿ� ����Ʈ Ȯ��"
            Height          =   495
            Left            =   120
            TabIndex        =   12
            Top             =   360
            Width           =   1815
         End
         Begin VB.CommandButton btnUnitCost 
            Caption         =   "��� �ܰ� Ȯ��"
            Height          =   495
            Left            =   120
            TabIndex        =   11
            Top             =   960
            Width           =   1815
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   " ��Ʈ�� ����"
         Height          =   1575
         Left            =   4275
         TabIndex        =   8
         Top             =   360
         Width           =   2535
         Begin VB.CommandButton btnGetPartnerBalance 
            Caption         =   "��Ʈ�� �ܿ� ����Ʈ Ȯ��"
            Height          =   495
            Left            =   120
            TabIndex        =   9
            Top             =   360
            Width           =   2295
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   " ��Ÿ"
         Height          =   1575
         Left            =   6930
         TabIndex        =   5
         Top             =   345
         Width           =   2175
         Begin VB.CommandButton btnGetPopbillURL 
            Caption         =   " �˺� �⺻ URL Ȯ��"
            Height          =   495
            Left            =   120
            TabIndex        =   7
            Top             =   840
            Width           =   1935
         End
         Begin VB.ComboBox cboPopbillTOGO 
            Height          =   300
            Left            =   120
            TabIndex        =   6
            Text            =   "LOGIN"
            Top             =   360
            Width           =   1935
         End
      End
   End
   Begin VB.TextBox txtUserID 
      Height          =   315
      Left            =   4560
      TabIndex        =   3
      Top             =   165
      Width           =   1935
   End
   Begin VB.TextBox txtCorpNum 
      Height          =   315
      Left            =   1335
      TabIndex        =   1
      Text            =   "1231212312"
      Top             =   180
      Width           =   1935
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "�˺����̵� : "
      Height          =   180
      Left            =   3480
      TabIndex        =   2
      Top             =   240
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "����ڹ�ȣ : "
      Height          =   180
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1080
   End
End
Attribute VB_Name = "frmExample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'�������̵�
Private Const linkID = "TESTER"
'���Ű. ���⿡ �����Ͻñ� �ٶ��ϴ�.
Private Const SecretKey = "zxv+w93mHiiQJTIkkiwLavhjmAFDjONY0NTAZjX+Q0s="

Private CashbillService As New PBCBService

Private Sub btn_GetURL_PBOX_Click()
    Dim url As String
    
    url = CashbillService.GetURL(txtCorpNum.Text, txtUserID.Text, "PBOX")
    
    If url = "" Then
         MsgBox ("[" + CStr(CashbillService.LastErrCode) + "] " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub


Private Sub btnCancelIssue_Click()
    Dim Response As PBResponse
  
    
    Set Response = CashbillService.CancelIssue(txtCorpNum.Text, txtMgtKey.Text, "���� ��� �޸�", txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(CashbillService.LastErrCode) + "] " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
End Sub

Private Sub btnCancelIssue_rev_Click()
    Dim Response As PBResponse
       
    Set Response = CashbillService.CancelIssue(txtCorpNum.Text, txtMgtKey.Text, "���� ��� �޸�", txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(CashbillService.LastErrCode) + "] " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
End Sub



Private Sub btnCheckIsMember_Click()
    Dim Response As PBResponse
    
    Set Response = CashbillService.CheckIsMember(txtCorpNum.Text, linkID)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(CashbillService.LastErrCode) + "] " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox (Response.message)
End Sub

Private Sub btnCheckMgtKeyInUse_Click()
    Dim Response As PBResponse
      
    Set Response = CashbillService.CheckMgtKeyInUse(txtCorpNum.Text, txtMgtKey.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(CashbillService.LastErrCode) + "] " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
    
End Sub

Private Sub btnDelete_Click()
    Dim Response As PBResponse
  
    
    Set Response = CashbillService.Delete(txtCorpNum.Text, txtMgtKey.Text, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(CashbillService.LastErrCode) + "] " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
End Sub


Private Sub btnGetBalance_Click()
    Dim balance As Double
    
    balance = CashbillService.GetBalance(txtCorpNum.Text)
    
    If balance < 0 Then
        
        MsgBox ("[" + CStr(CashbillService.LastErrCode) + "] " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "�ܿ�����Ʈ : " + CStr(balance)
    
    
End Sub

Private Sub btnGetDetailInfo_Click()

    Dim cbDetailInfo As PBCashbill
   
    
    Set cbDetailInfo = CashbillService.GetDetailInfo(txtCorpNum.Text, txtMgtKey.Text, txtUserID.Text)
     
    If cbDetailInfo Is Nothing Then
        MsgBox ("[" + CStr(CashbillService.LastErrCode) + "] " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = tmp + "mgtKey : " + cbDetailInfo.mgtKey + vbCrLf
    tmp = tmp + "tradeDate : " + cbDetailInfo.tradeDate + vbCrLf
    tmp = tmp + "tradeUsage : " + cbDetailInfo.tradeUsage + vbCrLf
    tmp = tmp + "tradeType : " + cbDetailInfo.tradeType + vbCrLf
    
    tmp = tmp + "taxationType : " + cbDetailInfo.taxationType + vbCrLf
    tmp = tmp + "supplyCost : " + cbDetailInfo.supplyCost + vbCrLf
    tmp = tmp + "tax : " + cbDetailInfo.tax + vbCrLf
    tmp = tmp + "serviceFee : " + cbDetailInfo.serviceFee + vbCrLf
    tmp = tmp + "totalAmount : " + cbDetailInfo.totalAmount + vbCrLf
    
    tmp = tmp + "franchiseCorpNum : " + cbDetailInfo.franchiseCorpNum + vbCrLf
    tmp = tmp + "franchiseCorpName : " + cbDetailInfo.franchiseCorpName + vbCrLf
    tmp = tmp + "franchiseCEOName : " + cbDetailInfo.franchiseCEOName + vbCrLf
    tmp = tmp + "franchiseAddr : " + cbDetailInfo.franchiseAddr + vbCrLf
    tmp = tmp + "franchiseTEL : " + cbDetailInfo.franchiseTEL + vbCrLf
    
    tmp = tmp + "identityNum : " + cbDetailInfo.identityNum + vbCrLf
    tmp = tmp + "customerName : " + cbDetailInfo.customerName + vbCrLf
    tmp = tmp + "itemName : " + cbDetailInfo.itemName + vbCrLf
    tmp = tmp + "orderNumber : " + cbDetailInfo.orderNumber + vbCrLf
    
    tmp = tmp + "email : " + cbDetailInfo.email + vbCrLf
    tmp = tmp + "hp : " + cbDetailInfo.hp + vbCrLf
    tmp = tmp + "fax : " + cbDetailInfo.fax + vbCrLf
    tmp = tmp + "smssendYN : " + CStr(cbDetailInfo.smssendYN) + vbCrLf
    tmp = tmp + "faxsendYN : " + CStr(cbDetailInfo.faxsendYN) + vbCrLf
    
    tmp = tmp + "confirmNum : " + cbDetailInfo.confirmNum + vbCrLf
    
    tmp = tmp + "orgConfirmNum : " + cbDetailInfo.orgConfirmNum + vbCrLf
    tmp = tmp + "orgTradeDate : " + cbDetailInfo.orgTradeDate + vbCrLf
    
    
    
    
    MsgBox tmp
    
End Sub


Private Sub btnGetEPrintUrl_Click()
    Dim url As String
    
    url = CashbillService.GetEPrintURL(txtCorpNum.Text, txtMgtKey.Text, txtUserID.Text)
    
    If url = "" Then
         MsgBox ("[" + CStr(CashbillService.LastErrCode) + "] " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub


Private Sub btnGetInfo_Click()
    Dim cbInfo As PBCbInfo
 
    Set cbInfo = CashbillService.GetInfo(txtCorpNum.Text, txtMgtKey.Text, txtUserID.Text)
     
    If cbInfo Is Nothing Then
        MsgBox ("[" + CStr(CashbillService.LastErrCode) + "] " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
   
    tmp = tmp + "itemKey : " + cbInfo.itemKey + vbCrLf
    tmp = tmp + "mgtKey : " + cbInfo.mgtKey + vbCrLf
    tmp = tmp + "tradeDate : " + cbInfo.tradeDate + vbCrLf
    tmp = tmp + "issueDT : " + cbInfo.issueDT + vbCrLf
    tmp = tmp + "customerName : " + cbInfo.customerName + vbCrLf
    tmp = tmp + "itemName : " + cbInfo.itemName + vbCrLf
    tmp = tmp + "identityNum : " + cbInfo.identityNum + vbCrLf
    tmp = tmp + "taxationType : " + cbInfo.taxationType + vbCrLf
    tmp = tmp + "totalAmount : " + cbInfo.totalAmount + vbCrLf
    tmp = tmp + "tradeUsage : " + cbInfo.tradeUsage + vbCrLf
    tmp = tmp + "tradeType : " + cbInfo.tradeType + vbCrLf
    tmp = tmp + "stateCode : " + CStr(cbInfo.stateCode) + vbCrLf
    tmp = tmp + "stateDT : " + cbInfo.stateDT + vbCrLf
    tmp = tmp + "printYN : " + CStr(cbInfo.printYN) + vbCrLf
    tmp = tmp + "confirmNum : " + cbInfo.confirmNum + vbCrLf
    tmp = tmp + "orgTradeDate : " + cbInfo.orgTradeDate + vbCrLf
    tmp = tmp + "orgConfirmNum : " + cbInfo.orgConfirmNum + vbCrLf
    tmp = tmp + "ntssendDT : " + cbInfo.ntssendDT + vbCrLf
    tmp = tmp + "ntsresult : " + cbInfo.ntsresult + vbCrLf
    tmp = tmp + "ntsresultDT : " + cbInfo.ntsresultDT + vbCrLf
    tmp = tmp + "ntsresultCode : " + cbInfo.ntsresultCode + vbCrLf
    tmp = tmp + "ntsresultMessage : " + cbInfo.ntsresultMessage + vbCrLf
   
    tmp = tmp + "regDT : " + cbInfo.regDT + vbCrLf
    
    MsgBox tmp
    
    
End Sub

Private Sub btnGetInfos_Click()
    Dim resultList As Collection
    Dim KeyList As New Collection
    
    KeyList.Add "123123"
    KeyList.Add "123123"
    KeyList.Add "123"
    KeyList.Add "123123123"
    
    Set resultList = CashbillService.GetInfos(txtCorpNum.Text, KeyList, txtUserID.Text)
     
    If resultList Is Nothing Then
        MsgBox ("[" + CStr(CashbillService.LastErrCode) + "] " + CashbillService.LastErrMessage)
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

Private Sub btnGetLogs_Click()
    Dim resultList As Collection
    
    Set resultList = CashbillService.GetLogs(txtCorpNum.Text, txtMgtKey.Text, txtUserID.Text)
     
    If resultList Is Nothing Then
        MsgBox ("[" + CStr(CashbillService.LastErrCode) + "] " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = "DocLogType | Log | ProcType | ProcCorpName | ProcMemo | RegDT | IP" + vbCrLf
    
    Dim log As PBCbLog
    
    For Each log In resultList
        tmp = tmp + CStr(log.docLogType) + " | " + log.log + " | " + log.procType + " | " + log.procMemo + " | " + log.regDT + " | " + log.ip + vbCrLf
    Next
    
    MsgBox tmp
End Sub

Private Sub btnGetMailURL_Click()
    Dim url As String
    
    url = CashbillService.GetMailURL(txtCorpNum.Text, txtMgtKey.Text, txtUserID.Text)
    
    If url = "" Then
         MsgBox ("[" + CStr(CashbillService.LastErrCode) + "] " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

Private Sub btnGetMassPrintURL_Click()
    Dim url As String
    Dim KeyList As New Collection
    
    KeyList.Add "123123"
    KeyList.Add "123123"
    KeyList.Add "123"
    KeyList.Add "123123123"
    
    url = CashbillService.GetMassPrintURL(txtCorpNum.Text, KeyList, txtUserID.Text)
     
    If url = "" Then
        MsgBox ("[" + CStr(CashbillService.LastErrCode) + "] " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    
End Sub

Private Sub btnGetPartnerBalance_Click()
    Dim balance As Double
    
    balance = CashbillService.GetPartnerBalance(txtCorpNum.Text)
    
    If balance < 0 Then
        MsgBox ("[" + CStr(CashbillService.LastErrCode) + "] " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "�ܿ�����Ʈ : " + CStr(balance)
    
End Sub

Private Sub btnGetPopbillURL_Click()
    Dim url As String
    
    url = CashbillService.GetPopbillURL(txtCorpNum.Text, txtUserID.Text, cboPopbillTOGO.Text)
    
    If url = "" Then
         MsgBox ("[" + CStr(CashbillService.LastErrCode) + "] " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

Private Sub btnGetPopUpURL_Click()
    Dim url As String
     
    url = CashbillService.GetPopUpURL(txtCorpNum.Text, txtMgtKey.Text, txtUserID.Text)
    
    If url = "" Then
         MsgBox ("[" + CStr(CashbillService.LastErrCode) + "] " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
    
End Sub

Private Sub btnGetPrintURL_Click()
    Dim url As String
    
    url = CashbillService.GetPrintURL(txtCorpNum.Text, txtMgtKey.Text, txtUserID.Text)
    
    If url = "" Then
         MsgBox ("[" + CStr(CashbillService.LastErrCode) + "] " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

Private Sub btnGetURL_SBOX_Click()
    Dim url As String
    
    url = CashbillService.GetURL(txtCorpNum.Text, txtUserID.Text, "SBOX")
    
    If url = "" Then
         MsgBox ("[" + CStr(CashbillService.LastErrCode) + "] " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

Private Sub btnGetURL_TBOX_Click()
    Dim url As String
    
    url = CashbillService.GetURL(txtCorpNum.Text, txtUserID.Text, "TBOX")
    
    If url = "" Then
         MsgBox ("[" + CStr(CashbillService.LastErrCode) + "] " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

Private Sub btnGetURL_WRITE_Click()
    Dim url As String
    
    url = CashbillService.GetURL(txtCorpNum.Text, txtUserID.Text, "WRITE")
    
    If url = "" Then
         MsgBox ("[" + CStr(CashbillService.LastErrCode) + "] " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

Private Sub btnIssue_Click()
    Dim Response As PBResponse
    
    Set Response = CashbillService.Issue(txtCorpNum.Text, txtMgtKey.Text, "����޸�", txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(CashbillService.LastErrCode) + "] " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
    
End Sub


Private Sub btnJoinMember_Click()
    Dim joinData As New PBJoinForm
    Dim Response As PBResponse
    
    joinData.linkID = linkID '���� ���̵�
    joinData.CorpNum = "1231212312" '����ڹ�ȣ "-" ����.
    joinData.CEOName = "��ǥ�ڼ���"
    joinData.CorpName = "ȸ����ȣ"
    joinData.Addr = "�ּ�"
    joinData.ZipCode = "500-100"
    joinData.BizType = "����"
    joinData.BizClass = "����"
    joinData.ID = "userid"      '6�� �̻� 20�� �̸�.
    joinData.PWD = "pwd_must_be_long_enough"    '6�� �̻� 20�� �̸�.
    joinData.ContactName = "����ڼ���"
    joinData.ContactTEL = "02-999-9999"
    joinData.ContactHP = "010-1234-5678"
    joinData.ContactFAX = "02-999-9998"
    joinData.ContactEmail = "test@test.com"
    
    Set Response = CashbillService.JoinMember(joinData)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(CashbillService.LastErrCode) + "] " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox (Response.message)
    
    
End Sub


Private Sub btnRegister_Click()
    Dim Cashbill As New PBCashbill
    
        
    Cashbill.mgtKey = txtMgtKey.Text        '�����ں� ������ȣ �Ҵ�, 1~24�ڸ� ����,������������ �ߺ����� ����.
    Cashbill.tradeType = "���ΰŷ�"         '���ΰŷ� or ��Ұŷ�
    Cashbill.franchiseCorpNum = "1231212312"
    Cashbill.franchiseCorpName = "������ ��ȣ"
    Cashbill.franchiseCEOName = "������ ��ǥ��"
    Cashbill.franchiseAddr = "������ �ּ�"
    Cashbill.franchiseTEL = "070-1234-1234"
    Cashbill.identityNum = "01041680206"
    Cashbill.customerName = "����"
    Cashbill.itemName = "��ǰ��"
    Cashbill.orderNumber = "�ֹ���ȣ"
    Cashbill.email = "test@test.com"
    Cashbill.hp = "111-1234-1234"
    Cashbill.fax = "777-444-3333"
    Cashbill.serviceFee = "0"
    Cashbill.supplyCost = "10000"
    Cashbill.tax = "1000"
    Cashbill.totalAmount = "11000"
    Cashbill.tradeUsage = "�ҵ������"      '�ҵ������ or ����������
    Cashbill.taxationType = "����"          '���� or �����
    
    Cashbill.smssendYN = False
    Cashbill.faxsendYN = False
    
    Dim Response As PBResponse
    
    Set Response = CashbillService.Register(txtCorpNum.Text, Cashbill, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(CashbillService.LastErrCode) + "] " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox (Response.message)
    

End Sub


Private Sub btnSendEmail_Click()
    Dim Response As PBResponse
    
    Set Response = CashbillService.SendEmail(txtCorpNum.Text, txtMgtKey.Text, "test@test.com", txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(CashbillService.LastErrCode) + "] " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
End Sub

Private Sub btnSendFAX_Click()
    Dim Response As PBResponse
    
    Set Response = CashbillService.SendFax(txtCorpNum.Text, txtMgtKey.Text, "07075106766", "111-2222-4444", txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(CashbillService.LastErrCode) + "] " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
End Sub

Private Sub btnSendSMS_Click()
    Dim Response As PBResponse
      
    Set Response = CashbillService.SendSMS(txtCorpNum.Text, txtMgtKey.Text, "07075106766", "111-2222-4444", "���� ���� �ִ� 90Byte", txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(CashbillService.LastErrCode) + "] " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
End Sub


Private Sub btnUnitCost_Click()
    Dim unitCost As Double
    
    unitCost = CashbillService.GetUnitCost(txtCorpNum.Text)
    
    If unitCost < 0 Then
        MsgBox ("[" + CStr(CashbillService.LastErrCode) + "] " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "����ܰ� : " + CStr(unitCost)
End Sub

Private Sub btnUpdate_Click()

   
    Dim Cashbill As New PBCashbill
    
    
    
    Dim Response As PBResponse
    
    Set Response = CashbillService.Update(txtCorpNum.Text, txtMgtKey.Text, Cashbill, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(CashbillService.LastErrCode) + "] " + CashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox (Response.message)
End Sub



Private Sub Form_Load()
    CashbillService.Initialize linkID, SecretKey
    CashbillService.IsTest = True
    
    
    cboPopbillTOGO.AddItem "LOGIN"
    cboPopbillTOGO.AddItem "CHRG"
    cboPopbillTOGO.AddItem "CERT"

  
    
End Sub

