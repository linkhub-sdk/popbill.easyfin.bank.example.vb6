VERSION 5.00
Begin VB.Form PopbillEasyFinBankExample 
   Caption         =   "�˺� ���� ������ȸ API Example"
   ClientHeight    =   10455
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17865
   LinkTopic       =   "Form1"
   ScaleHeight     =   10455
   ScaleWidth      =   17865
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.Frame Frame6 
      Caption         =   "������ȸ ���� API"
      Height          =   6735
      Left            =   240
      TabIndex        =   26
      Top             =   3360
      Width           =   17175
      Begin VB.ListBox searchInfo 
         Height          =   2940
         Left            =   360
         TabIndex        =   44
         Top             =   3360
         Width           =   15135
      End
      Begin VB.TextBox txtJobID 
         Height          =   375
         Left            =   1440
         TabIndex        =   42
         Top             =   2760
         Width           =   2655
      End
      Begin VB.Frame Frame10 
         Caption         =   "������ ����"
         Height          =   2295
         Left            =   12360
         TabIndex        =   30
         Top             =   360
         Width           =   2415
         Begin VB.CommandButton btnGetFlatRateState 
            Caption         =   "������ ���� ���� Ȯ��"
            Height          =   495
            Left            =   120
            TabIndex        =   40
            Top             =   960
            Width           =   2175
         End
         Begin VB.CommandButton btnGetFlatRatePopUpURL 
            Caption         =   "������ ���� ��û URL"
            Height          =   495
            Left            =   120
            TabIndex        =   39
            Top             =   360
            Width           =   2175
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "���°���"
         Height          =   2295
         Left            =   5160
         TabIndex        =   29
         Top             =   360
         Width           =   6735
         Begin VB.CommandButton btnRevokeCloseBankAccount 
            Caption         =   "������ ������û ���"
            Height          =   495
            Left            =   4440
            TabIndex        =   49
            Top             =   960
            Width           =   2175
         End
         Begin VB.CommandButton btnCloseBankAccount 
            Caption         =   "���� ������ ������û"
            Height          =   495
            Left            =   4440
            TabIndex        =   48
            Top             =   360
            Width           =   2175
         End
         Begin VB.CommandButton btnGetBankAccountInfo 
            Caption         =   "�������� Ȯ��"
            Height          =   495
            Left            =   2280
            TabIndex        =   47
            Top             =   360
            Width           =   2055
         End
         Begin VB.CommandButton btnUpdateBankAccount 
            Caption         =   "�������� ����"
            Height          =   495
            Left            =   120
            TabIndex        =   46
            Top             =   960
            Width           =   2055
         End
         Begin VB.CommandButton btnRegistBankAccount 
            Caption         =   "���� ���"
            Height          =   495
            Left            =   120
            TabIndex        =   45
            Top             =   360
            Width           =   2055
         End
         Begin VB.CommandButton btnListBankAccount 
            Caption         =   "���� ��� Ȯ��"
            Height          =   495
            Left            =   2280
            TabIndex        =   38
            Top             =   960
            Width           =   2055
         End
         Begin VB.CommandButton btnGetBankAccountMgtURL 
            Caption         =   "���°��� �˾� URL"
            Height          =   495
            Left            =   120
            TabIndex        =   37
            Top             =   1560
            Width           =   2055
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "�ŷ����� ����"
         Height          =   2295
         Left            =   2520
         TabIndex        =   28
         Top             =   360
         Width           =   2415
         Begin VB.CommandButton btnSearch 
            Caption         =   "�ŷ����� ��ȸ"
            Height          =   495
            Left            =   120
            TabIndex        =   36
            Top             =   360
            Width           =   2175
         End
         Begin VB.CommandButton btnSummary 
            Caption         =   "�ŷ����� ������� ��ȸ"
            Height          =   495
            Left            =   120
            TabIndex        =   35
            Top             =   960
            Width           =   2175
         End
         Begin VB.CommandButton btnSaveMemo 
            Caption         =   "�ŷ����� �޸�����"
            Height          =   495
            Left            =   120
            TabIndex        =   34
            Top             =   1560
            Width           =   2175
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "�ŷ����� ����"
         Height          =   2295
         Left            =   240
         TabIndex        =   27
         Top             =   360
         Width           =   2055
         Begin VB.CommandButton btnListActiveJob 
            Caption         =   "���� ���� ��� Ȯ��"
            Height          =   495
            Left            =   120
            TabIndex        =   33
            Top             =   1560
            Width           =   1815
         End
         Begin VB.CommandButton btnGetJobState 
            Caption         =   "���� ���� Ȯ��"
            Height          =   495
            Left            =   120
            TabIndex        =   32
            Top             =   960
            Width           =   1815
         End
         Begin VB.CommandButton btnRequestJob 
            Caption         =   "���� ��û"
            Height          =   495
            Left            =   120
            TabIndex        =   31
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Label Label4 
         Caption         =   "(�۾����̵�� '���� ��û' ȣ��� ��ȯ�˴ϴ� )"
         Height          =   255
         Left            =   4200
         TabIndex        =   43
         Top             =   2880
         Width           =   4095
      End
      Begin VB.Label Label3 
         Caption         =   "�۾����̵� : "
         Height          =   255
         Left            =   240
         TabIndex        =   41
         Top             =   2880
         Width           =   1095
      End
   End
   Begin VB.CommandButton btnUpdateCorpInfo 
      Caption         =   "ȸ������ ����"
      Height          =   410
      Left            =   14880
      TabIndex        =   5
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Frame Frame15 
      Caption         =   "ȸ������ ����"
      Height          =   1935
      Left            =   14760
      TabIndex        =   6
      Top             =   1080
      Width           =   2240
      Begin VB.CommandButton btnGetCorpInfo 
         Caption         =   "ȸ������ ��ȸ"
         Height          =   410
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.CommandButton btnListContact 
      Caption         =   "����� ��� ��ȸ"
      Height          =   410
      Left            =   9960
      TabIndex        =   3
      Top             =   1920
      Width           =   2055
   End
   Begin VB.CommandButton btnUpdateContact 
      Caption         =   "����� ���� ����"
      Height          =   410
      Left            =   9960
      TabIndex        =   4
      Top             =   2400
      Width           =   2055
   End
   Begin VB.CommandButton btnCheckID 
      Caption         =   "ID �ߺ� Ȯ��"
      Height          =   410
      Left            =   480
      TabIndex        =   2
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   " �˺� �⺻ API "
      Height          =   2535
      Left            =   240
      TabIndex        =   8
      Top             =   720
      Width           =   17175
      Begin VB.Frame Frame2 
         Caption         =   " ȸ������"
         Height          =   1935
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   1935
         Begin VB.CommandButton btnCheckIsMember 
            Caption         =   "���� ���� Ȯ��"
            Height          =   410
            Left            =   120
            TabIndex        =   23
            Top             =   360
            Width           =   1575
         End
         Begin VB.CommandButton btnJoinMember 
            Caption         =   "ȸ�� ����"
            Height          =   410
            Left            =   120
            TabIndex        =   22
            Top             =   1320
            Width           =   1575
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   " ����Ʈ ����"
         Height          =   1935
         Left            =   2160
         TabIndex        =   19
         Top             =   360
         Width           =   2175
         Begin VB.CommandButton btnGetChargeInfo 
            Caption         =   "�������� Ȯ��"
            Height          =   410
            Left            =   120
            TabIndex        =   20
            Top             =   360
            Width           =   1935
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "����� ����"
         Height          =   1935
         Left            =   9600
         TabIndex        =   17
         Top             =   360
         Width           =   2295
         Begin VB.CommandButton btnRegistContact 
            Caption         =   "����� �߰�"
            Height          =   410
            Left            =   120
            TabIndex        =   18
            Top             =   360
            Width           =   2055
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   " �˺� �⺻ URL"
         Height          =   1935
         Left            =   12000
         TabIndex        =   15
         Top             =   360
         Width           =   2295
         Begin VB.CommandButton btnGetAccessURL 
            Caption         =   " �˺� �α��� URL"
            Height          =   410
            Left            =   120
            TabIndex        =   16
            Top             =   360
            Width           =   2055
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "�������� ����Ʈ"
         Height          =   1935
         Left            =   4560
         TabIndex        =   12
         Top             =   360
         Width           =   2295
         Begin VB.CommandButton btnGetBalance 
            Caption         =   "�ܿ�����Ʈ Ȯ��"
            Height          =   410
            Left            =   120
            TabIndex        =   14
            Top             =   360
            Width           =   2055
         End
         Begin VB.CommandButton btnGetChargeURL 
            Caption         =   " ����Ʈ ���� URL"
            Height          =   410
            Left            =   120
            TabIndex        =   13
            Top             =   840
            Width           =   2055
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "��Ʈ�ʰ��� ����Ʈ"
         Height          =   1935
         Left            =   6960
         TabIndex        =   9
         Top             =   360
         Width           =   2535
         Begin VB.CommandButton btnGetPartnerBalance 
            Caption         =   "��Ʈ�� �ܿ�����Ʈ Ȯ��"
            Height          =   410
            Left            =   120
            TabIndex        =   11
            Top             =   360
            Width           =   2295
         End
         Begin VB.CommandButton btnGetPartnerURL_CHRG 
            Caption         =   "����Ʈ ���� URL"
            Height          =   410
            Left            =   120
            TabIndex        =   10
            Top             =   840
            Width           =   2295
         End
      End
   End
   Begin VB.TextBox txtUserID 
      Height          =   315
      Left            =   6120
      TabIndex        =   1
      Text            =   "testkorea"
      Top             =   240
      Width           =   1935
   End
   Begin VB.TextBox txtCorpNum 
      Height          =   315
      Left            =   2400
      TabIndex        =   0
      Text            =   "1234567890"
      Top             =   255
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "�˺�ȸ�� ���̵� : "
      Height          =   180
      Left            =   4560
      TabIndex        =   25
      Top             =   315
      Width           =   1500
   End
   Begin VB.Label Label1 
      Caption         =   "�˺�ȸ�� ����ڹ�ȣ :"
      Height          =   180
      Left            =   360
      TabIndex        =   24
      Top             =   315
      Width           =   1920
   End
End
Attribute VB_Name = "PopbillEasyFinBankExample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================================
'
' �˺� ������ȸ API VB 6.0 SDK Example
'
' - ������Ʈ ���� : 2020-07-20
' - ���� ������� ����ó : 1600-8536 / 070-4304-2991
' - ���� ������� �̸��� : code@linkhub.co.kr
'
'=========================================================================

Option Explicit

'��ũ���̵�
Private Const linkID = "TESTER"

'���Ű. ���⿡ �����Ͻñ� �ٶ��ϴ�.
Private Const SecretKey = "SwWxqU+0TErBXy/9TVjIPEnI0VTUMMSQZtJf3Ed8q3I="

'������ȸ ���� Ŭ���� ����
Private easyFinBankService As New PBEasyFinBankService

'=========================================================================
' ����ϰ��� �ϴ� ���̵��� �ߺ����θ� Ȯ���մϴ�.
' - https://docs.popbill.com/easyfinbank/vb/api#CheckID
'=========================================================================
Private Sub btnCheckID_Click()
    Dim Response As PBResponse
    
    Set Response = easyFinBankService.CheckID(txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(easyFinBankService.LastErrCode) + vbCrLf + "����޽��� : " + easyFinBankService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.Message)
End Sub

'=========================================================================
' ����ڹ�ȣ�� ��ȸ�Ͽ� ����ȸ�� ���Կ��θ� Ȯ���մϴ�.
' - LinkID�� ���������� �����Ǿ� �ִ� ��ũ���̵� ���Դϴ�.
' - https://docs.popbill.com/easyfinbank/vb/api#CheckIsMember
'=========================================================================
Private Sub btnCheckIsMember_Click()
    Dim Response As PBResponse
    
    Set Response = easyFinBankService.CheckIsMember(txtCorpNum.Text, linkID)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(easyFinBankService.LastErrCode) + vbCrLf + "����޽��� : " + easyFinBankService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.Message)
End Sub

'=========================================================================
' ������ ������ ������ ��û�մϴ�.
' - https://docs.popbill.com/easyfinbank/vb/api#CloseBankAccount
'=========================================================================
Private Sub btnCloseBankAccount_Click()
    
    Dim Response As PBResponse
    Dim BankCode As String
    Dim AccountNumber As String
    Dim CloseType As String
    
    ' [�ʼ�] �����ڵ�
    ' �������-0002 / �������-0003 / ��������-0004 /��������-0007 / ��������-0011 / �츮����-0020
    ' SC����-0023 / �뱸����-0031 / �λ�����-0032 / ��������-0034 / ��������-0035 / ��������-0037
    ' �泲����-0039 / �������ݰ�-0045 / ��������-0048 / ��ü��-0071 / KEB�ϳ�����-0081 / ��������-0088 /��Ƽ����-0027
    BankCode = ""
    
    ' [�ʼ�] ���¹�ȣ ������('-') ����
    AccountNumber = ""

    ' ��������, "�Ϲ�", "�ߵ�" �� ���� ����
    ' �Ϲ����� - �̿����� ������ ���Ⱓ���� �̿��� ����
    ' �ߵ����� - ��û�� �������� ����, ������ �ܿ��Ⱓ�� ���ҷ� ���Ǿ� ����Ʈ ȯ�� (���� �̿�Ⱓ �� �ߵ����� �� ���� ȯ��)
    CloseType = "�ߵ�"
    
    Set Response = easyFinBankService.CloseBankAccount(txtCorpNum.Text, BankCode, AccountNumber, CloseType)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(easyFinBankService.LastErrCode) + vbCrLf + "����޽��� : " + easyFinBankService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.Message)
End Sub

'=========================================================================
' �˺� ����Ʈ�� �α��� ���·� ������ �� �ִ� �������� �˾� URL�� ��ȯ�մϴ�.
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
' - https://docs.popbill.com/easyfinbank/vb/api#GetAccessURL
'=========================================================================
Private Sub btnGetAccessURL_Click()
    Dim url As String
           
    url = easyFinBankService.GetAccessURL(txtCorpNum.Text, txtUserID.Text)
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(easyFinBankService.LastErrCode) + vbCrLf + "����޽��� : " + easyFinBankService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' ����ȸ���� �ܿ�����Ʈ�� Ȯ���մϴ�.
' - ���ݹ���� ��Ʈ�ʰ����� ��� ��Ʈ�� �ܿ�����Ʈ(GetPartnerBalance API)�� ���� Ȯ���Ͻñ� �ٶ��ϴ�.
' - https://docs.popbill.com/easyfinbank/vb/api#GetBalance
'=========================================================================
Private Sub btnGetBalance_Click()
    Dim balance As Double
    
    balance = easyFinBankService.GetBalance(txtCorpNum.Text)
    
    If balance < 0 Then
        MsgBox ("�����ڵ� : " + CStr(easyFinBankService.LastErrCode) + vbCrLf + "����޽��� : " + easyFinBankService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "����ȸ�� �ܿ�����Ʈ : " + CStr(balance)
End Sub

'=========================================================================
' �˺��� ��ϵ� ���� ������ Ȯ���մϴ�.
' - https://docs.popbill.com/easyfinbank/vb/api#GetBankAccountInfo
'=========================================================================
Private Sub btnGetBankAccountInfo_Click()
    Dim AccountInfo As PBEasyFinBankAccount
    Dim tmp As String
    Dim BankCode As String
    Dim AccountNumber As String
    
    ' [�ʼ�] �����ڵ�
    ' �������-0002 / �������-0003 / ��������-0004 /��������-0007 / ��������-0011 / �츮����-0020
    ' SC����-0023 / �뱸����-0031 / �λ�����-0032 / ��������-0034 / ��������-0035 / ��������-0037
    ' �泲����-0039 / �������ݰ�-0045 / ��������-0048 / ��ü��-0071 / KEB�ϳ�����-0081 / ��������-0088 /��Ƽ����-0027
    BankCode = ""
    
    ' [�ʼ�] ���¹�ȣ ������('-') ����
    AccountNumber = ""
    
    Set AccountInfo = easyFinBankService.GetBankAccountInfo(txtCorpNum.Text, BankCode, AccountNumber)
     
    If AccountInfo Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(easyFinBankService.LastErrCode) + vbCrLf + "����޽��� : " + easyFinBankService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "bankCode (�����ڵ�) : " + AccountInfo.BankCode + vbCrLf
    tmp = tmp + "accountNumber (���¹�ȣ) : " + AccountInfo.AccountNumber + vbCrLf
    tmp = tmp + "accountName (���º�Ī) : " + AccountInfo.AccountName + vbCrLf
    tmp = tmp + "accountType (��������) : " + AccountInfo.AccountType + vbCrLf
    tmp = tmp + "state (������ ����) : " + CStr(AccountInfo.state) + vbCrLf
    tmp = tmp + "regDT (����Ͻ�) : " + AccountInfo.regDT + vbCrLf
    tmp = tmp + "contractDT (������ ���� �����Ͻ�) : " + AccountInfo.contractDT + vbCrLf
    tmp = tmp + "baseDate (�ڵ����� ������) : " + CStr(AccountInfo.baseDate) + vbCrLf
    tmp = tmp + "useEndDate (������ ���� ��������) : " + AccountInfo.useEndDate + vbCrLf
    tmp = tmp + "contractState (������ ���� ����) : " + CStr(AccountInfo.contractState) + vbCrLf
    tmp = tmp + "closeRequestYN (������ ������û ����) : " + CStr(AccountInfo.closeRequestYN) + vbCrLf
    tmp = tmp + "useRestrictYN (������ ������� ����) : " + CStr(AccountInfo.useRestrictYN) + vbCrLf
    tmp = tmp + "closeOnExpired (������ ����� ��������) : " + CStr(AccountInfo.closeOnExpired) + vbCrLf
    tmp = tmp + "unPaiedYN (�̼��� ���� ����) : " + CStr(AccountInfo.unPaidYN) + vbCrLf
    tmp = tmp + "memo (�޸�) : " + AccountInfo.Memo
    
    MsgBox tmp
End Sub

'=========================================================================
' ���� ���, ���� �� ������ �� �ִ� ���� ���� �˾� URL�� ��ȯ�մϴ�.
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
' - https://docs.popbill.com/easyfinbank/vb/api#GetBankAccountMgtURL
'=========================================================================
Private Sub btnGetBankAccountMgtURL_Click()
    Dim url As String
           
    url = easyFinBankService.GetBankAccountMgtURL(txtCorpNum.Text, txtUserID.Text)
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(easyFinBankService.LastErrCode) + vbCrLf + "����޽��� : " + easyFinBankService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' �˺� ������ȸ API ���� ���������� Ȯ���մϴ�.
' - https://docs.popbill.com/easyfinbank/vb/api#GetChargeInfo
'=========================================================================
Private Sub btnGetChargeInfo_Click()
    Dim ChargeInfo As PBChargeInfo
    Dim tmp As String
    
    Set ChargeInfo = easyFinBankService.GetChargeInfo(txtCorpNum.Text)
     
    If ChargeInfo Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(easyFinBankService.LastErrCode) + vbCrLf + "����޽��� : " + easyFinBankService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "unitCost (�����״ܰ�) : " + ChargeInfo.unitCost + vbCrLf
    tmp = tmp + "chargeMethod (��������) : " + ChargeInfo.chargeMethod + vbCrLf
    tmp = tmp + "rateSystem (��������) : " + ChargeInfo.rateSystem + vbCrLf
    
    MsgBox tmp
End Sub

'=========================================================================
' ����ȸ�� ����Ʈ ������ ���� �������� �˾� URL�� ��ȯ�մϴ�.
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
' - https://docs.popbill.com/easyfinbank/vb/api#GetChargeURL
'=========================================================================
Private Sub btnGetChargeURL_Click()
    Dim url As String
           
    url = easyFinBankService.GetChargeURL(txtCorpNum.Text, txtUserID.Text)
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(easyFinBankService.LastErrCode) + vbCrLf + "����޽��� : " + easyFinBankService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' ����ȸ���� ȸ�������� Ȯ���մϴ�.
' - https://docs.popbill.com/easyfinbank/vb/api#GetCorpInfo
'=========================================================================
Private Sub btnGetCorpInfo_Click()
    Dim CorpInfo As PBCorpInfo
    Dim tmp As String
    
    Set CorpInfo = easyFinBankService.GetCorpInfo(txtCorpNum.Text)
     
    If CorpInfo Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(easyFinBankService.LastErrCode) + vbCrLf + "����޽��� : " + easyFinBankService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "ceoname (��ǥ�ڸ�) : " + CorpInfo.ceoname + vbCrLf
    tmp = tmp + "corpName (��ȣ) : " + CorpInfo.corpName + vbCrLf
    tmp = tmp + "addr (�ּ�) : " + CorpInfo.addr + vbCrLf
    tmp = tmp + "bizType (����) : " + CorpInfo.bizType + vbCrLf
    tmp = tmp + "bizClass (����) : " + CorpInfo.bizClass + vbCrLf
    
    MsgBox tmp
End Sub

'=========================================================================
' ������ȸ ������ ���� ��û �������� �˾� URL�� ��ȯ�մϴ�.
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
' - https://docs.popbill.com/easyfinbank/vb/api#GetFlatRatePopUpURL
'=========================================================================
Private Sub btnGetFlatRatePopUpURL_Click()
    Dim url As String
           
    url = easyFinBankService.GetFlatRatePopUpURL(txtCorpNum.Text, txtUserID.Text)
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(easyFinBankService.LastErrCode) + vbCrLf + "����޽��� : " + easyFinBankService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' ������ȸ ������ ���� ���¸� Ȯ���մϴ�.
' - https://docs.popbill.com/easyfinbank/vb/api#GetFlatRateState
'=========================================================================
Private Sub btnGetFlatRateState_Click()
    Dim flatRateInfo As PBEasyFinBankFlatRate
    Dim tmp As String
    Dim BankCode As String
    Dim AccountNumber As String
    
    '�����ڵ�
    BankCode = "0048"
    
    '�˺��� ��ϵ� ���¹�ȣ
    AccountNumber = "131020538645"
    
    Set flatRateInfo = easyFinBankService.GetFlatRateState(txtCorpNum.Text, BankCode, AccountNumber)
     
    If flatRateInfo Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(easyFinBankService.LastErrCode) + vbCrLf + "����޽��� : " + easyFinBankService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "referencdeID (���¾��̵�) : " + flatRateInfo.referenceID + vbCrLf
    tmp = tmp + "contractDT (������ ���� �����Ͻ�) : " + flatRateInfo.contractDT + vbCrLf
    tmp = tmp + "useEndDate (������ ���� ������) : " + flatRateInfo.useEndDate + vbCrLf
    tmp = tmp + "baseDate (�ڵ����� ������) : " + CStr(flatRateInfo.baseDate) + vbCrLf
    tmp = tmp + "state (������ ���� ����) : " + CStr(flatRateInfo.state) + vbCrLf
    tmp = tmp + "closeRequestYN (���� ������û ����) : " + CStr(flatRateInfo.closeRequestYN) + vbCrLf
    tmp = tmp + "useRestrictYN (���� ������� ����) : " + CStr(flatRateInfo.useRestrictYN) + vbCrLf
    tmp = tmp + "closeOnExpired (���񽺸���� �������� ) : " + CStr(flatRateInfo.closeOnExpired) + vbCrLf
    tmp = tmp + "unPaidYN (�̼��� ���� ����) : " + CStr(flatRateInfo.unPaidYN) + vbCrLf
    
    MsgBox tmp
End Sub

'=========================================================================
' RequestJob(���� ��û)�� ���� ��ȯ ���� �۾����̵��� ���¸� Ȯ���մϴ�.
' - https://docs.popbill.com/easyfinbank/vb/api#GetJobState
'=========================================================================
Private Sub btnGetJobState_Click()
    Dim jobInfo As PBEasyFinBankJobState
    Dim tmp As String
    
    Set jobInfo = easyFinBankService.GetJobState(txtCorpNum.Text, txtJobID.Text)
     
    If jobInfo Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(easyFinBankService.LastErrCode) + vbCrLf + "����޽��� : " + easyFinBankService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "jobID(�۾����̵�) : " + jobInfo.jobID + vbCrLf
    tmp = tmp + "jobState(��������) : " + CStr(jobInfo.jobState) + vbCrLf
    tmp = tmp + "startDate(��������) : " + jobInfo.startDate + vbCrLf
    tmp = tmp + "endDate(��������) : " + jobInfo.endDate + vbCrLf
    tmp = tmp + "errorCode(�����ڵ�) : " + CStr(jobInfo.errorCode) + vbCrLf
    tmp = tmp + "errorReason(�����޽���) : " + jobInfo.errorReason + vbCrLf
    tmp = tmp + "jobStartDT(�۾� �����Ͻ�) : " + jobInfo.jobStartDT + vbCrLf
    tmp = tmp + "jobEndDT(�۾� �����Ͻ�) : " + jobInfo.jobEndDT + vbCrLf
    tmp = tmp + "regDT(���� ��û�Ͻ�) : " + jobInfo.regDT + vbCrLf
    
    MsgBox tmp
End Sub

'=========================================================================
' ����ȸ���� �ܿ�����Ʈ�� Ȯ���մϴ�.
' - ���ݹ���� ��Ʈ�ʰ����� ��� ��Ʈ�� �ܿ�����Ʈ(GetPartnerBalance API)�� ���� Ȯ���Ͻñ� �ٶ��ϴ�.
' - https://docs.popbill.com/easyfinbank/vb/api#GetPartnerBalance
'=========================================================================
Private Sub btnGetPartnerBalance_Click()
    Dim balance As Double
    
    balance = easyFinBankService.GetPartnerBalance(txtCorpNum.Text)
    
    If balance < 0 Then
        MsgBox ("�����ڵ� : " + CStr(easyFinBankService.LastErrCode) + vbCrLf + "����޽��� : " + easyFinBankService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "��Ʈ�� �ܿ�����Ʈ : " + CStr(balance)
End Sub

'=========================================================================
' ��Ʈ�� ����Ʈ ������ ���� �������� �˾� URL�� ��ȯ�մϴ�.
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
' - https://docs.popbill.com/easyfinbank/vb/api#GetPartnerURL
'=========================================================================
Private Sub btnGetPartnerURL_CHRG_Click()
    Dim url As String
           
    url = easyFinBankService.GetPartnerURL(txtCorpNum.Text, "CHRG")
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(easyFinBankService.LastErrCode) + vbCrLf + "����޽��� : " + easyFinBankService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' ����ڸ� ����ȸ������ ����ó���մϴ�.
' - https://docs.popbill.com/easyfinbank/vb/api#JoinMember
'=========================================================================
Private Sub btnJoinMember_Click()
    Dim joinData As New PBJoinForm
    Dim Response As PBResponse
    
    '��ũ ���̵�
    joinData.linkID = linkID
    
    '����ڹ�ȣ, '-'����, 10�ڸ�
    joinData.CorpNum = "1234567890"
    
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
    joinData.id = "userid"
    
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
    
    Set Response = easyFinBankService.JoinMember(joinData)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(easyFinBankService.LastErrCode) + vbCrLf + "����޽��� : " + easyFinBankService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.Message)
End Sub

'=========================================================================
' RequestJob(���� ��û)�� ���� ��ȯ ���� �۾����̵��� ����� Ȯ���մϴ�.
' - ���� ��û �� 1�ð��� ����� ���� ��û���� ���������� ��ȯ���� �ʽ��ϴ�.
' - https://docs.popbill.com/easyfinbank/vb/api#ListActiveJob
'=========================================================================
Private Sub btnListActiveJob_Click()
    Dim jobList As Collection
    Dim tmp As String
    Dim info As PBEasyFinBankJobState
    
    Set jobList = easyFinBankService.ListActiveJob(txtCorpNum.Text)
     
    If jobList Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(easyFinBankService.LastErrCode) + vbCrLf + "����޽��� : " + easyFinBankService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "�۾����̵�(jobID)�� ��ȿ�ð��� 1�ð��Դϴ�" + vbCrLf + vbCrLf
    tmp = tmp + "jobID(�۾����̵�) | jobState(��������) | startDate(��������) | endDate(��������) |" _
            + "errorCode(�����ڵ�) | errorReason(�����޽���) | jobStartDT(�۾� �����Ͻ�) | jobEndDT(�۾� �����Ͻ�) | regDT(���� ��û�Ͻ�) " + vbCrLf
    
    For Each info In jobList
        tmp = tmp + CStr(info.jobID) + " | "
        tmp = tmp + CStr(info.jobState) + " | "
        tmp = tmp + info.startDate + " | "
        tmp = tmp + info.endDate + " | "
        tmp = tmp + CStr(info.errorCode) + " | "
        tmp = tmp + info.errorReason + " | "
        tmp = tmp + info.jobStartDT + " | "
        tmp = tmp + info.jobEndDT + " | "
        tmp = tmp + info.regDT
        tmp = tmp + vbCrLf
    Next
    
    MsgBox tmp
    
    If jobList.count > 0 Then
        txtJobID.Text = jobList.Item(1).jobID
    End If
End Sub

'=========================================================================
' �˺��� ��ϵ� ������� ����� ��ȯ�Ѵ�.
' - https://docs.popbill.com/easyfinbank/vb/api#ListBankAccount
'=========================================================================
Private Sub btnListBankAccount_Click()
    Dim bankAccountList As Collection
    Dim tmp As String
    Dim info As PBEasyFinBankAccount
    
    Set bankAccountList = easyFinBankService.ListBankAccount(txtCorpNum.Text)
     
    If bankAccountList Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(easyFinBankService.LastErrCode) + vbCrLf + "����޽��� : " + easyFinBankService.LastErrMessage)
        Exit Sub
    End If
    
    
    tmp = "accountNumber(���¹�ȣ) | bankCode(�����ڵ�) | accountName(���º�Ī) | accountType(��������) | state(���� ����) | regDT(����Ͻ�) |" _
        + " contractState (������ ���� ����) | closeRequestYN (������ ������û ����) | useRestrictYN (������ ������� ����) | closeOnExpired (������ ����� ��������) | " _
        + " unPaiedYN (�̼��� ���� ����) | memo(�޸�)" + vbCrLf + vbCrLf
    
    For Each info In bankAccountList
        tmp = tmp + info.AccountNumber + " | "
        tmp = tmp + info.BankCode + " | "
        tmp = tmp + info.AccountName + " | "
        tmp = tmp + info.AccountType + " | "
        tmp = tmp + CStr(info.state) + " | "
        tmp = tmp + info.regDT + " | "
        
        tmp = tmp + info.contractDT + " | "
        tmp = tmp + CStr(info.baseDate) + " | "
        tmp = tmp + info.useEndDate + " | "
        tmp = tmp + CStr(info.contractState) + " | "
        tmp = tmp + CStr(info.closeRequestYN) + " | "
        tmp = tmp + CStr(info.useRestrictYN) + " | "
        tmp = tmp + CStr(info.closeOnExpired) + " | "
        tmp = tmp + CStr(info.unPaidYN) + " | "
        tmp = tmp + info.Memo
        tmp = tmp + vbCrLf
    Next
    
    MsgBox tmp
    
End Sub

'=========================================================================
' ����ȸ�� ����ڹ�ȣ�� ��ϵ� �����(�˺� �α��� ����) ����� Ȯ���մϴ�.
' - https://docs.popbill.com/easyfinbank/vb/api#ListContact
'=========================================================================
Private Sub btnListContact_Click()
    Dim resultList As Collection
    Dim tmp As String
    Dim info As PBContactInfo
    
    Set resultList = easyFinBankService.ListContact(txtCorpNum.Text)
     
    If resultList Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(easyFinBankService.LastErrCode) + vbCrLf + "����޽��� : " + easyFinBankService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "id(���̵�) | personName(����) | email(�̸���) | hp(�޴�����ȣ) |  fax(�ѽ���ȣ) | tel(����ó) | " _
         + "regDT(����Ͻ�) | searchAllAllowYN(ȸ����ȸ ���ѿ���) | mgrYN(������ ����) | state(����) " + vbCrLf
    
    For Each info In resultList
        tmp = tmp + info.id + " | " + info.personName + " | " + info.email + " | " + info.hp + " | " + info.fax _
        + info.tel + " | " + info.regDT + " | " + CStr(info.searchAllAllowYN) + " | " + CStr(info.mgrYN) + " | " + CStr(info.state) + vbCrLf
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' ������ȸ ���񽺸� �̿��� ���¸� �˺��� ����մϴ�.
' - https://docs.popbill.com/easyfinbank/vb/api#RegistBankAccount
'=========================================================================
Private Sub btnRegistBankAccount_Click()
    Dim AccountInfo As New PBEasyFinBankAccountForm
    Dim Response As PBResponse
    
    ' [�ʼ�] �����ڵ�
    ' �������-0002 / �������-0003 / ��������-0004 /��������-0007 / ��������-0011 / �츮����-0020
    ' SC����-0023 / �뱸����-0031 / �λ�����-0032 / ��������-0034 / ��������-0035 / ��������-0037
    ' �泲����-0039 / �������ݰ�-0045 / ��������-0048 / ��ü��-0071 / KEB�ϳ�����-0081 / ��������-0088 /��Ƽ����-0027
    AccountInfo.BankCode = ""
    
    ' [�ʼ�] ���¹�ȣ ������('-') ����
    AccountInfo.AccountNumber = ""

    ' [�ʼ�] ���º�й�ȣ
    AccountInfo.AccountPWD = ""

    ' [�ʼ�] ��������, "����" �Ǵ� "����" �Է�
    AccountInfo.AccountType = ""

    ' [�ʼ�] ������ �ĺ����� (��-�� ����)
    ' ���������� �����Ρ��� ��� : ����ڹ�ȣ(10�ڸ�)
    ' ���������� �����Ρ��� ��� : ������ ������� (6�ڸ�-YYMMDD)
    AccountInfo.IdentityNumber = ""

    ' ���� ��Ī
    AccountInfo.AccountName = ""

    ' ���ͳݹ�ŷ ���̵� (�������� �ʼ�)
    AccountInfo.BankID = ""

    ' ��ȸ���� ���� ���̵� (�뱸����, ����, �������� �ʼ�)
    AccountInfo.FastID = ""

    ' ��ȸ���� ���� ��й�ȣ (�뱸����, ����, �������� �ʼ�
    AccountInfo.FastPWD = ""

    ' �����Ⱓ(����), 1~12 �Է°���, �̱���� �⺻��(1) ó��
    ' - ��Ʈ�� ���ݹ���� ��� �Է°��� ������� 1���� ó��
    AccountInfo.UsePeriod = "1"

    ' �޸�
    AccountInfo.Memo = ""
    
   
    Set Response = easyFinBankService.RegistBankAccount(txtCorpNum.Text, AccountInfo)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(easyFinBankService.LastErrCode) + vbCrLf + "����޽��� : " + easyFinBankService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.Message)
End Sub

'=========================================================================
' ����ȸ�� ����ڹ�ȣ�� �����(�˺� �α��� ����)�� �߰��մϴ�.
' - https://docs.popbill.com/easyfinbank/vb/api#RegistContact
'=========================================================================
Private Sub btnRegistContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    '����� ���̵�, 6�� �̻� 50�� �̸�
    joinData.id = "testkorea"
    
    '��й�ȣ, 6�� �̻� 20�� �̸�
    joinData.pwd = "test@test.com"
    
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
    
    'ȸ����ȸ ���ѿ���, True-ȸ����ȸ / False-������ȸ
    joinData.searchAllAllowYN = True
    
    '������ ����, True-������ / False-�����
    joinData.mgrYN = False
        
    Set Response = easyFinBankService.RegistContact(txtCorpNum.Text, joinData)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(easyFinBankService.LastErrCode) + vbCrLf + "����޽��� : " + easyFinBankService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.Message)
End Sub

'=========================================================================
' ���� �ŷ������� Ȯ���ϱ� ���� �˺��� ������û�� �մϴ�. ��ȸ�Ⱓ�� ���� �������� 90�� �̳��θ� ���� �����մϴ�.
' - ��ȯ ���� �۾����̵�� �Լ� ȣ�� �������� 1�ð� ���� ��ȿ�մϴ�.
' - https://docs.popbill.com/easyfinbank/vb/api#RequestJob
'=========================================================================

Private Sub btnRequestJob_Click()
    Dim jobID As String
    Dim BankCode As String
    Dim AccountNumber As String
    Dim SDate As String
    Dim EDate As String
    
    '�����ڵ�
    BankCode = "0004"
    
    '�˺��� ��ϵ� ���¹�ȣ
    AccountNumber = "20700644024204"
    
    '��������, ǥ������(yyyyMMdd)
    SDate = "20210901"
    
    '��������, ǥ������(yyyyMMdd)
    EDate = "20210910"
    
    jobID = easyFinBankService.RequestJob(txtCorpNum.Text, BankCode, AccountNumber, SDate, EDate)
    
    If jobID = "" Then
        MsgBox ("�����ڵ� : " + CStr(easyFinBankService.LastErrCode) + vbCrLf + "����޽��� : " + easyFinBankService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "jobID(�۾����̵�) : " + jobID + vbCrLf
    
    txtJobID.Text = jobID
End Sub

'=========================================================================
' ��û�� ������ ������û�� ����մϴ�.
' - https://docs.popbill.com/easyfinbank/vb/api#RevokeCloseBankAccount
'=========================================================================
Private Sub btnRevokeCloseBankAccount_Click()

    Dim Response As PBResponse
    Dim BankCode As String
    Dim AccountNumber As String
    
    ' [�ʼ�] �����ڵ�
    ' �������-0002 / �������-0003 / ��������-0004 /��������-0007 / ��������-0011 / �츮����-0020
    ' SC����-0023 / �뱸����-0031 / �λ�����-0032 / ��������-0034 / ��������-0035 / ��������-0037
    ' �泲����-0039 / �������ݰ�-0045 / ��������-0048 / ��ü��-0071 / KEB�ϳ�����-0081 / ��������-0088 /��Ƽ����-0027
    BankCode = ""
    
    ' [�ʼ�] ���¹�ȣ ������('-') ����
    AccountNumber = ""
    
    Set Response = easyFinBankService.RevokeCloseBankAccount(txtCorpNum.Text, BankCode, AccountNumber)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(easyFinBankService.LastErrCode) + vbCrLf + "����޽��� : " + easyFinBankService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.Message)
End Sub

'=========================================================================
' �� ���� �ŷ� ������ �޸� �����մϴ�.
' - https://docs.popbill.com/easyfinbank/vb/api#SaveMemo
'=========================================================================
Private Sub btnSaveMemo_Click()
    Dim Response As PBResponse
    Dim tid As String
    Dim Memo As String
    
    ' �ŷ����� ���̵�, SeachAPI �����׸� �� tid
    tid = "02112181100000000120211210000003"
    
    '�޸�
    Memo = "�޸� �׽�Ʈ"
    
    Set Response = easyFinBankService.SaveMemo(txtCorpNum.Text, tid, Memo)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(easyFinBankService.LastErrCode) + vbCrLf + "����޽��� : " + easyFinBankService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.Message)
End Sub


'=========================================================================
' GetJobState(���� ���� Ȯ��)�� ���� ���� ������ Ȯ�ε� �۾����̵� Ȱ���Ͽ� ���� �ŷ� ������ ��ȸ�մϴ�.
' - https://docs.popbill.com/easyfinbank/vb/api#Search
'=========================================================================
Private Sub btnSearch_Click()
    Dim searchList As PBEasyFinBankSearch
    Dim TradeType As New Collection
    Dim page As Integer
    Dim perPage As Integer
    Dim order As String
    Dim tmp As String
    Dim listboxRow As String
    Dim SearchString As String
        
    '�ŷ����� �迭, I-�Ա�, O-����
    TradeType.Add "I"
    TradeType.Add "O"
    
    '��������ȣ, �⺻�� ��1��
    page = 1
    
    '�������� �˻�����, �⺻�� 500, �ִ� 1000
    perPage = 20
    
    '���� ����, D-��������, A-��������
    order = "D"
    
    '��ȸ �˻���, �Ա�/��ݾ�, �޸�, ���� like �˻�
    SearchString = ""
        
    Set searchList = easyFinBankService.Search(txtCorpNum.Text, txtJobID.Text, TradeType, SearchString, page, perPage, order, txtUserID.Text)
    
        
    If searchList Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(easyFinBankService.LastErrCode) + vbCrLf + "����޽��� : " + easyFinBankService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "code (�����ڵ�) : " + CStr(searchList.code) + vbCrLf
    tmp = tmp + "message (����޽���) : " + searchList.Message + vbCrLf
    tmp = tmp + "total (�� �˻���� �Ǽ�) : " + CStr(searchList.total) + vbCrLf
    tmp = tmp + "perPage (�������� �˻�����) : " + CStr(searchList.perPage) + vbCrLf
    tmp = tmp + "pageNum (������ ��ȣ) : " + CStr(searchList.pageNum) + vbCrLf
    tmp = tmp + "pageCount (������ ����) : " + CStr(searchList.pageCount) + vbCrLf
    tmp = tmp + "lastScrapDT (���� ��ȸ�Ͻ�) : " + searchList.lastScrapDT + vbCrLf + vbCrLf
    
    searchInfo.Clear
        
    searchInfo.AddItem "tid (�ŷ����� ���̵�) | trdate (�ŷ�����) | trserial (�ŷ����ں� �ŷ���������) | trdt (�ŷ��Ͻ�) | accIn (�Աݾ�) | accOut (��ݾ�) | balance (�ܾ�) | ", 0
    searchInfo.AddItem "remark1 (���) | remark2 (���) | remark3 (���) | remark4 (���) | regDT (����Ͻ�) | memo (�޸�)", 1
    
    Dim tiInfo As PBEasyFinBankSearchDetail
           
    For Each tiInfo In searchList.list
        listboxRow = ""
        listboxRow = tiInfo.tid + " | "
        listboxRow = listboxRow + tiInfo.trdate + " | "
        listboxRow = listboxRow + CStr(tiInfo.trserial) + " | "
        listboxRow = listboxRow + tiInfo.trdt + " | "
        listboxRow = listboxRow + tiInfo.accIn + " | "
        listboxRow = listboxRow + tiInfo.accOut + " | "
        listboxRow = listboxRow + tiInfo.balance + " | "
        listboxRow = listboxRow + tiInfo.remark1 + " | "
        listboxRow = listboxRow + tiInfo.remark2 + " | "
        listboxRow = listboxRow + tiInfo.remark3 + " | "
        listboxRow = listboxRow + tiInfo.remark4 + " | "
        listboxRow = listboxRow + tiInfo.regDT + " | "
        listboxRow = listboxRow + tiInfo.Memo
        searchInfo.AddItem listboxRow, searchInfo.ListCount
    Next
  
    MsgBox (tmp)
End Sub

'=========================================================================
' GetJobState(���� ���� Ȯ��)�� ���� ���� ������ Ȯ�ε� �۾����̵� Ȱ���Ͽ� ���� �ŷ������� ��� ������ ��ȸ�մϴ�.
' - https://docs.popbill.com/easyfinbank/vb/api#Summary
'=========================================================================
Private Sub btnSummary_Click()
    Dim summaryInfo As PBEasyFinBankSummary
    Dim TradeType As New Collection
    Dim SearchString As String
    Dim tmp As String
    
    '�ŷ����� �迭, I-�Ա�, O-����
    TradeType.Add "I"
    TradeType.Add "O"
    
    '��ȸ �˻���, �Ա�/��ݾ�, �޸�, ���� like �˻�
    SearchString = ""
    
    Set summaryInfo = easyFinBankService.Summary(txtCorpNum.Text, txtJobID.Text, TradeType, SearchString, txtUserID.Text)
        
    If summaryInfo Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(easyFinBankService.LastErrCode) + vbCrLf + "����޽��� : " + easyFinBankService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "count (������� �Ǽ�) : " + CStr(summaryInfo.count) + vbCrLf
    tmp = tmp + "cntAccIn (�Աݰŷ� �Ǽ�) : " + CStr(summaryInfo.cntAccIn) + vbCrLf
    tmp = tmp + "cntAccOut (��ݰŷ� �Ǽ�) : " + CStr(summaryInfo.cntAccOut) + vbCrLf
    tmp = tmp + "totalAccIn (�Աݾ� �հ�) : " + CStr(summaryInfo.totalAccIn) + vbCrLf
    tmp = tmp + "totalAccOut (��ݾ� �հ�) : " + CStr(summaryInfo.totalAccOut) + vbCrLf
       
    MsgBox (tmp)
End Sub

'=========================================================================
' �˺��� ��ϵ� ���������� �����մϴ�.
' - https://docs.popbill.com/easyfinbank/vb/api#UpdateBankAccount
'=========================================================================
Private Sub btnUpdateBankAccount_Click()
    Dim AccountInfo As New PBEasyFinBankAccountForm
    Dim Response As PBResponse
    
    ' [�ʼ�] �����ڵ�
    ' �������-0002 / �������-0003 / ��������-0004 /��������-0007 / ��������-0011 / �츮����-0020
    ' SC����-0023 / �뱸����-0031 / �λ�����-0032 / ��������-0034 / ��������-0035 / ��������-0037
    ' �泲����-0039 / �������ݰ�-0045 / ��������-0048 / ��ü��-0071 / KEB�ϳ�����-0081 / ��������-0088 /��Ƽ����-0027
    AccountInfo.BankCode = ""
    
    ' [�ʼ�] ���¹�ȣ ������('-') ����
    AccountInfo.AccountNumber = ""

    ' [�ʼ�] ���º�й�ȣ
    AccountInfo.AccountPWD = ""

    ' ���� ��Ī
    AccountInfo.AccountName = ""

    ' ���ͳݹ�ŷ ���̵� (�������� �ʼ�)
    AccountInfo.BankID = ""

    ' ��ȸ���� ���� ���̵� (�뱸����, ����, �������� �ʼ�)
    AccountInfo.FastID = ""

    ' ��ȸ���� ���� ��й�ȣ (�뱸����, ����, �������� �ʼ�
    AccountInfo.FastPWD = ""
    
    ' �޸�
    AccountInfo.Memo = ""
    
   
    Set Response = easyFinBankService.UpdateBankAccount(txtCorpNum.Text, AccountInfo)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(easyFinBankService.LastErrCode) + vbCrLf + "����޽��� : " + easyFinBankService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.Message)
End Sub

'=========================================================================
' ����ȸ�� ����ڹ�ȣ�� ��ϵ� �����(�˺� �α��� ����) ������ �����մϴ�.
' - https://docs.popbill.com/easyfinbank/vb/api#UpdateContact
'=========================================================================
Private Sub btnUpdateContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    '����� ���̵�
    joinData.id = txtUserID.Text
    
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

    'ȸ����ȸ ���ѿ���, True-ȸ����ȸ / False-������ȸ
    joinData.searchAllAllowYN = True
    
    '������ ����, True-������ / False-�����
    joinData.mgrYN = False
                
    Set Response = easyFinBankService.UpdateContact(txtCorpNum.Text, joinData, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(easyFinBankService.LastErrCode) + vbCrLf + "����޽��� : " + easyFinBankService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.Message)
End Sub

'=========================================================================
' ����ȸ���� ȸ�������� �����մϴ�.
' - https://docs.popbill.com/easyfinbank/vb/api#UpdateCorpInfo
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
    
    Set Response = easyFinBankService.UpdateCorpInfo(txtCorpNum.Text, CorpInfo)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(easyFinBankService.LastErrCode) + vbCrLf + "����޽��� : " + easyFinBankService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.Message)
End Sub

Private Sub Form_Load()

    '���� ������ȸ ���� �ʱ�ȭ
    easyFinBankService.Initialize linkID, SecretKey
    
    '����ȯ�� ������ True(���߿�), False(�����)
    easyFinBankService.IsTest = True
    
    '������ū IP���ѱ�� ��뿩��, True(����)
    easyFinBankService.IPRestrictOnOff = True
    
End Sub

