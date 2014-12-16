'****************************************************************************************
'�ý��۱���    : �ַ�Ǹ� /�ý��۸�/Server Control Class
'����   ȯ��    : COM+ Service Server Package
'���α׷���    : ccMDCMCUST_TRAN.vb
'��         ��    : - ����� ��� �մϴ�.
'Ư��  ����     : - Ư�̻��׿� ���� ǥ��
'                     -
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2009-07-03 ���� 10:32:13 By KTY
'****************************************************************************************

Imports System.Xml                  ' XMLó��
Imports SCGLControl                 ' ControlClass�� Base Class 
Imports SCGLUtil.cbSCGLConfig       ' ConfigurationClass
Imports SCGLUtil.cbSCGLErr          '����ó�� Ŭ����
Imports SCGLUtil.cbSCGLXml          'XMLó�� Ŭ����
Imports SCGLUtil.cbSCGLUtil         '��Ÿ��ƿ��Ƽ Ŭ����
Imports eSCCO '����Ƽ �߰�

' ��ƼƼ Ŭ���� ���� �ش� ��ƼƼ Ŭ������ ������Ʈ�� ������ �� Imports �Ͻʽÿ�. 
' Imports ��ƼƼ������Ʈ

Public Class ccSCCOBANKTYPE
    Inherits ccControl

#Region "GROUP BLOCK : ���� �Ǵ� ��ⷹ���� ����/��� ����"
    Private CLASS_NAME = "ccSCCOBANKTYPE"                  '�ڽ��� Ŭ������
    Private mobjceSC_BANKTYPE_MST As eSCCO.ceSC_BANKTYPE_MST     'TYPE �������

#End Region

#Region "GROUP BLOCK : Function Section"
   
    '=============== BANK_TYPE �ڷ� ��ȸ
    Public Function SelectRtn(ByVal strInfoXML As String, _
                              ByRef intRowCnt As Integer, _
                              ByRef intColCnt As Integer, _
                              ByVal strBUSINO As String) As Object

        Dim strWhere As String       'Where���� ����
        Dim strFormat As String      'SQL Format ����
        Dim strSQL As String          'SQL ����
        Dim vntData As Object        'Array Return ����(Array ������� �� ����)
        Dim Con1 As String

        If strBUSINO <> "" Then Con1 = String.Format(" AND (BUSINO = '{0}')", strBUSINO)

        strWhere = BuildFields(" ", Con1)

        strFormat = strFormat & " SELECT "
        strFormat = strFormat & " DBO.SC_BUSINO_CUSTNAME_FUN(BUSINO) CUSTNAME,"
        strFormat = strFormat & " substring(BUSINO,1,3) + '-' + substring(BUSINO,4,2) + '-' + substring(BUSINO,6,5) BUSINO,"
        strFormat = strFormat & " BANK_KEY,"
        strFormat = strFormat & " BANK_NUM,"
        strFormat = strFormat & " BANK_TYPE,"
        strFormat = strFormat & " BANK_USER,"
        strFormat = strFormat & " CASE USE_YN WHEN 'Y' THEN '1' ELSE '0' END USE_YN"
        strFormat = strFormat & " FROM SC_BANKTYPE_MST"
        strFormat = strFormat & " WHERE 1=1"
        strFormat = strFormat & " {0}"
        strFormat = strFormat & " ORDER BY DBO.SC_BUSINO_CUSTNAME_FUN(BUSINO) "

        SetConfig(strInfoXML) '�⺻���� Setting
        With mobjSCGLConfig '�⺻���� Config ��ü
            strSQL = String.Format(strFormat, strWhere)
            Try
                ' DB ����
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                ' ------ Array ������ ��ȸ (True �϶� ������� ���� ��ȸ(Sheet Data Binding �� ��� ���), False �϶� �����͸� ��ȸ)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)
                Return vntData

            Catch err As Exception
                ' ���� ����
                Throw RaiseSysErr(err, CLASS_NAME & ".SelectRtn")
            Finally
                ' DB ���� ����
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function


    ' =============== ProcessRtn_CUSTDTL    �ŷ�ó ������ ����
    Public Function ProcessRtn(ByVal strInfoXML As String, _
                               ByVal vntData As Object) As Object

        Dim intRtn As Integer
        Dim i, intColCnt, intRows As Integer
        Dim strBUSINO
        Dim strUSEYN

        SetConfig(strInfoXML)
        With mobjSCGLConfig
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                .mobjSCGLSql.SQLBeginTrans()
                If IsArray(vntData) Then
                    '''����� Entity ��ü����(Config ������ �Ѱܻ���)
                    mobjceSC_BANKTYPE_MST = New ceSC_BANKTYPE_MST(mobjSCGLConfig)
                    '''vntData�� �÷���, �ο���� �����Է�
                    intColCnt = UBound(vntData, 1) : intRows = UBound(vntData, 2)

                    '1��° �ο��� ����ڹ�ȣ�� ���� BANKTYPE �� ���� �Ŀ� ���ο� �����ͷ� ������Ʈ �Ѵ�.
                    strBUSINO = GetElement(vntData, "BUSINO", intColCnt, 1, OPTIONAL_STR)
                    intRtn = mobjceSC_BANKTYPE_MST.DeleteDo(strBUSINO)

                    For i = 1 To intRows
                        strUSEYN = GetElement(vntData, "USE_YN", intColCnt, 1, OPTIONAL_STR)

                        If strUSEYN = "1" Then
                            strUSEYN = "Y"
                        Else
                            strUSEYN = "N"
                        End If

                        intRtn = InsertRtn(vntData, intColCnt, i, strUSEYN)
                    Next
                End If
                .mobjSCGLSql.SQLCommitTrans()
                Return intRows
            Catch err As Exception
                .mobjSCGLSql.SQLRollbackTrans()
                Throw RaiseSysErr(err, CLASS_NAME & ".ProcessRtn")
            Finally
                .mobjSCGLSql.SQLDisconnect()
                mobjceSC_BANKTYPE_MST.Dispose()
            End Try
        End With
    End Function


#End Region

#Region "GROUP BLOCK : Entity Function Section"

    Private Function InsertRtn(ByVal vntData As Object, _
                               ByVal intColCnt As Integer, _
                               ByVal intRow As Integer, _
                               ByVal strUSEYN As String) As Integer
        Dim intRtn As Integer
        intRtn = mobjceSC_BANKTYPE_MST.InsertDo( _
                                       GetElement(vntData, "BUSINO", intColCnt, intRow), _
                                       GetElement(vntData, "BANK_KEY", intColCnt, intRow), _
                                       GetElement(vntData, "BANK_NUM", intColCnt, intRow), _
                                       GetElement(vntData, "BANK_TYPE", intColCnt, intRow), _
                                       GetElement(vntData, "BANK_USER", intColCnt, intRow), _
                                       strUSEYN, _
                                       GetElement(vntData, "ATTR01", intColCnt, intRow), _
                                       GetElement(vntData, "ATTR02", intColCnt, intRow), _
                                       GetElement(vntData, "ATTR03", intColCnt, intRow), _
                                       GetElement(vntData, "ATTR04", intColCnt, intRow), _
                                       GetElement(vntData, "ATTR05", intColCnt, intRow))
        Return intRtn
    End Function
#End Region
End Class