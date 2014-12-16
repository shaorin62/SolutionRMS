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

Public Class ccSCCOBMTIMLIST
    Inherits ccControl

#Region "GROUP BLOCK : ���� �Ǵ� ��ⷹ���� ����/��� ����"
    Private CLASS_NAME = "ccSCCOBMTIMLIST"                  '�ڽ��� Ŭ������
    Private mobjceSC_CCTR As eSCCO.ceSC_CCTR     'TYPE �������

#End Region

#Region "GROUP BLOCK : Function Section"
    Public Function Get_COMBO_VALUE(ByVal strInfoXML As String, _
                                     ByRef intRowCnt As Integer, _
                                     ByRef intColCnt As Integer) As Object

        Dim strSQL As String
        Dim vntData As Object

        SetConfig(strInfoXML)   '�⺻���� ����					

        With mobjSCGLConfig

            strSQL = " SELECT "
            strSQL = strSQL & " BMORDER, BMNAME+ ' - ' + BMDEFINE  BMDEFINE"
            strSQL = strSQL & " FROM SC_BM  "
            strSQL = strSQL & " WHERE 1=1  "
            strSQL = strSQL & " ORDER BY BMCODE "

            ''������ ��ȸ
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".Get_COMBO_VALUE")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

    '=============== bm�� ��ȸ
    Public Function SelectRtn(ByVal strInfoXML As String, _
                              ByRef intRowCnt As Integer, _
                              ByRef intColCnt As Integer, _
                              ByVal strBMNAME As String, _
                              ByVal strDEPTNAME As String, _
                              ByVal strUSE_YN As String) As Object

        Dim strCols As String         '�÷�����
        Dim strWhere As String       'Where���� ����
        Dim strFormat As String      'SQL Format ����
        Dim strSQL As String          'SQL ����
        Dim vntData As Object        'Array Return ����(Array ������� �� ����)
        Dim strXMLData As String    'XML  Return ����(XML  �� ����� �� ����)
        Dim Con1, Con2, Con3 As String


        If strBMNAME <> "" Then Con1 = String.Format(" AND (DBO.SC_GET_BMNAME_FUN(BMCODE) like '%{0}%')", strBMNAME)
        If strDEPTNAME <> "" Then Con2 = String.Format(" AND (DBO.SC_DEPT_NAME_FUN(DEPT_CD) like '%{0}%')", strDEPTNAME)
        If strUSE_YN <> "" Then Con3 = String.Format(" AND USE_YN = '{0}'", strUSE_YN)

        strWhere = BuildFields(" ", Con1, Con2, Con3)

        strFormat = "SELECT SEQ,"
        strFormat = strFormat & " BMCODE, "
        strFormat = strFormat & " HIGHDEPT_CD, "
        strFormat = strFormat & " DBO.SC_DEPT_NAME_FUN(HIGHDEPT_CD) HIGHDEPT_NAME,"
        strFormat = strFormat & " CCTR,"
        strFormat = strFormat & " BA,"
        strFormat = strFormat & " DEPT_CD,"
        strFormat = strFormat & " DBO.SC_DEPT_NAME_FUN(DEPT_CD) DEPT_NAME, "
        strFormat = strFormat & " FDATE, TDATE,"
        strFormat = strFormat & "  CASE USE_YN WHEN 'Y' THEN '1' ELSE '0' END USE_YN "
        strFormat = strFormat & " FROM SC_CCTR   "
        strFormat = strFormat & " WHERE 1=1 {0}"
        strFormat = strFormat & " ORDER BY BMCODE, CCTR, HIGHDEPT_NAME, DEPT_NAME"

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

    '�μ� BM ���� ���� 
    Public Function ProcessRtn(ByVal strInfoXML As String, _
                               ByVal vntData As Object) As Integer '������ INSERT/UPDATE
        Dim intRtn As Integer
        Dim i, intColCnt, intRows As Integer
        Dim strCODE As String
        Dim intCnt As Integer
        Dim strFDATE, strTDATE
        Dim strUSE_YN


        SetConfig(strInfoXML)
        With mobjSCGLConfig
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                .mobjSCGLSql.SQLBeginTrans()
                '�ϴܱ׸��� ����ó��
                If IsArray(vntData) Then
                    '''����� Entity ��ü����(Config ������ �Ѱܻ���)
                    mobjceSC_CCTR = New ceSC_CCTR(mobjSCGLConfig)
                    '''vntData�� �÷���, �ο���� �����Է�
                    intColCnt = UBound(vntData, 1) : intRows = UBound(vntData, 2)
                    intCnt = 0

                    For i = 1 To intRows
                        strFDATE = "" : strTDATE = ""
                        If GetElement(vntData, "FDATE", intColCnt, i, OPTIONAL_STR) <> "" Then strFDATE = GetElement(vntData, "FDATE", intColCnt, i, OPTIONAL_STR).Substring(0, 4) & GetElement(vntData, "FDATE", intColCnt, i, OPTIONAL_STR).Substring(5, 2) & GetElement(vntData, "FDATE", intColCnt, i, OPTIONAL_STR).Substring(8, 2)
                        If GetElement(vntData, "TDATE", intColCnt, i, OPTIONAL_STR) <> "" Then strTDATE = GetElement(vntData, "TDATE", intColCnt, i, OPTIONAL_STR).Substring(0, 4) & GetElement(vntData, "TDATE", intColCnt, i, OPTIONAL_STR).Substring(5, 2) & GetElement(vntData, "TDATE", intColCnt, i, OPTIONAL_STR).Substring(8, 2)

                        If GetElement(vntData, "USE_YN", intColCnt, i, OPTIONAL_NUM) = 1 Then
                            strUSE_YN = "Y"
                        Else
                            strUSE_YN = "N"
                        End If

                        If GetElement(vntData, "SEQ", intColCnt, i, NULL_NUM, True) = -999999 Then
                            intRtn = InsertRtn_SC_CCTR(vntData, intColCnt, i, strFDATE, strTDATE, strUSE_YN)
                        Else
                            intRtn = UpdateRtn_SC_CCTR(vntData, intColCnt, i, strFDATE, strTDATE, strUSE_YN)

                        End If
                    Next
                End If

                .mobjSCGLSql.SQLCommitTrans()
                Return intRows
            Catch err As Exception
                .mobjSCGLSql.SQLRollbackTrans()
                Throw RaiseSysErr(err, CLASS_NAME & ".ProcessRtn")
            Finally
                .mobjSCGLSql.SQLDisconnect()
                mobjceSC_CCTR.Dispose()

            End Try
        End With
    End Function

#End Region

#Region "GROUP BLOCK : Entity Function Section"
    'CE�ܿ� ������ ���� ����� ����
    Private Function InsertRtn_SC_CCTR(ByVal vntData As Object, _
                                       ByVal intColCnt As Integer, _
                                       ByVal intRow As Integer, _
                                       ByVal strFDATE As String, _
                                       ByVal strTDATE As String, _
                                       ByVal strUSE_YN As String) As Integer

        Dim intRtn As Integer

        intRtn = mobjceSC_CCTR.InsertDo( _
                                GetElement(vntData, "BMCODE", intColCnt, intRow), _
                                GetElement(vntData, "HIGHDEPT_CD", intColCnt, intRow), _
                                GetElement(vntData, "DEPT_CD", intColCnt, intRow), _
                                GetElement(vntData, "CCTR", intColCnt, intRow), _
                                GetElement(vntData, "BA", intColCnt, intRow), _
                                strFDATE, _
                                strTDATE, _
                                strUSE_YN)
        Return intRtn
    End Function

    Private Function UpdateRtn_SC_CCTR(ByVal vntData As Object, _
                                       ByVal intColCnt As Integer, _
                                       ByVal intRow As Integer, _
                                       ByVal strFDATE As String, _
                                       ByVal strTDATE As String, _
                                       ByVal strUSE_YN As String) As Integer

        Dim intRtn As Integer

        intRtn = mobjceSC_CCTR.UpdateDo( _
                                GetElement(vntData, "SEQ", intColCnt, intRow, NULL_NUM, True), _
                                GetElement(vntData, "BMCODE", intColCnt, intRow), _
                                GetElement(vntData, "HIGHDEPT_CD", intColCnt, intRow), _
                                GetElement(vntData, "DEPT_CD", intColCnt, intRow), _
                                GetElement(vntData, "CCTR", intColCnt, intRow), _
                                GetElement(vntData, "BA", intColCnt, intRow), _
                                strFDATE, _
                                strTDATE, _
                                strUSE_YN)
        Return intRtn
    End Function


#End Region
End Class