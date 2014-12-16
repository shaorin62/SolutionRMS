'****************************************************************************************
'Generated By: MakeSFAR V.2.0.0 - ��ƼƼ Ŭ���� ����Ŀ - ��ȭ S&C
'�ý��۱��� : �ַ�Ǹ�/�ý��۸�/Server Entity Class
'����  ȯ�� : GAC(Global Assembly Cache)
'���α׷��� : ceSC_REALMEDCODE_MST.vb ( SC_REALMEDCODE_MST Entity ó�� Class)
'��      �� : SC_REALMEDCODE_MST Entity�� ����Insert/Update/Delete/Select�� ó��
'             - �θ�ƼƼ ��ü�� SCGLUtil.ceEntity�� ���
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2008-01-14 ���� 11:09:29 By Making Entity Bean
'****************************************************************************************
Imports SCGLUtil.cbSCGLUtil     '���� ��ƿ��Ƽ ��ü
Imports SCGLUtil.cbSCGLErr      '���� ����ó�� ��ü
Imports SCGLEntity              '��ƼƼ ��ü�� �θ� ��ü

Public Class ceSC_CUST_EMP
    Inherits ceEntity

#Region "GROUP BLOCk : ���� �Ǵ� ��ⷹ���� ����/��� ����"
    Private Const CLASS_NAME = "ceSC_CUST_EMP"    '�ڽ��� Ŭ������
#End Region

#Region "GROUP BLOCk : �ܺο� ���� Method"
#Region "SQL Insert/Update/Delete/Select"

    '*****************************************************************
    '�Է� : strSQL = SQL ��
    '��ȯ : ó���Ǽ�
    '��� : �ش� Entity�� Update ó��
    '���� : Key ���ǰ� Value Field����������(OPTIONAL_STR/OPTIONAL_NUM/OPTIONAL_DTM)
    '*****************************************************************
    Public Function InsertDo(ByVal strCUSTCODE As String, _
            ByVal dblSEQ As Double, _
            Optional ByVal strHIGHCUSTCODE As String = OPTIONAL_STR, _
            Optional ByVal strEMP_NAME As String = OPTIONAL_STR, _
            Optional ByVal strEMP_HP As String = OPTIONAL_STR, _
            Optional ByVal strEMP_EMAIL As String = OPTIONAL_STR, _
            Optional ByVal strEMP_TEL As String = OPTIONAL_STR, _
            Optional ByVal strJONGNO As String = OPTIONAL_STR, _
            Optional ByVal strDEPT_NAME As String = OPTIONAL_STR, _
            Optional ByVal strDEF_GBN As String = OPTIONAL_STR, _
            Optional ByVal strVOCH_TYPE As String = OPTIONAL_STR, _
            Optional ByVal strURL As String = OPTIONAL_STR, _
            Optional ByVal strUSE_YN As String = OPTIONAL_STR, _
            Optional ByVal strMEMO As String = OPTIONAL_STR, _
            Optional ByVal strATTR01 As String = OPTIONAL_STR, _
            Optional ByVal strATTR02 As String = OPTIONAL_STR, _
            Optional ByVal strATTR03 As String = OPTIONAL_STR, _
            Optional ByVal strATTR04 As String = OPTIONAL_STR, _
            Optional ByVal strATTR05 As String = OPTIONAL_STR)

        Dim strSQL As String
        Dim strFields As New System.Text.StringBuilder
        Dim strValues As New System.Text.StringBuilder
        Dim strNOW As String '����Ʈ���� ó���� ������ �޾� �ؽ�Ʈ�� ó�� �Ѵ�.. 
        strNOW = Now

        Try
            BuildNameValues(",", "CUSTCODE", strCUSTCODE, strFields, strValues)
            BuildNameValues(",", "SEQ", dblSEQ, strFields, strValues)
            BuildNameValues(",", "HIGHCUSTCODE", strHIGHCUSTCODE, strFields, strValues)
            BuildNameValues(",", "EMP_NAME", strEMP_NAME, strFields, strValues)
            BuildNameValues(",", "EMP_HP", strEMP_HP, strFields, strValues)
            BuildNameValues(",", "EMP_EMAIL", strEMP_EMAIL, strFields, strValues)
            BuildNameValues(",", "EMP_TEL", strEMP_TEL, strFields, strValues)
            BuildNameValues(",", "JONGNO", strJONGNO, strFields, strValues)
            BuildNameValues(",", "DEPT_NAME", strDEPT_NAME, strFields, strValues)
            BuildNameValues(",", "DEF_GBN", strDEF_GBN, strFields, strValues)
            BuildNameValues(",", "VOCH_TYPE", strVOCH_TYPE, strFields, strValues)
            BuildNameValues(",", "URL", strURL, strFields, strValues)
            BuildNameValues(",", "USE_YN", strUSE_YN, strFields, strValues)
            BuildNameValues(",", "MEMO", strMEMO, strFields, strValues)
            BuildNameValues(",", "ATTR01", strATTR01, strFields, strValues)
            BuildNameValues(",", "ATTR02", strATTR02, strFields, strValues)
            BuildNameValues(",", "ATTR03", strATTR03, strFields, strValues)
            BuildNameValues(",", "ATTR04", strATTR04, strFields, strValues)
            BuildNameValues(",", "ATTR05", strATTR05, strFields, strValues)
            BuildNameValues(",", "CUSER", mobjSCGLConfig.WRKUSR, strFields, strValues)
            BuildNameValues(",", "CDATE", strNOW, strFields, strValues)
            BuildNameValues(",", "UUSER", mobjSCGLConfig.WRKUSR, strFields, strValues)
            BuildNameValues(",", "UDATE", strNOW, strFields, strValues)

            strSQL = String.Format("INSERT INTO {0} ({1}) VALUES({2})", EntityName, strFields, strValues)

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".InsertDo")
        End Try
    End Function

    Public Function UpdateDo(Optional ByVal strCUSTCODE As String = OPTIONAL_STR, _
            Optional ByVal dblSEQ As Double = OPTIONAL_NUM, _
            Optional ByVal strHIGHCUSTCODE As String = OPTIONAL_STR, _
            Optional ByVal strEMP_NAME As String = OPTIONAL_STR, _
            Optional ByVal strEMP_HP As String = OPTIONAL_STR, _
            Optional ByVal strEMP_EMAIL As String = OPTIONAL_STR, _
            Optional ByVal strEMP_TEL As String = OPTIONAL_STR, _
            Optional ByVal strJONGNO As String = OPTIONAL_STR, _
            Optional ByVal strDEPT_NAME As String = OPTIONAL_STR, _
            Optional ByVal strDEF_GBN As String = OPTIONAL_STR, _
            Optional ByVal strVOCH_TYPE As String = OPTIONAL_STR, _
            Optional ByVal strURL As String = OPTIONAL_STR, _
            Optional ByVal strUSE_YN As String = OPTIONAL_STR, _
            Optional ByVal strMEMO As String = OPTIONAL_STR, _
            Optional ByVal strATTR01 As String = OPTIONAL_STR, _
            Optional ByVal strATTR02 As String = OPTIONAL_STR, _
            Optional ByVal strATTR03 As String = OPTIONAL_STR, _
            Optional ByVal strATTR04 As String = OPTIONAL_STR, _
            Optional ByVal strATTR05 As String = OPTIONAL_STR) As Integer

        Dim strSQL As String
        Dim strNOW As String '����Ʈ���� ó���� ������ �޾� �ؽ�Ʈ�� ó�� �Ѵ�.. 
        strNOW = Now

        Try
            strSQL = String.Format("UPDATE {0} SET {1} WHERE {2}", EntityName, _
                     BuildFields(",", _
                        GetFieldNameValue("HIGHCUSTCODE", strHIGHCUSTCODE), _
                        GetFieldNameValue("EMP_NAME", strEMP_NAME), _
                        GetFieldNameValue("EMP_HP", strEMP_HP), _
                        GetFieldNameValue("EMP_EMAIL", strEMP_EMAIL), _
                        GetFieldNameValue("EMP_TEL", strEMP_TEL), _
                        GetFieldNameValue("JONGNO", strJONGNO), _
                        GetFieldNameValue("DEPT_NAME", strDEPT_NAME), _
                        GetFieldNameValue("DEF_GBN", strDEF_GBN), _
                        GetFieldNameValue("VOCH_TYPE", strVOCH_TYPE), _
                        GetFieldNameValue("URL", strURL), _
                        GetFieldNameValue("USE_YN", strUSE_YN), _
                        GetFieldNameValue("MEMO", strMEMO), _
                        GetFieldNameValue("ATTR01", strATTR01), _
                        GetFieldNameValue("ATTR02", strATTR02), _
                        GetFieldNameValue("ATTR03", strATTR03), _
                        GetFieldNameValue("ATTR04", strATTR04), _
                        GetFieldNameValue("ATTR05", strATTR05), _
                        GetFieldNameValue("UUSER", mobjSCGLConfig.WRKUSR), _
                        GetFieldNameValue("UDATE", strNOW)), _
                     BuildFields("AND", _
                        GetFieldNameValue("CUSTCODE", strCUSTCODE), GetFieldNameValue("SEQ", dblSEQ)))

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".UpdateDo")
        End Try
    End Function

    '*****************************************************************
    '�Է� : strSQL = SQL ��
    '��ȯ : ó���Ǽ�
    '��� : �ش� Entity�� Delete ó��
    '���� : Key ������ ��������(OPTIONAL_STR/OPTIONAL_NUM/OPTIONAL_DTM)
    '*****************************************************************
    Public Function DeleteDo(Optional ByVal strCUSTCODE As String = OPTIONAL_STR, Optional ByVal dblSEQ As Double = OPTIONAL_NUM) As Integer
        Dim strSQL As String

        Try
            strSQL = String.Format("DELETE FROM {0} WHERE {1}", EntityName, _
                     BuildFields("AND", _
                                   GetFieldNameValue("CUSTCODE", strCUSTCODE), GetFieldNameValue("SEQ", dblSEQ)))

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".DeleteDo")
        End Try
    End Function

#End Region

#Region "��ü ����/����"
    '*****************************************************************
    '�Է� : strInfoXML = ����⺻������ ���� XML
    'objSCGLSql = DB ó�� ��ü �ν��Ͻ� ����    '��ȯ : ����
    '��� : DB ó���� ���� ����⺻���� ����
    '*****************************************************************
    Public Sub New(Optional ByVal objSCGLConfig As SCGLUtil.cbSCGLConfig = Nothing, Optional ByVal strInfoXML As String = "")
        MyBase.SetConfig(objSCGLConfig, strInfoXML)
        MyBase.EntityName = "SC_CUST_EMP"     'Entity Name ����
    End Sub

    '���� ����� Base Class���� �����Ǿ� ����
#End Region
#End Region

End Class





