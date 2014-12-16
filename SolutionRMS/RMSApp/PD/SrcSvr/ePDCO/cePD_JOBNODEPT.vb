
'****************************************************************************************
'�ý��۱��� : RMS/PD/Server Entity Class
'����  ȯ�� : GAC(Global Assembly Cache)
'���α׷��� : cePD_JOBNO.vb (PD_JOBNO Entity ó�� Class)
'��      �� : PD_JOBNO Entity�� ����Insert/Update/Delete/Select�� ó��
'             - �θ�ƼƼ ��ü�� SCGLUtil.ceEntity�� ���
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2008-11-07 
'****************************************************************************************
Imports SCGLUtil.cbSCGLUtil     '���� ��ƿ��Ƽ ��ü
Imports SCGLUtil.cbSCGLErr      '���� ����ó�� ��ü
Imports SCGLEntity              '��ƼƼ ��ü�� �θ� ��ü

Public Class cePD_JOBNODEPT
    Inherits ceEntity

#Region "GROUP BLOCk : ���� �Ǵ� ��ⷹ���� ����/��� ����"
    Private Const CLASS_NAME = "cePD_JOBNODEPT"    '�ڽ��� Ŭ������
#End Region

#Region "GROUP BLOCk : �ܺο� ���� Method"


#Region "SQL Insert/Update/Delete/Select"
    '*****************************************************************
    '�Է� : strSQL = SQL ��
    '��ȯ : ó���Ǽ�
    '��� : �ش� Entity�� Insert ó��
    '*****************************************************************
    'SEQ|DEPTNAME|DEPTCODE|EMPNAME|EMPNO

    Public Function InsertDo(ByVal strJOBNO As String, _
                             Optional ByVal strSEQ As String = OPTIONAL_STR, _
                             Optional ByVal strDEPTCODE As String = OPTIONAL_STR, _
                             Optional ByVal strEMPNO As String = OPTIONAL_STR, _
                             Optional ByVal strJOBNOSEQ As String = OPTIONAL_STR, _
                             Optional ByVal strACTRATE As Double = OPTIONAL_NUM) As Integer

        Dim strSQL As String
        Dim strFields As New System.Text.StringBuilder
        Dim strValues As New System.Text.StringBuilder
        Dim strNOW As String '����Ʈ���� ó���� ������ �޾� �ؽ�Ʈ�� ó�� �Ѵ�.. 
        strNOW = Now
        Try
            BuildNameValues(",", "JOBNO", strJOBNO, strFields, strValues)
            BuildNameValues(",", "SEQ", strSEQ, strFields, strValues)
            BuildNameValues(",", "DEPTCODE", strDEPTCODE, strFields, strValues)
            BuildNameValues(",", "EMPNO", strEMPNO, strFields, strValues)
            BuildNameValues(",", "JOBNOSEQ", strJOBNOSEQ, strFields, strValues)
            BuildNameValues(",", "ACTRATE", strACTRATE, strFields, strValues)
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

    '*****************************************************************
    '�Է� : strSQL = SQL ��
    '��ȯ : ó���Ǽ�
    '��� : �ش� Entity�� Update ó��
    '���� : Key ���ǰ� Value Field����������(OPTIONAL_STR/OPTIONAL_NUM/OPTIONAL_DTM)
    '*****************************************************************
    'strJOBNO,seq,deptcode,empno,jobnoseq,actrate
    Public Function UpdateDo(Optional ByVal strJOBNO As String = OPTIONAL_STR, _
                             Optional ByVal strOLDSEQ As String = OPTIONAL_STR, _
                             Optional ByVal strDEPTCODE As String = OPTIONAL_STR, _
                             Optional ByVal strEMPNO As String = OPTIONAL_STR, _
                             Optional ByVal strJOBNOSEQ As Double = OPTIONAL_NUM, _
                             Optional ByVal strACTRATE As Double = OPTIONAL_NUM) As Integer

        Dim strSQL As String
        Dim strNOW As String '����Ʈ���� ó���� ������ �޾� �ؽ�Ʈ�� ó�� �Ѵ�.. 
        strNOW = Now

        Try
            strSQL = String.Format("UPDATE {0} SET {1} WHERE {2}", EntityName, _
                     BuildFields(",", _
                        GetFieldNameValue("DEPTCODE", strDEPTCODE), _
                        GetFieldNameValue("EMPNO", strEMPNO), _
                        GetFieldNameValue("ACTRATE", strACTRATE), _
                        GetFieldNameValue("UUSER", mobjSCGLConfig.WRKUSR), _
                        GetFieldNameValue("UDATE", strNOW)), _
                    BuildFields("AND", _
                        GetFieldNameValue("SEQ", strOLDSEQ), GetFieldNameValue("JOBNO", strJOBNO), GetFieldNameValue("JOBNOSEQ", strJOBNOSEQ)))

            'BuildFields("AND", _
            '           GetFieldNameValue("SEQ", strOLDSEQ)), BuildFields("AND", _
            '                     GetFieldNameValue("JOBNO", strJOBNO)), BuildFields("AND", _
            '                    GetFieldNameValue("JOBNOSEQ", strJOBNOSEQ)))

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
    Public Function DeleteDo(Optional ByVal strSEQ As String = OPTIONAL_STR, _
                             Optional ByVal strJOBNO As String = OPTIONAL_STR, _
                             Optional ByVal strJOBNOSEQ As Double = OPTIONAL_NUM) As Integer

        Dim strSQL As String

        Try
            strSQL = String.Format("DELETE FROM {0} WHERE {1}", EntityName, _
                     BuildFields("AND", _
                                   GetFieldNameValue("SEQ", strSEQ), _
                                   GetFieldNameValue("JOBNO", strJOBNO), _
                                   GetFieldNameValue("JOBNOSEQ", strJOBNOSEQ)))

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".DeleteDo")
        End Try
    End Function

    Public Function DeleteDo2(Optional ByVal strJOBNO As String = OPTIONAL_STR) As Integer

        Dim strSQL As String

        Try
            strSQL = String.Format("DELETE FROM {0} WHERE {1}", EntityName, _
                     BuildFields("AND", _
                                   GetFieldNameValue("JOBNO", strJOBNO)))

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".DeleteDo")
        End Try
    End Function


    '*****************************************************************
    '�Է� : strSQL = SQL ��
    '��ȯ : ó���Ǽ�
    '��� : �ش� Entity�� Delete ó��
    '���� : Key ������ ��������(OPTIONAL_STR/OPTIONAL_NUM/OPTIONAL_DTM)
    '*****************************************************************
    Public Function DeleteDo_JOBNO(Optional ByVal strJOBNO As String = OPTIONAL_STR) As Integer

        Dim strSQL As String

        Try
            strSQL = String.Format("DELETE FROM {0} WHERE {1}", EntityName, _
                     BuildFields("AND", _
                                   GetFieldNameValue("JOBNO", strJOBNO)))

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
        MyBase.EntityName = "PD_JOBNODEPT"     'Entity Name ����
    End Sub

    '���� ����� Base Class���� �����Ǿ� ����
#End Region
#End Region
End Class

