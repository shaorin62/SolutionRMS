'Public Class cePD_CONTRACT_SALE

'End Class
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

Public Class cePD_CONTRACT_DTL
    Inherits ceEntity

#Region "GROUP BLOCk : ���� �Ǵ� ��ⷹ���� ����/��� ����"
    Private Const CLASS_NAME = "cePD_CONTRACT_DTL"    '�ڽ��� Ŭ������
#End Region

#Region "GROUP BLOCk : �ܺο� ���� Method"
#Region "SQL Insert/Update/Delete/Select"
    '*****************************************************************
    '�Է� : strSQL = SQL ��
    '��ȯ : ó���Ǽ�
    '��� : �ش� Entity�� Insert ó��
    '*****************************************************************
    Public Function InsertDo(ByVal strOUTSCODE As String, _
                             Optional ByVal strJOBNO As String = OPTIONAL_STR, _
                             Optional ByVal dblAMT As Double = OPTIONAL_NUM, _
                             Optional ByVal strCONTRACTDAY As String = OPTIONAL_STR, _
                             Optional ByVal strCONTRACTNO As String = OPTIONAL_STR, _
                             Optional ByVal strMEMO As String = OPTIONAL_STR, _
                             Optional ByVal strCONFIRM_USER As String = OPTIONAL_STR, _
                             Optional ByVal strCONFIRM_DATE As String = OPTIONAL_STR, _
                             Optional ByVal strVOCHNO As String = OPTIONAL_STR, _
                             Optional ByVal strATTR01 As String = OPTIONAL_STR, _
                             Optional ByVal strATTR02 As String = OPTIONAL_STR, _
                             Optional ByVal strATTR03 As String = OPTIONAL_STR, _
                             Optional ByVal strATTR04 As String = OPTIONAL_STR, _
                             Optional ByVal strATTR05 As String = OPTIONAL_STR, _
                             Optional ByVal dblATTR06 As Double = OPTIONAL_NUM, _
                             Optional ByVal dblATTR07 As Double = OPTIONAL_NUM, _
                             Optional ByVal dblATTR08 As Double = OPTIONAL_NUM, _
                             Optional ByVal dblATTR09 As Double = OPTIONAL_NUM, _
                             Optional ByVal dblATTR10 As Double = OPTIONAL_NUM) As Integer

        Dim strSQL As String
        Dim strFields As New System.Text.StringBuilder
        Dim strValues As New System.Text.StringBuilder
        Dim strNOW As String '����Ʈ���� ó���� ������ �޾� �ؽ�Ʈ�� ó�� �Ѵ�.. 
        strNOW = Now


        Try
            BuildNameValues(",", "OUTSCODE", strOUTSCODE, strFields, strValues)
            BuildNameValues(",", "JOBNO", strJOBNO, strFields, strValues)
            BuildNameValues(",", "AMT", dblAMT, strFields, strValues)
            BuildNameValues(",", "CONTRACTDAY", strCONTRACTDAY, strFields, strValues)
            BuildNameValues(",", "CONTRACTNO", strCONTRACTNO, strFields, strValues)
            BuildNameValues(",", "MEMO", strMEMO, strFields, strValues)
            BuildNameValues(",", "CONFIRM_USER", strCONFIRM_USER, strFields, strValues)
            BuildNameValues(",", "CONFIRM_DATE", strCONFIRM_DATE, strFields, strValues)
            BuildNameValues(",", "VOCHNO", strVOCHNO, strFields, strValues)
            BuildNameValues(",", "ATTR01", strATTR01, strFields, strValues)
            BuildNameValues(",", "ATTR02", strATTR02, strFields, strValues)
            BuildNameValues(",", "ATTR03", strATTR03, strFields, strValues)
            BuildNameValues(",", "ATTR04", strATTR04, strFields, strValues)
            BuildNameValues(",", "ATTR05", strATTR05, strFields, strValues)
            BuildNameValues(",", "ATTR06", dblATTR06, strFields, strValues)
            BuildNameValues(",", "ATTR07", dblATTR07, strFields, strValues)
            BuildNameValues(",", "ATTR08", dblATTR08, strFields, strValues)
            BuildNameValues(",", "ATTR09", dblATTR09, strFields, strValues)
            BuildNameValues(",", "ATTR10", dblATTR10, strFields, strValues)
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
    Public Function UpdateDo(ByVal strSEQ As String, _
                           Optional ByVal strOUTSCODE As String = OPTIONAL_STR, _
                           Optional ByVal strJOBNO As String = OPTIONAL_STR, _
                           Optional ByVal dblAMT As Double = OPTIONAL_NUM, _
                           Optional ByVal strCONTRACTDAY As String = OPTIONAL_STR, _
                           Optional ByVal strCONTRACTNO As String = OPTIONAL_STR, _
                           Optional ByVal strMEMO As String = OPTIONAL_STR, _
                           Optional ByVal strCONFIRM_USER As String = OPTIONAL_STR, _
                           Optional ByVal strCONFIRM_DATE As String = OPTIONAL_STR, _
                           Optional ByVal strVOCHNO As String = OPTIONAL_STR, _
                           Optional ByVal strATTR01 As String = OPTIONAL_STR, _
                           Optional ByVal strATTR02 As String = OPTIONAL_STR, _
                           Optional ByVal strATTR03 As String = OPTIONAL_STR, _
                           Optional ByVal strATTR04 As String = OPTIONAL_STR, _
                           Optional ByVal strATTR05 As String = OPTIONAL_STR, _
                           Optional ByVal dblATTR06 As Double = OPTIONAL_NUM, _
                           Optional ByVal dblATTR07 As Double = OPTIONAL_NUM, _
                           Optional ByVal dblATTR08 As Double = OPTIONAL_NUM, _
                           Optional ByVal dblATTR09 As Double = OPTIONAL_NUM, _
                           Optional ByVal dblATTR10 As Double = OPTIONAL_NUM) As Integer

        Dim strSQL As String
        Dim strNOW As String '����Ʈ���� ó���� ������ �޾� �ؽ�Ʈ�� ó�� �Ѵ�.. 
        strNOW = Now
        Try

            strSQL = String.Format("UPDATE {0} SET {1} WHERE {2}", EntityName, _
                     BuildFields(",", _
                        GetFieldNameValue("OUTSCODE", strOUTSCODE), _
                        GetFieldNameValue("JOBNO", strJOBNO), _
                        GetFieldNameValue("AMT", dblAMT), _
                        GetFieldNameValue("CONTRACTDAY", strCONTRACTDAY), _
                        GetFieldNameValue("CONTRACTNO", strCONTRACTNO), _
                        GetFieldNameValue("MEMO", strMEMO), _
                        GetFieldNameValue("CONFIRM_USER", strCONFIRM_USER), _
                        GetFieldNameValue("CONFIRM_DATE", strCONFIRM_DATE), _
                        GetFieldNameValue("VOCHNO", strVOCHNO), _
                        GetFieldNameValue("ATTR01", strATTR01), _
                        GetFieldNameValue("ATTR02", strATTR02), _
                        GetFieldNameValue("ATTR03", strATTR03), _
                        GetFieldNameValue("ATTR04", strATTR04), _
                        GetFieldNameValue("ATTR05", strATTR05), _
                        GetFieldNameValue("ATTR06", dblATTR06), _
                        GetFieldNameValue("ATTR07", dblATTR07), _
                        GetFieldNameValue("ATTR08", dblATTR08), _
                        GetFieldNameValue("ATTR09", dblATTR09), _
                        GetFieldNameValue("ATTR10", dblATTR10), _
                        GetFieldNameValue("UUSER", mobjSCGLConfig.WRKUSR), _
                        GetFieldNameValue("UDATE", strNOW)), _
                     BuildFields("AND", _
                        GetFieldNameValue("SEQ", strSEQ)))

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
    Public Function DeleteDo(ByVal strSEQ As String) As Integer
        Dim strSQL As String

        Try
            strSQL = "DELETE FROM PD_CONTRACT_DTL WHERE SEQ = '" & strSEQ & "'"

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
    Public Function UpdateDo_CONTRACTNO(ByVal strSEQ As Integer, _
                                        ByVal strCONTRACTNO As String) As Integer
        Dim strSQL As String

        Try
            strSQL = "UPDATE PD_CONTRACT_DTL SET CONTRACTNO = '" & strCONTRACTNO & "' WHERE SEQ = '" & strSEQ & "'"

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
        MyBase.EntityName = "PD_CONTRACT_DTL"     'Entity Name ����
    End Sub

    '���� ����� Base Class���� �����Ǿ� ����
#End Region
#End Region
End Class
