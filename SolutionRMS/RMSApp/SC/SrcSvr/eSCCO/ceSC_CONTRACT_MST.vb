
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2008-01-14 ���� 11:09:29 By Making Entity Bean
'****************************************************************************************
Imports SCGLUtil.cbSCGLUtil     '���� ��ƿ��Ƽ ��ü
Imports SCGLUtil.cbSCGLErr      '���� ����ó�� ��ü
Imports SCGLEntity              '��ƼƼ ��ü�� �θ� ��ü
Public Class ceSC_CONTRACT_MST
    Inherits ceEntity

#Region "GROUP BLOCk : ���� �Ǵ� ��ⷹ���� ����/��� ����"
    Private Const CLASS_NAME = "ceSC_CONTRACT_MST"    '�ڽ��� Ŭ������
#End Region

#Region "SQL Insert/Update/Delete/Select"
    '*****************************************************************
    '�Է� : strSQL = SQL ��
    '��ȯ : ó���Ǽ�
    '��� : �ش� Entity�� Update ó��
    '���� : Key ���ǰ� Value Field����������(OPTIONAL_STR/OPTIONAL_NUM/OPTIONAL_DTM)
    '*****************************************************************

    Public Function InsertDo(Optional ByVal strSEQ As Double = OPTIONAL_NUM, _
                             Optional ByVal strCUSTCODE As String = OPTIONAL_STR, _
                             Optional ByVal strCUSTNAME As String = OPTIONAL_STR, _
                             Optional ByVal strCONTRACTNO As String = OPTIONAL_STR, _
                             Optional ByVal strCONTRACTNAME As String = OPTIONAL_STR, _
                             Optional ByVal strGBN As String = OPTIONAL_STR, _
                             Optional ByVal strAMT As Double = OPTIONAL_NUM, _
                             Optional ByVal strSTDATE As String = OPTIONAL_STR, _
                             Optional ByVal strEDDATE As String = OPTIONAL_STR, _
                             Optional ByVal strCONTRACTDAY As String = OPTIONAL_STR, _
                             Optional ByVal strCONFIRMFLAG As String = OPTIONAL_STR, _
                             Optional ByVal strCONFIRMDATE As String = OPTIONAL_STR, _
                             Optional ByVal strCONFIRM_USER As String = OPTIONAL_STR, _
                             Optional ByVal strMEMO As String = OPTIONAL_STR, _
                             Optional ByVal strCONDITION As String = OPTIONAL_STR, _
                             Optional ByVal strATTR01 As String = OPTIONAL_STR, _
                             Optional ByVal strATTR02 As String = OPTIONAL_STR, _
                             Optional ByVal strATTR03 As String = OPTIONAL_STR, _
                             Optional ByVal strATTR04 As String = OPTIONAL_STR, _
                             Optional ByVal strATTR05 As String = OPTIONAL_STR, _
                             Optional ByVal strATTR06 As Double = OPTIONAL_NUM, _
                             Optional ByVal strATTR07 As Double = OPTIONAL_NUM, _
                             Optional ByVal strATTR08 As Double = OPTIONAL_NUM, _
                             Optional ByVal strATTR09 As Double = OPTIONAL_NUM, _
                             Optional ByVal strATTR10 As Double = OPTIONAL_NUM)


        Dim strSQL As String
        Dim strFields As New System.Text.StringBuilder
        Dim strValues As New System.Text.StringBuilder
        Dim strNOW As String '����Ʈ���� ó���� ������ �޾� �ؽ�Ʈ�� ó�� �Ѵ�.. 
        strNOW = Now
        Try
            BuildNameValues(",", "SEQ", strSEQ, strFields, strValues)
            BuildNameValues(",", "CUSTCODE", strCUSTCODE, strFields, strValues)
            BuildNameValues(",", "CUSTNAME", strCUSTNAME, strFields, strValues)
            BuildNameValues(",", "CONTRACTNO", strCONTRACTNO, strFields, strValues)
            BuildNameValues(",", "CONTRACTNAME", strCONTRACTNAME, strFields, strValues)
            BuildNameValues(",", "GBN", strGBN, strFields, strValues)
            BuildNameValues(",", "AMT", strAMT, strFields, strValues)
            BuildNameValues(",", "STDATE", strSTDATE, strFields, strValues)
            BuildNameValues(",", "EDDATE", strEDDATE, strFields, strValues)
            BuildNameValues(",", "CONTRACTDAY", strCONTRACTDAY, strFields, strValues)
            BuildNameValues(",", "CONFIRMFLAG", strCONFIRMFLAG, strFields, strValues)
            BuildNameValues(",", "CONFIRMDATE", strCONFIRMDATE, strFields, strValues)
            BuildNameValues(",", "CONFIRM_USER", strCONFIRM_USER, strFields, strValues)
            BuildNameValues(",", "MEMO", strMEMO, strFields, strValues)
            BuildNameValues(",", "CONDITION", strCONDITION, strFields, strValues)
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

    Public Function UpdateDo(Optional ByVal strSEQ As Double = OPTIONAL_NUM, _
                             Optional ByVal strCUSTCODE As String = OPTIONAL_STR, _
                             Optional ByVal strCUSTNAME As String = OPTIONAL_STR, _
                             Optional ByVal strCONTRACTNO As String = OPTIONAL_STR, _
                             Optional ByVal strCONTRACTNAME As String = OPTIONAL_STR, _
                             Optional ByVal strGBN As String = OPTIONAL_STR, _
                             Optional ByVal strAMT As Double = OPTIONAL_NUM, _
                             Optional ByVal strSTDATE As String = OPTIONAL_STR, _
                             Optional ByVal strEDDATE As String = OPTIONAL_STR, _
                             Optional ByVal strCONTRACTDAY As String = OPTIONAL_STR, _
                             Optional ByVal strCONFIRMFLAG As String = OPTIONAL_STR, _
                             Optional ByVal strCONFIRMDATE As String = OPTIONAL_STR, _
                             Optional ByVal strCONFIRM_USER As String = OPTIONAL_STR, _
                             Optional ByVal strMEMO As String = OPTIONAL_STR, _
                             Optional ByVal strCONDITION As String = OPTIONAL_STR, _
                             Optional ByVal strATTR01 As String = OPTIONAL_STR, _
                             Optional ByVal strATTR02 As String = OPTIONAL_STR, _
                             Optional ByVal strATTR03 As String = OPTIONAL_STR, _
                             Optional ByVal strATTR04 As String = OPTIONAL_STR, _
                             Optional ByVal strATTR05 As String = OPTIONAL_STR, _
                             Optional ByVal strATTR06 As Double = OPTIONAL_NUM, _
                             Optional ByVal strATTR07 As Double = OPTIONAL_NUM, _
                             Optional ByVal strATTR08 As Double = OPTIONAL_NUM, _
                             Optional ByVal strATTR09 As Double = OPTIONAL_NUM, _
                             Optional ByVal strATTR10 As Double = OPTIONAL_NUM)

        Dim strSQL As String
        Dim strNOW As String '����Ʈ���� ó���� ������ �޾� �ؽ�Ʈ�� ó�� �Ѵ�.. 
        strNOW = Now
        Try
            strSQL = String.Format("UPDATE {0} SET {1} WHERE {2}", EntityName, _
                     BuildFields(",", _
                        GetFieldNameValue("CUSTCODE", strCUSTCODE), _
                        GetFieldNameValue("CUSTNAME", strCUSTNAME), _
                        GetFieldNameValue("CONTRACTNO", strCONTRACTNO), _
                        GetFieldNameValue("CONTRACTNAME", strCONTRACTNAME), _
                        GetFieldNameValue("GBN", strGBN), _
                        GetFieldNameValue("AMT", strAMT), _
                        GetFieldNameValue("STDATE", strSTDATE), _
                        GetFieldNameValue("EDDATE", strEDDATE), _
                        GetFieldNameValue("CONTRACTDAY", strCONTRACTDAY), _
                        GetFieldNameValue("CONFIRMFLAG", strCONFIRMFLAG), _
                        GetFieldNameValue("CONFIRMDATE", strCONFIRMDATE), _
                        GetFieldNameValue("CONFIRM_USER", strCONFIRM_USER), _
                        GetFieldNameValue("MEMO", strMEMO), _
                        GetFieldNameValue("CONDITION", strCONDITION), _
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
    Public Function DeleteDo(Optional ByVal strSEQ As String = OPTIONAL_STR) As Integer
        Dim strSQL As String

        Try
            strSQL = String.Format("DELETE FROM {0} WHERE {1}", EntityName, _
                     BuildFields("AND", _
                                   GetFieldNameValue("SEQ", strSEQ)))

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
    Public Function Update_ConfOK(Optional ByVal strSEQNO As String = OPTIONAL_STR, _
                                  Optional ByVal strCONTRACTNO As String = OPTIONAL_STR) As Integer
        Dim strSQL As String

        Try
            Dim strNOW As String '����Ʈ���� ó���� ������ �޾� �ؽ�Ʈ�� ó�� �Ѵ�.. 
            strNOW = Now

            strSQL = "  UPDATE SC_CONTRACT_MST"
            strSQL = strSQL & "  SET CONFIRMFLAG = 'Y',"
            strSQL = strSQL & "  CONFIRM_USER = '" & mobjSCGLConfig.WRKUSR & "', "
            strSQL = strSQL & "  CONFIRMDATE =  '" & strNOW & "', "
            strSQL = strSQL & "  CONTRACTNO = '" & strCONTRACTNO & "'"
            strSQL = strSQL & "  WHERE SEQ = '" & strSEQNO & "'"

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
    Public Function Update_ConfCAN(Optional ByVal strSEQNO As String = OPTIONAL_STR) As Integer
        Dim strSQL As String

        Try
            Dim strNOW As String '����Ʈ���� ó���� ������ �޾� �ؽ�Ʈ�� ó�� �Ѵ�.. 
            strNOW = Now

            strSQL = "  UPDATE SC_CONTRACT_MST"
            strSQL = strSQL & "  SET CONFIRMFLAG = 'N',"
            strSQL = strSQL & "  CONFIRM_USER = '', "
            strSQL = strSQL & "  CONFIRMDATE =  '', "
            strSQL = strSQL & "  CONTRACTNO = '' "
            strSQL = strSQL & "  WHERE SEQ = '" & strSEQNO & "'"

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
        MyBase.EntityName = "SC_CONTRACT_MST"     'Entity Name ����
    End Sub

    '���� ����� Base Class���� �����Ǿ� ����
#End Region

End Class