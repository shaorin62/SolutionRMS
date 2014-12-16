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

Public Class cePD_CONTRACT_BASE
    Inherits ceEntity

#Region "GROUP BLOCk : ���� �Ǵ� ��ⷹ���� ����/��� ����"
    Private Const CLASS_NAME = "cePD_CONTRACT_BASE"    '�ڽ��� Ŭ������
#End Region

#Region "GROUP BLOCk : �ܺο� ���� Method"
#Region "SQL Insert/Update/Delete/Select"
    '*****************************************************************
    '�Է� : strSQL = SQL ��
    '��ȯ : ó���Ǽ�
    '��� : �ش� Entity�� Insert ó��
    '*****************************************************************
    Public Function InsertDo(ByVal strCONTRACTNO As String, _
            Optional ByVal strOUTSCODE As String = OPTIONAL_STR, _
            Optional ByVal strCONTRACTNAME As String = OPTIONAL_STR, _
            Optional ByVal strGBN As String = OPTIONAL_STR, _
            Optional ByVal dblAMT As Double = OPTIONAL_NUM, _
            Optional ByVal strSTDATE As String = OPTIONAL_STR, _
            Optional ByVal strEDDATE As String = OPTIONAL_STR, _
            Optional ByVal strCONTRACTDAY As String = OPTIONAL_STR, _
            Optional ByVal strMANAGER As String = OPTIONAL_STR, _
            Optional ByVal strCONFIRMFLAG As String = OPTIONAL_STR, _
            Optional ByVal strCOMENT As String = OPTIONAL_STR, _
            Optional ByVal strGUBUN As String = OPTIONAL_STR, _
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
            BuildNameValues(",", "CONTRACTNO", strCONTRACTNO, strFields, strValues)
            BuildNameValues(",", "OUTSCODE", strOUTSCODE, strFields, strValues)
            BuildNameValues(",", "CONTRACTNAME", strCONTRACTNAME, strFields, strValues)
            BuildNameValues(",", "GBN", strGBN, strFields, strValues)
            BuildNameValues(",", "AMT", dblAMT, strFields, strValues)
            BuildNameValues(",", "STDATE", strSTDATE, strFields, strValues)
            BuildNameValues(",", "EDDATE", strEDDATE, strFields, strValues)
            BuildNameValues(",", "CONTRACTDAY", strCONTRACTDAY, strFields, strValues)
            BuildNameValues(",", "MANAGER", strMANAGER, strFields, strValues)
            BuildNameValues(",", "CONFIRMFLAG", strCONFIRMFLAG, strFields, strValues)
            BuildNameValues(",", "COMENT", strCOMENT, strFields, strValues)
            BuildNameValues(",", "GUBUN", strGUBUN, strFields, strValues)
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
    Public Function UpdateDo(ByVal strCONTRACTNO As String, _
            Optional ByVal strOUTSCODE As String = OPTIONAL_STR, _
            Optional ByVal strCONTRACTNAME As String = OPTIONAL_STR, _
            Optional ByVal strGBN As String = OPTIONAL_STR, _
            Optional ByVal dblAMT As Double = OPTIONAL_NUM, _
            Optional ByVal strSTDATE As String = OPTIONAL_STR, _
            Optional ByVal strEDDATE As String = OPTIONAL_STR, _
            Optional ByVal strCONTRACTDAY As String = OPTIONAL_STR, _
            Optional ByVal strMANAGER As String = OPTIONAL_STR, _
            Optional ByVal strCONFIRMFLAG As String = OPTIONAL_STR, _
            Optional ByVal strCOMENT As String = OPTIONAL_STR, _
            Optional ByVal strGUBUN As String = OPTIONAL_STR) As Integer

        Dim strSQL As String
        Dim strNOW As String '����Ʈ���� ó���� ������ �޾� �ؽ�Ʈ�� ó�� �Ѵ�.. 
        strNOW = Now
        Try

            strSQL = String.Format("UPDATE {0} SET {1} WHERE {2}", EntityName, _
                     BuildFields(",", _
                        GetFieldNameValue("CONTRACTNAME", strCONTRACTNAME), _
                        GetFieldNameValue("GBN", strGBN), _
                        GetFieldNameValue("OUTSCODE", strOUTSCODE), _
                        GetFieldNameValue("AMT", dblAMT), _
                        GetFieldNameValue("STDATE", strSTDATE), _
                        GetFieldNameValue("EDDATE", strEDDATE), _
                        GetFieldNameValue("CONTRACTDAY", strCONTRACTDAY), _
                        GetFieldNameValue("MANAGER", strMANAGER), _
                        GetFieldNameValue("CONFIRMFLAG", strCONFIRMFLAG), _
                        GetFieldNameValue("COMENT", strCOMENT), _
                        GetFieldNameValue("GUBUN", strGUBUN), _
                        GetFieldNameValue("UUSER", mobjSCGLConfig.WRKUSR), _
                        GetFieldNameValue("UDATE", strNOW)), _
                     BuildFields("AND", _
                        GetFieldNameValue("CONTRACTNO", strCONTRACTNO)))

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".UpdateDo")
        End Try
    End Function

    '*****************************************************************
    '�Է� : strSQL = SQL ��
    '��ȯ : ó���Ǽ�
    '��� : �ش� Entity�� Delete ó��
    '���� : Key ������ ��������
    '*****************************************************************
    Public Function UpdateDo_CONFIRM(ByVal strCONTRACTNO As String, _
                                     ByVal strCONFIRMFLAG As String) As Integer
        Dim strSQL As String
        Dim strNOW As String
        strNOW = Now

        Try
            strSQL = " UPDATE PD_CONTRACT_BASE "
            strSQL = strSQL & " SET CONFIRMFLAG = '" & strCONFIRMFLAG & "' "
            strSQL = strSQL & " WHERE CONTRACTNO = '" & strCONTRACTNO & "'"

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".UpdateDo_CONFIRM")
        End Try
    End Function

    '*****************************************************************
    '�Է� : strSQL = SQL ��
    '��ȯ : ó���Ǽ�
    '��� : �ش� Entity�� Delete ó��
    '���� : Key ������ ��������
    '*****************************************************************
    Public Function DeleteDo(ByVal strCONTRACTNO As String) As Integer
        Dim strSQL As String
        Dim strNOW As String
        strNOW = Now

        Try
            strSQL = "DELETE FROM PD_CONTRACT_BASE WHERE CONTRACTNO = '" & strCONTRACTNO & "';"

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
        MyBase.EntityName = "PD_CONTRACT_BASE"     'Entity Name ����
    End Sub

#End Region
#End Region
End Class
