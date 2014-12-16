
'****************************************************************************************
'�ý��۱��� : RMS/PD/Server Entity Class
'����  ȯ�� : 
'���α׷��� : cePD_SUBITEM_DTL.vb (PD_SUBITEM_DTL Entity ó�� Class)
'��      �� : PD_SUBITEM_DTL Entity�� ����Insert/Update/Delete/Select�� ó��
'             - �θ�ƼƼ ��ü�� SCGLUtil.ceEntity�� ���
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2009-10-19 
'****************************************************************************************
Imports SCGLUtil.cbSCGLUtil     '���� ��ƿ��Ƽ ��ü
Imports SCGLUtil.cbSCGLErr      '���� ����ó�� ��ü
Imports SCGLEntity              '��ƼƼ ��ü�� �θ� ��ü

Public Class cePD_ESTTYPE_HDR
    Inherits ceEntity

#Region "GROUP BLOCk : ���� �Ǵ� ��ⷹ���� ����/��� ����"
    Private Const CLASS_NAME = "cePD_ESTTYPE_HDR"    '�ڽ��� Ŭ������
#End Region

#Region "GROUP BLOCk : �ܺο� ���� Method"
#Region "SQL Insert/Update/Delete/Select"
    '*****************************************************************
    '�Է� : strSQL = SQL ��
    '��ȯ : ó���Ǽ�
    '��� : �ش� Entity�� Insert ó��
    '*****************************************************************
    'SEQ,TYPENAME,CLIENTCODE,COMMISSION,NONECOMMISSION,SUSURATE,SUSUAMT,AMT,SUMAMT
    Public Function InsertDo(ByVal dblSEQ As Double, _
            Optional ByVal strTYPENAME As String = OPTIONAL_STR, _
            Optional ByVal strCLIENTCODE As String = OPTIONAL_STR, _
            Optional ByVal dblCOMMISSION As Double = OPTIONAL_NUM, _
            Optional ByVal dblNONECOMMISSION As Double = OPTIONAL_NUM, _
            Optional ByVal dblSUSURATE As Double = OPTIONAL_NUM, _
            Optional ByVal dblSUSUAMT As Double = OPTIONAL_NUM, _
            Optional ByVal dblAMT As Double = OPTIONAL_NUM, _
            Optional ByVal dblSUMAMT As Double = OPTIONAL_NUM, _
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
            BuildNameValues(",", "SEQ", dblSEQ, strFields, strValues)
            BuildNameValues(",", "TYPENAME", strTYPENAME, strFields, strValues)
            BuildNameValues(",", "CLIENTCODE", strCLIENTCODE, strFields, strValues)
            BuildNameValues(",", "COMMISSION", dblCOMMISSION, strFields, strValues)
            BuildNameValues(",", "NONECOMMISSION", dblNONECOMMISSION, strFields, strValues)
            BuildNameValues(",", "SUSURATE", dblSUSURATE, strFields, strValues)
            BuildNameValues(",", "SUSUAMT", dblSUSUAMT, strFields, strValues)
            BuildNameValues(",", "AMT", dblAMT, strFields, strValues)
            BuildNameValues(",", "SUMAMT", dblSUMAMT, strFields, strValues)
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
    'SEQ,TYPENAME,CLIENTCODE,COMMISSION,NONECOMMISSION,SUSURATE,SUSUAMT,AMT,SUMAMT
    Public Function UpdateDo(ByVal dblSEQ As Double, _
            Optional ByVal strTYPENAME As String = OPTIONAL_STR, _
            Optional ByVal strCLIENTCODE As String = OPTIONAL_STR, _
            Optional ByVal dblCOMMISSION As Double = OPTIONAL_NUM, _
            Optional ByVal dblNONECOMMISSION As Double = OPTIONAL_NUM, _
            Optional ByVal dblSUSURATE As Double = OPTIONAL_NUM, _
            Optional ByVal dblSUSUAMT As Double = OPTIONAL_NUM, _
            Optional ByVal dblAMT As Double = OPTIONAL_NUM, _
            Optional ByVal dblSUMAMT As Double = OPTIONAL_NUM, _
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
                        GetFieldNameValue("TYPENAME", strTYPENAME), _
                        GetFieldNameValue("CLIENTCODE", strCLIENTCODE), _
                        GetFieldNameValue("COMMISSION", dblCOMMISSION), _
                        GetFieldNameValue("NONECOMMISSION", dblNONECOMMISSION), _
                        GetFieldNameValue("SUSURATE", dblSUSURATE), _
                        GetFieldNameValue("SUSUAMT", dblSUSUAMT), _
                        GetFieldNameValue("AMT", dblAMT), _
                        GetFieldNameValue("SUMAMT", dblSUMAMT), _
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
                        GetFieldNameValue("SEQ", dblSEQ)))

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
    'strPREESTNO,dblSEQ,dblSUBITEMCODESEQ,dblITEMCODESEQ,strITEMCODE
    Public Function DeleteDo(ByVal dblSEQ As Double) As Integer
        Dim strSQL As String

        Try
            strSQL = "DELETE FROM PD_ESTTYPE_HDR WHERE SEQ = " & dblSEQ
            strSQL = strSQL & ";DELETE FROM PD_ESTTYPE_DTL WHERE HDRSEQ = " & dblSEQ
            strSQL = strSQL & ";DELETE FROM PD_ESTSUBTYPE WHERE HDRSEQ = " & dblSEQ
            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".DeleteDo")
        End Try
    End Function

    '*****************************************************************
    '�Է� : strSQL = SQL ��
    '��ȯ : ó���Ǽ�
    '��� : �ش� Entity�� Select ó��
    '*****************************************************************
    Public Function SelectDo(ByRef intRowCnt As Integer, ByRef intColCnt As Integer, _
                             Optional ByVal strSEQNO As String = OPTIONAL_STR, _
                             Optional ByVal strSelFields As String = "*", _
                             Optional ByVal intLimitRow As Integer = 0, _
                             Optional ByVal intSelMode As Integer = SELMODE.ARR, _
                             Optional ByVal blnBindingHeader As Boolean = False) As Object
        Dim strSQL As String
        Dim strKeyFields As String

        Try
            strKeyFields = BuildFields("AND", _
                                    GetFieldNameValue("SEQNO", strSEQNO))

            Return SelectDoExt(intRowCnt, intColCnt, strSelFields, strKeyFields, intLimitRow, intSelMode, blnBindingHeader)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".SeleteDo")
        End Try
    End Function
    '*****************************************************************
    '�Է� : strSQL = SQL ��
    '��ȯ : ó���Ǽ�
    '��� : strSQL �ش繮�� �״�� ó�� û������ ����� ���ֺ� ���� ��� ������ ������Ʈ
    '*****************************************************************
    Public Function SqlExe(ByVal strTEMPSQL As String) As Integer
        Dim strSQL As String
        Try
            strSQL = strTEMPSQL
            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".SqlExe")
        End Try
    End Function

    '*****************************************************************
    '�Է� : strSQL = SQL ��
    '��ȯ : ó���Ǽ�
    '��� : strSQL �ش繮�� �״�� ó�� TEMP���� �����Ѵ�.
    '*****************************************************************
    Public Function TempDeleteDo(ByVal strUSER) As String
        Dim strSQL As String
        Try
            strSQL = "DELETE FROM PD_ESTSUBTYPE_INPUT WHERE CUSER = '" & strUSER & "'"
            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".TempDeleteDo")
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
        MyBase.EntityName = "PD_ESTTYPE_HDR"     'Entity Name ����
    End Sub

    '���� ����� Base Class���� �����Ǿ� ����
#End Region
#End Region
End Class