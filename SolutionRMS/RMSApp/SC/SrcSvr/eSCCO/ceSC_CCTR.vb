'Public Class ceSC_BM

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

Public Class ceSC_CCTR
    Inherits ceEntity

#Region "GROUP BLOCk : ���� �Ǵ� ��ⷹ���� ����/��� ����"
    Private Const CLASS_NAME = "ceSC_CCTR"    '�ڽ��� Ŭ������
#End Region

#Region "GROUP BLOCk : �ܺο� ���� Method"
#Region "SQL Insert/Update/Delete/Select"
    '*****************************************************************
    '�Է� : strSQL = SQL ��
    '��ȯ : ó���Ǽ�
    '��� : �ش� Entity�� Insert ó��
    '*****************************************************************
    Public Function InsertDo(ByVal strBMCODE As String, _
                             Optional ByVal strHIGHDEPT_CD As String = OPTIONAL_STR, _
                             Optional ByVal strDEPT_CD As String = OPTIONAL_STR, _
                             Optional ByVal strCCTR As String = OPTIONAL_STR, _
                             Optional ByVal strBA As String = OPTIONAL_STR, _
                             Optional ByVal strFDATE As String = OPTIONAL_STR, _
                             Optional ByVal strTDATE As String = OPTIONAL_STR, _
                             Optional ByVal strUSE_YN As String = OPTIONAL_STR) As Integer

        Dim strSQL As String
        Dim strFields As New System.Text.StringBuilder
        Dim strValues As New System.Text.StringBuilder
        Dim strNOW As String '����Ʈ���� ó���� ������ �޾� �ؽ�Ʈ�� ó�� �Ѵ�.. 
        strNOW = Now
        Try
            BuildNameValues(",", "BMCODE", strBMCODE, strFields, strValues)
            BuildNameValues(",", "HIGHDEPT_CD", strHIGHDEPT_CD, strFields, strValues)
            BuildNameValues(",", "DEPT_CD", strDEPT_CD, strFields, strValues)
            BuildNameValues(",", "CCTR", strCCTR, strFields, strValues)
            BuildNameValues(",", "BA", strBA, strFields, strValues)
            BuildNameValues(",", "FDATE", strFDATE, strFields, strValues)
            BuildNameValues(",", "TDATE", strTDATE, strFields, strValues)
            BuildNameValues(",", "USE_YN", strUSE_YN, strFields, strValues)
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
    Public Function UpdateDo(ByVal dblSEQ As Double, _
                             ByVal strBMCODE As String, _
                             Optional ByVal strHIGHDEPT_CD As String = OPTIONAL_STR, _
                             Optional ByVal strDEPT_CD As String = OPTIONAL_STR, _
                             Optional ByVal strCCTR As String = OPTIONAL_STR, _
                             Optional ByVal strBA As String = OPTIONAL_STR, _
                             Optional ByVal strFDATE As String = OPTIONAL_STR, _
                             Optional ByVal strTDATE As String = OPTIONAL_STR, _
                             Optional ByVal strUSE_YN As String = OPTIONAL_STR) As Integer


        Dim strSQL As String
        Dim strNOW As String '����Ʈ���� ó���� ������ �޾� �ؽ�Ʈ�� ó�� �Ѵ�.. 
        strNOW = Now
        Try
            strSQL = ""

            strSQL = String.Format("UPDATE {0} SET {1} WHERE {2}", EntityName, _
                   BuildFields(",", _
                       GetFieldNameValue("BMCODE", strBMCODE), _
                       GetFieldNameValue("HIGHDEPT_CD", strHIGHDEPT_CD), _
                       GetFieldNameValue("DEPT_CD", strDEPT_CD), _
                       GetFieldNameValue("CCTR", strCCTR), _
                       GetFieldNameValue("BA", strBA), _
                       GetFieldNameValue("FDATE", strFDATE), _
                       GetFieldNameValue("TDATE", strTDATE), _
                       GetFieldNameValue("USE_YN", strUSE_YN)), _
                   BuildFields("AND", _
                     GetFieldNameValue("SEQ", dblSEQ)))

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".UpdateDo")
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
        MyBase.EntityName = "SC_CCTR"     'Entity Name ����
    End Sub

    '���� ����� Base Class���� �����Ǿ� ����
#End Region
#End Region
End Class