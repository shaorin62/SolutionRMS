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

Public Class ceMD_SENDTAXFILE_MST
    Inherits ceEntity

#Region "GROUP BLOCk : ���� �Ǵ� ��ⷹ���� ����/��� ����"
    Private Const CLASS_NAME = "ceMD_SENDTAXFILE_MST"    '�ڽ��� Ŭ������
#End Region

#Region "GROUP BLOCk : �ܺο� ���� Method"
#Region "SQL Insert/Update/Delete/Select"

    '*****************************************************************
    '�Է� : strSQL = SQL ��
    '��ȯ : ó���Ǽ�
    '��� : �ش� Entity�� Update ó��
    '���� : Key ���ǰ� Value Field����������(OPTIONAL_STR/OPTIONAL_NUM/OPTIONAL_DTM)
    '*****************************************************************
    Public Function FileInsertDo(Optional ByVal strYEARMON As String = OPTIONAL_STR, _
                                 Optional ByVal intSEQ As Double = OPTIONAL_NUM, _
                                 Optional ByVal strRMSNO As String = OPTIONAL_STR, _
                                 Optional ByVal strENDFLAG As String = OPTIONAL_STR, _
                                 Optional ByVal strSEND_GBN As String = OPTIONAL_STR) As Integer


        Dim strSQL As String
        Dim strFields As New System.Text.StringBuilder
        Dim strValues As New System.Text.StringBuilder
        Dim strNOW As String '����Ʈ���� ó���� ������ �޾� �ؽ�Ʈ�� ó�� �Ѵ�.. 
        strNOW = Now
        Try
            BuildNameValues(",", "YEARMON", strYEARMON, strFields, strValues)
            BuildNameValues(",", "SEQ", intSEQ, strFields, strValues)
            BuildNameValues(",", "RMSNO", strRMSNO, strFields, strValues)
            BuildNameValues(",", "ENDFLAG", strENDFLAG, strFields, strValues)
            BuildNameValues(",", "SEND_GBN", strSEND_GBN, strFields, strValues)
            BuildNameValues(",", "CUSER", mobjSCGLConfig.WRKUSR, strFields, strValues)
            BuildNameValues(",", "CDATE", strNOW, strFields, strValues)
            strSQL = String.Format("INSERT INTO {0} ({1}) VALUES({2})", EntityName, strFields, strValues)

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".InsertDo")
        End Try
    End Function

    '*****************************************************************
    '�Է� : strSQL = SQL ��
    '��ȯ : ó���Ǽ�
    '��� : �ش� Entity�� Delete ó��
    '���� : Key ������ ��������(OPTIONAL_STR/OPTIONAL_NUM/OPTIONAL_DTM)
    '*****************************************************************
    Public Function UpdateRtn_RMSNO_MD_TRUTAX_HDR(ByVal strTAXYEARMON, _
                                                  ByVal intTAXNO, _
                                                  ByVal strMEDFLAG, _
                                                  ByVal strRMSNO, _
                                                  ByVal strCANCEL_YN) As Integer
        Dim strSQL As String

        Try
            strSQL = " UPDATE MD_TRUTAX_HDR "
            strSQL = strSQL & "   SET RMSNO = '" & strRMSNO & "', ATTR04 = '" & strCANCEL_YN & "'"
            strSQL = strSQL & "   WHERE TAXYEARMON = '" & strTAXYEARMON & "' AND TAXNO = '" & intTAXNO & "' AND MEDFLAG = '" & strMEDFLAG & "'"

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".UPDATE_TRANS_DO")
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
        MyBase.EntityName = "MD_SENDTAXFILE_MST"     'Entity Name ����
    End Sub

    '���� ����� Base Class���� �����Ǿ� ����
#End Region
#End Region

End Class





