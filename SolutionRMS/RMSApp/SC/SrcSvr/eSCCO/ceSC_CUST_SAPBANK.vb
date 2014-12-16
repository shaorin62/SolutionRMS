'****************************************************************************************
'Generated By: JNF
'�ý��۱��� : RMS/SC/Server Entity Class
'����  ȯ�� : GAC(Global Assembly Cache)
'���α׷��� : ceSC_FEE_MST.vb ( SC_FEE_MST Entity ó�� Class)
'��      �� : SC_FEE_MST Entity�� ����Insert/Update/Delete/Select�� ó��
'             - �θ�ƼƼ ��ü�� SCGLUtil.ceEntity�� ���
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2009-08-20 ���� 06:18 By Making Entity Bean
'****************************************************************************************
Imports SCGLUtil.cbSCGLUtil     '���� ��ƿ��Ƽ ��ü
Imports SCGLUtil.cbSCGLErr      '���� ����ó�� ��ü
Imports SCGLEntity              '��ƼƼ ��ü�� �θ� ��ü

Public Class ceSC_CUST_SAPBANK
    Inherits ceEntity

#Region "GROUP BLOCk : ���� �Ǵ� ��ⷹ���� ����/��� ����"
    Private Const CLASS_NAME = "ceSC_CUST_SAPBANK"    '�ڽ��� Ŭ������
#End Region

#Region "GROUP BLOCk : �ܺο� ���� Method"
#Region "SQL Insert/Update/Delete/Select"

    '*****************************************************************
    '�Է� : strSQL = SQL ��
    '��ȯ : ó���Ǽ�
    '��� : �ش� Entity�� Insert ó��
    '*****************************************************************
    Public Function InsertDo(ByVal strSAUPNO As String, _
            Optional ByVal strBVTYP As String = OPTIONAL_STR, _
            Optional ByVal strBANKL As String = OPTIONAL_STR, _
            Optional ByVal strBANKN As String = OPTIONAL_STR, _
            Optional ByVal strKOINH As String = OPTIONAL_STR)

        Dim strSQL As String
        Dim strFields As New System.Text.StringBuilder
        Dim strValues As New System.Text.StringBuilder
        Dim strNOW As String '����Ʈ���� ó���� ������ �޾� �ؽ�Ʈ�� ó�� �Ѵ�.. 
        strNOW = Now
        Try
            BuildNameValues(",", "SAUPNO", strSAUPNO, strFields, strValues)
            BuildNameValues(",", "BVTYP", strBVTYP, strFields, strValues)
            BuildNameValues(",", "BANKL", strBANKL, strFields, strValues)
            BuildNameValues(",", "BANKN", strBANKN, strFields, strValues)
            BuildNameValues(",", "KOINH", strKOINH, strFields, strValues)

            strSQL = String.Format("INSERT INTO {0} ({1}) VALUES({2})", EntityName, strFields, strValues)

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".InsertDo")
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
        MyBase.EntityName = "SC_CUST_SAPBANK"     'Entity Name ����
    End Sub

    '���� ����� Base Class���� �����Ǿ� ����
#End Region
#End Region

End Class






