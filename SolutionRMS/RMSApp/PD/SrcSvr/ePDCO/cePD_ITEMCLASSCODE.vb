'****************************************************************************************
'�ý��۱��� : RMS/PD/Server Entity Class
'����  ȯ�� : GAC(Global Assembly Cache)
'���α׷��� : cePD_ITEMCLASSCODE.vb (PD_ITEMCODE Entity ó�� Class)
'��      �� : PD_ITEMCLASSCODE Entity�� ����Insert/Update/Delete/Select�� ó��
'             - �θ�ƼƼ ��ü�� SCGLUtil.ceEntity�� ���
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2008-12-15 
'****************************************************************************************
Imports SCGLUtil.cbSCGLUtil     '���� ��ƿ��Ƽ ��ü
Imports SCGLUtil.cbSCGLErr      '���� ����ó�� ��ü
Imports SCGLEntity              '��ƼƼ ��ü�� �θ� ��ü

Public Class cePD_ITEMCLASSCODE
    Inherits ceEntity

#Region "GROUP BLOCk : ���� �Ǵ� ��ⷹ���� ����/��� ����"
    Private Const CLASS_NAME = "cePD_ITEMCLASSCODE"    '�ڽ��� Ŭ������
#End Region

#Region "GROUP BLOCk : �ܺο� ���� Method"
#Region "SQL Insert/Update/Delete/Select"
    '*****************************************************************
    '�Է� : strSQL = SQL ��
    '��ȯ : ó���Ǽ�
    '��� : �ش� Entity�� Insert ó��
    '*****************************************************************
    'DIVCD,CLASSCD,DIVNM,CLASSNM
    Public Function InsertDo(ByVal strDIVCD As String, _
                             Optional ByVal strCLASSCD As String = OPTIONAL_STR, _
                             Optional ByVal strDIVNM As String = OPTIONAL_STR, _
                             Optional ByVal strCLASSNM As String = OPTIONAL_STR) As Integer

        Dim strSQL As String
        Dim strFields As New System.Text.StringBuilder
        Dim strValues As New System.Text.StringBuilder
        Dim strNOW As String '����Ʈ���� ó���� ������ �޾� �ؽ�Ʈ�� ó�� �Ѵ�.. 
        strNOW = Now
        Try
            BuildNameValues(",", "DIVCD", strDIVCD, strFields, strValues)
            BuildNameValues(",", "CLASSCD", strCLASSCD, strFields, strValues)
            BuildNameValues(",", "DIVNM", strDIVNM, strFields, strValues)
            BuildNameValues(",", "CLASSNM", strCLASSNM, strFields, strValues)
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
    'DIVCD,CLASSCD,DIVNM,CLASSNM
    Public Function UpdateDo(ByVal strDIVCD As String, _
                             Optional ByVal strCLASSCD As String = OPTIONAL_STR, _
                             Optional ByVal strDIVNM As String = OPTIONAL_STR, _
                             Optional ByVal strCLASSNM As String = OPTIONAL_STR) As Integer

        Dim strSQL As String
        Dim strNOW As String '����Ʈ���� ó���� ������ �޾� �ؽ�Ʈ�� ó�� �Ѵ�.. 
        strNOW = Now
        Try
            strSQL = String.Format("UPDATE {0} SET {1} WHERE {2}", EntityName, _
                     BuildFields(",", _
                        GetFieldNameValue("DIVCD", strDIVCD), _
                        GetFieldNameValue("CLASSCD", strCLASSCD), _
                        GetFieldNameValue("DIVNM", strDIVNM), _
                        GetFieldNameValue("CLASSNM", strCLASSNM), _
                        GetFieldNameValue("UUSER", mobjSCGLConfig.WRKUSR), _
                        GetFieldNameValue("UDATE", strNOW)), _
                     BuildFields("AND", _
                        GetFieldNameValue("DIVCD", strDIVCD), GetFieldNameValue("CLASSCD", strCLASSCD)))

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
    'dblSEQ, strJOBNO, strPREESTNO
    Public Function DeleteDo(ByVal strDIVCD As String, _
                             ByVal strCLASSCD As String) As Integer
        Dim strSQL As String
        Try
            strSQL = String.Format("DELETE FROM {0} WHERE {1}", EntityName, _
                     BuildFields("AND", _
                                   GetFieldNameValue("DIVCD", strDIVCD), GetFieldNameValue("CLASSCD", strCLASSCD)))

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
    'Update_Batch
    '*****************************************************************
    '�Է� : strSQL = SQL ��
    '��ȯ : ó���Ǽ�
    '��� : strSQL �ش繮�� �״�� ó�� (�������� Ȯ�� ó���Ѵ�)
    '*****************************************************************
    Public Function Update_Batch(ByVal strDIVCD As String, _
                                 ByVal strCODE As String, _
                                 ByVal strNAME As String) As Integer
        Dim strSQL As String
        Try
            strSQL = "UPDATE PD_ITEMCLASSCODE SET CLASSNM = '" & strNAME & "' WHERE DIVCD = '" & strDIVCD & "' AND CLASSCD = '" & strCODE & "'"
            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".Update_Batch")
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
        MyBase.EntityName = "PD_ITEMCLASSCODE"     'Entity Name ����
    End Sub

    '���� ����� Base Class���� �����Ǿ� ����
#End Region
#End Region
End Class

'------->>��ƼƼ INSERT/UPDATE �����Դϴ�. �ݵ�� �ڽ��� ȯ�濡 ���߾ �����Ͻñ� �ٶ��ϴ�.
'=========================================================
'       'vntData Array�� ����� �� Insert/Update �Դϴ�.
'=========================================================
'        Dim intRtn As Integer
'        intRtn = mobjceSC_JOBCUST.InsertDo( _
'                                       GetElement(vntData,"SEQNO", intColCnt, intRow), _
'                                       GetElement(vntData,"SEQNAME", intColCnt, intRow), _
'                                       GetElement(vntData,"CUSTCODE", intColCnt, intRow), _
'                                       GetElement(vntData,"ACCCUSTCODE", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"DEPTCD", intColCnt, intRow), _
'                                       GetElement(vntData,"ATTR01", intColCnt, intRow), _
'                                       GetElement(vntData,"ATTR02", intColCnt, intRow), _
'                                       GetElement(vntData,"ATTR03", intColCnt, intRow), _
'                                       GetElement(vntData,"ATTR04", intColCnt, intRow), _
'                                       GetElement(vntData,"ATTR05", intColCnt, intRow), _
'                                       GetElement(vntData,"ATTR06", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"ATTR07", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"ATTR08", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"ATTR09", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"ATTR10", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"CUSER", intColCnt, intRow), _
'                                       GetElement(vntData,"CDATE", intColCnt, intRow, NULL_DTM, true ), _
'                                       GetElement(vntData,"UUSER", intColCnt, intRow), _
'                                       GetElement(vntData,"UDATE", intColCnt, intRow, NULL_DTM, true ) _
'                                       )
'        Return intRtn

'        Dim intRtn As Integer
'        intRtn = mobjceSC_JOBCUST.UpdateDo( _
'                                       GetElement(vntData,"SEQNO", intColCnt, intRow), _
'                                       GetElement(vntData,"SEQNAME", intColCnt, intRow), _
'                                       GetElement(vntData,"CUSTCODE", intColCnt, intRow), _
'                                       GetElement(vntData,"ACCCUSTCODE", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"DEPTCD", intColCnt, intRow), _
'                                       GetElement(vntData,"ATTR01", intColCnt, intRow), _
'                                       GetElement(vntData,"ATTR02", intColCnt, intRow), _
'                                       GetElement(vntData,"ATTR03", intColCnt, intRow), _
'                                       GetElement(vntData,"ATTR04", intColCnt, intRow), _
'                                       GetElement(vntData,"ATTR05", intColCnt, intRow), _
'                                       GetElement(vntData,"ATTR06", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"ATTR07", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"ATTR08", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"ATTR09", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"ATTR10", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"CUSER", intColCnt, intRow), _
'                                       GetElement(vntData,"CDATE", intColCnt, intRow, NULL_DTM, true ), _
'                                       GetElement(vntData,"UUSER", intColCnt, intRow), _
'                                       GetElement(vntData,"UDATE", intColCnt, intRow, NULL_DTM, true ) _
'                                       )
'        Return intRtn


'=========================================================
'       'XmlData �� ����� �� Insert/Update �Դϴ�.
'=========================================================
'        Dim intRtn As Integer
'        intRtn = mobjceSC_JOBCUST.InsertDo( _
'                                       XMLGetElement(xmlRoot,"SEQNO"), _
'                                       XMLGetElement(xmlRoot,"SEQNAME"), _
'                                       XMLGetElement(xmlRoot,"CUSTCODE"), _
'                                       XMLGetElement(xmlRoot,"ACCCUSTCODE", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"DEPTCD"), _
'                                       XMLGetElement(xmlRoot,"ATTR01"), _
'                                       XMLGetElement(xmlRoot,"ATTR02"), _
'                                       XMLGetElement(xmlRoot,"ATTR03"), _
'                                       XMLGetElement(xmlRoot,"ATTR04"), _
'                                       XMLGetElement(xmlRoot,"ATTR05"), _
'                                       XMLGetElement(xmlRoot,"ATTR06", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"ATTR07", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"ATTR08", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"ATTR09", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"ATTR10", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"CUSER"), _
'                                       XMLGetElement(xmlRoot,"CDATE", NULL_DTM, true ), _
'                                       XMLGetElement(xmlRoot,"UUSER"), _
'                                       XMLGetElement(xmlRoot,"UDATE", NULL_DTM, true ) _
'                                       )
'        Return intRtn

'        Dim intRtn As Integer
'        intRtn = mobjceSC_JOBCUST.UpdateDo( _
'                                       XMLGetElement(xmlRoot,"SEQNO"), _
'                                       XMLGetElement(xmlRoot,"SEQNAME"), _
'                                       XMLGetElement(xmlRoot,"CUSTCODE"), _
'                                       XMLGetElement(xmlRoot,"ACCCUSTCODE", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"DEPTCD"), _
'                                       XMLGetElement(xmlRoot,"ATTR01"), _
'                                       XMLGetElement(xmlRoot,"ATTR02"), _
'                                       XMLGetElement(xmlRoot,"ATTR03"), _
'                                       XMLGetElement(xmlRoot,"ATTR04"), _
'                                       XMLGetElement(xmlRoot,"ATTR05"), _
'                                       XMLGetElement(xmlRoot,"ATTR06", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"ATTR07", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"ATTR08", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"ATTR09", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"ATTR10", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"CUSER"), _
'                                       XMLGetElement(xmlRoot,"CDATE", NULL_DTM, true ), _
'                                       XMLGetElement(xmlRoot,"UUSER"), _
'                                       XMLGetElement(xmlRoot,"UDATE", NULL_DTM, true ) _
'                                       )
'        Return intRtn


