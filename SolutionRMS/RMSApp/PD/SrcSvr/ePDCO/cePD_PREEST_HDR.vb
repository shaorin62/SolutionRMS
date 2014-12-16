'Public Class cePD_PREEST_HDR

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

Public Class cePD_PREEST_HDR
    Inherits ceEntity

#Region "GROUP BLOCk : ���� �Ǵ� ��ⷹ���� ����/��� ����"
    Private Const CLASS_NAME = "cePD_PREEST_HDR"    '�ڽ��� Ŭ������
#End Region

#Region "GROUP BLOCk : �ܺο� ���� Method"
#Region "SQL Insert/Update/Delete/Select"
    '*****************************************************************
    '�Է� : strSQL = SQL ��
    '��ȯ : ó���Ǽ�
    '��� : �ش� Entity�� Insert ó��
    '*****************************************************************
    'PREESTNO,PREESTNAME,JOBNO,JOBNAME,AMT,MEMO,CREDAY
    Public Function InsertDo(ByVal strPREESTNO As String, _
            Optional ByVal strPREESTNAME As String = OPTIONAL_STR, _
            Optional ByVal strJOBNO As String = OPTIONAL_STR, _
            Optional ByVal strMEMO As String = OPTIONAL_STR, _
            Optional ByVal strCREDAY As String = OPTIONAL_STR, _
            Optional ByVal strCLIENTSUBCODE As String = OPTIONAL_STR, _
            Optional ByVal dblSUSURATE As Double = OPTIONAL_NUM, _
            Optional ByVal strCLIENTCODE As String = OPTIONAL_STR, _
            Optional ByVal strSUBSEQ As String = OPTIONAL_STR, _
            Optional ByVal strATTR01 As String = OPTIONAL_STR, _
            Optional ByVal strATTR02 As String = OPTIONAL_STR, _
            Optional ByVal strATTR03 As String = OPTIONAL_STR, _
            Optional ByVal strATTR04 As String = OPTIONAL_STR, _
            Optional ByVal strATTR05 As String = OPTIONAL_STR, _
            Optional ByVal dblATTR06 As Double = OPTIONAL_NUM, _
            Optional ByVal dblATTR07 As Double = OPTIONAL_NUM, _
            Optional ByVal dblATTR08 As Double = OPTIONAL_NUM, _
            Optional ByVal dblATTR09 As Double = OPTIONAL_NUM, _
            Optional ByVal dblATTR10 As Double = OPTIONAL_NUM)

        Dim strSQL As String
        Dim strFields As New System.Text.StringBuilder
        Dim strValues As New System.Text.StringBuilder
        Dim strNOW As String '����Ʈ���� ó���� ������ �޾� �ؽ�Ʈ�� ó�� �Ѵ�.. 
        strNOW = Now
        Try
            BuildNameValues(",", "PREESTNO", strPREESTNO, strFields, strValues)
            BuildNameValues(",", "PREESTNAME", strPREESTNAME, strFields, strValues)
            BuildNameValues(",", "JOBNO", strJOBNO, strFields, strValues)
            BuildNameValues(",", "MEMO", strMEMO, strFields, strValues)
            BuildNameValues(",", "CREDAY", strCREDAY, strFields, strValues)
            BuildNameValues(",", "CLIENTSUBCODE", strCLIENTSUBCODE, strFields, strValues)
            BuildNameValues(",", "SUSURATE", dblSUSURATE, strFields, strValues)
            BuildNameValues(",", "CLIENTCODE", strCLIENTCODE, strFields, strValues)
            BuildNameValues(",", "SUBSEQ", strSUBSEQ, strFields, strValues)
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
    'PREESTNO,PREESTNAME,JOBNO,JOBNAME,AMT,MEMO,CREDAY
    Public Function UpdateDo(Optional ByVal strPREESTNO As String = OPTIONAL_STR, _
            Optional ByVal strPREESTNAME As String = OPTIONAL_STR, _
            Optional ByVal strJOBNO As String = OPTIONAL_STR, _
            Optional ByVal strMEMO As String = OPTIONAL_STR, _
            Optional ByVal strCREDAY As String = OPTIONAL_STR, _
            Optional ByVal strCLIENTSUBCODE As String = OPTIONAL_STR, _
            Optional ByVal dblSUSURATE As Double = OPTIONAL_NUM, _
            Optional ByVal strCLIENTCODE As String = OPTIONAL_STR, _
            Optional ByVal strSUBSEQ As String = OPTIONAL_STR, _
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
                        GetFieldNameValue("PREESTNAME", strPREESTNAME), _
                        GetFieldNameValue("JOBNO", strJOBNO), _
                        GetFieldNameValue("MEMO", strMEMO), _
                        GetFieldNameValue("CREDAY", strCREDAY), _
                        GetFieldNameValue("CLIENTSUBCODE", strCLIENTSUBCODE), _
                        GetFieldNameValue("SUSURATE", dblSUSURATE), _
                        GetFieldNameValue("CLIENTCODE", strCLIENTCODE), _
                        GetFieldNameValue("SUBSEQ", strSUBSEQ), _
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
                        GetFieldNameValue("PREESTNO", strPREESTNO)))

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".UpdateDo")
        End Try
    End Function
    Public Function InsertDo_HDR(ByVal strPREESTNO As String, _
                                 Optional ByVal strPREESTNAME As String = OPTIONAL_STR, _
                                 Optional ByVal strJOBNO As String = OPTIONAL_STR, _
                                 Optional ByVal strTIMCODE As String = OPTIONAL_STR, _
                                 Optional ByVal dblSUSURATE As Double = OPTIONAL_NUM, _
                                 Optional ByVal strCLIENTCODE As String = OPTIONAL_STR, _
                                 Optional ByVal dblSUSUAMT As Double = OPTIONAL_NUM, _
                                 Optional ByVal dblCOMMITION As Double = OPTIONAL_NUM, _
                                 Optional ByVal dblNONCOMMITION As Double = OPTIONAL_NUM, _
                                 Optional ByVal dblSUMAMT As Double = OPTIONAL_NUM, _
                                 Optional ByVal strSUBSEQ As String = OPTIONAL_STR, _
                                 Optional ByVal strMEMO As String = OPTIONAL_STR, _
                                 Optional ByVal strCONFIRMFLAG As String = OPTIONAL_STR, _
                                 Optional ByVal strCONFIRMGBN As String = OPTIONAL_STR, _
                                 Optional ByVal dblAMT As Double = OPTIONAL_NUM) As Integer
        'strPREESTNO,PREESTNAME,JOBNO,CREDAY,CLIENTSUBCODE,SUSURATE,CLIENTCODE
        'SUSUAMT,COMMITION,NONCOMMITION,SUMAMT
        Dim strSQL As String
        Dim strFields As New System.Text.StringBuilder
        Dim strValues As New System.Text.StringBuilder
        Dim strNOW As String '����Ʈ���� ó���� ������ �޾� �ؽ�Ʈ�� ó�� �Ѵ�.. 
        strNOW = Now
        Try
            BuildNameValues(",", "PREESTNO", strPREESTNO, strFields, strValues)
            BuildNameValues(",", "PREESTNAME", strPREESTNAME, strFields, strValues)
            BuildNameValues(",", "JOBNO", strJOBNO, strFields, strValues)
            BuildNameValues(",", "TIMCODE", strTIMCODE, strFields, strValues)
            BuildNameValues(",", "SUSURATE", dblSUSURATE, strFields, strValues)
            BuildNameValues(",", "CLIENTCODE", strCLIENTCODE, strFields, strValues)
            BuildNameValues(",", "SUSUAMT", dblSUSUAMT, strFields, strValues)
            BuildNameValues(",", "COMMITION", dblCOMMITION, strFields, strValues)
            BuildNameValues(",", "NONCOMMITION", dblNONCOMMITION, strFields, strValues)
            BuildNameValues(",", "SUMAMT", dblSUMAMT, strFields, strValues)
            BuildNameValues(",", "SUBSEQ", strSUBSEQ, strFields, strValues)
            BuildNameValues(",", "MEMO", strMEMO, strFields, strValues)
            BuildNameValues(",", "CONFIRMFLAG", strCONFIRMFLAG, strFields, strValues)
            BuildNameValues(",", "CONFIRMGBN", strCONFIRMGBN, strFields, strValues)
            BuildNameValues(",", "AMT", dblAMT, strFields, strValues)
            BuildNameValues(",", "CUSER", mobjSCGLConfig.WRKUSR, strFields, strValues)
            BuildNameValues(",", "CDATE", strNOW, strFields, strValues)
            BuildNameValues(",", "UUSER", mobjSCGLConfig.WRKUSR, strFields, strValues)
            BuildNameValues(",", "UDATE", strNOW, strFields, strValues)

            strSQL = String.Format("INSERT INTO {0} ({1}) VALUES({2})", EntityName, strFields, strValues)

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".InsertDo_HDR")
        End Try
    End Function
    '*****************************************************************
    '�Է� : strSQL = SQL ��
    '��ȯ : ó���Ǽ�
    '��� : �ش� Entity�� Update ó��
    '���� : Key ���ǰ� Value Field����������(OPTIONAL_STR/OPTIONAL_NUM/OPTIONAL_DTM)
    '*****************************************************************
    'strPREESTNO,SUSURATE,SUSUAMT,COMMITION,NONCOMMITION,SUMAMT
    Public Function UpdateDo_CALCUL(Optional ByVal strPREESTNO As String = OPTIONAL_STR, _
            Optional ByVal strAGREEYEARMON As String = OPTIONAL_STR, _
            Optional ByVal dblSUSURATE As Double = OPTIONAL_NUM, _
            Optional ByVal dblSUSUAMT As Double = OPTIONAL_NUM, _
            Optional ByVal dblCOMMITION As Double = OPTIONAL_NUM, _
            Optional ByVal dblNONCOMMITION As Double = OPTIONAL_NUM, _
            Optional ByVal dblSUMAMT As Double = OPTIONAL_NUM, _
            Optional ByVal strMEMO As String = OPTIONAL_STR, _
            Optional ByVal strPREESTNAME As String = OPTIONAL_STR, _
            Optional ByVal dblAMT As Double = OPTIONAL_NUM) As Integer

        Dim strSQL As String
        Dim strNOW As String '����Ʈ���� ó���� ������ �޾� �ؽ�Ʈ�� ó�� �Ѵ�.. 
        strNOW = Now

        Try
            strSQL = String.Format("UPDATE {0} SET {1} WHERE {2}", EntityName, _
                     BuildFields(",", _
                        GetFieldNameValue("CONFIRMFLAG", strAGREEYEARMON), _
                        GetFieldNameValue("SUSURATE", dblSUSURATE), _
                        GetFieldNameValue("SUSUAMT", dblSUSUAMT), _
                        GetFieldNameValue("COMMITION", dblCOMMITION), _
                        GetFieldNameValue("NONCOMMITION", dblNONCOMMITION), _
                        GetFieldNameValue("SUMAMT", dblSUMAMT), _
                        GetFieldNameValue("MEMO", strMEMO), _
                        GetFieldNameValue("AMT", dblAMT), _
                        GetFieldNameValue("PREESTNAME", strPREESTNAME), _
                        GetFieldNameValue("UUSER", mobjSCGLConfig.WRKUSR), _
                        GetFieldNameValue("UDATE", strNOW)), _
                     BuildFields("AND", _
                        GetFieldNameValue("PREESTNO", strPREESTNO)))

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".UpdateDo_CALCUL")
        End Try
    End Function

    '*****************************************************************
    '�Է� : strSQL = SQL ��
    '��ȯ : ó���Ǽ�
    '��� : �ش� Entity�� Delete ó��
    '���� : Key ������ ��������(OPTIONAL_STR/OPTIONAL_NUM/OPTIONAL_DTM)
    '*****************************************************************
    Public Function DeleteDo(ByVal strCODE As String) As Integer
        Dim strSQL As String

        Try
            strSQL = "DELETE FROM PD_PREEST_DTL WHERE PREESTNO = '" & strCODE & "'; DELETE FROM PD_PREEST_HDR WHERE PREESTNO = '" & strCODE & "'"

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
    '��� : strSQL �ش繮�� �״�� ó��
    '*****************************************************************
    Public Function InsertRtn_COPY(ByVal strTEMPSQL As String) As Integer
        Dim strSQL As String
        Try
            strSQL = strTEMPSQL
            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".InsertRtn_COPY")
        End Try
    End Function
    '*****************************************************************
    '�Է� : strSQL = SQL ��
    '��ȯ : ó���Ǽ�
    '��� : ���۹�ȣ �� ENDFLAG �� ���� �Ѵ� �Ƿ� >> ����
    '*****************************************************************
    Public Function UpdateENDFLAG(ByVal strJOBNO As String) As Integer
        Dim strSQL As String
        Try
            strSQL = "UPDATE PD_JOBNO SET ENDFLAG = 'PF02' WHERE JOBNO = '" & strJOBNO & "' "
            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".UpdateENDFLAG")
        End Try
    End Function
    '*****************************************************************
    '�Է� : strSQL = SQL ��
    '��ȯ : ó���Ǽ�
    '��� : strSQL �ش繮�� �״�� ó�� û������ ����� ���ֺ� ���� ��� ������ ������Ʈ
    '*****************************************************************
    Public Function NewUpdate(ByVal strTEMPSQL As String) As Integer
        Dim strSQL As String
        Try
            strSQL = strTEMPSQL
            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".NewUpdate")
        End Try
    End Function
    '*****************************************************************
    '�Է� : strSQL = SQL ��
    '��ȯ : ó���Ǽ�
    '��� : strSQL �ش繮�� �״�� ó�� û������ ����� ���ֺ� ���� ��� ������ ������Ʈ
    '*****************************************************************
    Public Function NewDelete(ByVal strTEMPSQL As String) As Integer
        Dim strSQL As String
        Try
            strSQL = strTEMPSQL
            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".NewDelete")
        End Try
    End Function
    '*****************************************************************
    '�Է� : strSQL = SQL ��
    '��ȯ : ó���Ǽ�
    '��� : strSQL �ش繮�� �״�� ó�� TEMP���� �����Ѵ�.
    '*****************************************************************
    Public Function TempDeleteDo(ByVal strUSER As String)
        Dim strSQL As String
        Try
            'strSQL = "DELETE FROM PD_SUBITEM_INPUT; DELETE FROM PD_PRODUCTIONCOMMI_INPUT; DELETE FROM PD_PRODUCTIONCOMMI_HDRINPUT"
            strSQL = "DELETE FROM PD_SUBITEM_INPUT where cuser = '" & strUSER & "' "
            strSQL = strSQL & " DELETE FROM PD_PRODUCTIONCOMMI_INPUT where cuser = '" & strUSER & "' "
            strSQL = strSQL & " DELETE FROM PD_PRODUCTIONCOMMI_HDRINPUT where cuser = '" & strUSER & "'"

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".TempDeleteDo")
        End Try
    End Function

    '*****************************************************************
    '�Է� : strSQL = SQL ��
    '��ȯ : ó���Ǽ�
    '��� : strSQL �ش繮�� �״�� ó�� û������ ����� ���ֺ� ���� ��� ������ ������Ʈ
    '*****************************************************************
    Public Function SubItemDeleteDo(ByVal strTEMPSQL As String) As Integer
        Dim strSQL As String
        Try
            strSQL = strTEMPSQL
            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".TempDeleteDo2")
        End Try
    End Function

    '*****************************************************************
    '�Է� : strSQL = SQL ��
    '��ȯ : ó���Ǽ�
    '��� : �ش� Entity�� Update ó��
    '���� : Key ���ǰ� Value Field����������(OPTIONAL_STR/OPTIONAL_NUM/OPTIONAL_DTM)
    '*****************************************************************
    Public Function UpdateDo_BASIC(Optional ByVal strPREESTNO As String = OPTIONAL_STR, _
                                    Optional ByVal strAGREEYEARMON As String = OPTIONAL_STR, _
                                    Optional ByVal dblSUSURATE As Double = OPTIONAL_NUM, _
                                    Optional ByVal dblSUSUAMT As Double = OPTIONAL_NUM, _
                                    Optional ByVal strMEMO As String = OPTIONAL_STR, _
                                    Optional ByVal strPREESTNAME As String = OPTIONAL_STR) As Integer

        Dim strSQL As String
        Dim strNOW As String '����Ʈ���� ó���� ������ �޾� �ؽ�Ʈ�� ó�� �Ѵ�.. 
        strNOW = Now

        Try
            strSQL = String.Format("UPDATE {0} SET {1} WHERE {2}", EntityName, _
                     BuildFields(",", _
                        GetFieldNameValue("CONFIRMFLAG", strAGREEYEARMON), _
                        GetFieldNameValue("SUSURATE", dblSUSURATE), _
                        GetFieldNameValue("SUSUAMT", dblSUSUAMT), _
                        GetFieldNameValue("MEMO", strMEMO), _
                        GetFieldNameValue("PREESTNAME", strPREESTNAME), _
                        GetFieldNameValue("UUSER", mobjSCGLConfig.WRKUSR), _
                        GetFieldNameValue("UDATE", strNOW)), _
                     BuildFields("AND", _
                        GetFieldNameValue("PREESTNO", strPREESTNO)))

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".UpdateDo_CALCUL")
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
        MyBase.EntityName = "PD_PREEST_HDR"     'Entity Name ����
    End Sub

    '���� ����� Base Class���� �����Ǿ� ����
#End Region
#End Region
End Class

