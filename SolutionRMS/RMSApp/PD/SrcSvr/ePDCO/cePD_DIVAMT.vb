'Public Class cePD_DIVAMT

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

Public Class cePD_DIVAMT
    Inherits ceEntity

#Region "GROUP BLOCk : ���� �Ǵ� ��ⷹ���� ����/��� ����"
    Private Const CLASS_NAME = "cePD_DIVAMT"    '�ڽ��� Ŭ������
#End Region

#Region "GROUP BLOCk : �ܺο� ���� Method"
#Region "SQL Insert/Update/Delete/Select"
    '*****************************************************************
    '�Է� : strSQL = SQL ��
    '��ȯ : ó���Ǽ�
    '��� : �ش� Entity�� Insert ó��
    '*****************************************************************
    'YEARMON,JOBNO,SEQ,PREESTNO,JOBNAME,INCJOBNOSEQ,INCJOBNO,CLIENTCODE,TIMCODE,SUBSEQ,CREDAY,
    'DIVAMT,ADJAMT,CHARGE,ENDFLAG,
    'TRANSYEARMON,TRANSNO,TRANSNOSEQ,TAXYEARMON,TAXNO,TAXNOSEQ,TAXCODE,DEMANDFLAG,CONFIRMFLAG,CLOSINGMONTH,ENDMONTH,CONFIRMDEMEND,
    'CHECK_POINT,MEMO,USENO,MANAGER
    Public Function InsertDo(ByVal strYEARMON As String, _
            Optional ByVal strJOBNO As String = OPTIONAL_STR, _
            Optional ByVal dblSEQ As Double = OPTIONAL_NUM, _
            Optional ByVal strPREESTNO As String = OPTIONAL_STR, _
            Optional ByVal strJOBNAME As String = OPTIONAL_STR, _
            Optional ByVal dblINCJOBNOSEQ As Double = OPTIONAL_NUM, _
            Optional ByVal strINCJOBNO As String = OPTIONAL_STR, _
            Optional ByVal strCLIENTCODE As String = OPTIONAL_STR, _
            Optional ByVal strTIMCODE As String = OPTIONAL_STR, _
            Optional ByVal strSUBSEQ As String = OPTIONAL_STR, _
            Optional ByVal strCREDAY As String = OPTIONAL_STR, _
            Optional ByVal dblDIVAMT As Double = OPTIONAL_NUM, _
            Optional ByVal dblADJAMT As Double = OPTIONAL_NUM, _
            Optional ByVal dblCHARGE As Double = OPTIONAL_NUM, _
            Optional ByVal strENDFLAG As String = OPTIONAL_STR, _
            Optional ByVal strTRANSYEARMON As String = OPTIONAL_STR, _
            Optional ByVal dblTRANSNO As Double = OPTIONAL_NUM, _
            Optional ByVal dblTRANSNOSEQ As Double = OPTIONAL_NUM, _
            Optional ByVal strTAXYEARMON As String = OPTIONAL_STR, _
            Optional ByVal dblTAXNO As Double = OPTIONAL_NUM, _
            Optional ByVal dblTAXNOSEQ As Double = OPTIONAL_NUM, _
            Optional ByVal strTAXCODE As String = OPTIONAL_STR, _
            Optional ByVal strDEMANDFLAG As String = OPTIONAL_STR, _
            Optional ByVal strCONFIRMFLAG As String = OPTIONAL_STR, _
            Optional ByVal strCLOSINGMONTH As String = OPTIONAL_STR, _
            Optional ByVal strENDMONTH As String = OPTIONAL_STR, _
            Optional ByVal strCONFIRMDEMEND As String = OPTIONAL_STR, _
            Optional ByVal strCHECK_POINT As String = OPTIONAL_STR, _
            Optional ByVal strMEMO As String = OPTIONAL_STR, _
            Optional ByVal strUSENO As String = OPTIONAL_STR, _
            Optional ByVal strMANAGER As String = OPTIONAL_STR, _
            Optional ByVal strCHARGEHISTORY As Double = OPTIONAL_NUM, _
            Optional ByVal strDATAYEARMON As String = OPTIONAL_STR, _
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
            BuildNameValues(",", "YEARMON", strYEARMON, strFields, strValues)
            BuildNameValues(",", "JOBNO", strJOBNO, strFields, strValues)
            BuildNameValues(",", "SEQ", dblSEQ, strFields, strValues)
            BuildNameValues(",", "PREESTNO", strPREESTNO, strFields, strValues)
            BuildNameValues(",", "JOBNAME", strJOBNAME, strFields, strValues)
            BuildNameValues(",", "INCJOBNOSEQ", dblINCJOBNOSEQ, strFields, strValues)
            BuildNameValues(",", "INCJOBNO", strINCJOBNO, strFields, strValues)
            BuildNameValues(",", "CLIENTCODE", strCLIENTCODE, strFields, strValues)
            BuildNameValues(",", "TIMCODE", strTIMCODE, strFields, strValues)
            BuildNameValues(",", "SUBSEQ", strSUBSEQ, strFields, strValues)
            BuildNameValues(",", "CREDAY", strCREDAY, strFields, strValues)
            BuildNameValues(",", "DIVAMT", dblDIVAMT, strFields, strValues)
            BuildNameValues(",", "ADJAMT", dblADJAMT, strFields, strValues)
            BuildNameValues(",", "CHARGE", dblCHARGE, strFields, strValues)
            BuildNameValues(",", "ENDFLAG", strENDFLAG, strFields, strValues)
            BuildNameValues(",", "TRANSYEARMON", strTRANSYEARMON, strFields, strValues)
            BuildNameValues(",", "TRANSNO", dblTRANSNO, strFields, strValues)
            BuildNameValues(",", "TRANSNOSEQ", dblTRANSNOSEQ, strFields, strValues)
            BuildNameValues(",", "TAXYEARMON", strTAXYEARMON, strFields, strValues)
            BuildNameValues(",", "TAXNO", dblTAXNO, strFields, strValues)
            BuildNameValues(",", "TAXNOSEQ", dblTAXNOSEQ, strFields, strValues)
            BuildNameValues(",", "TAXCODE", strTAXCODE, strFields, strValues)
            BuildNameValues(",", "DEMANDFLAG", strDEMANDFLAG, strFields, strValues)
            BuildNameValues(",", "CONFIRMFLAG", strCONFIRMFLAG, strFields, strValues)
            BuildNameValues(",", "CLOSINGMONTH", strCLOSINGMONTH, strFields, strValues)
            BuildNameValues(",", "ENDMONTH", strENDMONTH, strFields, strValues)
            BuildNameValues(",", "CONFIRMDEMEND", strCONFIRMDEMEND, strFields, strValues)
            BuildNameValues(",", "CHECK_POINT", strCHECK_POINT, strFields, strValues)
            BuildNameValues(",", "MEMO", strMEMO, strFields, strValues)
            BuildNameValues(",", "USENO", strUSENO, strFields, strValues)
            BuildNameValues(",", "MANAGER", strMANAGER, strFields, strValues)
            BuildNameValues(",", "CHARGEHISTORY", strCHARGEHISTORY, strFields, strValues)
            BuildNameValues(",", "DATAYEARMON", strDATAYEARMON, strFields, strValues)
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
    'PREESTNO,SEQ,JOBNO,YEARMON,CREDAY,CLIENTCODE,CLIENTSUBCODE,DIVAMT,JOBNAME
    Public Function UpdateDo(ByVal strYEARMON As String, _
            Optional ByVal strJOBNO As String = OPTIONAL_STR, _
            Optional ByVal dblSEQ As Double = OPTIONAL_NUM, _
            Optional ByVal strPREESTNO As String = OPTIONAL_STR, _
            Optional ByVal strJOBNAME As String = OPTIONAL_STR, _
            Optional ByVal dblINCJOBNOSEQ As Double = OPTIONAL_NUM, _
            Optional ByVal strINCJOBNO As String = OPTIONAL_STR, _
            Optional ByVal strCLIENTCODE As String = OPTIONAL_STR, _
            Optional ByVal strTIMCODE As String = OPTIONAL_STR, _
            Optional ByVal strSUBSEQ As String = OPTIONAL_STR, _
            Optional ByVal strCREDAY As String = OPTIONAL_STR, _
            Optional ByVal dblDIVAMT As Double = OPTIONAL_NUM, _
            Optional ByVal dblADJAMT As Double = OPTIONAL_NUM, _
            Optional ByVal dblCHARGE As Double = OPTIONAL_NUM, _
            Optional ByVal strENDFLAG As String = OPTIONAL_STR, _
            Optional ByVal strTRANSYEARMON As String = OPTIONAL_STR, _
            Optional ByVal dblTRANSNO As Double = OPTIONAL_NUM, _
            Optional ByVal dblTRANSNOSEQ As Double = OPTIONAL_NUM, _
            Optional ByVal strTAXYEARMON As String = OPTIONAL_STR, _
            Optional ByVal dblTAXNO As Double = OPTIONAL_NUM, _
            Optional ByVal dblTAXNOSEQ As Double = OPTIONAL_NUM, _
            Optional ByVal strTAXCODE As String = OPTIONAL_STR, _
            Optional ByVal strDEMANDFLAG As String = OPTIONAL_STR, _
            Optional ByVal strCONFIRMFLAG As String = OPTIONAL_STR, _
            Optional ByVal strCLOSINGMONTH As String = OPTIONAL_STR, _
            Optional ByVal strENDMONTH As String = OPTIONAL_STR, _
            Optional ByVal strCONFIRMDEMEND As String = OPTIONAL_STR, _
            Optional ByVal strCHECK_POINT As String = OPTIONAL_STR, _
            Optional ByVal strMEMO As String = OPTIONAL_STR, _
            Optional ByVal strUSENO As String = OPTIONAL_STR, _
            Optional ByVal strMANAGER As String = OPTIONAL_STR, _
            Optional ByVal strCHARGEHISTORY As Double = OPTIONAL_NUM, _
            Optional ByVal strDATAYEARMON As String = OPTIONAL_STR, _
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
                        GetFieldNameValue("YEARMON", strYEARMON), _
                        GetFieldNameValue("JOBNO", strJOBNO), _
                        GetFieldNameValue("SEQ", dblSEQ), _
                        GetFieldNameValue("PREESTNO", strPREESTNO), _
                        GetFieldNameValue("JOBNAME", strJOBNAME), _
                        GetFieldNameValue("INCJOBNOSEQ", dblINCJOBNOSEQ), _
                        GetFieldNameValue("INCJOBNO", strINCJOBNO), _
                        GetFieldNameValue("CLIENTCODE", strCLIENTCODE), _
                        GetFieldNameValue("TIMCODE", strTIMCODE), _
                        GetFieldNameValue("SUBSEQ", strSUBSEQ), _
                        GetFieldNameValue("CREDAY", strCREDAY), _
                        GetFieldNameValue("DIVAMT", dblDIVAMT), _
                        GetFieldNameValue("ADJAMT", dblADJAMT), _
                        GetFieldNameValue("CHARGE", dblCHARGE), _
                        GetFieldNameValue("ENDFLAG", strENDFLAG), _
                        GetFieldNameValue("TRANSYEARMON", strTRANSYEARMON), _
                        GetFieldNameValue("TRANSNO", dblTRANSNO), _
                        GetFieldNameValue("TRANSNOSEQ", dblTRANSNOSEQ), _
                        GetFieldNameValue("TAXYEARMON", strTAXYEARMON), _
                        GetFieldNameValue("TAXNO", dblTAXNO), _
                        GetFieldNameValue("TAXNOSEQ", dblTAXNOSEQ), _
                        GetFieldNameValue("TAXCODE", strTAXCODE), _
                        GetFieldNameValue("DEMANDFLAG", strDEMANDFLAG), _
                        GetFieldNameValue("CONFIRMFLAG", strCONFIRMFLAG), _
                        GetFieldNameValue("CLOSINGMONTH", strCLOSINGMONTH), _
                        GetFieldNameValue("ENDMONTH", strENDMONTH), _
                        GetFieldNameValue("CONFIRMDEMEND", strCONFIRMDEMEND), _
                        GetFieldNameValue("CHECK_POINT", strCHECK_POINT), _
                        GetFieldNameValue("MEMO", strMEMO), _
                        GetFieldNameValue("USENO", strUSENO), _
                        GetFieldNameValue("MANAGER", strMANAGER), _
                        GetFieldNameValue("CHARGEHISTORY", strCHARGEHISTORY), _
                        GetFieldNameValue("DATAYEARMON", strDATAYEARMON), _
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
                        GetFieldNameValue("PREESTNO", strPREESTNO), GetFieldNameValue("SEQ", dblSEQ), GetFieldNameValue("JOBNO", strJOBNO)))

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
    Public Function DeleteDo(ByVal dblSEQ As Integer, _
                             ByVal strJOBNO As String, _
                             ByVal strPREESTNO As String) As Integer
        Dim strSQL As String

        Try
            strSQL = String.Format("DELETE FROM {0} WHERE {1}", EntityName, _
                     BuildFields("AND", _
                                   GetFieldNameValue("SEQ", dblSEQ), GetFieldNameValue("JOBNO", strJOBNO), GetFieldNameValue("PREESTNO", strPREESTNO)))

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
    '��� : �ش� Entity�� Select ó��
    '*****************************************************************
    Public Function DeleteUpdateDo(Optional ByVal strTRANSYEARMON As String = OPTIONAL_STR, _
                                   Optional ByVal strTRANSNO As String = OPTIONAL_STR) As Integer

        Dim strSQL As String

        Try

            strSQL = " UPDATE A SET  "
            strSQL = strSQL & " A.CONFIRMFLAG = '3',A.ENDFLAG = 'PF02', A.TRANSYEARMON = '', A.TRANSNO = NULL, A.TRANSNOSEQ = NULL"
            strSQL = strSQL & " FROM PD_DIVAMT A , PD_TRANS_DTL B "
            strSQL = strSQL & " WHERE B.TRANSYEARMON = A.TRANSYEARMON AND B.TRANSNO = A.TRANSNO "
            strSQL = strSQL & " AND B.TRANSYEARMON = '" & strTRANSYEARMON & "' AND  B.TRANSNO = '" & strTRANSNO & "'"

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".DeleteUpdateDo")
        End Try
    End Function


    '*****************************************************************
    '�Է� : strSQL = SQL ��
    '��ȯ : ó���Ǽ�
    '��� : �ش� Entity�� ����ݾ� ������Ʈ ó��
    '*****************************************************************
    Public Function ADJAMTUpdateDo(Optional ByVal strJOBNO As String = OPTIONAL_STR, _
                                   Optional ByVal dblSEQ As Double = OPTIONAL_NUM, _
                                   Optional ByVal dblADJAMT As Double = OPTIONAL_NUM, _
                                   Optional ByVal strTRANSYEARMON As String = OPTIONAL_STR, _
                                   Optional ByVal intTRANSNO As Double = OPTIONAL_NUM, _
                                   Optional ByVal intTRANSNOSEQ As Double = OPTIONAL_NUM) As Integer

        Dim strSQL As String

        Try
            strSQL = " UPDATE PD_DIVAMT SET CONFIRMFLAG = '4', ENDFLAG = 'PF03', "
            strSQL = strSQL & " TRANSYEARMON = '" & strTRANSYEARMON & "', "
            strSQL = strSQL & " TRANSNO = " & intTRANSNO & ","
            strSQL = strSQL & " TRANSNOSEQ = " & intTRANSNOSEQ & ""
            strSQL = strSQL & " WHERE SEQ = '" & dblSEQ & "' AND JOBNO = '" & strJOBNO & "'"
            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".ADJAMTUpdateDo")
        End Try
    End Function

    '*****************************************************************
    '�Է� : strSQL = SQL ��
    '��ȯ : ó���Ǽ�
    '��� :
    '*****************************************************************
    Public Function Update_CHECK_POINT(Optional ByVal strPREESTNO As String = OPTIONAL_STR, _
                                       Optional ByVal strSEQ As Integer = OPTIONAL_NUM, _
                                       Optional ByVal strCHECK_POINT As String = OPTIONAL_STR) As Integer
        Dim strSQL As String

        Try
            strSQL = String.Format(" UPDATE {0} SET {1} WHERE {2}", EntityName, GetFieldNameValue("CHECK_POINT", strCHECK_POINT), _
                     BuildFields("AND", _
                        GetFieldNameValue("PREESTNO", strPREESTNO), GetFieldNameValue("SEQ", strSEQ)))

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".UpdateRtn_ENDFLAG")
        End Try
    End Function
    '*****************************************************************
    '�Է� : strSQL = SQL �� By KTH
    '��ȯ : ó���Ǽ�
    '��� : ��Ź���̺��� ���ݰ�꼭 ��ȣ ������Ʈ
    '*****************************************************************

    Public Function Update_TruTax(ByVal strJOBNO As String, _
                                  ByVal intJOBNOSEQ As Integer, _
                                  ByVal strTAXYEARMON As String, _
                                  ByVal intTAXNO As Integer, _
                                  ByVal intTAXNOSEQ As Integer) As Integer
        'strTRANSNO, strTAXNO
        Dim strSQL As String
        Try
            strSQL = "UPDATE PD_DIVAMT SET TAXYEARMON = '" & strTAXYEARMON & "' , TAXNO = " & intTAXNO & ", TAXNOSEQ = " & intTAXNOSEQ & ""
            strSQL = strSQL & "  WHERE JOBNO = '" & strJOBNO & "' AND SEQ = " & intJOBNOSEQ & ""
            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".Update_TruTax")
        End Try
    End Function

    '*****************************************************************
    '�Է� : strSQL = SQL �� By KTH
    '��ȯ : ó���Ǽ�
    '��� : ��Ź���̺��� ���ݰ�꼭 ������ ��ȣ ������Ʈ
    '*****************************************************************
    Public Function TruTaxDeleteUpdateDo(ByVal strTAXYEARMON As String, ByVal strTAXNO As Integer) As Integer
        Dim strSQL As String
        Try
            strSQL = "UPDATE PD_DIVAMT SET TAXYEARMON = '', TAXNO = NULL, TAXNOSEQ = NULL "
            strSQL = strSQL & "  WHERE TAXYEARMON = '" & strTAXYEARMON & "' AND TAXNO = " & strTAXNO & ""
            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".TruTaxDeleteUpdateDo")
        End Try
    End Function

    '*****************************************************************
    '�Է� : strSQL = SQL �� By KTH
    '��ȯ : ó���Ǽ�
    '��� : EXEAMT , EXEFLAG ������Ʈ
    '*****************************************************************
    Public Function Update_EXEAMTCONFIRM(ByVal strJOBNO As String, ByVal dblSEQ As Double, ByVal IntEXEAMT As Integer) As Integer
        Dim strSQL As String
        Try
            strSQL = "UPDATE PD_DIVAMT SET  EXE_FLAG = '1', EXEAMT = " & IntEXEAMT & ""
            strSQL = strSQL & "  WHERE JOBNO = '" & strJOBNO & "' AND SEQ = " & dblSEQ & ""
            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".Update_EXEAMTCONFIRM")
        End Try
    End Function


    '*****************************************************************
    '�Է� : strSQL = SQL �� By KTH
    '��ȯ : ó���Ǽ�
    '��� : EXEAMT , EXEFLAG ������Ʈ
    '*****************************************************************
    Public Function Update_EXEAMTCONFIRMCANCEL(ByVal strJOBNO As String, ByVal dblSEQ As Double) As Integer
        Dim strSQL As String
        Try
            strSQL = "UPDATE PD_DIVAMT SET  EXE_FLAG = '0', EXEAMT = 0 "
            strSQL = strSQL & "  WHERE JOBNO = '" & strJOBNO & "' AND SEQ = " & dblSEQ & ""
            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".Update_EXEAMTCONFIRMCANCEL")
        End Try
    End Function

    '*****************************************************************
    '�Է� : strSQL = SQL ��
    '��ȯ : ó���Ǽ�
    '��� : strSQL �ش繮�� �״�� ó�� 
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
#End Region

#Region "��ü ����/����"
    '*****************************************************************
    '�Է� : strInfoXML = ����⺻������ ���� XML
    'objSCGLSql = DB ó�� ��ü �ν��Ͻ� ����    '��ȯ : ����
    '��� : DB ó���� ���� ����⺻���� ����
    '*****************************************************************
    Public Sub New(Optional ByVal objSCGLConfig As SCGLUtil.cbSCGLConfig = Nothing, Optional ByVal strInfoXML As String = "")
        MyBase.SetConfig(objSCGLConfig, strInfoXML)
        MyBase.EntityName = "PD_DIVAMT"     'Entity Name ����
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

