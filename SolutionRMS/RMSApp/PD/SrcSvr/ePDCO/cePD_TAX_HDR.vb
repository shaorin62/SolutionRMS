'****************************************************************************************
'Generated By: MakeSFAR V.2.0.0 - ��ƼƼ Ŭ���� ����Ŀ - ��ȭ S&C
'�ý��۱��� : �ַ�Ǹ�/�ý��۸�/Server Entity Class
'����  ȯ�� : GAC(Global Assembly Cache)
'���α׷��� : cePD_TAX_HDR.vb ( PD_TAX_HDR Entity ó�� Class)
'��      �� : PD_TAX_HDR Entity�� ����Insert/Update/Delete/Select�� ó��
'             - �θ�ƼƼ ��ü�� SCGLUtil.ceEntity�� ���
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2008-03-07 ���� 12:37:08 By Making Entity Bean
'****************************************************************************************
Imports SCGLUtil.cbSCGLUtil     '���� ��ƿ��Ƽ ��ü
Imports SCGLUtil.cbSCGLErr      '���� ����ó�� ��ü
Imports SCGLEntity              '��ƼƼ ��ü�� �θ� ��ü

Public Class cePD_TAX_HDR
    Inherits ceEntity

#Region "GROUP BLOCk : ���� �Ǵ� ��ⷹ���� ����/��� ����"
    Private Const CLASS_NAME = "cePD_TAX_HDR"    '�ڽ��� Ŭ������
#End Region

#Region "GROUP BLOCk : �ܺο� ���� Method"
#Region "SQL Insert/Update/Delete/Select"
    '*****************************************************************
    '�Է� : strSQL = SQL ��
    '��ȯ : ó���Ǽ�
    '��� : �ش� Entity�� Insert ó��
    '*****************************************************************
    Public Function InsertDo(ByVal strTAXYEARMON As String, _
            ByVal dblTAXNO As Double, _
            ByVal strCLIENTCODE As String, _
            ByVal strTIMCODE As String, _
            Optional ByVal strSUBSEQ As String = OPTIONAL_STR, _
            Optional ByVal dblAMT As Double = OPTIONAL_NUM, _
            Optional ByVal dblVAT As Double = OPTIONAL_NUM, _
            Optional ByVal dblSUMAMT As Double = OPTIONAL_NUM, _
            Optional ByVal strDEMANDDAY As String = OPTIONAL_STR, _
            Optional ByVal strPRINTDAY As String = OPTIONAL_STR, _
            Optional ByVal strVOCHNO As String = OPTIONAL_STR, _
            Optional ByVal strCLIENTBUSINO As String = OPTIONAL_STR, _
            Optional ByVal strREALBUSINO As String = OPTIONAL_STR, _
            Optional ByVal strSUMM As String = OPTIONAL_STR, _
            Optional ByVal strVATFLAG As String = OPTIONAL_STR, _
            Optional ByVal strTAXFLAG As String = OPTIONAL_STR, _
            Optional ByVal strTAXCODE As String = OPTIONAL_STR, _
            Optional ByVal strPAYCODE As String = OPTIONAL_STR, _
            Optional ByVal strCLIENTNAME As String = OPTIONAL_STR, _
            Optional ByVal strCLIENTOWNER As String = OPTIONAL_STR, _
            Optional ByVal strCLIENTADDR1 As String = OPTIONAL_STR, _
            Optional ByVal strCLIENTADDR2 As String = OPTIONAL_STR, _
            Optional ByVal strJOBGUBN As String = OPTIONAL_STR, _
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
            BuildNameValues(",", "TAXYEARMON", strTAXYEARMON, strFields, strValues)
            BuildNameValues(",", "TAXNO", dblTAXNO, strFields, strValues)
            BuildNameValues(",", "CLIENTCODE", strCLIENTCODE, strFields, strValues)
            BuildNameValues(",", "TIMCODE", strTIMCODE, strFields, strValues)
            BuildNameValues(",", "SUBSEQ", strSUBSEQ, strFields, strValues)
            BuildNameValues(",", "AMT", dblAMT, strFields, strValues)
            BuildNameValues(",", "VAT", dblVAT, strFields, strValues)
            BuildNameValues(",", "SUMAMT", dblSUMAMT, strFields, strValues)
            BuildNameValues(",", "DEMANDDAY", strDEMANDDAY, strFields, strValues)
            BuildNameValues(",", "PRINTDAY", strPRINTDAY, strFields, strValues)
            BuildNameValues(",", "VOCHNO", strVOCHNO, strFields, strValues)
            BuildNameValues(",", "CLIENTBUSINO", strCLIENTBUSINO, strFields, strValues)
            BuildNameValues(",", "REALBUSINO", strREALBUSINO, strFields, strValues)
            BuildNameValues(",", "SUMM", strSUMM, strFields, strValues)
            BuildNameValues(",", "VATFLAG", strVATFLAG, strFields, strValues)
            BuildNameValues(",", "TAXFLAG", strTAXFLAG, strFields, strValues)
            BuildNameValues(",", "TAXCODE", strTAXCODE, strFields, strValues)
            BuildNameValues(",", "PAYCODE", strPAYCODE, strFields, strValues)
            BuildNameValues(",", "CLIENTNAME", strCLIENTNAME, strFields, strValues)
            BuildNameValues(",", "CLIENTOWNER", strCLIENTOWNER, strFields, strValues)
            BuildNameValues(",", "CLIENTADDR1", strCLIENTADDR1, strFields, strValues)
            BuildNameValues(",", "CLIENTADDR2", strCLIENTADDR2, strFields, strValues)
            BuildNameValues(",", "JOBGUBN", strJOBGUBN, strFields, strValues)
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
    Public Function UpdateDo(Optional ByVal strTAXYEARMON As String = OPTIONAL_STR, _
            Optional ByVal dblTAXNO As Double = OPTIONAL_NUM, _
            Optional ByVal strCLIENTCODE As String = OPTIONAL_STR, _
            Optional ByVal strCLIENTSUBCODE As String = OPTIONAL_STR, _
            Optional ByVal strDEMANDDAY As String = OPTIONAL_STR, _
            Optional ByVal strPRINTDAY As String = OPTIONAL_STR, _
            Optional ByVal strVATFLAG As String = OPTIONAL_STR, _
            Optional ByVal dblAMT As Double = OPTIONAL_NUM, _
            Optional ByVal dblVAT As Double = OPTIONAL_NUM, _
            Optional ByVal dblSUMAMT As Double = OPTIONAL_NUM, _
            Optional ByVal strTAXFLAG As String = OPTIONAL_STR, _
            Optional ByVal strTAXCODE As String = OPTIONAL_STR, _
            Optional ByVal strSUMM As String = OPTIONAL_STR, _
            Optional ByVal strVOCHNO As String = OPTIONAL_STR, _
            Optional ByVal strPAYCODE As String = OPTIONAL_STR, _
            Optional ByVal strBUSINO As String = OPTIONAL_STR, _
            Optional ByVal strSUBSEQ As String = OPTIONAL_STR, _
            Optional ByVal strJOBGUBN As String = OPTIONAL_STR, _
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
                        GetFieldNameValue("CLIENTCODE", strCLIENTCODE), _
                        GetFieldNameValue("CLIENTSUBCODE", strCLIENTSUBCODE), _
                        GetFieldNameValue("DEMANDDAY", strDEMANDDAY), _
                        GetFieldNameValue("PRINTDAY", strPRINTDAY), _
                        GetFieldNameValue("VATFLAG", strVATFLAG), _
                        GetFieldNameValue("AMT", dblAMT), _
                        GetFieldNameValue("VAT", dblVAT), _
                        GetFieldNameValue("SUMAMT", dblSUMAMT), _
                        GetFieldNameValue("TAXFLAG", strTAXFLAG), _
                        GetFieldNameValue("TAXCODE", strTAXCODE), _
                        GetFieldNameValue("SUMM", strSUMM), _
                        GetFieldNameValue("VOCHNO", strVOCHNO), _
                        GetFieldNameValue("PAYCODE", strPAYCODE), _
                        GetFieldNameValue("BUSINO", strBUSINO), _
                        GetFieldNameValue("SUBSEQ", strSUBSEQ), _
                        GetFieldNameValue("JOBGUBN", strJOBGUBN), _
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
                        GetFieldNameValue("TAXYEARMON", strTAXYEARMON), GetFieldNameValue("TAXNO", dblTAXNO)))

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
    Public Function DeleteDo(Optional ByVal strTAXYEARMON As String = OPTIONAL_STR, Optional ByVal dblTAXNO As Double = OPTIONAL_NUM) As Integer
        Dim strSQL As String

        Try
            strSQL = String.Format("DELETE FROM {0} WHERE {1}", EntityName, _
                     BuildFields("AND", _
                                   GetFieldNameValue("TAXYEARMON", strTAXYEARMON), GetFieldNameValue("TAXNO", dblTAXNO)))

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
                            Optional ByVal strTAXYEARMON As String = OPTIONAL_STR, _
                            Optional ByVal dblTAXNO As Double = OPTIONAL_NUM, _
                            Optional ByVal strSelFields As String = "*", _
                            Optional ByVal intLimitRow As Integer = 0, _
                            Optional ByVal intSelMode As Integer = SELMODE.ARR, _
                            Optional ByVal blnBindingHeader As Boolean = False) As Object
        Dim strSQL As String
        Dim strKeyFields As String

        Try
            strKeyFields = BuildFields("AND", _
                                    GetFieldNameValue("TAXYEARMON", strTAXYEARMON), GetFieldNameValue("TAXNO", dblTAXNO))

            Return SelectDoExt(intRowCnt, intColCnt, strSelFields, strKeyFields, intLimitRow, intSelMode, blnBindingHeader)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".SeleteDo")
        End Try
    End Function
    '*****************************************************************
    '�Է� : strSQL = SQL ��
    '��ȯ : ó���Ǽ�
    '��� : ���� ������Ʈ
    '*************  ****************************************************
    Public Function UpdateRtn_SUMM(ByVal strTAXYEARMON As String, _
                                   ByVal strTAXNO As Double, _
                                   ByVal strSUMM As String, _
                                   ByVal lngVAT As Double) As Integer
        'strTAXYEARMON, strTAXNO, strSUMM
        Dim strSQL As String

        Try
            strSQL = "UPDATE PD_TAX_HDR SET SUMM = '" & strSUMM & "',VAT = " & lngVAT
            strSQL = strSQL & " WHERE TAXYEARMON = '" & strTAXYEARMON & "' AND TAXNO = " & strTAXNO
            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".UpdateRtn_SUMM")
        End Try
    End Function


    '*****************************************************************
    '�Է� : strSQL = SQL ��
    '��ȯ : ó���Ǽ�
    '��� : ���ݰ�꼭 �Ϸ�ǿ� ���Ͽ� ����/�ΰ��� ����
    '���� : Key ������ ��������(OPTIONAL_STR/OPTIONAL_NUM/OPTIONAL_DTM)
    '*****************************************************************
    Public Function Update_Vat(ByVal strTAXYEARMON As String, _
                              ByVal intTAXNO As Double, _
                              ByVal strSUMM As String, _
                              ByVal strVAT As Double) As Integer

        Dim strSQL As String 'strTAXYEARMON, intTAXNO, strTRUST_SEQ

        Try
            strSQL = " UPDATE PD_TAX_HDR "
            strSQL = strSQL & " SET SUMM = '" & strSUMM & "',"
            strSQL = strSQL & " VAT = '" & strVAT & "'"
            strSQL = strSQL & " WHERE TAXYEARMON = '" & strTAXYEARMON & "' "
            strSQL = strSQL & " AND TAXNO = " & intTAXNO

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".Update_Vat")
        End Try
    End Function


    '*****************************************************************
    '�Է� : strSQL = SQL ��
    '��ȯ : ó���Ǽ�
    '��� : ���ݰ�꼭 ��� ����
    '*****************************************************************
    Public Function DeleteDo_Tax(ByVal strTAXYEARMON As String, _
                                 ByVal strTAXNO As Double) As Integer
        Dim strSQL As String
        Try
            strSQL = "DELETE FROM  PD_TAX_HDR "
            strSQL = strSQL & " WHERE TAXYEARMON = '" & strTAXYEARMON & "' AND TAXNO = " & strTAXNO
            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".DeleteDo_Tax")
        End Try
    End Function


    Public Function DeleteMark(Optional ByVal strTAXYEARMON As String = OPTIONAL_STR, Optional ByVal dblTAXNO As Double = OPTIONAL_NUM) As Integer
        Dim strSQL As String

        Try
            strSQL = String.Format("UPDATE {0} SET ATTR10 = 999999 WHERE {1}", EntityName, _
                     BuildFields("AND", _
                                   GetFieldNameValue("TAXYEARMON", strTAXYEARMON), GetFieldNameValue("TAXNO", dblTAXNO)))

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
    Public Function UpdateRtn_TAX_HDR_MERGE(ByVal strTAXYEARMON, ByVal intTAXNO) As Integer
        Dim strSQL As String 'strTAXYEARMON, intTAXNO, strTRUST_SEQ

        Try
            strSQL = " UPDATE PD_TAX_HDR "
            strSQL = strSQL & "   SET MERGEFLAG = '1'"
            strSQL = strSQL & "   WHERE TAXYEARMON = '" & strTAXYEARMON & "' AND TAXNO = '" & intTAXNO & "'"
            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".UpdateRtn_COMMITAX_HDR_MERGE")
        End Try
    End Function


    Public Function UpdateRtn_TAX_HDR_MERGE_CANCEL(ByVal strTAXYEARMON, ByVal intTAXNO) As Integer
        Dim strSQL As String 'strTAXYEARMON, intTAXNO, strTRUST_SEQ

        Try
            strSQL = " UPDATE PD_TAX_HDR "
            strSQL = strSQL & "   SET MERGEFLAG = ''"
            strSQL = strSQL & "   WHERE TAXYEARMON = '" & strTAXYEARMON & "' AND TAXNO = '" & intTAXNO & "'"
            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".UpdateRtn_TAX_HDR_MERGE_CANCEL")
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
        MyBase.EntityName = "PD_TAX_HDR"     'Entity Name ����
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
'        intRtn = mobjcePD_TAX_HDR.InsertDo( _
'                                       GetElement(vntData,"TAXYEARMON", intColCnt, intRow), _
'                                       GetElement(vntData,"TAXNO", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"CLIENTCODE", intColCnt, intRow), _
'                                       GetElement(vntData,"REAL_MED_CODE", intColCnt, intRow), _
'                                       GetElement(vntData,"DEMANDDAY", intColCnt, intRow), _
'                                       GetElement(vntData,"PRINTDAY", intColCnt, intRow), _
'                                       GetElement(vntData,"VATFLAG", intColCnt, intRow), _
'                                       GetElement(vntData,"AMT", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"VAT", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"SUMAMT", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"TAXFLAG", intColCnt, intRow), _
'                                       GetElement(vntData,"TAXCODE", intColCnt, intRow), _
'                                       GetElement(vntData,"SUMM", intColCnt, intRow), _
'                                       GetElement(vntData,"VOCHNO", intColCnt, intRow), _
'                                       GetElement(vntData,"PAYCODE", intColCnt, intRow), _
'                                       GetElement(vntData,"BUSINO", intColCnt, intRow), _
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
'        intRtn = mobjcePD_TAX_HDR.UpdateDo( _
'                                       GetElement(vntData,"TAXYEARMON", intColCnt, intRow), _
'                                       GetElement(vntData,"TAXNO", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"CLIENTCODE", intColCnt, intRow), _
'                                       GetElement(vntData,"REAL_MED_CODE", intColCnt, intRow), _
'                                       GetElement(vntData,"DEMANDDAY", intColCnt, intRow), _
'                                       GetElement(vntData,"PRINTDAY", intColCnt, intRow), _
'                                       GetElement(vntData,"VATFLAG", intColCnt, intRow), _
'                                       GetElement(vntData,"AMT", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"VAT", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"SUMAMT", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"TAXFLAG", intColCnt, intRow), _
'                                       GetElement(vntData,"TAXCODE", intColCnt, intRow), _
'                                       GetElement(vntData,"SUMM", intColCnt, intRow), _
'                                       GetElement(vntData,"VOCHNO", intColCnt, intRow), _
'                                       GetElement(vntData,"PAYCODE", intColCnt, intRow), _
'                                       GetElement(vntData,"BUSINO", intColCnt, intRow), _
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
'        intRtn = mobjcePD_TAX_HDR.InsertDo( _
'                                       XMLGetElement(xmlRoot,"TAXYEARMON"), _
'                                       XMLGetElement(xmlRoot,"TAXNO", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"CLIENTCODE"), _
'                                       XMLGetElement(xmlRoot,"REAL_MED_CODE"), _
'                                       XMLGetElement(xmlRoot,"DEMANDDAY"), _
'                                       XMLGetElement(xmlRoot,"PRINTDAY"), _
'                                       XMLGetElement(xmlRoot,"VATFLAG"), _
'                                       XMLGetElement(xmlRoot,"AMT", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"VAT", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"SUMAMT", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"TAXFLAG"), _
'                                       XMLGetElement(xmlRoot,"TAXCODE"), _
'                                       XMLGetElement(xmlRoot,"SUMM"), _
'                                       XMLGetElement(xmlRoot,"VOCHNO"), _
'                                       XMLGetElement(xmlRoot,"PAYCODE"), _
'                                       XMLGetElement(xmlRoot,"BUSINO"), _
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
'        intRtn = mobjcePD_TAX_HDR.UpdateDo( _
'                                       XMLGetElement(xmlRoot,"TAXYEARMON"), _
'                                       XMLGetElement(xmlRoot,"TAXNO", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"CLIENTCODE"), _
'                                       XMLGetElement(xmlRoot,"REAL_MED_CODE"), _
'                                       XMLGetElement(xmlRoot,"DEMANDDAY"), _
'                                       XMLGetElement(xmlRoot,"PRINTDAY"), _
'                                       XMLGetElement(xmlRoot,"VATFLAG"), _
'                                       XMLGetElement(xmlRoot,"AMT", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"VAT", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"SUMAMT", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"TAXFLAG"), _
'                                       XMLGetElement(xmlRoot,"TAXCODE"), _
'                                       XMLGetElement(xmlRoot,"SUMM"), _
'                                       XMLGetElement(xmlRoot,"VOCHNO"), _
'                                       XMLGetElement(xmlRoot,"PAYCODE"), _
'                                       XMLGetElement(xmlRoot,"BUSINO"), _
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

