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

Public Class cePD_EXE_DTL
    Inherits ceEntity

#Region "GROUP BLOCk : ���� �Ǵ� ��ⷹ���� ����/��� ����"
    Private Const CLASS_NAME = "cePD_EXE_DTL"    '�ڽ��� Ŭ������
#End Region

#Region "GROUP BLOCk : �ܺο� ���� Method"
#Region "SQL Insert/Update/Delete/Select"
    '*****************************************************************
    '�Է� : strSQL = SQL ��
    '��ȯ : ó���Ǽ�
    '��� : �ش� Entity�� Insert ó��
    '*****************************************************************
    'JOBNO,SEQ,SORTSEQ,PREESTNO,ITEMCODESEQ,ITEMCODE,QTY,PRICE,AMT,OUTSCODE,STD,ADJAMT,VOCHNO,ADJDAY

    Public Function InsertDo(ByVal strJOBNO As String, _
            Optional ByVal dblSEQ As Double = OPTIONAL_NUM, _
            Optional ByVal strPREESTNO As String = OPTIONAL_STR, _
            Optional ByVal dblITEMCODESEQ As Double = OPTIONAL_NUM, _
            Optional ByVal strITEMCODE As String = OPTIONAL_STR, _
            Optional ByVal dblQTY As Double = OPTIONAL_NUM, _
            Optional ByVal dblPRICE As Double = OPTIONAL_NUM, _
            Optional ByVal dblAMT As Double = OPTIONAL_NUM, _
            Optional ByVal strOUTSCODE As String = OPTIONAL_STR, _
            Optional ByVal strSTD As String = OPTIONAL_STR, _
            Optional ByVal dblADJAMT As Double = OPTIONAL_NUM, _
            Optional ByVal strVOCHNO As String = OPTIONAL_STR, _
            Optional ByVal strADJDAY As String = OPTIONAL_STR, _
            Optional ByVal strREGDATE As String = OPTIONAL_STR, _
            Optional ByVal strAMTFLAG As String = OPTIONAL_STR, _
            Optional ByVal strADDFLAG As String = OPTIONAL_STR, _
            Optional ByVal strVATCODE As String = OPTIONAL_STR, _
            Optional ByVal strINCOMCODE As String = OPTIONAL_STR, _
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
            BuildNameValues(",", "JOBNO", strJOBNO, strFields, strValues)
            BuildNameValues(",", "SEQ", dblSEQ, strFields, strValues)
            BuildNameValues(",", "PREESTNO", strPREESTNO, strFields, strValues)
            BuildNameValues(",", "ITEMCODESEQ", dblITEMCODESEQ, strFields, strValues)
            BuildNameValues(",", "ITEMCODE", strITEMCODE, strFields, strValues)
            BuildNameValues(",", "QTY", dblQTY, strFields, strValues)
            BuildNameValues(",", "PRICE", dblPRICE, strFields, strValues)
            BuildNameValues(",", "AMT", dblAMT, strFields, strValues)
            BuildNameValues(",", "OUTSCODE", strOUTSCODE, strFields, strValues)
            BuildNameValues(",", "STD", strSTD, strFields, strValues)
            BuildNameValues(",", "ADJAMT", dblADJAMT, strFields, strValues)
            BuildNameValues(",", "VOCHNO", strVOCHNO, strFields, strValues)
            BuildNameValues(",", "ADJDAY", strADJDAY, strFields, strValues)
            BuildNameValues(",", "REGDATE", strREGDATE, strFields, strValues)
            BuildNameValues(",", "AMTFLAG", strAMTFLAG, strFields, strValues)
            BuildNameValues(",", "ADDFLAG", strADDFLAG, strFields, strValues)
            BuildNameValues(",", "VATCODE", strVATCODE, strFields, strValues)
            BuildNameValues(",", "INCOMCODE", strINCOMCODE, strFields, strValues)
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
    '
    Public Function UpdateDo(ByVal strJOBNO As String, _
            Optional ByVal dblSEQ As Double = OPTIONAL_NUM, _
            Optional ByVal strPREESTNO As String = OPTIONAL_STR, _
            Optional ByVal dblITEMCODESEQ As Double = OPTIONAL_NUM, _
            Optional ByVal strITEMCODE As String = OPTIONAL_STR, _
            Optional ByVal dblQTY As Double = OPTIONAL_NUM, _
            Optional ByVal dblPRICE As Double = OPTIONAL_NUM, _
            Optional ByVal dblAMT As Double = OPTIONAL_NUM, _
            Optional ByVal strOUTSCODE As String = OPTIONAL_STR, _
            Optional ByVal strSTD As String = OPTIONAL_STR, _
            Optional ByVal dblADJAMT As Double = OPTIONAL_NUM, _
            Optional ByVal strVOCHNO As String = OPTIONAL_STR, _
            Optional ByVal strADJDAY As String = OPTIONAL_STR, _
            Optional ByVal strREGDATE As String = OPTIONAL_STR, _
            Optional ByVal strAMTFLAG As String = OPTIONAL_STR, _
            Optional ByVal strADDFLAG As String = OPTIONAL_STR, _
            Optional ByVal strVATCODE As String = OPTIONAL_STR, _
            Optional ByVal strINCOMCODE As String = OPTIONAL_STR, _
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
        Dim strNOW As String '����Ʈ���� ó���� ������ �޾� �ؽ�Ʈ�� ó�� �Ѵ�.. 
        strNOW = Now
        Try
            'JOBNO,SEQ,SORTSEQ,PREESTNO,ITEMCODESEQ,ITEMCODE,QTY,PRICE,AMT,OUTSCODE,STD,ADJAMT,VOCHNO,ADJDAY
            strSQL = String.Format("UPDATE {0} SET {1} WHERE {2}", EntityName, _
                     BuildFields(",", _
                        GetFieldNameValue("PREESTNO", strPREESTNO), _
                        GetFieldNameValue("ITEMCODESEQ", dblITEMCODESEQ), _
                        GetFieldNameValue("ITEMCODE", strITEMCODE), _
                        GetFieldNameValue("QTY", dblQTY), _
                        GetFieldNameValue("PRICE", dblPRICE), _
                        GetFieldNameValue("AMT", dblAMT), _
                        GetFieldNameValue("OUTSCODE", strOUTSCODE), _
                        GetFieldNameValue("STD", strSTD), _
                        GetFieldNameValue("ADJAMT", dblADJAMT), _
                        GetFieldNameValue("VOCHNO", strVOCHNO), _
                        GetFieldNameValue("ADJDAY", strADJDAY), _
                        GetFieldNameValue("REGDATE", strREGDATE), _
                        GetFieldNameValue("AMTFLAG", strAMTFLAG), _
                        GetFieldNameValue("ADDFLAG", strADDFLAG), _
                        GetFieldNameValue("VATCODE", strVATCODE), _
                        GetFieldNameValue("INCOMCODE", strINCOMCODE), _
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
                        GetFieldNameValue("JOBNO", strJOBNO), GetFieldNameValue("SEQ", dblSEQ)))

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
    'strJOBNO, dblSEQ, dblSORTSEQ
    Public Function DeleteDo(ByVal strJOBNO As String, _
                             ByVal dblSEQ As Double) As Integer
        Dim strSQL As String

        Try
            strSQL = String.Format("DELETE FROM {0} WHERE {1}", EntityName, _
                     BuildFields("AND", _
                                   GetFieldNameValue("JOBNO", strJOBNO), GetFieldNameValue("SEQ", dblSEQ)))

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
    '��� : strSQL �ش繮�� �״�� ó�� (�������� Ȯ�� ó���Ѵ�)
    '*****************************************************************
    Public Function UpdateDo_Confirm(ByVal strTEMPSQL As String) As Integer
        Dim strSQL As String
        Try
            strSQL = strTEMPSQL
            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".UpdateDo_Confirm")
        End Try
    End Function
    '*****************************************************************
    '�Է� : strSQL = SQL ��
    '��ȯ : ó���Ǽ�
    '��� : strSQL �ش繮�� �״�� ó�� (�������� Ȯ����� ó���Ѵ�)
    '*****************************************************************
    Public Function UpdateDo_ConfirmCancel(ByVal strTEMPSQL As String) As Integer
        Dim strSQL As String
        Try
            strSQL = strTEMPSQL
            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".UpdateDo_ConfirmCancel")
        End Try
    End Function
    '*****************************************************************
    '�Է� : strSQL = SQL ��
    '��ȯ : ó���Ǽ�
    '��� : strSQL �ش繮�� �״�� ó�� ������� �����Ѵ�
    '*****************************************************************
    Public Function DeleteDo_ALL(ByVal strTEMPSQL As String) As Integer
        Dim strSQL As String
        Try
            strSQL = strTEMPSQL
            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".DeleteDo_ALL")
        End Try
    End Function
    '*****************************************************************
    '�Է� : strSQL = SQL ��
    '��ȯ : ó���Ǽ�
    '��� : strSQL �ش繮�� �״�� ó�� �������� Ȯ�� 
    '*****************************************************************
    Public Function UpdateRtn_PurchaseNo(ByVal strTEMPSQL As String) As Integer
        Dim strSQL As String
        Try
            strSQL = strTEMPSQL
            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".UpdateRtn_PurchaseNo")
        End Try
    End Function
    '*****************************************************************
    '�Է� : strSQL = SQL ��
    '��ȯ : ó���Ǽ�
    '��� : strSQL �ش繮�� �״�� ó�� ��Ʈ���� ����� �ϰ�������Ʈ 
    '*****************************************************************
    Public Function UpdateSort(ByVal strTEMPSQL As String) As Integer
        Dim strSQL As String
        Try
            strSQL = strTEMPSQL
            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".UpdateSort")
        End Try
    End Function
    'UpdateRtn_VochNo
    '*****************************************************************
    '�Է� : strSQL = SQL ��
    '��ȯ : ó���Ǽ�
    '��� : strSQL �ش繮�� �״�� ó�� ��Ʈ���� ����� �ϰ�������Ʈ 
    '*****************************************************************
    Public Function UpdateRtn_VochNo(ByVal strTEMPSQL As String) As Integer
        Dim strSQL As String
        Try
            strSQL = strTEMPSQL
            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".UpdateRtn_VochNo")
        End Try
    End Function
    'UpdateRtn_VochNo
    '*****************************************************************
    '�Է� : strSQL = SQL ��
    '��ȯ : ó���Ǽ�
    '��� : strSQL �ش繮�� �״�� ó�� ��Ʈ���� ����� �ϰ�������Ʈ 
    '*****************************************************************
    Public Function UpdateRtn_Confirm(ByVal strTEMPSQL As String) As Integer
        Dim strSQL As String
        Try
            strSQL = strTEMPSQL
            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".UpdateRtn_Confirm")
        End Try
    End Function


    '*****************************************************************
    '�Է� : strSQL = SQL ��
    '��ȯ : ó���Ǽ�
    '��� : ���ݰ�꼭 �������� ���� ���۰�����ȣ�� �÷��� ���� (û��PF03)
    '���� : Key ������ �������� �ƴ�
    '*****************************************************************
    'UpdateRtn_SUMM(strTAXYEARMON, strTAXNO, strSUMM)
    Public Function Update_EndFlag(ByVal strJOBNO As String) As Integer
        'strTAXYEARMON, intTAXNO
        Dim strSQL As String

        Try
            strSQL = "UPDATE PD_JOBNO SET ENDFLAG='PF02'"
            strSQL = strSQL & " WHERE 1=1"
            strSQL = strSQL & " AND (ENDFLAG = 'PF01' OR ENDFLAG = 'PF06')"
            strSQL = strSQL & " and JOBNO = '" & strJOBNO & "'"

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".Update_EndFlag")
        End Try
    End Function



    '*****************************************************************
    '�Է� : strSQL = SQL ��
    '��ȯ : ó���Ǽ�
    '��� : ���ݰ�꼭 �������� ���� ���۰�����ȣ�� �÷��� ���� (û��PF03)
    '���� : Key ������ �������� �ƴ�
    '*****************************************************************
    'UpdateRtn_SUMM(strTAXYEARMON, strTAXNO, strSUMM)
    Public Function Delete_EndFlag(ByVal strJOBNO As String) As Integer
        'strTAXYEARMON, intTAXNO
        Dim strSQL As String

        Try
            strSQL = "UPDATE PD_JOBNO SET ENDFLAG='PF01'"
            strSQL = strSQL & " WHERE JOBNO = '" & strJOBNO & "'"
            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".Delete_EndFlag")
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
        MyBase.EntityName = "PD_EXE_DTL"     'Entity Name ����
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


