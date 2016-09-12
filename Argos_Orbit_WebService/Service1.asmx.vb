Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel
Imports System.Data
Imports System.Data.Odbc
Imports System.Data.Sql
Imports System.Security.Cryptography
Imports System.Text
Imports System.IO
' To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line.
' <System.Web.Script.Services.ScriptService()> _
<System.Web.Services.WebService(Namespace:="http://10.64.1.52/", Description:="Argos Web interface testing", Name:="InterfaceOrbitArgos")> _
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<ToolboxItem(False)> _
Public Class Service1
    Inherits System.Web.Services.WebService
    Dim connStr As String
    Dim conn As OdbcConnection
    Dim cmd As OdbcCommand
    Dim queryString As String
    Dim queryStr As String
    Dim reader As OdbcDataReader
    Public User As UserDetails
    Public OrbitUserID As String

    Public Structure customerDetails
        Public Account_number As String
        Public First_name As String
        Public Middle_name As String
        Public Last_Name As String
        Public Birthday As String
        Public street As String
        Public House_number As String
        Public Neighbourhood As String
        Public Community_Territory As String
        Public Free_address_field As String
        Public Home_phone_number As String
        Public Mobile_phone_number As String
        Public Savings_balance As Decimal
        Public Province As String
        Public SIC_Code As String
        Public Birth_Place As String
        Public Birth_Province As String
        Public Gender As String
        Public Business_Phone As String
        Public ID_Number As String
        Public Identification_Type As String
        Public Branch_Name As String
        Public TDPleage As Decimal  'added for the new argos
        Public previousLoanAmount As Decimal  'added for the new argos
    End Structure
    '
    '
    'Functions
    '
    '
    <WebMethod(Description:="Display Client informations"), SoapHeader("User")> _
    Public Function GetCustomerDetails(ByVal acct_no As String) As customerDetails

        Dim errorMsg As String = "No error"
        Dim resStr As String = "Sucess with account " & acct_no

        Dim cust As customerDetails = Nothing
        connStr = My.Settings.DBConnectionString
        '
        If CheckUserPwd(User) = False Then
            resStr = "Failure with user or pwd invalid"
            errorMsg = "No exeption"
            WriteToLog("GetCustomerDetails", resStr, errorMsg)
            Return Nothing
        End If
        Try
            queryString = "select 'Account number' = a.acct_no, 'First name' = b.first_name, 'Middle  name' = isnull(b.middle_initial,' ')," & _
            "'Last Name' = b.last_name, Birthday = isnull(b.birth_dt,' '), street = isnull(c.address_line_1,' ')," & _
            "'House number' = isnull(c.address_line_1,' '), Neighbourhood = isnull(c.address_line_2,' '), 'Community-Territory' = isnull(c.address_line_2,' ')," & _
            "'Free address field' = isnull(c.address_line_3,' '), 'Home_phone_number' = isnull(c.phone_1,' '), 'Mobile_phone_number' = isnull(c.phone_3,' '), 'Savings balance' = a.cur_bal ," & _
            "'Business_Phone' = isnull(c.phone_2,' '), 'Province' = isnull(c.province,' '), 'SIC_Code' = isnull(b.SIC_Code,0), 'Birth_Place' = isnull(b.city_of_birth,' ')," & _
            "'Birth_Province' = isnull(c.region, ' '), b.sex, 'ID_Number' = isnull(id_value,' ') , " & _
            "'Identification_Type' = isnull((select identification from ad_rm_ident where ident_id = b.ident_id),' ')," & _
            "'Branch_Name' = isnull((select name_1 from ad_gb_branch where branch_no = b.branch_no ),' ')," & _
            " isnull((select cur_bal from dp_display where acct_no = (select max(acct_no) from dp_display " & _
            " where acct_type = 'TD' and status in ('Active','RenewPending') and class_code = 806 and cur_bal > 0 " & _
            " and rim_no = a.rim_no)),0) as tdpledge, " & _
            " isnull((select amt from ln_display where acct_no = (select max(acct_no) from ln_display where acct_type = 'CL' and status not in ('incomplete','Active') " & _
            " and rim_no = a.rim_no)),0) as previousLoanAmount " & _
            " from dp_display a " & _
            " join rm_acct b on a.rim_no = b.rim_no " & _
            " join rm_address c on b.rim_no = c.rim_no " & _
            " where acct_no = '" & acct_no & "' " & _
            " and a.crncy_id = 11 and a.class_code = 100"
            conn = New OdbcConnection(connStr)
            conn.Open()
            cmd = New OdbcCommand(queryString, conn)
            reader = cmd.ExecuteReader
            reader.Read()
            cust.Account_number = reader("Account number").ToString
            cust.First_name = reader("First name").ToString
            cust.Middle_name = reader("Middle  name").ToString
            cust.Last_Name = reader("Last Name").ToString
            cust.Birthday = reader("Birthday").ToString
            '
            Dim sep() As String = {" ", ",", ";"}
            Dim tempStr() As String = reader("street").ToString.Split(sep, StringSplitOptions.None)
            If tempStr.Count >= 1 Then
                cust.House_number = tempStr(0)
            Else
                cust.House_number = reader("House number").ToString
            End If
            If tempStr.Count >= 2 Then
                cust.street = tempStr(1)
            Else
                cust.street = reader("street").ToString
            End If
            'cust.House_number = reader("House number").ToString
            cust.Neighbourhood = reader("Neighbourhood").ToString
            cust.Community_Territory = reader("Community-Territory").ToString
            cust.Free_address_field = reader("Free address field").ToString
            cust.Home_phone_number = reader("Home_phone_number").ToString
            cust.Mobile_phone_number = reader("Mobile_phone_number").ToString
            cust.Province = reader("Province").ToString()
            cust.SIC_Code = reader("SIC_Code").ToString()
            cust.Birth_Place = reader("Birth_Place").ToString()
            cust.Birth_Province = reader("Birth_Province").ToString()
            cust.Gender = reader("sex").ToString()
            cust.Business_Phone = reader("Business_Phone").ToString()
            cust.ID_Number = reader("ID_Number").ToString()
            cust.Identification_Type = reader("Identification_Type").ToString()
            cust.Branch_Name = reader("Branch_Name").ToString()
            cust.Savings_balance = reader("Savings balance").ToString
            cust.TDPleage = reader("tdpledge").ToString
            cust.previousLoanAmount = reader("previousLoanAmount").ToString
            WriteToLog("GetCustomerDetails", "Success with account " & acct_no, "No Error")
        Catch ex As Exception
            'do nothing
            resStr = "Failure with exception"
            errorMsg = ex.Message
            WriteToLog("GetCustomerDetails", resStr, errorMsg)
        End Try
        Return cust
    End Function

    <WebMethod(Description:="Attach a member to an existing group loan, it returns an ExitResponse structure, see the readme for details"), SoapHeader("User")> _
    Public Function InsertGroupMember(ByVal AccountNumber As String, ByVal MemberLoanAmount As Integer, _
                                      ByVal GroupLoanNo As String, ByVal SicCode As String) As Integer
        'Return value
        '0 : success
        '1 : connection error (unable to connect to the database or the database is not available)
        '2 : account number is not valid (does not exist or is not active)
        '3 : Member already exist
        '4 : Group loan no invalid
        '5 : member amount more than available loan amount
        'Attach a group member to the group loan 
        'Get more informations about each member
        'Additional field to get from ln_group_loan_members
        Dim CONTRACT_DT As String
        Dim LOAN_TRM As String
        Dim LOAN_PERIOD As String
        Dim MAT_DT As String
        Dim DISB_TYPE As String
        Dim DISB_ACCT_TYPE As String
        Dim DISB_ACCT_NO As String
        Dim PMT_ACCT_METHOD As String
        Dim PMT_ACCT_TYPE As String
        Dim PMT_ACCT_NO As String
        Dim PMT_PER_YEAR As String
        Dim FIRST_PMT_DT As String
        Dim PMT_EVERY_TRM As String
        Dim PMT_EVERY_PERIOD As String
        Dim LAST_PMT_DT As String
        Dim EFFECTIVE_DT As String
        Dim STATUS As String
        Dim STATUS_SORT As String
        Dim NO_PMTS As String
        Dim CHG_ALT_ACCT_TYPE As String
        Dim CHG_ALT_ACCT_NO As String
        Dim EMPL_ID As String
        Dim PTID As String
        Dim GroupLoanRim As String
        Dim undisbursed As Integer
        'additional field to get for the customer
        Dim RIM_NO As String
        Dim FullName As String
        '
        Dim res As Integer
        '
        Dim errorMsg As String = "No error"
        Dim resStr As String = "Sucess with account " & AccountNumber.Trim
        '
        connStr = My.Settings.DBConnectionString
        conn = New OdbcConnection(connStr)

        If CheckUserPwd(User) = False Then
            resStr = "Failure with user or pwd invalid"
            WriteToLog("InsertGroupMember", resStr, errorMsg)
            Return 1
        End If

        Try
            conn.Open()
            'check if account number valid
            queryStr = "select count(*) from dp_display where status = 'Active' and acct_no = '" & AccountNumber.Trim & "'"
            cmd = New OdbcCommand(queryStr, conn)
            res = CInt(cmd.ExecuteScalar())
            If res <> 1 Then
                'error, account number invalid
                resStr = "Failure with account number invalid state, only active account accepted, " & AccountNumber.Trim
                WriteToLog("InsertGroupMember", resStr, errorMsg)
                Return 2
            End If
            'check if account is already attached to the group loan
            queryStr = "select count(*) from LN_GRP_LOAN_MEMBERS where grp_loan_no = '" & GroupLoanNo & "' and  DISB_ACCT_NO = '" & AccountNumber.Trim & "'"
            cmd.CommandText = queryStr
            res = CInt(cmd.ExecuteScalar())
            If res > 0 Then
                'error, member already exist
                resStr = "Failure with member already exist"
                WriteToLog("InsertGroupMember", resStr, errorMsg)
                Return 3
            End If
        Catch ex As Exception
            errorMsg = ex.Message
            'connection error
            resStr = "Failure with exception"
            WriteToLog("InsertGroupMember", resStr, errorMsg)
            Return 1
        End Try
        'more info on the group loan
        queryStr = "select RIM_NO,CONTRACT_DT,LOAN_TRM,LOAN_PERIOD,MAT_DT,DISB_TYPE,DISB_ACCT_TYPE,DISB_ACCT_NO," & _
                        " PMT_ACCT_METHOD,PMT_ACCT_TYPE,PMT_ACCT_NO,PMT_PER_YEAR,FIRST_PMT_DT,PMT_EVERY_TRM,PMT_EVERY_PERIOD," & _
                        " LAST_PMT_DT,EFFECTIVE_DT,STATUS,STATUS_SORT,RIM_NO,GRP_LOAN_NO,NO_PMTS,CHG_ALT_ACCT_TYPE," & _
                        " CHG_ALT_ACCT_NO,CREATE_DT,EMPL_ID,ROW_VERSION,PTID,MEMO_TYPE,undisbursed " & _
                        " from LN_GRP_LOAN where GRP_LOAN_NO = '" & GroupLoanNo & "'"
        Try
            cmd.CommandText = queryStr
            reader = cmd.ExecuteReader()
            If reader.HasRows = False Then
                resStr = "Failure with invalid group loan no"
                WriteToLog("InsertGroupMember", resStr, errorMsg)
                'Goup loan no invalid
                Return 4
            End If
            reader.Read()
            CONTRACT_DT = reader("CONTRACT_DT").ToString
            LOAN_TRM = reader("LOAN_TRM").ToString
            LOAN_PERIOD = reader("LOAN_PERIOD").ToString
            MAT_DT = reader("MAT_DT").ToString
            DISB_TYPE = reader("DISB_TYPE").ToString
            DISB_ACCT_TYPE = reader("DISB_ACCT_TYPE").ToString
            DISB_ACCT_NO = reader("DISB_ACCT_NO").ToString
            PMT_ACCT_METHOD = reader("PMT_ACCT_METHOD").ToString
            PMT_ACCT_TYPE = reader("PMT_ACCT_TYPE").ToString
            PMT_ACCT_NO = reader("PMT_ACCT_NO").ToString
            PMT_PER_YEAR = reader("PMT_PER_YEAR").ToString
            FIRST_PMT_DT = reader("FIRST_PMT_DT").ToString
            PMT_EVERY_TRM = reader("PMT_EVERY_TRM").ToString
            PMT_EVERY_PERIOD = reader("PMT_EVERY_PERIOD").ToString
            LAST_PMT_DT = reader("LAST_PMT_DT").ToString
            EFFECTIVE_DT = reader("EFFECTIVE_DT").ToString
            STATUS = reader("STATUS").ToString
            STATUS_SORT = reader("STATUS_SORT").ToString
            NO_PMTS = reader("NO_PMTS").ToString
            CHG_ALT_ACCT_TYPE = reader("CHG_ALT_ACCT_TYPE").ToString
            CHG_ALT_ACCT_NO = reader("CHG_ALT_ACCT_NO").ToString
            EMPL_ID = reader("EMPL_ID").ToString
            GroupLoanRim = reader("RIM_NO").ToString
            undisbursed = reader("undisbursed").ToString
            If MemberLoanAmount > undisbursed Then
                resStr = "Failure with member limit > undisbursed"
                WriteToLog("InsertGroupMember", resStr, errorMsg)
                'member amount more than available loan amount
                Return 5
            End If
        Catch ex As Exception
            resStr = "Failure with execption"
            errorMsg = ex.Message
            WriteToLog("InsertGroupMember", resStr, errorMsg)
            'Error, connexion error
            Return 1
        End Try
        queryStr = "select fullname = last_name + ' ' + first_name, rim_no from rm_acct where rim_no in " & _
            " (select rim_no from dp_display where acct_no = '" & AccountNumber.Trim & "')"
        Try
            reader.Close()
            cmd.CommandText = queryStr
            reader = cmd.ExecuteReader()
            reader.Read()
            FullName = reader("fullname").ToString
            RIM_NO = reader("RIM_NO").ToString
            'Get the member Id (will be used also as ptid)
            reader.Close()
            cmd.CommandType = CommandType.StoredProcedure
            cmd.CommandText = "{call psp_class_get_ptid(?,?)}"
            cmd.Parameters.Add("@psTable", OdbcType.VarChar).Value = "LN_GRP_LOAN_MEMBERS"
            cmd.Parameters.Add("@pnCount", OdbcType.Int).Value = 1
            PTID = cmd.ExecuteScalar()
            'insert into the table
            queryStr = " Insert into  LN_GRP_LOAN_MEMBERS (MEMBER_ID,  MEMBER_NAME,  LOAN_AMT,  CONTRACT_DT, " & _
                       " LOAN_TRM,  LOAN_PERIOD,  MAT_DT, SIC_CODE, MULTI_DISB, DISB_TYPE, DISB_ACCT_TYPE, " & _
                       " DISB_ACCT_NO, DISB_ACCT_METHOD, CHECK_NO, CHECK_PRINTED_BY, " & _
                       " PMT_ACCT_METHOD, PMT_ACCT_TYPE, PMT_ACCT_NO, PMT_PER_YEAR, FIRST_PMT_DT, PMT_EVERY_TRM," & _
                       " PMT_EVERY_PERIOD, LAST_PMT_DT, EFFECTIVE_DT, STATUS,  STATUS_SORT,  RIM_NO,  GRP_LOAN_NO, " & _
                       " LOAN_STATUS, NO_PMTS, CHG_ALT_ACCT_TYPE, CHG_ALT_ACCT_NO,CREATE_DT,  EMPL_ID, ROW_VERSION,PTID,MEMO_TYPE) " & _
                       " Values ( '" & PTID & "','" & FullName.Replace("'", "''") & "'," & MemberLoanAmount.ToString & ",'" & CONTRACT_DT & "'," & _
                       LOAN_TRM & ",'" & LOAN_PERIOD & "','" & MAT_DT & "'," & SicCode & ",'N', 'Separate Account'," & _
                       "'SV','" & AccountNumber.Trim & "','Transfer', NULL,  NULL,  'Separate Account', 'SV','" & AccountNumber & "'," & _
                       "12,'" & FIRST_PMT_DT & "'," & PMT_EVERY_TRM & ",'" & PMT_EVERY_PERIOD & "','" & LAST_PMT_DT & "'," & _
                       "'" & EFFECTIVE_DT & " ','Active', 10," & RIM_NO & ",'" & GroupLoanNo & "','UnProcessed'," & NO_PMTS & "," & _
                       "'" & CHG_ALT_ACCT_TYPE & "','" & CHG_ALT_ACCT_NO & "',getdate(),0,1," & PTID & ",' ')"
            reader.Close()
            cmd.CommandType = CommandType.Text
            cmd.CommandText = queryStr
            cmd.ExecuteNonQuery()
            'update the undisbursed value in the table ln_grp_loan
            undisbursed -= MemberLoanAmount
            queryStr = "update ln_grp_loan set undisbursed = " & undisbursed.ToString & _
                        " where rim_no = " & GroupLoanRim & " and grp_loan_no = '" & GroupLoanNo & "'"
            cmd.CommandType = CommandType.Text
            cmd.CommandText = queryStr
            cmd.ExecuteNonQuery()
            resStr = "Sucess with account " & AccountNumber.Trim
            WriteToLog("InsertGroupMember", resStr, errorMsg)
            Return 0
        Catch ex As Exception
            'error
            resStr = "Failure with exception"
            errorMsg = ex.Message
            WriteToLog("InsertGroupMember", resStr, errorMsg)
            Return 1
        End Try
    End Function

    <WebMethod(Description:="Create a group loan in Orbit"), SoapHeader("User")> _
    Public Function CreateGroupLoan(ByVal group_RIM As String, ByVal Credit_Officer_Code As String, ByVal class_code As String, _
                                    ByVal description As String, ByVal surpvisor_code As String, _
                                    ByVal branch_no As String, ByVal loan_amount As String, ByVal contract_date As String, _
                                    ByVal loan_trm As String, ByVal location As String, ByVal meet_every_trm As String, _
                                    ByVal next_meeting_dt As String, ByVal groupe_account As String, ByVal loan_amt_limit As String, _
                                    ByVal first_payment_date As String, ByVal pmt_every_trm As String) As String
        'return the group loan no of the newly created 
        'error code
        'connection error = 1
        'invalid rim = 2
        'invalid class code = 3
        'last payment exceeds the maturity date = 4
        Const userCode As Integer = 0
        Dim res As Integer
        Dim NextGrpLoanNo As String
        Dim NewNextGrpLoanNo As String
        Dim effectiveGroupLoanNo As String
        Dim rate_type As String
        Dim mat_index_id As String
        Dim mat_margin As String
        Dim calc_bal_type As String
        Dim accounting_method As String
        Dim index_id As String
        Dim rate As String
        Dim advance As String
        Dim crncy_id As String
        Dim accr_basis As String
        Dim post_option As String

        Dim errorMsg As String = "No error"
        Dim resStr As String = "Sucess with group rim " & group_RIM

        connStr = My.Settings.DBConnectionString
        conn = New OdbcConnection(connStr)

        OrbitUserID = My.Settings.OrbitUserID

        If CheckUserPwd(User) = False Then
            resStr = "Failure with user or pwd invalid"
            errorMsg = "No exeption"
            WriteToLog("CreateGroupLoan", resStr, errorMsg)
            Return 1
        End If

        'check if the rim is valid
        queryStr = "select count(*) from rm_acct where rim_no = " & group_RIM
        Try
            conn.Open()
            cmd = New OdbcCommand(queryStr, conn)
            res = CInt(cmd.ExecuteScalar())
            If res <> 1 Then
                'error, invalid rim
                resStr = "Failure with invalid rim"
                errorMsg = "No exception"
                WriteToLog("CreateGroupLoan", resStr, errorMsg)
                Return 2
            End If
        Catch ex As Exception
            'connection error
            resStr = "Failure with exception"
            errorMsg = String.Format("{0}  {1}", ex.Message, ex.StackTrace)
            WriteToLog("CreateGroupLoan", resStr, errorMsg)
            Return 1
        End Try
        Try
            'get additional infomation : rate, mat_dt, final pmt , etc
            queryStr = "Select a.rate_type,  a.int_type, a.mat_index_id, a.margin,   a.mat_margin,  a.calc_bal_type," & _
                       " a.accounting_method, b.index_id,  b.rate, c.advance, c.crncy_id, a.accr_basis, a.post_option " & _
                       " From ad_ln_cls_int_opt a, ad_gb_rate_index b,ad_ln_cls c,ad_gb_crncy d " & _
                       " Where  a.acct_type =  'CL' " & _
                       " And  a.class_code = " & class_code & _
                       " And  a.index_id = b.index_id " & _
                       " And  c.crncy_id = d.crncy_id " & _
                       " and a.class_code = c.class_code " & _
                       " And a.class_code in (9,11)"
            'WriteToLog("CreateGroupLoan", "query1", queryStr)
            cmd.CommandText = queryStr
            reader = cmd.ExecuteReader()
            If reader.HasRows = False Then
                'invalid class code 
                resStr = "Failure with invalid class code"
                errorMsg = "No exception"
                WriteToLog("CreateGroupLoan", resStr, errorMsg)
                Return 3
            End If
            reader.Read()
            rate_type = reader("rate_type").ToString
            mat_index_id = reader("mat_index_id").ToString
            mat_margin = reader("mat_margin").ToString
            calc_bal_type = reader("calc_bal_type").ToString
            accounting_method = reader("accounting_method").ToString
            index_id = reader("index_id").ToString
            rate = reader("rate").ToString
            advance = reader("advance").ToString
            crncy_id = reader("crncy_id").ToString
            accr_basis = reader("accr_basis").ToString
            post_option = reader("post_option").ToString
            'get the ptid by calling the procedure psp_class_get_ptid
            reader.Close()
            cmd.CommandType = CommandType.StoredProcedure
            cmd.CommandText = "{call psp_class_get_ptid(?,?)}"
            cmd.Parameters.Add("@psTable", OdbcType.VarChar).Value = "LN_GRP_LOAN"
            cmd.Parameters.Add("@pnCount", OdbcType.Int).Value = 1
            Dim PTID As Integer = cmd.ExecuteScalar()
            'get the next group loan no
            queryStr = "select  grp_loan_next_no, IsNull( grp_loan_next_no, '0') from  ad_ln_control"
            cmd.CommandType = CommandType.Text
            cmd.CommandText = queryStr
            reader = cmd.ExecuteReader()
            reader.Read()
            NextGrpLoanNo = reader("grp_loan_next_no").ToString
            reader.Close()
            'compute a the value of group loan no
            NewNextGrpLoanNo = GetNewNextGrpLoanNo(NextGrpLoanNo)
            effectiveGroupLoanNo = GetEffectiveGroupLoanNo(NextGrpLoanNo)
            'insert into the table
            Dim matDt As DateTime = ComputeMatDt(CType(contract_date, DateTime), loan_trm)
            Dim noPmts As Integer = loan_trm / pmt_every_trm
            Dim lastPmtDt As DateTime = CType(first_payment_date, DateTime).AddDays((pmt_every_trm * (noPmts - 1)) * 7)
            If lastPmtDt > matDt Then
                'last payment exceeds the maturity date
                resStr = "Failure with last payment exceeds the maturity date"
                errorMsg = "No exception"
                WriteToLog("CreateGroupLoan", resStr, errorMsg)
                Return 4
            End If
            queryStr = "Insert into LN_GRP_LOAN  (BRANCH_NO,  CHG_ALT_ACCT_TYPE,  CHG_ALT_ACCT_NO,  GRP_LOAN_NO,  ACCT_TYPE,  CLASS_CODE,  " & _
            " CRNCY_ID,  DESCRIPTION,  RSM_ID,  PURPOSE_ID,  CUR_BAL, LOAN_AMT,  ADVANCE_TYPE,  CONTRACT_DT,  LOAN_TRM,  LOAN_PERIOD, " & _
            " MAT_DT, RATE_TYPE,  INDEX_ID, CURRENT_RATE,  MARGIN,  INT_TYPE,  CALC_BAL_TYPE,  POST_OPTION,  ACCR_BASIS, MAT_MARGIN,  " & _
            " MAT_INDEX_ID, ACCOUNTING_METHOD, DISB_TYPE,  DISB_ACCT_TYPE,  DISB_ACCT_NO,  LOAN_AMT_LIMIT, LOAN_AMT_LIMIT_TYPE, PMT_ACCT_METHOD,  " & _
            " PMT_ACCT_TYPE, PMT_ACCT_NO, PMT_TYPE, PMT_PER_YEAR, NO_PMTS, FIRST_PMT_DT, PMT_EVERY_TRM, PMT_EVERY_PERIOD, LAST_PMT_DT, EFFECTIVE_DT,  " & _
            " STATUS,   STATUS_SORT,  GROUP_ACCT_TYPE,  UNDISBURSED, BASE_RATE, RIM_NO, SUPERVISOR_ID, LOCATION, GROUP_ACCT_NO, MEET_EVERY_TRM,  " & _
            " MEET_EVERY_PERIOD, NEXT_MEETING_DT,  CREATE_DT,  EMPL_ID,  ROW_VERSION,   PTID,  MEMO_TYPE ) " & _
            " Values ( " & branch_no & ",'SV','" & groupe_account & "','" & effectiveGroupLoanNo & "','CL'," & class_code & _
            "," & crncy_id & ",'" & description.Replace("'", "''") & "'," & Credit_Officer_Code & ", NULL, NULL," & loan_amount & ",'Single','" & contract_date & "', " & _
            loan_trm & ", 'Week(s)','" & matDt.ToShortDateString & "','" & rate_type & "'," & index_id & "," & rate & ",NULL,'Simple','" & calc_bal_type & "','" & post_option & "','" & accr_basis & "', NULL, NULL," & _
            "'" & accounting_method & "','Separate Account',NULL,NULL," & loan_amt_limit & ",'Amount','Separate Account',NULL,NULL," & _
            " 'Level',12, " & noPmts.ToString & ",'" & first_payment_date & "'," & pmt_every_trm & ",'Week(s)','" & lastPmtDt.ToShortDateString & "',getdate(),'Active',10,'SV'," & _
            loan_amount.ToString & "," & rate & "," & group_RIM & "," & surpvisor_code & ",'" & location.Replace("'", "''") & "','" & groupe_account & "'," & meet_every_trm & _
            ",'Week(s)','" & next_meeting_dt & "',getdate()," & userCode & ",  1, " & PTID & ",' ' )"
            cmd.CommandType = CommandType.Text
            cmd.CommandText = queryStr
            cmd.ExecuteNonQuery()
            'WriteToLog("CreateGroupLoan", "query2", queryStr)
            'update the value of next group loan no
            queryStr = String.Format("update  ad_ln_control set  grp_loan_next_no = '{0}' where  grp_loan_next_no   = '{1}'", NewNextGrpLoanNo, NextGrpLoanNo)
            'WriteToLog("CreateGroupLoan", "query3", queryStr)
            cmd.CommandType = CommandType.Text
            cmd.CommandText = queryStr
            cmd.ExecuteNonQuery()
            'WriteToLog("CreateGroupLoan", " after query3", queryStr)
            'audit
            queryStr = " Insert Into GB_AUDIT (SCREEN_ID,SCREEN_PTID,TABLE_NAME,COLUMN_NAME,TABLE_PTID,PREV_VALUE," & _
                       " ACTUAL_PREV_VALUE,	NEW_VALUE,ACTUAL_NEW_VALUE,	EMPL_ID,CREATE_DT,EFFECTIVE_DT, " & _
                       " BRANCH_NO, DESCRIPTION) Values	(7122," & PTID.ToString & ",'LN_GRP_LOAN','STATUS'," & _
                       PTID.ToString & ",""'<New>'"", ""'<New>'"", ""'Active'"", ""'Active'"", " & OrbitUserID & "," & _
                       " GetDate( ),GetDate( )," & branch_no & ",'" & description & "')"
            'WriteToLog("CreateGroupLoan", "query 4", queryStr)
            cmd.CommandType = CommandType.Text
            cmd.CommandText = queryStr
            cmd.ExecuteNonQuery()
            resStr = "Sucess with group loan no " & effectiveGroupLoanNo
            errorMsg = "No exception"
            WriteToLog("CreateGroupLoan", resStr, errorMsg)
        Catch ex As Exception
            'error
            resStr = "Failure with exception"
            errorMsg = ex.Message
            WriteToLog("CreateGroupLoan", resStr, errorMsg)
            Return 1
        End Try
        Return effectiveGroupLoanNo
    End Function

    '<WebMethod(Description:="Update customer info in Orbit with info from argos")> _
    Private Function UpdateCustomerDetails(ByVal Account_number As String, ByVal First_name As String, ByVal Middle_name As String, _
    ByVal Last_Name As String, ByVal Birthday As String, ByVal street As String, ByVal House_number As String, _
    ByVal Neighbourhood As String, ByVal Community_Territory As String, _
    ByVal Home_phone_number As String, ByVal Mobile_phone_number As String, _
    ByVal Province As String, ByVal SIC_Code As String, ByVal Birth_Place As String, ByVal Birth_Province As String, _
    ByVal Gender As String, ByVal Business_Phone As String, ByVal ID_Number As String, ByVal Identification_Type As String, _
    ByVal Branch_Name As String) As Integer
        '
        Dim Rim As String
        Dim branch_no As String
        Dim Description As String

        If CheckUserPwd(User) = False Then
            Return 1
        End If

        'Get previous values
        Dim custD As customerDetails = GetCustomerDetails(Account_number)
        If custD.First_name = "" Then
            'error
            Return 1
        End If
        Try
            'Get the rim of the customer
            queryString = String.Format("select rim_no from dp_acct where acct_no = '{0}'", Account_number)
            conn = New OdbcConnection(connStr)
            conn.Open()
            cmd = New OdbcCommand(queryString, conn)
            reader = cmd.ExecuteReader
            reader.Read()
            Rim = reader("rim_no").ToString
            'Get the customer branch_no
            reader.Close()
            queryString = String.Format("select branch_no from ad_gb_branch where short_name = '{0}'", custD.Branch_Name)
            cmd.CommandText = queryString
            reader = cmd.ExecuteReader
            reader.Read()
            branch_no = reader("branch_no").ToString
            reader.Close()
            'other info
            OrbitUserID = My.Settings.OrbitUserID
            Description = String.Format("Customer:  {0} - {1} {2},{3}", Rim, custD.Last_Name, custD.Middle_name, custD.First_name)

            'update one field at a time and audit 
            'update fisrt_name
            If First_name <> "" AndAlso First_name <> custD.First_name Then
                'update first_name
                queryString = String.Format("update rm_acct set first_name = '{0}' where rim_no = {1}", First_name, Rim)
                cmd.CommandText = queryString
                cmd.ExecuteNonQuery()
                'audit
                queryStr = " Insert Into GB_AUDIT (SCREEN_ID,SCREEN_PTID,TABLE_NAME,COLUMN_NAME,TABLE_PTID,PREV_VALUE," & _
                           " ACTUAL_PREV_VALUE,	NEW_VALUE,ACTUAL_NEW_VALUE,	EMPL_ID,CREATE_DT,EFFECTIVE_DT, " & _
                           " BRANCH_NO, DESCRIPTION) Values	(327," & Rim & ",'RM_ACCT','FIRST_NAME'," & _
                           Rim & ",""'<" & custD.First_name & ">'"", ""'<" & custD.First_name & ">'"", ""'" & _
                           First_name & "'"", ""'" & First_name & "'"", " & OrbitUserID & "," & _
                           " GetDate( ),GetDate( )," & branch_no & ",'" & Description & "')"
                cmd.CommandType = CommandType.Text
                cmd.CommandText = queryStr
                cmd.ExecuteNonQuery()
            End If

            'update last_name
            If Last_Name <> "" AndAlso Last_Name <> custD.Last_Name Then
                queryString = String.Format("update rm_acct set last_name = '{0}' where rim_no = {1}", First_name, Rim)
                cmd.CommandText = queryString
                cmd.ExecuteNonQuery()
                'audit
                queryStr = " Insert Into GB_AUDIT (SCREEN_ID,SCREEN_PTID,TABLE_NAME,COLUMN_NAME,TABLE_PTID,PREV_VALUE," & _
                           " ACTUAL_PREV_VALUE,	NEW_VALUE,ACTUAL_NEW_VALUE,	EMPL_ID,CREATE_DT,EFFECTIVE_DT, " & _
                           " BRANCH_NO, DESCRIPTION) Values	(327," & Rim & ",'RM_ACCT','LAST_NAME'," & _
                           Rim & ",""'<" & custD.Last_Name & ">'"", ""'<" & custD.Last_Name & ">'"", ""'" & _
                           Last_Name & "'"", ""'" & Last_Name & "'"", " & OrbitUserID & "," & _
                           " GetDate( ),GetDate( )," & branch_no & ",'" & Description & "')"
                cmd.CommandType = CommandType.Text
                cmd.CommandText = queryStr
                cmd.ExecuteNonQuery()
            End If

            'update middle_name
            If Middle_name <> "" AndAlso Middle_name <> custD.Middle_name Then
                queryString = String.Format("update rm_acct set middle_initial = '{0}' where rim_no = {1}", Middle_name, Rim)
                cmd.CommandText = queryString
                cmd.ExecuteNonQuery()
                'audit
                queryStr = " Insert Into GB_AUDIT (SCREEN_ID,SCREEN_PTID,TABLE_NAME,COLUMN_NAME,TABLE_PTID,PREV_VALUE," & _
                           " ACTUAL_PREV_VALUE,	NEW_VALUE,ACTUAL_NEW_VALUE,	EMPL_ID,CREATE_DT,EFFECTIVE_DT, " & _
                           " BRANCH_NO, DESCRIPTION) Values	(327," & Rim & ",'RM_ACCT','middle_initial'," & _
                           Rim & ",""'<" & custD.Middle_name & ">'"", ""'<" & custD.Middle_name & ">'"", ""'" & _
                           Middle_name & "'"", ""'" & Middle_name & "'"", " & OrbitUserID & "," & _
                           " GetDate( ),GetDate( )," & branch_no & ",'" & Description & "')"
                cmd.CommandType = CommandType.Text
                cmd.CommandText = queryStr
                cmd.ExecuteNonQuery()
            End If

            'update Birthday
            If Birthday <> "" AndAlso Birthday <> custD.Birthday Then
                queryString = String.Format("update rm_acct set Birth_dt = '{0}' where rim_no = {1}", Birthday, Rim)
                cmd.CommandText = queryString
                cmd.ExecuteNonQuery()
                'audit
                queryStr = " Insert Into GB_AUDIT (SCREEN_ID,SCREEN_PTID,TABLE_NAME,COLUMN_NAME,TABLE_PTID,PREV_VALUE," & _
                           " ACTUAL_PREV_VALUE,	NEW_VALUE,ACTUAL_NEW_VALUE,	EMPL_ID,CREATE_DT,EFFECTIVE_DT, " & _
                           " BRANCH_NO, DESCRIPTION) Values	(327," & Rim & ",'RM_ACCT','BIRTH_DT'," & _
                           Rim & ",""'<" & custD.Birthday & ">'"", ""'<" & custD.Birthday & ">'"", ""'" & _
                           Birthday & "'"", ""'" & Birthday & "'"", " & OrbitUserID & "," & _
                           " GetDate( ),GetDate( )," & branch_no & ",'" & Description & "')"
                cmd.CommandType = CommandType.Text
                cmd.CommandText = queryStr
                cmd.ExecuteNonQuery()
            End If

            'update Business_Phone
            If Business_Phone <> "" AndAlso Business_Phone <> custD.Business_Phone Then
                queryString = String.Format("update rm_address set phone_2 = '{0}' where rim_no = {1}", Business_Phone, Rim)
                cmd.CommandText = queryString
                cmd.ExecuteNonQuery()
                'audit
                queryStr = " Insert Into GB_AUDIT (SCREEN_ID,SCREEN_PTID,TABLE_NAME,COLUMN_NAME,TABLE_PTID,PREV_VALUE," & _
                           " ACTUAL_PREV_VALUE,	NEW_VALUE,ACTUAL_NEW_VALUE,	EMPL_ID,CREATE_DT,EFFECTIVE_DT, " & _
                           " BRANCH_NO, DESCRIPTION) Values	(327," & Rim & ",'RM_ADDRESS','PHONE_2'," & _
                           Rim & ",""'<" & custD.Business_Phone & ">'"", ""'<" & custD.Business_Phone & ">'"", ""'" & _
                           Business_Phone & "'"", ""'" & Business_Phone & "'"", " & OrbitUserID & "," & _
                           " GetDate( ),GetDate( )," & branch_no & ",'" & Description & "')"
                cmd.CommandType = CommandType.Text
                cmd.CommandText = queryStr
                cmd.ExecuteNonQuery()
            End If



            'update Home_phone_number
            If Home_phone_number <> "" AndAlso Home_phone_number <> custD.Home_phone_number Then
                queryString = String.Format("update rm_address set phone_1 = '{0}' where rim_no = {1}", Home_phone_number, Rim)
                cmd.CommandText = queryString
                cmd.ExecuteNonQuery()
                'audit
                queryStr = " Insert Into GB_AUDIT (SCREEN_ID,SCREEN_PTID,TABLE_NAME,COLUMN_NAME,TABLE_PTID,PREV_VALUE," & _
                           " ACTUAL_PREV_VALUE,	NEW_VALUE,ACTUAL_NEW_VALUE,	EMPL_ID,CREATE_DT,EFFECTIVE_DT, " & _
                           " BRANCH_NO, DESCRIPTION) Values	(327," & Rim & ",'RM_ADDRESS','PHONE_1'," & _
                           Rim & ",""'<" & custD.Home_phone_number & ">'"", ""'<" & custD.Home_phone_number & ">'"", ""'" & _
                           Home_phone_number & "'"", ""'" & Home_phone_number & "'"", " & OrbitUserID & "," & _
                           " GetDate( ),GetDate( )," & branch_no & ",'" & Description & "')"
                cmd.CommandType = CommandType.Text
                cmd.CommandText = queryStr
                cmd.ExecuteNonQuery()
            End If

            'update Mobile_phone_number
            If Mobile_phone_number <> "" AndAlso Mobile_phone_number <> custD.Mobile_phone_number Then
                queryString = String.Format("update rm_address set phone_3 = '{0}' where rim_no = {1}", Mobile_phone_number, Rim)
                cmd.CommandText = queryString
                cmd.ExecuteNonQuery()
                'audit
                queryStr = " Insert Into GB_AUDIT (SCREEN_ID,SCREEN_PTID,TABLE_NAME,COLUMN_NAME,TABLE_PTID,PREV_VALUE," & _
                           " ACTUAL_PREV_VALUE,	NEW_VALUE,ACTUAL_NEW_VALUE,	EMPL_ID,CREATE_DT,EFFECTIVE_DT, " & _
                           " BRANCH_NO, DESCRIPTION) Values	(327," & Rim & ",'RM_ADDRESS','PHONE_3'," & _
                           Rim & ",""'<" & custD.Mobile_phone_number & ">'"", ""'<" & custD.Mobile_phone_number & ">'"", ""'" & _
                           Mobile_phone_number & "'"", ""'" & Mobile_phone_number & "'"", " & OrbitUserID & "," & _
                           " GetDate( ),GetDate( )," & branch_no & ",'" & Description & "')"
                cmd.CommandType = CommandType.Text
                cmd.CommandText = queryStr
                cmd.ExecuteNonQuery()
            End If

            'update Province
            If Province <> "" AndAlso Province <> custD.Province Then
                queryString = String.Format("update rm_address set Province = '{0}' where rim_no = {1}", Province, Rim)
                cmd.CommandText = queryString
                cmd.ExecuteNonQuery()
                'audit
                queryStr = " Insert Into GB_AUDIT (SCREEN_ID,SCREEN_PTID,TABLE_NAME,COLUMN_NAME,TABLE_PTID,PREV_VALUE," & _
                           " ACTUAL_PREV_VALUE,	NEW_VALUE,ACTUAL_NEW_VALUE,	EMPL_ID,CREATE_DT,EFFECTIVE_DT, " & _
                           " BRANCH_NO, DESCRIPTION) Values	(327," & Rim & ",'RM_ADDRESS','Province'," & _
                           Rim & ",""'<" & custD.Province & ">'"", ""'<" & custD.Province & ">'"", ""'" & _
                           Province & "'"", ""'" & Province & "'"", " & OrbitUserID & "," & _
                           " GetDate( ),GetDate( )," & branch_no & ",'" & Description & "')"
                cmd.CommandType = CommandType.Text
                cmd.CommandText = queryStr
                cmd.ExecuteNonQuery()
            End If

            'update Birth_Province
            If Birth_Province <> "" AndAlso Birth_Province <> custD.Birth_Province Then
                queryString = String.Format("update rm_address set region = '{0}' where rim_no = {1}", Birth_Province, Rim)
                cmd.CommandText = queryString
                cmd.ExecuteNonQuery()
                'audit
                queryStr = " Insert Into GB_AUDIT (SCREEN_ID,SCREEN_PTID,TABLE_NAME,COLUMN_NAME,TABLE_PTID,PREV_VALUE," & _
                           " ACTUAL_PREV_VALUE,	NEW_VALUE,ACTUAL_NEW_VALUE,	EMPL_ID,CREATE_DT,EFFECTIVE_DT, " & _
                           " BRANCH_NO, DESCRIPTION) Values	(327," & Rim & ",'RM_ADDRESS','region'," & _
                           Rim & ",""'<" & custD.Birth_Province & ">'"", ""'<" & custD.Birth_Province & ">'"", ""'" & _
                           Birth_Province & "'"", ""'" & Birth_Province & "'"", " & OrbitUserID & "," & _
                           " GetDate( ),GetDate( )," & branch_no & ",'" & Description & "')"
                cmd.CommandType = CommandType.Text
                cmd.CommandText = queryStr
                cmd.ExecuteNonQuery()
            End If

            'update SIC_Code
            If SIC_Code <> "" AndAlso SIC_Code <> custD.SIC_Code Then
                queryString = String.Format("update rm_acct set SIC_Code = '{0}' where rim_no = {1}", SIC_Code, Rim)
                cmd.CommandText = queryString
                cmd.ExecuteNonQuery()
                'audit
                queryStr = " Insert Into GB_AUDIT (SCREEN_ID,SCREEN_PTID,TABLE_NAME,COLUMN_NAME,TABLE_PTID,PREV_VALUE," & _
                           " ACTUAL_PREV_VALUE,	NEW_VALUE,ACTUAL_NEW_VALUE,	EMPL_ID,CREATE_DT,EFFECTIVE_DT, " & _
                           " BRANCH_NO, DESCRIPTION) Values	(327," & Rim & ",'RM_ACCT','SIC_Code'," & _
                           Rim & ",""'<" & custD.SIC_Code & ">'"", ""'<" & custD.SIC_Code & ">'"", ""'" & _
                           SIC_Code & "'"", ""'" & SIC_Code & "'"", " & OrbitUserID & "," & _
                           " GetDate( ),GetDate( )," & branch_no & ",'" & Description & "')"
                cmd.CommandType = CommandType.Text
                cmd.CommandText = queryStr
                cmd.ExecuteNonQuery()
            End If

            'update Birth_Place
            If Birth_Place <> "" AndAlso Birth_Place <> custD.Birth_Place Then
                queryString = String.Format("update rm_acct set city_of_birth = '{0}' where rim_no = {1}", Birth_Place, Rim)
                cmd.CommandText = queryString
                cmd.ExecuteNonQuery()
                'audit
                queryStr = " Insert Into GB_AUDIT (SCREEN_ID,SCREEN_PTID,TABLE_NAME,COLUMN_NAME,TABLE_PTID,PREV_VALUE," & _
                           " ACTUAL_PREV_VALUE,	NEW_VALUE,ACTUAL_NEW_VALUE,	EMPL_ID,CREATE_DT,EFFECTIVE_DT, " & _
                           " BRANCH_NO, DESCRIPTION) Values	(327," & Rim & ",'RM_ACCT','city_of_birth'," & _
                           Rim & ",""'<" & custD.Birth_Place & ">'"", ""'<" & custD.Birth_Place & ">'"", ""'" & _
                           Birth_Place & "'"", ""'" & Birth_Place & "'"", " & OrbitUserID & "," & _
                           " GetDate( ),GetDate( )," & branch_no & ",'" & Description & "')"
                cmd.CommandType = CommandType.Text
                cmd.CommandText = queryStr
                cmd.ExecuteNonQuery()
            End If

            'update Gender
            If Gender <> "" AndAlso Gender <> custD.Gender Then
                queryString = String.Format("update rm_acct set sex = '{0}' where rim_no = {1}", Gender, Rim)
                cmd.CommandText = queryString
                cmd.ExecuteNonQuery()
                'audit
                queryStr = " Insert Into GB_AUDIT (SCREEN_ID,SCREEN_PTID,TABLE_NAME,COLUMN_NAME,TABLE_PTID,PREV_VALUE," & _
                           " ACTUAL_PREV_VALUE,	NEW_VALUE,ACTUAL_NEW_VALUE,	EMPL_ID,CREATE_DT,EFFECTIVE_DT, " & _
                           " BRANCH_NO, DESCRIPTION) Values	(327," & Rim & ",'RM_ACCT','sex'," & _
                           Rim & ",""'<" & custD.Gender & ">'"", ""'<" & custD.Gender & ">'"", ""'" & _
                           Gender & "'"", ""'" & Gender & "'"", " & OrbitUserID & "," & _
                           " GetDate( ),GetDate( )," & branch_no & ",'" & Description & "')"
                cmd.CommandType = CommandType.Text
                cmd.CommandText = queryStr
                cmd.ExecuteNonQuery()
            End If

            'update ID_Number
            If ID_Number <> "" AndAlso ID_Number <> custD.ID_Number Then
                queryString = String.Format("update rm_acct set id_value = '{0}' where rim_no = {1}", ID_Number, Rim)
                cmd.CommandText = queryString
                cmd.ExecuteNonQuery()
                'audit
                queryStr = " Insert Into GB_AUDIT (SCREEN_ID,SCREEN_PTID,TABLE_NAME,COLUMN_NAME,TABLE_PTID,PREV_VALUE," & _
                           " ACTUAL_PREV_VALUE,	NEW_VALUE,ACTUAL_NEW_VALUE,	EMPL_ID,CREATE_DT,EFFECTIVE_DT, " & _
                           " BRANCH_NO, DESCRIPTION) Values	(327," & Rim & ",'RM_ACCT','id_value'," & _
                           Rim & ",""'<" & custD.ID_Number & ">'"", ""'<" & custD.ID_Number & ">'"", ""'" & _
                           ID_Number & "'"", ""'" & ID_Number & "'"", " & OrbitUserID & "," & _
                           " GetDate( ),GetDate( )," & branch_no & ",'" & Description & "')"
                cmd.CommandType = CommandType.Text
                cmd.CommandText = queryStr
                cmd.ExecuteNonQuery()
            End If

            'update Identification_Type
            If Identification_Type <> "" AndAlso Identification_Type <> custD.Identification_Type Then
                queryString = String.Format("update rm_acct set ident_id = (select ident_id from ad_rm_ident where identification = '{0}') where rim_no = {1}", Identification_Type, Rim)
                cmd.CommandText = queryString
                cmd.ExecuteNonQuery()
                'audit
                queryStr = " Insert Into GB_AUDIT (SCREEN_ID,SCREEN_PTID,TABLE_NAME,COLUMN_NAME,TABLE_PTID,PREV_VALUE," & _
                           " ACTUAL_PREV_VALUE,	NEW_VALUE,ACTUAL_NEW_VALUE,	EMPL_ID,CREATE_DT,EFFECTIVE_DT, " & _
                           " BRANCH_NO, DESCRIPTION) Values	(327," & Rim & ",'RM_ACCT','id_value'," & _
                           Rim & ",""'<" & custD.Identification_Type & ">'"", ""'<" & custD.Identification_Type & ">'"", ""'" & _
                           Identification_Type & "'"", ""'" & Identification_Type & "'"", " & OrbitUserID & "," & _
                           " GetDate( ),GetDate( )," & branch_no & ",'" & Description & "')"
                cmd.CommandType = CommandType.Text
                cmd.CommandText = queryStr
                cmd.ExecuteNonQuery()
            End If

            'update Branch_Name
            If Identification_Type <> "" AndAlso Branch_Name <> custD.Branch_Name Then
                queryString = String.Format("update rm_acct set branch_no = (select branch_no from ad_gb_branch where name_1 = '{0}') where rim_no = {1}", Branch_Name, Rim)
                cmd.CommandText = queryString
                cmd.ExecuteNonQuery()
                'audit
                queryStr = " Insert Into GB_AUDIT (SCREEN_ID,SCREEN_PTID,TABLE_NAME,COLUMN_NAME,TABLE_PTID,PREV_VALUE," & _
                           " ACTUAL_PREV_VALUE,	NEW_VALUE,ACTUAL_NEW_VALUE,	EMPL_ID,CREATE_DT,EFFECTIVE_DT, " & _
                           " BRANCH_NO, DESCRIPTION) Values	(327," & Rim & ",'RM_ACCT','id_value'," & _
                           Rim & ",""'<" & custD.Branch_Name & ">'"", ""'<" & custD.Branch_Name & ">'"", ""'" & _
                           Branch_Name & "'"", ""'" & Branch_Name & "'"", " & OrbitUserID & "," & _
                           " GetDate( ),GetDate( )," & branch_no & ",'" & Description & "')"
                cmd.CommandType = CommandType.Text
                cmd.CommandText = queryStr
                cmd.ExecuteNonQuery()
            End If

        Catch ex As Exception
            'do nothing
            Return 1
        End Try
    End Function

    <WebMethod(Description:="Show the log file content"), SoapHeader("User")> _
    Public Function ViewLog()
        Dim logPath As String = My.Settings.LogPath
        Dim fileReader As StreamReader = File.OpenText(logPath)
        Dim result As String = fileReader.ReadToEnd()
        fileReader.Close()
        Return result
    End Function


    Private Function GetNewNextGrpLoanNo(ByVal Cur As String) As String
        Dim NextGrpLoanNo As String = Cur
        Dim intNewNextGrpLoanNo As String = CStr(CInt(NextGrpLoanNo.TrimStart()) + 1) 'get the value without zero
        Dim nbZero As Integer = Cur.Length - intNewNextGrpLoanNo.Length
        Dim c As Char = "0"
        Dim tmpStr As String = New String(c, nbZero)
        Return tmpStr + intNewNextGrpLoanNo.ToString
    End Function

    Private Function GetEffectiveGroupLoanNo(ByVal cur As String) As String
        Dim NextGrpLoanNo As String = cur
        Dim intNewNextGrpLoanNo As String = CStr(CInt(NextGrpLoanNo.TrimStart()) + 1) 'get the value without zero
        Dim nbZero As Integer = cur.Length - intNewNextGrpLoanNo.Length
        Dim c As Char = "0"
        Dim tmpStr As String = New String(c, nbZero - 3)
        Return "000-" + tmpStr + intNewNextGrpLoanNo.ToString
    End Function

    Private Function ComputeMatDt(ByVal contractdt As DateTime, ByVal loan_term As Integer) As DateTime
        Dim mat_dt As DateTime
        mat_dt = contractdt.AddDays(loan_term * 7)
        Return mat_dt
    End Function

    Private Function CheckUserPwd(ByVal userd As UserDetails) As Boolean
        If (My.Settings.UserName = userd.UserName) And (My.Settings.Password = userd.password) Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Sub WriteToLog(ByVal caller As String, ByVal result As String, ByVal errorMessage As String)
        'Write to the log
        Dim logPath As String = My.Settings.LogPath
        Dim logFile As FileInfo = New FileInfo(logPath)
        Dim strData As String = String.Format("{0} : {1} , {2} , {3} {4}", My.Computer.Clock.LocalTime, caller, result, errorMessage, ControlChars.NewLine)
        Dim info As Byte() = New UTF8Encoding(True).GetBytes(strData)


        Dim fs As FileStream = File.Open(logPath, FileMode.Append, FileAccess.Write, FileShare.ReadWrite)
        fs.Write(info, 0, info.Length)
        fs.Flush()
        fs.Close()


        'Dim fileWritter As StreamWriter = logFile.AppendText()  'File.CreateText(logPath)
        'fileWritter.WriteLine(String.Format("{0} : {1} , {2} , {3}", My.Computer.Clock.LocalTime, caller, result, errorMessage))

        'fileWritter.Flush()
        'fileWritter.Close()
    End Sub
End Class