using System;
using System.IO;
using System.Data;
using System.Data.SqlClient;

using ImportUser.Properties;
using System.Diagnostics;


namespace ImportUser
{
    /// <summary>
    /// This class provides the core application capability and hence provides the 
    /// implementaiton for the BMO IAAP web application user import. This class
    /// is written to support the ImportUser command line utiltiy, and hence
    /// directly reads from and writes to the Console. 
    /// This class also relies on the settings file accompanying the ImportUser applicaiton.
    /// </summary>
    public class ImportUserImpl
    {
        public string UserFileType { get; set; }

        /// <summary>
        /// This is the only externally exposed method of the class and is meant to be 
        /// invoked by a clietn to request the services of this class. This method determines
        /// the file type of the user import file type provided, and then imports the users
        /// based on the file type, if the user doesn't already exist. Finally this method
        /// also adds the user to the appropriate approval role.
        /// </summary>
        /// <param name="args">command line arguments</param>
        /// <returns></returns>
        public bool Import( string[] args )
        {
            string sDebug = String.Empty;
            string fileTypeText = string.Empty;
            int existingUserCount = 0;
            int newUserCount = 0;
            bool bError = false;


            #region Get Filename and Type Then Read User data file
            sDebug = "Finding file";
            Console.WriteLine( "filename: {0}", args[0] );

            FileStream fs = File.OpenRead( args[0] );

            sDebug = "Reading file";
            StreamReader sr = new StreamReader( fs );
            string fileContent = sr.ReadToEnd();


            sDebug = "Determining File Type";
            UserImportFileType fileType = UserImportFileType.Undefined;
            Enum.TryParse<UserImportFileType>( args[1], out fileType );
            if ( fileType == UserImportFileType.Undefined )
            {
                throw new ArgumentException( "Incorrect Filetype argument" );
            }
            this.UserFileType = fileTypeText = fileType.ToString();
            

            sDebug = "Reading rows";
            string[] fileRows = fileContent.Split( new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries );
            Console.WriteLine( "Rows read: {0}", fileRows.Length );
            Console.WriteLine();
            Console.WriteLine();
            #endregion


            #region connect to database
            sDebug = "connect to database";
            SqlConnection conn = new SqlConnection( Settings.Default.dbConnectionString );
            conn.Open();
            #endregion


            // Loop through the rows - skipping the first row
            for ( int row = Settings.Default.csaRowToStart; row < fileRows.Length; row++ )
            {
                try
                {
                    string rowNum = " when reading row " + row;

                    // read columns
                    sDebug = "read columns" + rowNum;
                    string[] columns = fileRows[row].Split( new char[] { ',' } );
                    string userId;


                    // Create new user or reuse if user already exists
                    EvaluateUser( out sDebug, ref existingUserCount, ref newUserCount, conn, rowNum, columns, out userId );


                    // Import file based on file type
                    if ( fileType == UserImportFileType.CSA )
                    {
                        if ( !ImportCsaUsers( conn, row, rowNum, columns, userId, out sDebug ) )
                        {
                            break;
                        }
                    }
                    else if ( fileType == UserImportFileType.OpsRisk )
                    {
                        if ( !ImportOpsRiskUsers( conn, row, rowNum, columns, userId, out sDebug ) )
                        {
                            break;
                        }
                    }
                    else if ( fileType == UserImportFileType.Governance )
                    {
                        if ( !ImportGovernanceUsers( conn, row, rowNum, columns, userId, out sDebug ) )
                        {
                            break;
                        }
                    }
                }
                catch ( Exception innerE )
                {
                    string errorMsg = string.Format( "Exception occured when {0}{1}{2}", sDebug, "\r\n", innerE.Message );
                    Console.WriteLine( errorMsg );
                    if ( Settings.Default.stopAtError )
                    {
                        throw new ApplicationException(
                            "Application is currently configured to stop at errors.\r\n"
                            + errorMsg, innerE
                        );
                    }

                    bError = true;
                }
            }


            Console.WriteLine();
            Console.WriteLine( "New Users Imported: {0}", newUserCount );
            Console.WriteLine( "Existing Users Imported: {0}", existingUserCount );
            Console.WriteLine();

            return bError;
        }


        private void EvaluateUser( out string sDebug, ref int existingUserCount, ref int newUserCount, SqlConnection conn, string rowNum, string[] columns, out string userId )
        {
            #region check if user exists and get user Id
            bool isExistingUser = false;
            userId = String.Empty;

            sDebug = "checking if the user already exists " + rowNum;
            string userName = columns[Settings.Default.NameColumn].Trim().Replace( "'", "''" );
            string userEmail = columns[Settings.Default.csaEmailColumn].Trim().Replace( "'", "" );
            SqlCommand checkUserSql = conn.CreateCommand();

            checkUserSql.CommandText = "select id from AspNetUsers where Email = @userEmail";
            checkUserSql.Parameters.Add( "@userEmail", SqlDbType.VarChar ).Value = userEmail;


            sDebug += "\r\n" + checkUserSql.CommandText;
            userId = ( string )(checkUserSql.ExecuteScalar());
            checkUserSql.Dispose();
            if ( string.IsNullOrEmpty( userId ) )
            {
                userId = Guid.NewGuid().ToString();
                newUserCount++;
            }
            else
            {
                isExistingUser = true;
                existingUserCount++;
            }

            #endregion


            if ( !isExistingUser )
            {
                #region create user record
                sDebug = "creating user record query" + rowNum;

                SqlCommand userRecordSql = conn.CreateCommand();

                userRecordSql.CommandText = "insert into AspNetUsers " +
                    "( Id, Name, LastModifiedDate, EmailConfirmed, PhoneNumberConfirmed, " +
                    "TwoFactorEnabled, LockoutEnabled, AccessFailedCount, UserName, Email ) " +
                    "values ( @userId, @Name, @modDate, 0, 0, 0, 0, 0, @userName, @userEmail)";


                userRecordSql.Parameters.Add( "@userId", SqlDbType.VarChar, 128 ).Value = userId;
                userRecordSql.Parameters.Add( "@Name", SqlDbType.VarChar ).Value = userName.Replace( "'", "" );
                userRecordSql.Parameters.Add( "@modDate", SqlDbType.DateTime ).Value = DateTime.Now.ToString();
                userRecordSql.Parameters.Add( "@userName", SqlDbType.VarChar, 256 ).Value = userEmail;
                userRecordSql.Parameters.Add( "@userEmail", SqlDbType.VarChar, 256 ).Value = userEmail;


                sDebug = "executing create user query" + rowNum + "\r\n" + userRecordSql.CommandText;
                userRecordSql.ExecuteNonQuery();
                userRecordSql.Dispose();
                #endregion
            }
        }


        private bool ImportCsaUsers( 
            SqlConnection conn, 
            int row,
            string rowNum, 
            string[] columns, 
            string userId, 
            out string sDebug )
        {

            // check if this is the first invocation
            if ( row == Settings.Default.csaRowToStart )
            {

                // Check CsaAttributes
                sDebug = "Checking if CsaAttributes records already exist";
                if ( !ShouldWeContinueImport( conn, "CsaAttributes" ) )
                {
                    Console.WriteLine();
                    Console.WriteLine( "Terminating application as user requested to abort the operation" );
                    return false;
                }

            }


            // Write CsaAttributes
            sDebug = "creating CsaAttributes record for the user";
            SqlCommand csaAttributeSql = conn.CreateCommand();

            csaAttributeSql.CommandText = "insert into CsaAttributes(AppUserId, OperatingGroup_Id, " +
                "SubOperatingGroup_Id, LineOfBusiness_Id, Jurisdiction_Id, CorporateServiceArea_Id, DateCreated, " +
                "LastUpdated) Values(@userId, @opGroup, @subOpGroup, @lob, @jurisdiction, " +
                "@corpArea, @dateCreated, @dateUpdated)";

            csaAttributeSql.Parameters.Add( "@userId", SqlDbType.VarChar, 128 ).Value = userId;
            csaAttributeSql.Parameters.Add( "@opGroup", SqlDbType.Int ).Value =
                GetOpId( columns[Settings.Default.csaOpGroupColumn] );
            csaAttributeSql.Parameters.Add( "@subOpGroup", SqlDbType.Int ).Value =
                GetSubOpId( columns[Settings.Default.csaSubOpGroupColumn] );
            csaAttributeSql.Parameters.Add( "@lob", SqlDbType.Int ).Value =
                GetLobId( columns[Settings.Default.csaLobColumn] );
            csaAttributeSql.Parameters.Add( "@jurisdiction", SqlDbType.Int ).Value =
                GetJuriId( columns[Settings.Default.csaJurisdictionColumn] );
            csaAttributeSql.Parameters.Add( "@corpArea", SqlDbType.Int ).Value =
                GetCorpId( columns[Settings.Default.csaCorpAreaColumn] ).CorporateServiceArea;
            csaAttributeSql.Parameters.Add( "@dateCreated", SqlDbType.DateTime ).Value =
                DateTime.Now.ToString();
            csaAttributeSql.Parameters.Add( "@dateUpdated", SqlDbType.DateTime ).Value =
                DateTime.Now.ToString();


            sDebug = "executing create CsaAttributes query" + rowNum + "\r\n" + csaAttributeSql.CommandText;
            csaAttributeSql.ExecuteNonQuery();
            csaAttributeSql.Dispose();



            // Check if this user is a CsaApprover
            sDebug = "Check if this user is a CsaApprover";
            SqlCommand isOpsRiskApproverSql = conn.CreateCommand();
            isOpsRiskApproverSql.CommandText = "select count(*) from CsaApproval where ApproverId=" +
                "@userId and CorporateServiceAreaId = @corpAreaId";

            isOpsRiskApproverSql.Parameters.Add( "@userId", SqlDbType.VarChar, 128 ).Value = userId;
            isOpsRiskApproverSql.Parameters.Add( "@corpAreaId", SqlDbType.Int ).Value =
                GetCorpId( columns[Settings.Default.csaCorpAreaColumn] ).CorporateServiceArea;

            sDebug = "executing is User a CSA Approver query" + rowNum + "\r\n" +
                isOpsRiskApproverSql.CommandText;
            int approvalCount = ( int )(isOpsRiskApproverSql.ExecuteScalar());
            isOpsRiskApproverSql.Dispose();


            if ( approvalCount == 0 )
            {
                // Write record in CsaApprovalRole
                sDebug = "Add record in CsaApproval Role";
                SqlCommand opsRiskApproverRoleSql = conn.CreateCommand();
                opsRiskApproverRoleSql.CommandText = "insert into CsaApproval(ApproverId, Status, CreateDate, " +
                    "CorporateServiceAreaId) Values(@userId, 0, @createDate, @corpSvcAreaId)";

                opsRiskApproverRoleSql.Parameters.Add( "@userId", SqlDbType.VarChar, 128 ).Value = userId;
                opsRiskApproverRoleSql.Parameters.Add( "@createDate", SqlDbType.DateTime ).Value =
                    DateTime.Now.ToString();
                opsRiskApproverRoleSql.Parameters.Add( "@corpSvcAreaId", SqlDbType.Int ).Value =
                    GetCorpId( columns[Settings.Default.csaCorpAreaColumn] ).CorporateServiceArea;

                sDebug = "executing add CSA Approver Role query" + rowNum + "\r\n" +
                    opsRiskApproverRoleSql.CommandText;
                opsRiskApproverRoleSql.ExecuteNonQuery();
                opsRiskApproverRoleSql.Dispose();
            }

            return true;
        }



        private bool ImportOpsRiskUsers( 
            SqlConnection conn, 
            int row,
            string rowNum, 
            string[] columns, 
            string userId, 
            out string sDebug )
        {

            // check if this is the first invocation
            if ( row == Settings.Default.opsRiskRowToStart )
            {

                // Check if there are already OpsRisk Records in the database
                sDebug = "Check if there are already OpsRisk Records in the database";
                if ( !ShouldWeContinueImport( conn, "OpsRiskAttributes" ) )
                {
                    Console.WriteLine();
                    Console.WriteLine( "Terminating application as user requested to abort the operation" );
                    return false;
                }

            }


            // Write OpsRisk Record
            sDebug = "creating OpsRiskAttributes record for the user";
            SqlCommand opsRiskSql = conn.CreateCommand();
            opsRiskSql.CommandText = "insert into OpsRiskAttributes(AppUserId, LineOfBusiness) Values(" +
                "@userId, @lobId)";

            opsRiskSql.Parameters.Add( "@userid", SqlDbType.VarChar, 128 ).Value = userId;
            opsRiskSql.Parameters.Add( "@lobId", SqlDbType.Int ).Value =
                GetLobId( columns[Settings.Default.opsRiskLobColumn] );

            sDebug = "executing create OpsRiskAttributes query" + rowNum + "\r\n" + opsRiskSql.CommandText;
            opsRiskSql.ExecuteNonQuery();
            opsRiskSql.Dispose();


            // Check to see if the user is already an OpsRisk Approver
            sDebug = "Check to see if the user is already an OpsRisk Approver";
            SqlCommand isUserOpsRiskApproverSql = conn.CreateCommand();
            isUserOpsRiskApproverSql.CommandText = "select count(*) from OpsRiskApproval where ApproverId=@userId";
            isUserOpsRiskApproverSql.Parameters.Add( "@userId", SqlDbType.VarChar, 128 ).Value = userId;

            sDebug = "executing check if user is an OpsRisk Approver query" + rowNum + 
                "\r\n" + isUserOpsRiskApproverSql.CommandText;
            int opsRiskApproverCount = (int)isUserOpsRiskApproverSql.ExecuteScalar();

            if ( opsRiskApproverCount == 0 )
            {

                // Write record in OpsRiskApprovalRole
                sDebug = "Write record in OpsRiskApprovalRole";
                SqlCommand opsRiskApproverRoleSql = conn.CreateCommand();
                opsRiskApproverRoleSql.CommandText = "insert into OpsRiskApproval(ApproverId, Status, CreateDate, " +
                    "CorporateServiceAreaId) Values(@userId, 0, @createDate, @corpSvcAreaId)";

                opsRiskApproverRoleSql.Parameters.Add( "@userId", SqlDbType.VarChar, 128 ).Value = userId;
                opsRiskApproverRoleSql.Parameters.Add( "@createDate", SqlDbType.DateTime ).Value = DateTime.Now.ToString();
                opsRiskApproverRoleSql.Parameters.Add( "@corpSvcAreaId", SqlDbType.Int ).Value =
                    GetCorpId( columns[Settings.Default.opsRiskCorpAreaColumn] ).CorporateServiceArea;

                sDebug = "executing add user to OpsRiskApproval Role query" + rowNum + "\r\n" +
                    opsRiskApproverRoleSql.CommandText;
                opsRiskApproverRoleSql.ExecuteNonQuery();
                opsRiskApproverRoleSql.Dispose();

            }

            return true;
        }


        private bool ImportGovernanceUsers(
            SqlConnection conn,
            int row,
            string rowNum,
            string[] columns,
            string userId,
            out string sDebug )
        {

            // check to see if this is the first invocation
            if ( row == Settings.Default.govRowToStart )
            {

                // Check if there are already OpsRisk Records in the database
                sDebug = "Check if there are already GovernanceUser Records in the database";
                if ( !ShouldWeContinueImport( conn, "GovernanceAttributes" ) )
                {
                    Console.WriteLine();
                    Console.WriteLine( "Terminating application as user requested to abort the operation" );
                    return false;
                }

            }


            // Write Governance Record
            sDebug = "creating GovernanceAttributes record for the user";
            SqlCommand governanceAttributeSql = conn.CreateCommand();
            governanceAttributeSql.CommandText = "insert into GovernanceAttributes(AppUserId, OperatingGroup, " +
                "SubOperatingGroup, LineOfBusiness) Values(@userId, @opGroup, @subOpGroup, @lob)";

            governanceAttributeSql.Parameters.Add( "@userId", SqlDbType.VarChar, 128 ).Value = userId;
            governanceAttributeSql.Parameters.Add( "@opGroup", SqlDbType.Int ).Value =
                GetOpId( columns[Settings.Default.govOpGroupColumn] );
            governanceAttributeSql.Parameters.Add( "@subOpGroup", SqlDbType.Int ).Value =
                GetSubOpId( columns[Settings.Default.govSubOpGroupColumn] );
            governanceAttributeSql.Parameters.Add( "@lob", SqlDbType.Int ).Value =
                GetLobId( columns[Settings.Default.govLobColumn] );

            sDebug = "executing create GovernanceAttributes query" + rowNum + "\r\n" +
                governanceAttributeSql.CommandText;
            governanceAttributeSql.ExecuteNonQuery();
            governanceAttributeSql.Dispose();


            // check to see if the user is already a Governance Approver
            sDebug = "check to see if the user is already a Governance Approver";
            SqlCommand isGovApproverSql = conn.CreateCommand();
            isGovApproverSql.CommandText = "select count(*) from GovernanceApproval where ApproverId = @userId";

            isGovApproverSql.Parameters.Add( "@userId", SqlDbType.VarChar, 128 ).Value = userId;

            int govApproverCount = ( int )isGovApproverSql.ExecuteScalar();

            if ( govApproverCount == 0 )
            {

                // Write record in OpsRiskApprovalRole
                sDebug = "Write record in govApproverRole";
                SqlCommand govApproverRoleSql = conn.CreateCommand();
                govApproverRoleSql.CommandText = "insert into GovernanceApproval(ApproverId, Status, CreateDate, " +
                    "CorporateServiceAreaId) Values(@userId, 0, @createDate, @corpSvcAreaId)";

                govApproverRoleSql.Parameters.Add( "@userId", SqlDbType.VarChar, 128 ).Value = userId;
                govApproverRoleSql.Parameters.Add( "@createDate", SqlDbType.DateTime ).Value = DateTime.Now.ToString();
                govApproverRoleSql.Parameters.Add( "@corpSvcAreaId", SqlDbType.Int ).Value =
                    GetCorpId( columns[Settings.Default.govCorpAreaColumn] ).CorporateServiceArea;

                sDebug = "executing create govApproverRole query" + rowNum + "\r\n" +
                    govApproverRoleSql.CommandText;
                govApproverRoleSql.ExecuteNonQuery();
                govApproverRoleSql.Dispose();

            }

            return true;
        }


        private bool ShouldWeContinueImport( SqlConnection conn, string tableName )
        {
            SqlCommand opsRiskCountSql = conn.CreateCommand();
            opsRiskCountSql.CommandText = "select count(*) from " + tableName;
            int opsRiskCount = ( int )opsRiskCountSql.ExecuteScalar();
            if ( opsRiskCount > 0 )
            {
                Console.WriteLine( "Attention: There are existing records in {0}. " +
                    "Do you wish to Continue?\r\nPress x to continue importing and any other key to abort",
                    tableName);
                char keyPressed = Console.ReadKey().KeyChar;
                if ( keyPressed != 'x' && keyPressed != 'X' )
                {
                    return false;
                }
            }

            Console.WriteLine();
            Console.WriteLine( "Working..." );
            Console.WriteLine();

            return true;
        }



        #region Name to ID mapping functions for user properties

        private int GetOpId( string opText )
        {
            int result = 0;

            switch ( opText )
            {
                case "Canadian P&C":
                    result = 1;
                    break;
                case "Capital Markets":
                    result = 2;
                    break;
                case "Corporate":
                    result = 3;
                    break;
                case "T&O":
                    result = 4;
                    break;
                case "US P&C":
                    result = 5;
                    break;
                case "Wealth Management":
                    result = 6;
                    break;
                default:
                    result = 0;
                    Console.WriteLine( "Encountered unknown OperatingGroup: {0}", opText );
                    break;
            }
            return result;
        }


        private int GetSubOpId( string opText )
        {
            int result = 0;

            switch ( opText )
            {
                case "Canadian P&C":
                    result = 1;
                    break;
                case "Asset Management":
                    result = 2;
                    break;
                case "T&O":
                    result = 3;
                    break;
                case "NA Commercial":
                    result = 4;
                    break;
                case "Wealth Management":
                    result = 5;
                    break;
                default:
                    result = 0;
                    Console.WriteLine( "Encountered unknown SubOperatingGroup: {0}", opText );
                    break;
            }
            return result;
        }


        private int GetLobId( string lobText )
        {
            int result = 0;

            switch ( lobText )
            {
                case "Asia Growth":
                    result = 1;
                    break;
                case "BMO Insurance (Creditor & Reinsurance)":
                    result = 2;
                    break;
                case "BMO InvestorLine":
                    result = 3;
                    break;
                case "BMO Life Assurance":
                    result = 4;
                    break;
                case "BMO Private Banking":
                    result = 5;
                    break;
                case "Canadian P&C - Business Banking":
                    result = 6;
                    break;
                case "Canadian P&C - Personal":
                    result = 7;
                    break;
                case "Corporate Finance":
                    result = 8;
                    break;
                case "Corporate Real Estate":
                    result = 9;
                    break;
                case "Data Analytics":
                    result = 10;
                    break;
                case "Full-Service Investing":
                    result = 11;
                    break;
                case "Global Asset Management":
                    result = 12;
                    break;
                case "Global Product Operations":
                    result = 13;
                    break;
                case "HQ":
                    result = 14;
                    break;
                case "IT Risk Management":
                    result = 15;
                    break;
                case "North American Integrated Channels":
                    result = 16;
                    break;
                case "North American Retail Payments":
                    result = 17;
                    break;
                case "North American Treasury & Payment Solutions":
                    result = 18;
                    break;
                case "Procurement":
                case "Procurements":
                    result = 19;
                    break;
                case "Technology":
                    result = 20;
                    break;
                case "U.S. Personal Wealth (in $USD)":
                    result = 21;
                    break;
                default:
                    result = 0;
                    Console.WriteLine( "Encountered unknown LOB: {0}", lobText );
                    break;
            }

            return result;
        }


        private int GetJuriId( string juriText )
        {
            int result = 0;

            switch ( juriText )
            {
                case "All":
                case "Other":
                    result = 1;
                    break;
                case "Brazil":
                    result = 2;
                    break;
                case "Canada":
                    result = 3;
                    break;
                case "China":
                    result = 4;
                    break;
                case "Hong Kong":
                    result = 5;
                    break;
                case "India":
                    result = 6;
                    break;
                case "Ireland":
                    result = 7;
                    break;
                case "Italy":
                    result = 8;
                    break;
                case "Mexico":
                    result = 9;
                    break;
                case "Singapore":
                    result = 10;
                    break;
                case "UK - United Kingdom":
                    result = 11;
                    break;
                case "US":
                    result = 12;
                    break;
                default:
                    result = 0;
                    Console.WriteLine( "Encountered unknown Jurisdiction: {0}", juriText );
                    break;
            }

            return result;
        }


        private CorporateServiceAreaAndGroup GetCorpId( string corpText )
        {
            CorporateServiceAreaAndGroup result = new CorporateServiceAreaAndGroup();

            switch ( corpText )
            {
                case "Accounting & Financial Management":
                case "Accounting and Financial Management":
                case "Accounting & Financial Management Risk":
                    result.CorporateServiceArea = 13;
                    result.CorporateServiceAreaGroup = 3;
                    break;

                case "AML":
                case "Anti Money Laundering":
                    result.CorporateServiceArea = 21;
                    result.CorporateServiceAreaGroup = 6;
                    break;

                case "Business Continuity":
                case "Business Continuity Risk":
                    result.CorporateServiceArea = 12;
                    result.CorporateServiceAreaGroup = 2;
                    break;

                case "Credit Risk":
                case "Credit Risk & CCR":
                case "Credit Risk - Commercial":
                case "Credit Risk - Retail":
                case "Credit Risk - Business Banking":
                    result.CorporateServiceArea = 7;
                    result.CorporateServiceAreaGroup = 2;
                    break;

                case "Fraud/Criminal":
                    result.CorporateServiceArea = 5;
                    result.CorporateServiceAreaGroup = 1;
                    break;

                case "Human Resources":
                    result.CorporateServiceArea = 20;
                    result.CorporateServiceAreaGroup = 5;
                    break;

                case "Information Security & Management":
                case "Information Security & Information Management":
                    result.CorporateServiceArea = 16;
                    result.CorporateServiceAreaGroup = 4;
                    break;

                case "Legal":
                    result.CorporateServiceArea = 1;
                    result.CorporateServiceAreaGroup = 1;
                    break;

                case "Liquidity and Funding":
                case "Liquidity and Funding Risk":
                    result.CorporateServiceArea = 9;
                    result.CorporateServiceAreaGroup = 2;
                    break;

                case "Market Risk":
                case "Market Risk (Trading, Structural)":
                case "Market Risk (Trading, Structural, CCR)":
                    result.CorporateServiceArea = 8;
                    result.CorporateServiceAreaGroup = 2;
                    break;

                case "Model Risk":
                    result.CorporateServiceArea = 10;
                    result.CorporateServiceAreaGroup = 2;
                    break;

                case "Outsourcing & Supplier":
                    result.CorporateServiceArea = 17;
                    result.CorporateServiceAreaGroup = 4;
                    break;

                case "Physical Security":
                    result.CorporateServiceArea = 6;
                    result.CorporateServiceAreaGroup = 1;
                    break;

                case "Privacy":
                    result.CorporateServiceArea = 3;
                    result.CorporateServiceAreaGroup = 1;
                    break;

                case "Process & Operational Risk":
                    result.CorporateServiceArea = 11;
                    result.CorporateServiceAreaGroup = 2;
                    break;

                case "Project Management":
                case "T&O - Project Management":
                    result.CorporateServiceArea = 19;
                    result.CorporateServiceAreaGroup = 4;
                    break;

                case "Property":
                case "T&O Property":
                    result.CorporateServiceArea = 18;
                    result.CorporateServiceAreaGroup = 4;
                    break;

                case "Regulatory Compliance":
                    result.CorporateServiceArea = 2;
                    result.CorporateServiceAreaGroup = 1;
                    break;

                case "Reputation Risk":
                    result.CorporateServiceArea = 4;
                    result.CorporateServiceAreaGroup = 1;
                    break;

                case "Tax":
                    result.CorporateServiceArea = 14;
                    result.CorporateServiceAreaGroup = 3;
                    break;

                case "Technology":
                    result.CorporateServiceArea = 15;
                    result.CorporateServiceAreaGroup = 4;
                    break;

                default:
                    result.CorporateServiceArea = 0;
                    result.CorporateServiceAreaGroup = 0;
                    Console.WriteLine( "Encountered unknown CorporateServiceArea: {0}", corpText );
                    break;
            }

            return result;
        }

        #endregion
    }
}
