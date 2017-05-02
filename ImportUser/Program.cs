using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.Sql;
using System.Data.SqlClient;
using ImportUser.Properties;
using System.Diagnostics;

namespace ImportUser
{
    class Program
    {
        public enum UserImportFileType
        {
            CSA = 0,
            OpsRisk,
            Governance,
            Undefined
        }

        static void Main( string[] args )
        {
            string sDebug = String.Empty;
            bool bError = false;

            try
            {

                if ( args.Length == 2 && !string.IsNullOrEmpty( args[0] ) )
                {
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

                    sDebug = "Reading rows";
                    string[] fileRows = fileContent.Split( new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries );
                    Console.WriteLine( "Rows read: {0}", fileRows.Length );
                    #endregion


                    #region connect to database
                    sDebug = "connect to database";
                    SqlConnection conn = new SqlConnection( Settings.Default.dbConnectionString );
                    conn.Open();
                    #endregion


                    // Loop through the rows - skipping the first row
                    for ( int row = Settings.Default.rowToStart; row < fileRows.Length; row++ )
                    {
                        string rowNum = " when reading row " + row;

                        // read columns
                        sDebug = "read columns" + rowNum;
                        string[] columns = fileRows[row].Split( new char[] { ',' } );


                        #region check if user exists and get user Id
                        bool isExistingUser = false;
                        string userId = String.Empty;

                        
                        sDebug = "checking if the user already exists " + rowNum;
                        string userName = columns[Settings.Default.nameColumn].Trim().Replace( "'", "''" );
                        string userEmail = columns[Settings.Default.emailColumn];
                        SqlCommand checkUserSql = conn.CreateCommand();
                        checkUserSql.CommandText = "select id from AspNetUsers where Email = '" + userEmail + "'";

                        sDebug += "\r\n" + checkUserSql.CommandText;
                        userId = ( string )(checkUserSql.ExecuteScalar());
                        checkUserSql.Dispose();
                        if ( string.IsNullOrEmpty( userId ) )
                        {
                            userId = Guid.NewGuid().ToString();
                        }
                        else
                        {
                            isExistingUser = true;
                        }

                        #endregion


                        #region not requried anymore
                        // check if user data file has email address for them
                        //string userEmail = columns[Settings.Default.emailColumn];
                        //if ( string.IsNullOrEmpty( userEmail ) && Settings.Default.createEmail )
                        //{
                        //    userEmail = columns[Settings.Default.nameColumn].Trim().Replace( "'", "" ).Replace( ' ', '.' ) + "@bmo.com";
                        //}
                        #endregion


                        if ( !isExistingUser )
                        {
                            #region create user record
                            sDebug = "creating user record query" + rowNum;

                            SqlCommand userRecordSql = conn.CreateCommand();
                            userRecordSql.CommandText = "insert into AspNetUsers " +
                                "( Id, Name, LastModifiedDate, EmailConfirmed, PhoneNumberConfirmed, " +
                                "TwoFactorEnabled, LockoutEnabled, AccessFailedCount, UserName, Email ) " +
                                "values " +
                                "( '" + userId + "', '" + userName +
                                "', '" + DateTime.Now.ToString() + "', 0, 0, 0, 0, 0, '" +
                                userEmail + "', '" + userEmail + "')";

                            sDebug = "executing create user query" + rowNum + "\r\n" + userRecordSql.CommandText;
                            userRecordSql.ExecuteNonQuery();
                            userRecordSql.Dispose();
                            #endregion                            
                        }



                        if ( fileType == UserImportFileType.CSA )
                        {

                            // Write CsaAttributes
                            sDebug = "creating CsaAttributes record for the user";
                            SqlCommand csaAttributeSql = conn.CreateCommand();
                            csaAttributeSql.CommandText = "insert into CsaAttributes(AppUserId, OperatingGroup, SubOperatingGroup, " +
                                "LineOfBusiness, Jurisdiction, CorporateServiceArea) Values('" + userId + "', '" + 
                                GetOpId( columns[Settings.Default.opGroupColumn] ) + "', '" + 
                                GetSubOpId( columns[Settings.Default.subOpGroupColumn] ) + "', '" + 
                                GetLobId( columns[Settings.Default.lobColumn] ) + "', '" + 
                                GetJuriId( columns[Settings.Default.jurisdictionColumn] ) + "', '" + 
                                GetCorpId( columns[Settings.Default.corpAreaColumn] ).CorporateServiceArea + "')";

                            sDebug = "executing create CsaAttributes query" + rowNum + "\r\n" + csaAttributeSql.CommandText;
                            csaAttributeSql.ExecuteNonQuery();
                            csaAttributeSql.Dispose();


                            // Write record in CsaApprovalRole
                            sDebug = "Add record in CsaApproval Role";
                            SqlCommand opsRiskApproverRoleSql = conn.CreateCommand();
                            opsRiskApproverRoleSql.CommandText = "insert into CsaApproval(ApproverId, Status, CreateDate, " +
                                "CorporateServiceAreaId) Values('" + userId + "', 0, '" + DateTime.Now + 
                                "', " + GetCorpId( columns[Settings.Default.corpAreaColumn] ).CorporateServiceArea + ")";

                            sDebug = "executing add CSA Approver Role query" + rowNum + "\r\n" +
                                opsRiskApproverRoleSql.CommandText;
                            opsRiskApproverRoleSql.ExecuteNonQuery();
                            opsRiskApproverRoleSql.Dispose();

                            #region Not Required Anymore

                            //#region create Line of Business App Users record(s)
                            //if ( string.IsNullOrEmpty( columns[Settings.Default.lobColumn] ) )
                            //{
                            //    sDebug = "creating all possible LineOfBusinessAppUsers" + rowNum;
                            //    for ( int i = 1; i < 39; i++ )
                            //    {
                            //        SqlCommand lobUserSql = conn.CreateCommand();
                            //        lobUserSql.CommandText = "insert into LineofBusinessAppUsers" +
                            //            "    ( LineOfBusiness_id, AppUser_Id ) " +
                            //            "values" +
                            //            "    ( " + i + ", '" + userId + "' )";

                            //        sDebug = "executing create LineOfBusinessAppUser query" + rowNum + "\r\n" + lobUserSql.CommandText;
                            //        lobUserSql.ExecuteNonQuery();
                            //        lobUserSql.Dispose();
                            //    }
                            //}
                            //else
                            //{
                            //    sDebug = "creating specific LineOfBusinessappUsers" + rowNum;
                            //    int lobId = GetLobId( columns[Settings.Default.lobColumn].Trim() );
                            //    SqlCommand lobUserSql = conn.CreateCommand();
                            //    lobUserSql.CommandText = "insert into LineofBusinessAppUsers" +
                            //        "    ( LineOfBusiness_id, AppUser_Id ) " +
                            //        "values" +
                            //        "    ( " + lobId + ", '" + userId + "' )";

                            //    sDebug = "executing create LineOfBusinessAppUser query" + rowNum + "\r\n" + lobUserSql.CommandText;
                            //    lobUserSql.ExecuteNonQuery();
                            //    lobUserSql.Dispose();
                            //}
                            //#endregion


                            //#region create Jurisdiction App Users record(s)
                            //if ( string.IsNullOrEmpty( columns[Settings.Default.jurisdictionColumn] ) )
                            //{
                            //    sDebug = "creating all possible JurisdictionAppUsers" + rowNum;
                            //    for ( int i = 1; i < 9; i++ )
                            //    {
                            //        SqlCommand jurisdictionUserSql = conn.CreateCommand();
                            //        jurisdictionUserSql.CommandText = "insert into JurisdictionAppUsers" +
                            //        "    ( Jurisdiction_Id, AppUser_Id ) " +
                            //        "values" +
                            //        "    ( " + i + ", '" + userId + "' )";

                            //        sDebug = "executing create JurisdictionAppUsers query" + rowNum + "\r\n" + jurisdictionUserSql.CommandText;
                            //        jurisdictionUserSql.ExecuteNonQuery();
                            //        jurisdictionUserSql.Dispose();
                            //    }
                            //}
                            //else
                            //{
                            //    sDebug = "creating specific JurisdictionAppUsers" + rowNum;
                            //    int juriId = GetJuriId( columns[Settings.Default.jurisdictionColumn].Trim() );
                            //    SqlCommand jurisdictionUserSql = conn.CreateCommand();
                            //    jurisdictionUserSql.CommandText = "insert into JurisdictionAppUsers" +
                            //    "    ( Jurisdiction_Id, AppUser_Id ) " +
                            //    "values" +
                            //    "    ( " + juriId + ", '" + userId + "' )";

                            //    sDebug = "executing create JurisdictionAppUsers query" + rowNum + "\r\n" + jurisdictionUserSql.CommandText;
                            //    jurisdictionUserSql.ExecuteNonQuery();
                            //    jurisdictionUserSql.Dispose();
                            //}
                            //#endregion


                            //#region create Corporate Service Area User record(s)
                            //if ( string.IsNullOrEmpty( columns[Settings.Default.corpAreaColumn] ) )
                            //{
                            //    sDebug = "creating all possible CorporateServiceAreaAppUsers" + rowNum;
                            //    for ( int i = 1; i < 29; i++ )
                            //    {
                            //        SqlCommand corpServAreaUserSql = conn.CreateCommand();
                            //        corpServAreaUserSql.CommandText = "insert into CorporateServiceAreaAppUsers" +
                            //        "    ( CorporateServiceArea_Id, AppUser_Id ) " +
                            //        "values" +
                            //        "    ( " + i + ", '" + userId + "' )";

                            //        sDebug = "executing create CorporateServiceAreaAppUsers query" + rowNum + "\r\n" + corpServAreaUserSql.CommandText;
                            //        corpServAreaUserSql.ExecuteNonQuery();
                            //        corpServAreaUserSql.Dispose();
                            //    }
                            //}
                            //else
                            //{
                            //    sDebug = "creating specific CorporateServiceAreaAppUsers" + rowNum;
                            //    int corpId = GetCorpId( columns[Settings.Default.corpAreaColumn].Trim() );
                            //    SqlCommand corpServAreaUserSql = conn.CreateCommand();
                            //    corpServAreaUserSql.CommandText = "insert into CorporateServiceAreaAppUsers" +
                            //    "    ( CorporateServiceArea_Id, AppUser_Id ) " +
                            //    "values" +
                            //    "    ( " + corpId + ", '" + userId + "' )";

                            //    sDebug = "executing create CorporateServiceAreaAppUsers query" + rowNum + "\r\n" + corpServAreaUserSql.CommandText;
                            //    corpServAreaUserSql.ExecuteNonQuery();
                            //    corpServAreaUserSql.Dispose();
                            //}
                            //#endregion

                            #endregion

                        }
                        else if ( fileType == UserImportFileType.OpsRisk )
                        {

                            // Write OpsRisk Record
                            sDebug = "creating OpsRiskAttributes record for the user";
                            SqlCommand opsRiskSql = conn.CreateCommand();
                            opsRiskSql.CommandText = "insert into OpsRiskAttributes(AppUserId, LineOfBusiness) Values('" + 
                                userId + "', '" + GetLobId( columns[Settings.Default.lobColumn] ) + "')";

                            sDebug = "executing create OpsRiskAttributes query" + rowNum + "\r\n" + opsRiskSql.CommandText;
                            opsRiskSql.ExecuteNonQuery();
                            opsRiskSql.Dispose();


                            //TODO: Write record in OpsRiskApprovalRole
                            sDebug = "Write record in OpsRiskApprovalRole";
                            SqlCommand opsRiskApproverRoleSql = conn.CreateCommand();
                            opsRiskApproverRoleSql.CommandText = "insert into OpsRiskAttributes(AppUserId, LineOfBusiness) Values('" +
                                columns[0] + "', '" + columns[0] + "')";

                            sDebug = "executing create OpsRiskApproval query" + rowNum + "\r\n" + 
                                opsRiskApproverRoleSql.CommandText;
                            opsRiskApproverRoleSql.ExecuteNonQuery();
                            opsRiskApproverRoleSql.Dispose();

                        }
                        else if ( fileType == UserImportFileType.Governance )
                        {

                            // Write Governance Record
                            sDebug = "creating GovernanceAttributes record for the user";
                            SqlCommand governanceAttributeSql = conn.CreateCommand();
                            governanceAttributeSql.CommandText = "insert into GovernanceAttributes(AppUserId, OperatingGroup, " +
                                "SubOperatingGroup, LineOfBusiness) Values('" + userId + "', '" + 
                                GetOpId( columns[Settings.Default.opGroupColumn] ) + "', '" + 
                                GetSubOpId( columns[Settings.Default.subOpGroupColumn] ) + "', '" + 
                                GetLobId( columns[Settings.Default.lobColumn] ) + "')";

                            sDebug = "executing create GovernanceAttributes query" + rowNum + "\r\n" + 
                                governanceAttributeSql.CommandText;
                            governanceAttributeSql.ExecuteNonQuery();
                            governanceAttributeSql.Dispose();


                            //TODO: Write record in OpsRiskApprovalRole
                            sDebug = "Write record in govApproverRole";
                            SqlCommand govApproverRoleSql = conn.CreateCommand();
                            govApproverRoleSql.CommandText = "insert into GovernanceApproval(AppUserId, LineOfBusiness) Values('" +
                                columns[0] + "', '" + columns[0] + "')";

                            sDebug = "executing create govApproverRole query" + rowNum + "\r\n" +
                                govApproverRoleSql.CommandText;
                            govApproverRoleSql.ExecuteNonQuery();
                            govApproverRoleSql.Dispose();

                        }
                    }

                }
                else
                {
                    Console.WriteLine( "Incorrect Arguments" );
                    DisplayUsage();
                    bError = true;
                }

            }
            catch ( ArgumentException ae )
            {
                Console.WriteLine( "exception occured when {0}:\r\n{1}", sDebug, ae.Message );
                DisplayUsage();
                bError = true;
            }
            catch ( Exception e )
            {
                Console.WriteLine( "exception occured when {0}:\r\n{1}", sDebug, e.Message );
                bError = true;
            }
            finally
            {
                Console.WriteLine( "User Import process completed " + (bError ? "with errors" : "successfully") + "..." );
            }
            
            Console.ReadKey();
        }

        static void DisplayUsage( )
        {
            Console.WriteLine( "Usage: Program.exe <fileName> <fileType: 0=CSA, 1=OpsRisk, 2=Governance>" );
            Console.WriteLine( "Note: User must have access to the Iaap database" );
        }


        #region Name to ID mapping functions for user properties

        static int GetOpId( string opText )
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
                    break;
            }
            return result;
        }

        static int GetSubOpId( string opText )
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
                    break;
            }
            return result;
        }

        static int GetLobId( string lobText )
        {
            int result = 0;

            switch ( lobText )
            {
                case "Asia Growth":
                    result = 1;
                    break;
                case "BMO Insurance( Creditor & Reinsurance)":
                    result = 2;
                    break;
                case "BMO InvestorLine":
                    result = 3;
                    break;
                case "BMO LIfe Assurance":
                    result = 4;
                    break;
                case "BMO Private Banking":
                    result = 5;
                    break;
                case "Canadian P&C - Business Bnaking":
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
                case "Data Analyitcs":
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
                    break;
            }

            return result;
        }

        static int GetJuriId( string juriText )
        {
            int result = 0;

            switch ( juriText )
            {
                case "All":
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
                    break;
            }

            return result;
        }

        static CorporateServiceAreaAndGroup GetCorpId( string corpText )
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
                    break;
            }

            return result;
        }

        #endregion

    }

    public struct CorporateServiceAreaAndGroup
    {
        public int CorporateServiceArea { get; set; }
        public int CorporateServiceAreaGroup { get; set; }
        public string CorporateServiceAreaName { get; set; }
        public string CorporateServiceAreaGroupName { get; set; }
    }

}

