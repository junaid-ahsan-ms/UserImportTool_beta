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
        static void Main( string[] args )
        {
            string sDebug = String.Empty;
            bool bError = false;

            try
            {

                if ( args.Length > 0 )
                {
                    #region Read User data file
                    sDebug = "Finding file";
                    Console.WriteLine( "filename: {0}", args[0] );
                    //@"D:\My Engagements\BMO IAAP\User Import\CSA Data Master Unpivtoed (March 13 2017)_Edited_v0.1.csv"
                    FileStream fs = File.OpenRead( args[0] );

                    sDebug = "Reading file";
                    StreamReader sr = new StreamReader( fs );
                    string fileContent = sr.ReadToEnd();

                    sDebug = "Reading rows";
                    string [] fileRows = fileContent.Split( new string[] { "\r\n" },StringSplitOptions.RemoveEmptyEntries );
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
                        SqlCommand checkUserSql = conn.CreateCommand();
                        checkUserSql.CommandText = "select id from AspNetUsers where Name = '" + userName + "'";

                        sDebug += "\r\n" + checkUserSql.CommandText;
                        userId = (string)(checkUserSql.ExecuteScalar());
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


                        // check if user data file has email address for them
                        string userEmail = columns[Settings.Default.emailColumn];
                        if ( string.IsNullOrEmpty( userEmail ) && Settings.Default.createEmail )
                        {
                            userEmail = columns[Settings.Default.nameColumn].Trim().Replace( "'", "" ).Replace( ' ', '.' ) + "@bmo.com";
                        }


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


                            #region create Operating Group record(s)                        
                            if ( string.IsNullOrEmpty( columns[Settings.Default.opGroupColumn] ) )
                            {
                                //sDebug = "creating all possible OperatingGroups" + rowNum;
                                //for ( int i = 1; i < 11; i++ )
                                //{
                                //    SqlCommand operatingGroupSql = conn.CreateCommand();
                                //    operatingGroupSql.CommandText = "insert into OperatingGroupAppUsers" +
                                //    "    ( OperatingGroup_Id, AppUser_Id ) " +
                                //    "values" +
                                //    "    ( " + i + ", '" + userId + "' )";

                                //    sDebug = "executing create OperatingGroup query" + rowNum + "\r\n" + operatingGroupSql.CommandText;
                                //    operatingGroupSql.ExecuteNonQuery();
                                //    operatingGroupSql.Dispose();
                                //}
                            }
                            else
                            {
                                sDebug = "creating specific OperatingGroup" + rowNum;
                                int opId = GetOpId( columns[Settings.Default.opGroupColumn].Trim() );
                                SqlCommand operatingGroupSql = conn.CreateCommand();
                                operatingGroupSql.CommandText = "insert into OperatingGroupAppUsers" +
                                    "    ( OperatingGroup_Id, AppUser_Id ) " +
                                    "values" +
                                    "    ( " + opId + ", '" + userId + "' )";

                                sDebug = "executing create OperatingGroup query" + rowNum + "\r\n" + operatingGroupSql.CommandText;
                                operatingGroupSql.ExecuteNonQuery();
                                operatingGroupSql.Dispose();
                            }
                            #endregion
                        }


                        #region create Line of Business App Users record(s)
                        if ( string.IsNullOrEmpty( columns[Settings.Default.lobColumn] ) )
                        {
                            sDebug = "creating all possible LineOfBusinessAppUsers" + rowNum;
                            for ( int i = 1; i < 39; i++ )
                            {
                                SqlCommand lobUserSql = conn.CreateCommand();
                                lobUserSql.CommandText = "insert into LineofBusinessAppUsers" +
                                    "    ( LineOfBusiness_id, AppUser_Id ) " +
                                    "values" +
                                    "    ( " + i + ", '" + userId + "' )";

                                sDebug = "executing create LineOfBusinessAppUser query" + rowNum + "\r\n" + lobUserSql.CommandText;
                                lobUserSql.ExecuteNonQuery();
                                lobUserSql.Dispose();
                            }
                        }
                        else
                        {
                            sDebug = "creating specific LineOfBusinessappUsers" + rowNum;
                            int lobId = GetLobId( columns[Settings.Default.lobColumn].Trim() );
                            SqlCommand lobUserSql = conn.CreateCommand();
                            lobUserSql.CommandText = "insert into LineofBusinessAppUsers" +
                                "    ( LineOfBusiness_id, AppUser_Id ) " +
                                "values" +
                                "    ( " + lobId + ", '" + userId + "' )";

                            sDebug = "executing create LineOfBusinessAppUser query" + rowNum + "\r\n" + lobUserSql.CommandText;
                            lobUserSql.ExecuteNonQuery();
                            lobUserSql.Dispose();
                        }
                        #endregion


                        #region create Jurisdiction App Users record(s)
                        if ( string.IsNullOrEmpty( columns[Settings.Default.jurisdictionColumn] ) )
                        {
                            sDebug = "creating all possible JurisdictionAppUsers" + rowNum;
                            for ( int i = 1; i < 9; i++ )
                            {
                                SqlCommand jurisdictionUserSql = conn.CreateCommand();
                                jurisdictionUserSql.CommandText = "insert into JurisdictionAppUsers" +
                                "    ( Jurisdiction_Id, AppUser_Id ) " +
                                "values" +
                                "    ( " + i + ", '" + userId + "' )";

                                sDebug = "executing create JurisdictionAppUsers query" + rowNum + "\r\n" + jurisdictionUserSql.CommandText;
                                jurisdictionUserSql.ExecuteNonQuery();
                                jurisdictionUserSql.Dispose();
                            }
                        }
                        else
                        {
                            sDebug = "creating specific JurisdictionAppUsers" + rowNum;
                            int juriId = GetJuriId( columns[Settings.Default.jurisdictionColumn].Trim() );
                            SqlCommand jurisdictionUserSql = conn.CreateCommand();
                            jurisdictionUserSql.CommandText = "insert into JurisdictionAppUsers" +
                            "    ( Jurisdiction_Id, AppUser_Id ) " +
                            "values" +
                            "    ( " + juriId + ", '" + userId + "' )";

                            sDebug = "executing create JurisdictionAppUsers query" + rowNum + "\r\n" + jurisdictionUserSql.CommandText;
                            jurisdictionUserSql.ExecuteNonQuery();
                            jurisdictionUserSql.Dispose();
                        }
                        #endregion


                        #region create Corporate Service Area User record(s)
                        if ( string.IsNullOrEmpty( columns[Settings.Default.corpAreaColumn] ) )
                        {
                            sDebug = "creating all possible CorporateServiceAreaAppUsers" + rowNum;
                            for ( int i = 1; i < 29; i++ )
                            {
                                SqlCommand corpServAreaUserSql = conn.CreateCommand();
                                corpServAreaUserSql.CommandText = "insert into CorporateServiceAreaAppUsers" +
                                "    ( CorporateServiceArea_Id, AppUser_Id ) " +
                                "values" +
                                "    ( " + i + ", '" + userId + "' )";

                                sDebug = "executing create CorporateServiceAreaAppUsers query" + rowNum + "\r\n" + corpServAreaUserSql.CommandText;
                                corpServAreaUserSql.ExecuteNonQuery();
                                corpServAreaUserSql.Dispose();
                            }
                        }
                        else
                        {
                            sDebug = "creating specific CorporateServiceAreaAppUsers" + rowNum;
                            int corpId = GetCorpId( columns[Settings.Default.corpAreaColumn].Trim() );
                            SqlCommand corpServAreaUserSql = conn.CreateCommand();
                            corpServAreaUserSql.CommandText = "insert into CorporateServiceAreaAppUsers" +
                            "    ( CorporateServiceArea_Id, AppUser_Id ) " +
                            "values" +
                            "    ( " + corpId + ", '" + userId + "' )";

                            sDebug = "executing create CorporateServiceAreaAppUsers query" + rowNum + "\r\n" + corpServAreaUserSql.CommandText;
                            corpServAreaUserSql.ExecuteNonQuery();
                            corpServAreaUserSql.Dispose();
                        }
                        #endregion
                    }

                }
                else
                {
                    Console.WriteLine( "no filename" );
                    bError = true;
                }
                
            }
            catch ( Exception e )
            {
                Console.WriteLine( "exception occured when {0}:\r\n{1}", sDebug, e.Message );
                bError = true;
            }

            Console.WriteLine( "User Import process completed " + (bError ? "with errors" : "successfully") + "..." );
            Console.ReadKey();
        }

        #region Name to ID mapping functions for user properties

        static int GetOpId( string opText )
        {
            int result = 0;

            switch ( opText )
            {
                case "P&C CA":
                    result = 1;
                    break;
                //case "P&C US":
                //    result = 2;     // same as 9
                //    break;
                case "BMO CM":
                    result = 3;
                    break;
                //case "Wealth Management":   // same as 7
                //    result = 4;
                //    break;
                case "T&O":
                    result = 5;
                    break;
                case "Corporate":
                    result = 6;
                    break;
                case "Wealth Management":
                    result = 7;
                    break;
                case "Asset Management":
                    result = 8;
                    break;
                case "US P&C":
                    result = 9;
                    break;
                case "NA Commercial":
                    result = 10;
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
                case "AML":
                    result = 1;
                    break;
                case "Audit":
                    result = 2;
                    break;
                case "BMO InvestorLine":
                    result = 3;
                    break;
                case "BMO Life Insurance":
                    result = 4;
                    break;
                case "BMO Nesbitt Burns":
                    result = 5;
                    break;
                case "BMO Partners":
                    result = 6;
                    break;
                case "BMO Private Bank":
                    result = 7;
                    break;
                case "Channels":
                    result = 8;
                    break;
                case "Commercial Banking":
                    result = 9;
                    break;
                case "Corporate Banking":
                    result = 10;
                    break;
                case "CorporateFinance Banking":
                    result = 11;
                    break;
                case "Everyday Banking":
                    result = 12;
                    break;
                case "Finance & Communications":
                    result = 13;
                    break;
                case "GAM":
                    result = 14;
                    break;
                case "Global Trade & Banking":
                    result = 15;
                    break;
                case "HR":
                    result = 16;
                    break;
                case "Institutional Trust Services":
                    result = 17;
                    break;
                case "Investment Banking":
                    result = 18;
                    break;
                case "Legal":
                    result = 19;
                    break;
                case "Marketing":
                    result = 20;
                    break;
                case "North American Retail Payments":
                    result = 21;
                    break;
                case "Payment Cards":
                    result = 22;
                    break;
                case "Payment Cards (US)":
                    result = 23;
                    break;
                case "Personal Lending":
                    result = 24;
                    break;
                case "Retail & Business Banking":
                    result = 25;
                    break;
                case "Retail Collections":
                    result = 26;
                    break;
                case "Risk":
                    result = 27;
                    break;
                case "SAMU":
                    result = 28;
                    break;
                case "SBSA":
                    result = 29;
                    break;
                case "Strategy":
                    result = 30;
                    break;
                case "Structured Notes":
                    result = 31;
                    break;
                //case "Treasury & Payment Solutions":
                //    result = 32;
                //    break;
                case "Term Investments":
                    result = 33;
                    break;
                case "Trading Products":
                    result = 34;
                    break;
                case "Trading Products (Equity Derivatives, Commodities)":
                    result = 35;
                    break;
                case "Trading Products (Other Derivatives)":
                    result = 36;
                    break;
                case "Trading Products (Securities Lending and Repos)":
                    result = 37;
                    break;
                case "Treasury & Payment Solutions":
                    result = 38;
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
                case "Asia":
                    result = 1;
                    break;
                case "Canada":
                    result = 2;
                    break;
                case "China":
                    result = 3;
                    break;
                case "EMEA":
                    result = 4;
                    break;
                case "Europe":
                    result = 5;
                    break;
                case "Hong Kong":
                    result = 6;
                    break;
                case "International":
                    result = 7;
                    break;
                case "Singapore":
                    result = 8;
                    break;
                case "US":
                    result = 9;
                    break;
                default:
                    result = 0;
                    break;
            }

            return result;
        }

        static int GetCorpId( string corpText )
        {
            int result = 0;

            switch ( corpText )
            {
                case "Accounting & Financial Management":                    
                case "Accounting and Financial Management":
                    result = 1;
                    break;
                case "Accounting & Financial Management Risk":
                    result = 2;
                    break;
                case "AML":
                    result = 3;
                    break;
                case "Business Continuity":
                    result = 4;
                    break;
                case "Business Continuity Risk":
                    result = 5;
                    break;
                case "Credit Risk":
                    result = 6;
                    break;
                case "Credit Risk - Business Banking":
                    result = 7;
                    break;
                case "Credit Risk - Commercial":
                    result = 8;
                    break;
                case "Credit Risk - Retail":
                    result = 9;
                    break;
                case "Credit Risk & CCR":
                    result = 10;
                    break;
                case "Fraud/Criminal":
                    result = 11;
                    break;
                case "Human Resources":
                    result = 12;
                    break;
                case "Information Security & Information Management":
                    result = 13;
                    break;
                case "Legal":
                    result = 14;
                    break;
                case "Liquidity and Funding":
                    result = 15;
                    break;
                case "Liquidity and Funding Risk":
                    result = 16;
                    break;
                case "Market Risk (Trading, Structural)":
                    result = 17;
                    break;
                case "Market Risk (Trading, Structural, CCR)":
                    result = 18;
                    break;
                case "Model Risk":
                    result = 19;
                    break;
                case "Outsourceing & Supplier":
                    result = 20;
                    break;
                case "Physical Security":
                    result = 21;
                    break;
                case "Privacy":
                    result = 22;
                    break;
                case "Process & Operation Risk":
                    result = 23;
                    break;
                case "Project Management":
                    result = 24;
                    break;
                case "Property":
                    result = 25;
                    break;
                case "Regulatory Compliance":
                    result = 26;
                    break;
                case "Reputation Risk":
                    result = 27;
                    break;
                case "Tax":
                    result = 28;
                    break;
                default:
                    result = 0;
                    break;
            }

            return result;
        }

        #endregion

    }

    
}
