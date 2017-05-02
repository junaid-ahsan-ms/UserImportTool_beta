using System;
using System.IO;
using System.Data;
using System.Data.SqlClient;

using ImportUser.Properties;
using System.Diagnostics;

namespace ImportUser
{

    /// <summary>
    /// This is a console utility written for BMO IAAP project to help import Users
    /// for the IAAP web application and assign them to appropriate roles
    /// </summary>
    class Program
    {        

        static void Main( string[] args )
        {
            string sDebug = String.Empty;
            string fileTypeText = string.Empty;
            bool bError = false;

            ImportUserImpl userImporter = new ImportUserImpl();

            try
            {

                if ( args.Length == 2 && !string.IsNullOrEmpty( args[0] ) )
                {                    
                    bError = userImporter.Import( args );                    
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
                Console.WriteLine( "User Import process for file type {0} completed " + 
                    (bError ? "with errors" : "successfully") + "...", userImporter.UserFileType );
            }

            
            Console.WriteLine();
            Console.WriteLine( "Press any key to continue..." );
            Console.ReadKey();
        }

        static void DisplayUsage( )
        {
            Console.WriteLine( "Usage: Program.exe <fileName> <fileType: 0=CSA, 1=OpsRisk, 2=Governance>" );
            Console.WriteLine( "Note: User must have access to the Iaap database" );
        }   

    }    

}

