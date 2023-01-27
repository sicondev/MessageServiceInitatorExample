using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MessageServiceExample
{
    /// <summary>
    /// Main program
    /// </summary>
    internal class Program
    {
        private static readonly ConsoleColor DEFAULT_FORECOLOR;
        private static readonly Sage.Common.Messaging.CrossCutMessageSource SiconOrderCreatedMessageSource = new Sage.Common.Messaging.CrossCutMessageSource("SiconSalesOrder", "Created", Sage.Common.Messaging.ProcessPoint.PostMethod);
       
        /// <summary>
        /// Static Constructor
        /// </summary>
        static Program()
        {
            try
            {
                DEFAULT_FORECOLOR = Console.ForegroundColor;
            }
            catch (Exception e)
            {
                LogError(e);
            }
        }

        /// <summary>
        /// Main entry point
        /// </summary>
        /// <param name="args"></param>
        static void Main(string[] args)
        {
            try
            {
                Console.Clear();

                LogGeneral("Resolving Sage Assemblies...");

                //Resolve all sage assemblies
                AssemblyResolver.ResolveAssemblies();

                LogGeneral("Sage Assemblies Resolved.");

                string companyName = ConfigurationManager.AppSettings["SageCompanyName"];

                //Syncronises any addon dlls from the server to local app data for the current user account (so addon dlls are in the same directory,
                //if your app is outside of local app data you may need to pull the assemblies in from the sage directory in IIS)
                //This also ensure any X-Table extensions are loaded
                bool syncroniseClientFiles = true;

                //Create client initators ensures initiators in any addons are set up, including Sage Manufacturing
                bool createClientInitiators = true;

                LogInfo($"Attempting to connect to company '{companyName}'.");

                //Connect to Sage 200 using the company name, and ensuring Syncronise client files and create client initiators is true
                using (SageConnection sageConnection = new SageConnection(companyName, syncroniseClientFiles, createClientInitiators))
                {
                    LogSuccess($"Connected to '{companyName}'.");

                    Program.CreateSalesOrder();

                    Program.CreateSalesOrderAdvanced();

                    LogGeneral($"Disconnecting from '{companyName}'.");
                }

                LogGeneral($"Disconnected from '{companyName}'.");

                LogGeneral($"Shutting Down.");

                LogInfo("Press any key to exit.");

                Console.ReadKey();
            }
            catch (Exception e)
            {
                LogError(e);
            }
        }

        /// <summary>
        /// Creates a sales order
        /// </summary>
        private static void CreateSalesOrder()
        {
            try
            {
                LogGeneral("Creating Sales Order...");

                using (Sage.Accounting.SOP.SOPOrder sopOrder = new Sage.Accounting.SOP.SOPOrder())
                {
                    sopOrder.Customer = Sage.Accounting.SalesLedger.CustomerFactory.Factory.Fetch("Abb001");

                    sopOrder.Update();

                    sopOrder.Post(true, true);

                    //Notify the Cross cut message source
                    Sage.Common.Messaging.MessageService.GetInstance()?.Notify(SiconOrderCreatedMessageSource, sopOrder, new Sage.Common.Messaging.MessageArgs());

                    LogSuccess($"Sales Order '{sopOrder.DocumentNo}' posted successfully.");
                }
            }
            catch (Exception)
            {
                throw;
            }
        }

        /// <summary>
        /// Creates a sales order using the advanced option
        /// </summary>
        private static void CreateSalesOrderAdvanced()
        {
            try
            {
                LogGeneral("Creating Sales Order (Advanced)...");

                using (Sage.Accounting.SOP.SOPOrder sopOrder = new Sage.Accounting.SOP.SOPOrder())
                {
                    sopOrder.Customer = Sage.Accounting.SalesLedger.CustomerFactory.Factory.Fetch("Abb001");

                    sopOrder.Update();

                    sopOrder.Post(true, true);

                    Hashtable hashtable = new Hashtable();

                    hashtable["SalesOrder"] = sopOrder;
                    hashtable["CourierService"] = "DHL";
                    hashtable["CourerServiceDescription"] = "Special Delivery Instructions";
                    hashtable["ProjectNumber"] = "J00000001";
                    hashtable["ProjectHeaderNumber"] = "Revenue";

                    //Notify the Cross cut message source
                    Sage.Common.Messaging.MessageService.GetInstance()?.Notify(SiconOrderCreatedMessageSource, hashtable, new Sage.Common.Messaging.MessageArgs());

                    LogSuccess($"Sales Order '{sopOrder.DocumentNo}' posted successfully.");
                }
            }
            catch (Exception)
            {
                throw;
            }
        }

        #region Logging

        /// <summary>
        /// Logs a Warning
        /// </summary>
        /// <param name="message">The Message</param>
        private static void LoadWarning(string message) => Log($"Warning: {message}", ConsoleColor.Yellow);

        /// <summary>
        /// Logs a Success Message
        /// </summary>
        /// <param name="message">The Message</param>
        private static void LogSuccess(string message) => Log($"Success: {message}", ConsoleColor.Green);

        /// <summary>
        /// Logs an Error Message
        /// </summary>
        /// <param name="e">The Exception</param>
        private static void LogError(Exception e) => Log($"Error: {e}", ConsoleColor.Red);

        /// <summary>
        /// Logs an Info Message
        /// </summary>
        /// <param name="message">The Message</param>
        private static void LogInfo(string message) => Log($"Info: {message}", ConsoleColor.Cyan);

        /// <summary>
        /// Logs an Info Message
        /// </summary>
        /// <param name="message">The Message</param>
        private static void LogGeneral(string message) => Log($"Info: {message}", null);

        /// <summary>
        /// Writes a message to the console
        /// </summary>
        /// <param name="message">The Message</param>
        /// <param name="color">The Color</param>
        private static void Log(string message, ConsoleColor? color = null)
        {
            Console.ForegroundColor = color ?? DEFAULT_FORECOLOR;
            Console.WriteLine($"{DateTime.Now}: {message}");
            Console.ForegroundColor = DEFAULT_FORECOLOR;
        }

        #endregion Logging
    }
}
