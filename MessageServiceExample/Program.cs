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
    internal static class Program
    {
        private static readonly ConsoleColor DEFAULT_FORECOLOR;
        private static readonly Sage.Common.Messaging.CrossCutMessageSource SiconOrderCreatedMessageSource = new Sage.Common.Messaging.CrossCutMessageSource("SiconSalesOrder", "Created", Sage.Common.Messaging.ProcessPoint.PostMethod);
        private static readonly Sage.Common.Messaging.CrossCutMessageSource SiconOrderLineCreatedMessageSource = new Sage.Common.Messaging.CrossCutMessageSource("SiconSalesOrderLine", "Created", Sage.Common.Messaging.ProcessPoint.PostMethod);

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

                //Synchronises any addon dlls from the server to local app data for the current user account (so addon dlls are in the same directory,
                //if your app is outside of local app data you may need to pull the assemblies in from the sage directory in IIS)
                //This also ensure any X-Table extensions are loaded
                const bool synchroniseClientFiles = true;

                //Create client initiators ensures initiators in any addons are set up, including Sage Manufacturing
                const bool createClientInitiators = true;

                LogInfo($"Attempting to connect to company '{companyName}'.");

                //Connect to Sage 200 using the company name, and ensuring Synchronise client files and create client initiators is true
                using (SageConnection sageConnection = new SageConnection(companyName, synchroniseClientFiles, createClientInitiators))
                {
                    LogSuccess($"Connected to '{companyName}'.");

                    Program.CreateSalesOrder("ABB001", "ACS/BLENDER");

                    Program.CreateSalesOrderAdvanced("ABB001", "ACS/BLENDER", "DPD NEXT DAY", "Special Instructions");
                    Program.CreateSalesOrderAdvanced("ABB001", "ACS/BLENDER", "", "");

                    LogGeneral($"Disconnecting from '{companyName}'.");
                }

                LogGeneral($"Disconnected from '{companyName}'.");

                LogGeneral($"Shutting Down.");

                LogInfo("Press any key to exit.");


            }
            catch (Exception e)
            {
                LogError(e);
            }
            finally
            {
                Console.ReadKey();
            }
        }

        /// <summary>
        /// Creates a sales order
        /// </summary>
        private static void CreateSalesOrder(string customerRef, string itemCode)
        {
            try
            {
                LogGeneral("Creating Sales Order...");

                using (Sage.Accounting.SOP.SOPOrder sopOrder = new Sage.Accounting.SOP.SOPOrder())
                {
                    sopOrder.Customer = Sage.Accounting.SalesLedger.CustomerFactory.Factory.Fetch(customerRef);

                    sopOrder.Update();

                    AddNonTraceableStandardItemLine(sopOrder, sopOrder.SOPLedger, itemCode);

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
        private static void CreateSalesOrderAdvanced(string customerRef, string itemCode, string courierService, string deliveryInstructions)
        {
            try
            {
                LogGeneral("Creating Sales Order (Advanced)...");

                using (Sage.Accounting.SOP.SOPOrder sopOrder = new Sage.Accounting.SOP.SOPOrder())
                {
                    sopOrder.Customer = Sage.Accounting.SalesLedger.CustomerFactory.Factory.Fetch(customerRef);

                    sopOrder.Update();

                    AddNonTraceableStandardItemLine(sopOrder, sopOrder.SOPLedger, itemCode);

                    sopOrder.Post(true, true);

                    Hashtable hashtable = new Hashtable
                    {
                        ["SalesOrder"] = sopOrder,
                        ["SiconCourierDelServiceDesc"] = courierService,
                        ["DeliveryInstructions"] = deliveryInstructions,
                        ["ReadyToPick"] = bool.TrueString
                    };

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

        /// <summary>
        /// Adds a Non Traceable Standard Item Line
        /// </summary>
        /// <param name="sopOrder">The SOP Order</param>
        /// <param name="sopLedger">The SOP Ledger</param>
        /// <param name="itemCode">The Item Code</param>
        private static void AddNonTraceableStandardItemLine(Sage.Accounting.SOP.SOPOrderReturn sopOrder, Sage.Accounting.SOP.SOPLedger sopLedger, string itemCode)
        {
            try
            {
                using (Sage.Accounting.Stock.StockItems stockItems = new Sage.Accounting.Stock.StockItems())
                {
                    // Find the non-traceable item "TILE/WHT/20X20"
                    Sage.Accounting.Stock.StockItem stockItem = stockItems[itemCode];

                    if (stockItem != null)
                    {
                        // Use the WarehouseStockItemViews to get the WarehouseItem that links the
                        // stock item "TILE/WHT/20X20" and warehouse "WAREHOUSE"
                        Sage.Accounting.Stock.Views.WarehouseStockItemViews warehouseStockItemViews = new Sage.Accounting.Stock.Views.WarehouseStockItemViews();

                        Sage.Accounting.Stock.Views.WarehouseStockItemView warehouseStockItemView = null;

                        // Add the stock item filter
                        warehouseStockItemViews.Query.Filters.Add(new Sage.ObjectStore.Filter(Sage.Accounting.Stock.Views.WarehouseStockItemView.FIELD_ITEM, stockItem.Item));
                        warehouseStockItemViews.Query.Filters.Add(new Sage.ObjectStore.Filter(Sage.Accounting.Stock.Views.WarehouseStockItemView.FIELD_WAREHOUSEID, sopOrder.WarehouseDbKey));

                        warehouseStockItemViews.Query.Find();

                        // Find the first warehouse bin that has available stock
                        foreach (Sage.Accounting.Stock.Views.WarehouseStockItemView view in warehouseStockItemViews)
                        {
                            if (view.FreeStockAvailable > System.Decimal.Zero)
                            {
                                warehouseStockItemView = view;

                                break;
                            }
                        }

                        if (warehouseStockItemView != null)
                        {
                            // Instantiate the line type
                            Sage.Accounting.SOP.SOPStandardItemLine sopStandardItemLine =
                                new Sage.Accounting.SOP.SOPStandardItemLine();

                            // Set the order
                            sopStandardItemLine.SOPOrderReturn = sopOrder;

                            try
                            {
                                // Set the stock item
                                sopStandardItemLine.Item = stockItem;

                                // Set the warehouse item
                                sopStandardItemLine.WarehouseItem = warehouseStockItemView.WarehouseItem;

                                // Set the fulfilment method
                                if (sopOrder.SOPUserPermission.OverrideFulfilmentMethod)
                                {
                                    // Optional: this can either be set to From Stock, From Supplier Via Stock
                                    // or From Supplier Direct to Customer
                                    sopStandardItemLine.FulfilmentMethod =
                                        Sage.Accounting.Stock.SOPOrderFulfilmentMethodEnum.EnumFulfilmentFromStock;
                                }

                                // If the SOP ledger is configured to allocate on order entry setting
                                // the LineQuantity will also set the ToAllocateQuantity property
                                sopStandardItemLine.LineQuantity = 1M;

                                // Make sure user has permission to set price and discount
                                if (sopOrder.SOPUserPermission.OverridePricesDiscounts)
                                {
                                    // Add the allowable warning for Ex20319Exception so that the
                                    // UnitSellingPrice can be set without requesting confirmation
                                    Sage.Accounting.Application.AllowableWarnings.Add(sopStandardItemLine,
                                        typeof(Sage.Accounting.Exceptions.Ex20319Exception));

                                    // Set the new price
                                    sopStandardItemLine.UnitSellingPrice = 199.99M;

                                    // Add the allowable warning for Ex20320Exception so that the
                                    // UnitDiscountPercent can be set without requesting confirmation                    
                                    Sage.Accounting.Application.AllowableWarnings.Add(sopStandardItemLine,
                                        typeof(Sage.Accounting.Exceptions.Ex20320Exception));

                                    // Set discount
                                    sopStandardItemLine.UnitDiscountPercent = 5M;
                                }

                                // Add the line to the order
                                sopOrder.Lines.Add(sopStandardItemLine);

                                // Save the line
                                // This also sets the document number to next auto generated order number
                                sopStandardItemLine.Post(sopLedger.Setting.RecordCancelledOrdLines);

                                //Notify Message Source
                                Sage.Common.Messaging.MessageService.GetInstance().Notify(SiconOrderLineCreatedMessageSource, sopStandardItemLine, new Sage.Common.Messaging.MessageArgs());
                            }
                            finally
                            {
                                // Remove any allowable warnings
                                if (Sage.Accounting.Application.AllowableWarnings.IsAllowed(sopStandardItemLine,
                                    typeof(Sage.Accounting.Exceptions.Ex20319Exception)))
                                {
                                    Sage.Accounting.Application.AllowableWarnings.Delete(sopStandardItemLine,
                                        typeof(Sage.Accounting.Exceptions.Ex20319Exception));
                                }

                                if (Sage.Accounting.Application.AllowableWarnings.IsAllowed(sopStandardItemLine,
                                    typeof(Sage.Accounting.Exceptions.Ex20320Exception)))
                                {
                                    Sage.Accounting.Application.AllowableWarnings.Delete(sopStandardItemLine,
                                        typeof(Sage.Accounting.Exceptions.Ex20320Exception));
                                }
                            }
                        }
                    }
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
