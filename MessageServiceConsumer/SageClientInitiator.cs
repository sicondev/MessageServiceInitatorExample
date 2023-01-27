using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace MessageServiceConsumer
{
    /// <summary>
    /// Initiator for the Client
    /// </summary>
    public class ClientInitiator
    {
        //Delegate for the SOPOrderSaved Handler
        private Sage.Common.Messaging.MessageHandler SOPOrderSavedHander = new Sage.Common.Messaging.MessageHandler(SOPOrderSaved);

        //Delegate for Sicon Order Created
        private Sage.Common.Messaging.MessageHandler SiconOrderCreatedHander = new Sage.Common.Messaging.MessageHandler(SiconOrderCreated);

        //Delegate for Kititng Sicon Order Created
        private Sage.Common.Messaging.MessageHandler KittingSiconOrderCreatedHander = new Sage.Common.Messaging.MessageHandler(KittingSiconOrderCreated);

        //Message Source for Sicon Order Created
        private static readonly Sage.Common.Messaging.CrossCutMessageSource SiconOrderCreatedMessageSource = new Sage.Common.Messaging.CrossCutMessageSource("SiconSalesOrder", "Created", Sage.Common.Messaging.ProcessPoint.PostMethod);

        //Message Source for Kitting Order Created
        private static readonly Sage.Common.Messaging.CrossCutMessageSource KittingOrderCreatedMessageSource = new Sage.Common.Messaging.CrossCutMessageSource("KittingSiconSalesOrder", "Created", Sage.Common.Messaging.ProcessPoint.PostMethod);

        /// <summary>
        /// Constructor
        /// </summary>
        public ClientInitiator()
        {
            try
            {
                this.Subscribe();

                Application.ApplicationExit += (s, e) => this.UnSubscribe();
            }
            catch (Exception)
            {
                throw;
            }
        }

        /// <summary>
        /// Subscribes to Messages
        /// </summary>
        private void Subscribe()
        {
            try
            {
                Sage.Common.Messaging.MessageService messageService = Sage.Common.Messaging.MessageService.GetInstance();

                messageService.Subscribe(Sage.Accounting.SOP.SOPLedgerMessageSource.SOPOrderSaved, SOPOrderSavedHander);
                messageService.Subscribe(SiconOrderCreatedMessageSource, SiconOrderCreatedHander);
                messageService.Subscribe(KittingOrderCreatedMessageSource, KittingSiconOrderCreatedHander);

                messageService = null;
            }
            catch (Exception)
            {
                throw;
            }
        }

        /// <summary>
        /// Unsubscribes to messages
        /// </summary>
        private void UnSubscribe()
        {
            try
            {
                Sage.Common.Messaging.MessageService messageService = Sage.Common.Messaging.MessageService.GetInstance();

                messageService.Unsubscribe(Sage.Accounting.SOP.SOPLedgerMessageSource.SOPOrderSaved, SOPOrderSavedHander);
                messageService.Unsubscribe(SiconOrderCreatedMessageSource, SiconOrderCreatedHander);
                messageService.Unsubscribe(KittingOrderCreatedMessageSource, KittingSiconOrderCreatedHander);

                messageService = null;
                SOPOrderSavedHander = null;
            }
            catch (Exception)
            {

                throw;
            }
        }

        /// <summary>
        /// Handles the SOPOrderSaved Message
        /// </summary>
        /// <param name="sender">The sender</param>
        /// <param name="args">The Message Args</param>
        /// <returns><Message Response/returns>
        private static Sage.Common.Messaging.Response SOPOrderSaved(object sender, Sage.Common.Messaging.MessageArgs args)
        {
            if (sender is Sage.Accounting.SOP.SOPOrder order)
            {
                Console.ForegroundColor = ConsoleColor.Magenta;
                Console.WriteLine($"Message Service Consumer: SOP Order '{order.DocumentNo}' saved.");
            }

            return new Sage.Common.Messaging.Response(new Sage.Common.Messaging.ResponseArgs());
        }

        /// <summary>
        /// Handles the SiconOrderCreated Message
        /// </summary>
        /// <param name="sender">The sender</param>
        /// <param name="args">The Message Args</param>
        /// <returns><Message Response/returns>
        private static Sage.Common.Messaging.Response SiconOrderCreated(object sender, Sage.Common.Messaging.MessageArgs args)
        {
            if (sender is Sage.Accounting.SOP.SOPOrder order)
            {
                Console.ForegroundColor = ConsoleColor.Blue;
                Console.WriteLine($"Message Service Consumer: Sicon Sales Order '{order.DocumentNo}' Created.");
            }
            else if (sender is Hashtable hashtable)
            {
                if (hashtable["SalesOrder"] is Sage.Accounting.SOP.SOPOrder salesOrder)
                {
                    Console.ForegroundColor = ConsoleColor.Blue;
                    Console.WriteLine($"Message Service Consumer: Sicon Sales Order (Advanced) '{salesOrder.DocumentNo}' Created.");
                    Console.WriteLine($"Message Service Consumer: Courier Service: '{hashtable["CourierService"]}' Created.");
                    Console.WriteLine($"Message Service Consumer: Delivery Instructions: '{hashtable["CourerServiceDescription"]}' Created.");
                    Console.WriteLine($"Message Service Consumer: Project Nunber: '{hashtable["ProjectNumber"]}' Created.");
                    Console.WriteLine($"Message Service Consumer: Project Header: '{hashtable["ProjectHeaderNumber"]}' Created.");
                }

                //Alternate way to get the sales order out of the hashtable
                Sage.Accounting.SOP.SOPOrder sopOrder = hashtable.OfType<Sage.Accounting.SOP.SOPOrder>().FirstOrDefault();
            }

            return new Sage.Common.Messaging.Response(new Sage.Common.Messaging.ResponseArgs());
        }

        /// <summary>
        /// Handles the KittingSOPOrderSaved Message
        /// </summary>
        /// <param name="sender">The sender</param>
        /// <param name="args">The Message Args</param>
        /// <returns><Message Response/returns>
        private static Sage.Common.Messaging.Response KittingSiconOrderCreated(object sender, Sage.Common.Messaging.MessageArgs args)
        {
            if (sender is Sage.Accounting.SOP.SOPOrder order)
            {
                Console.ForegroundColor = ConsoleColor.DarkBlue;
                Console.WriteLine($"Message Service Consumer: Kitting - Sicon Sales Order '{order.DocumentNo}' Created.");
            }
            else if (sender is Hashtable hashtable)
            {
                if (hashtable["SalesOrder"] is Sage.Accounting.SOP.SOPOrder salesOrder)
                {
                    Console.ForegroundColor = ConsoleColor.Blue;
                    Console.WriteLine($"Message Service Consumer: Sicon Sales Order (Advanced) '{salesOrder.DocumentNo}' Created.");
                }

                //Alternate way to get the sales order out of the hashtable
                Sage.Accounting.SOP.SOPOrder sopOrder = hashtable.OfType<Sage.Accounting.SOP.SOPOrder>().FirstOrDefault();
            }

            return new Sage.Common.Messaging.Response(new Sage.Common.Messaging.ResponseArgs());
        }
    }
}
