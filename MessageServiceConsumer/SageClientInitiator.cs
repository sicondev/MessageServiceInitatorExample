using System;
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
    }
}
