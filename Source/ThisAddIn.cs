using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;
using System.Drawing;
using CalScanner;

namespace MOMOutlookAddIn
{
    public partial class ThisAddIn
    {
        public static Microsoft.Office.Interop.Outlook.Application app = null;
        public static Dictionary<String, String> emailNameMappingDict = null;
        public static Microsoft.Office.Interop.Outlook._NameSpace ns = null;
        public static Microsoft.Office.Interop.Outlook.MAPIFolder calendar = null;
        Microsoft.Office.Interop.Outlook.MAPIFolder inboxFolder = null;
        Microsoft.Office.Interop.Outlook.MAPIFolder sentFolder = null;
        Microsoft.Office.Interop.Outlook.MAPIFolder taskFolder = null;
        Microsoft.Office.Interop.Outlook.MAPIFolder todoFolder = null;
        Microsoft.Office.Interop.Outlook.MAPIFolder contactFolder = null;

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon1();
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            populateAddressListFromOutlook();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }
            
        protected void populateAddressListFromOutlook()
        {

            DateTime startDate, endDate;
            endDate = DateTime.Today;
            startDate = endDate.AddDays(-30);

           // app = new Microsoft.Office.Interop.Outlook.Application(); - this line will create problem with Outlook 2013 sometimes
            app = this.Application;
            //}
            ns = app.GetNamespace("MAPI");


            calendar = ns.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderCalendar);

            String StringToCheck = "";
            StringToCheck = "[Start] >= " + "\"" + startDate.ToString().Substring(0, startDate.ToString().IndexOf(" ")) + "\""
                + " AND [End] <= \"" + endDate.ToString().Substring(0, endDate.ToString().IndexOf(" ")) + "\"";


            Microsoft.Office.Interop.Outlook.Items oItems = (Microsoft.Office.Interop.Outlook.Items)calendar.Items;
            Microsoft.Office.Interop.Outlook.Items restricted;

            oItems.Sort("[Start]", false);
            oItems.IncludeRecurrences = true;

            restricted = oItems.Restrict(StringToCheck);
            restricted.Sort("[Start]", false);

            restricted.IncludeRecurrences = true;
            Microsoft.Office.Interop.Outlook.AppointmentItem oAppt = (Microsoft.Office.Interop.Outlook.AppointmentItem)restricted.GetFirst();

            Dictionary<String, String> comboDict = new Dictionary<string, string>();
            comboDict.Add("ALL", "ALL");

            //Loop through each appointment item to find out the unique recipient list and add to the combo box
            while (oAppt != null)
            {
                oAppt = (Microsoft.Office.Interop.Outlook.AppointmentItem)restricted.GetNext();

                if (oAppt != null)
                    foreach (Microsoft.Office.Interop.Outlook.Recipient rcp in oAppt.Recipients)
                    {
                        //Display the email id in bracket along with the name in case the name rcp.Name does not contain the email address
                        //This will help in situtation where the the same recipient with multiple emails address need to be distinguised
                        String recpDisplayString = rcp.Name.IndexOf("@") < 0 ? rcp.Name + "(" + rcp.Address + ")" : rcp.Name;

                        if (!comboDict.ContainsKey(recpDisplayString))
                        {
                            comboDict.Add(recpDisplayString, recpDisplayString);
                            if (recpDisplayString != null && !emailNameMappingDict.ContainsKey(recpDisplayString))
                                emailNameMappingDict.Add(recpDisplayString, rcp.Address != null ? rcp.Address : rcp.Name);
                            if (rcp.Name != null && !MOM_Form.autoCompleteList.Contains(recpDisplayString))
                                MOM_Form.autoCompleteList.Add(recpDisplayString);
                        }
                    }
            }

            //listBox_AddrList.DataSource = new BindingSource(comboDict, null);
            //listBox_AddrList.DisplayMember = "Value";
            //listBox_AddrList.ValueMember = "Key";
            //listBox_AddrList.SelectedValue = "ALL";

            populateAutoCompleteList(startDate, endDate);
        }

        protected void populateAutoCompleteList(DateTime startDate, DateTime endDate)
        {
            inboxFolder = ns.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderInbox);
            sentFolder = ns.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderSentMail);
            String StringToCheck = "";
            StringToCheck = "[Received] >= " + "\"" + startDate.ToString().Substring(0, startDate.ToString().IndexOf(" ")) + "\""
                + " AND [Received] <= \"" + endDate.ToString().Substring(0, endDate.ToString().IndexOf(" ")) + "\"";


            Microsoft.Office.Interop.Outlook.Items oItemsInbox = (Microsoft.Office.Interop.Outlook.Items)inboxFolder.Items;
            Microsoft.Office.Interop.Outlook.Items oItemsSentBox = (Microsoft.Office.Interop.Outlook.Items)sentFolder.Items;

            Microsoft.Office.Interop.Outlook.Items restricted;
            Microsoft.Office.Interop.Outlook.MailItem mailObject = null;
            //First scan the inbox items for unique contacts
            oItemsInbox.Sort("[Received]", false);
            oItemsInbox.IncludeRecurrences = true;
            restricted = oItemsInbox.Restrict(StringToCheck);
            restricted.Sort("[Received]", false);

            restricted.IncludeRecurrences = true;

             Object emailItem=null;
            if (restricted.GetFirst() is Microsoft.Office.Interop.Outlook.MailItem)
                emailItem = (Microsoft.Office.Interop.Outlook.MailItem)restricted.GetFirst();

            while (emailItem != null)
            {
                if (emailItem is Microsoft.Office.Interop.Outlook.MailItem)
                    mailObject = (Microsoft.Office.Interop.Outlook.MailItem)emailItem;

                if (!MOM_Form.autoCompleteList.Contains(mailObject.SenderEmailAddress.Trim()))
                    MOM_Form.autoCompleteList.Add(mailObject.SenderEmailAddress.Trim());

                emailItem = restricted.GetNext();
            }

            StringToCheck = "[SentOn] >= " + "\"" + startDate.ToString().Substring(0, startDate.ToString().IndexOf(" ")) + "\""
                + " AND [SentOn] <= \"" + endDate.ToString().Substring(0, endDate.ToString().IndexOf(" ")) + "\"";

            oItemsSentBox.Sort("[SentOn]", false);
            oItemsSentBox.IncludeRecurrences = true;
            restricted = oItemsSentBox.Restrict(StringToCheck);
            restricted.Sort("[SentOn]", false);

            restricted.IncludeRecurrences = true;
            if (restricted.GetFirst() is MailItem)
                emailItem = (Microsoft.Office.Interop.Outlook.MailItem)restricted.GetFirst();

            while (emailItem != null)
            {
                if (emailItem is Microsoft.Office.Interop.Outlook.MailItem)
                    mailObject = (Microsoft.Office.Interop.Outlook.MailItem)emailItem;

                if (!MOM_Form.autoCompleteList.Contains(mailObject.SenderEmailAddress.Trim()))
                    MOM_Form.autoCompleteList.Add(mailObject.SenderEmailAddress.Trim());

                emailItem = restricted.GetNext();
            }

        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
