using Microsoft.Office.Interop.Outlook;
using SendSecureAddIn.Properties;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;


namespace SendSecureAddIn
{
    [ComVisible(true)]
    public class Ribbon1 : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public Ribbon1()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            string ribbonXML = null;

            if (ribbonID == "Microsoft.Outlook.Mail.Compose")
            {
                ribbonXML = GetResourceText("SendSecureAddIn.Ribbon1.xml");
            }

            return ribbonXML;
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit http://go.microsoft.com/fwlink/?LinkID=271226

        //callback for when the ribbon loads
        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        //callback for when the button is clicked
        public void OnButtonSendSecure(Office.IRibbonControl control)
        {
            //get the active inspector from the application
            Inspector inspector = Globals.ThisAddIn.Application.ActiveInspector();

            //save a variable to use for the mail item
            MailItem mailItem = null;

            //if the current item is a mailitem, set the mail item
            if (inspector.CurrentItem is MailItem)
            {
                mailItem = inspector.CurrentItem;
            }

            //if the mailitem is not null, add the send secure flag and then send
            if (mailItem != null)
            {
                mailItem.Subject = "[SEND SECURE] " + mailItem.Subject;
                mailItem.Send(); 
            }
        }

        //callback for getting image for the button
        public Bitmap ButtonSendSecure_GetImage(Office.IRibbonControl control)
        {
            return Resources.lock_icon;
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
