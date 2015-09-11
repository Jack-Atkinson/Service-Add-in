using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Windows.Forms;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new Ribbon();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace Service_Add_in
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;
        ThisAddIn addIn;
        Outlook.AppointmentItem myapptitem_;
        Outlook.Explorer currentExplorer;

        public Ribbon()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("Service_Add_in.Ribbon.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit http://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            currentExplorer = Globals.ThisAddIn.Application.ActiveExplorer();
            currentExplorer.SelectionChange += new Outlook
                .ExplorerEvents_10_SelectionChangeEventHandler
                (CurrentExplorer_Event);
            addIn = Globals.ThisAddIn;
            this.ribbon = ribbonUI;
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

        public void punchin_Click(Office.IRibbonControl control)
        {
            AddDateTextToBody("IN");
        }

        public void punchout_Click(Office.IRibbonControl control)
        {
            AddDateTextToBody("OUT");
        }

        public void AddDateTextToBody(string status)
        {
            if (myapptitem_ != null)
            {
                DateTime currenttime = DateTime.Now;
                DateTime lastcheckin = new DateTime();
                string punchstatus = Environment.NewLine + DateTime.Now.ToString() + " - " + status;

                switch (status)
                {
                    case "IN":
                        if (LastPunchSame("IN", ref lastcheckin))
                        {
                            return;
                        }
                        break;
                    case "OUT":
                        if (LastPunchSame("OUT", ref lastcheckin))
                        {
                            return;
                        }
                        else
                        {
                            double minutes = currenttime.Subtract(lastcheckin).Minutes;
                            punchstatus += " Total: " + minutes;
                        }
                        break;
                    default:
                        break;
                }

                myapptitem_.Body += punchstatus;

                myapptitem_.Save();
                MessageBox.Show("You punched " + status + " at " + DateTime.Now.ToString(),
                                "Alert",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Information);
            }
        }

        public bool LastPunchSame(string status, ref DateTime lastcheckin)
        {
            string body = myapptitem_.Body;
            if (body == null)
                return false;

            string[] lines = body.Replace("\r", "").Split('\n');
            if (lines.Length != 0 && lines[lines.Length - 1].Length > 20)
            {
                string lastpunch = lines[lines.Length - 1];
                lastcheckin = DateTime.Parse(lastpunch.Substring(0, 20));
                if (lastpunch.Contains(status))
                {
                    MessageBox.Show("You need to check " + (status == "IN" ? "out" : "in") +
                                    " before you can check " + (status == "IN" ? "in" : "out") + " again",
                                    "Alert",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Information);
                    return true;
                }
            }
            return false;
        }

        private void CurrentExplorer_Event()
        {
            Outlook.MAPIFolder selectedFolder =
                Globals.ThisAddIn.Application.ActiveExplorer().CurrentFolder;
            if (selectedFolder.Name == "Calendar")
            {
                if (Globals.ThisAddIn.Application.ActiveExplorer().Selection.Count > 0)
                {
                    Object selObject = Globals.ThisAddIn.Application.ActiveExplorer().Selection[1];
                    if (selObject is Outlook.AppointmentItem)
                    {
                        Outlook.AppointmentItem apptItem =
                            (selObject as Outlook.AppointmentItem);
                        myapptitem_ = apptItem;
                    }
                }
            }
        }

    }
}
