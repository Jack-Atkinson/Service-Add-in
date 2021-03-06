﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Windows.Forms;
using System.Text.RegularExpressions;

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
        string regexpattern =
            "(\\d{1,2}\\/\\d{1,2}\\/\\d{4}\\s" +
            "\\d{1,2}:\\d{1,2}:\\d{1,2}\\s[AP]M)\\s" +
            "[-|–]\\s(IN|OUT)(\\sSession\\stime:\\s\\d+\\sMinutes)?";

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

        public void getTotal_Click(Office.IRibbonControl control)
        {
            string[] times = getPunchTimes();
            if (times[0].Equals(null))
            {
                MessageBox.Show("No times in calendar item",
                                "Alert",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Exclamation);
                return;
            }

            if (times[times.Length - 1].Contains("IN"))
            {
                MessageBox.Show("You need to punch out before you can get the total time!",
                                "Alert",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Exclamation);
                return;
            }

            DialogResult result =
                MessageBox.Show("Are you sure you want to calculate total time?",
                                "Question",
                                MessageBoxButtons.YesNo,
                                MessageBoxIcon.Question);
            if (result != DialogResult.Yes)
                return;
            
            MatchCollection matches;
            string punchtype;
            DateTime intime = new DateTime();
            DateTime outtime = new DateTime();
            List<double> differences = new List<double>();

            foreach (string time in times.ToArray())
            {
                matches = Regex.Matches(time, regexpattern);
                punchtype = matches[0].Groups[2].Value;
                switch(punchtype) //this snippet looks a little...wrong, must test possible cases that could cause errors
                {
                    case "IN":
                        intime = DateTime.Parse(matches[0].Groups[1].Value);
                        break;
                    case "OUT":
                        outtime = DateTime.Parse(matches[0].Groups[1].Value);
                        differences.Add(GetHoursWorked(intime, outtime));
                        break;
                    default:
                        break;
                }
            }

            double totalTime = differences.Sum();
            string message = string.Format("\n\nTotal: {0:0.00} Hour(s)\n\n", totalTime);
            myapptitem_.Body += message;
            myapptitem_.Save();

            MessageBox.Show(message,
                            "Information",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
        }

        public void AddDateTextToBody(string status)
        {
            if (!myapptitem_.Equals(null))
            {
                DateTime currenttime = DateTime.Now;
                DateTime lastcheckin = new DateTime();

                if(!ValidPunch(status, ref lastcheckin))
                {
                    MessageBox.Show("You need to check " + (status == "IN" ? "out" : "in") +
                                    " before you can check " + (status == "IN" ? "in" : "out") + " again",
                                    "Alert",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Exclamation);
                    return;
                }

                string punchstatus = Environment.NewLine + DateTime.Now.ToString() + " - " + status;

                if (status.Equals("OUT"))
                {
                    double totalTime = GetHoursWorked(lastcheckin, currenttime);
                    punchstatus += string.Format(" Session time: {0:0.00} Hour(s)", totalTime);
                }


                myapptitem_.Body += punchstatus;
                
                myapptitem_.Save();
                MessageBox.Show("You punched " + status + " at " + DateTime.Now.ToString(),
                                "Information",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Information);
            }
        }

        public bool ValidPunch(string status, ref DateTime lastcheckin)
        {
            string[] times = getPunchTimes();
            string lastpunch = "";
            string lastline = times[times.Length - 1];

            if (!Equals(lastline, null) && Regex.IsMatch(lastline, regexpattern))
            {
                MatchCollection matches = Regex.Matches(lastline, regexpattern);
                lastpunch = matches[0].Groups[2].Value;
                lastcheckin = DateTime.Parse(matches[0].Groups[1].Value);
            }

            if (status.Equals("OUT"))
            {
                if (lastpunch.Equals("") || lastpunch.Equals(status))
                    return false;
            }
            else if (status.Equals("IN"))
            {
                if (lastpunch == status)
                    return false;
            }
            else
                return false;
            return true;
        }

        public string[] getPunchTimes()
        {
            string[] times = new string[1];

            if (Equals(myapptitem_, null))
                return times;

            string apptbody = myapptitem_.Body;

            if (Equals(apptbody, null))
                return times;

            times = apptbody.Replace("\r", "").Split('\n');

            times = times.Where(time => Regex.IsMatch(time, regexpattern)).ToArray();

            if (times.Length == 0)
                times = new string[1];

            return times;
        }

        private double GetHoursWorked(DateTime startTime, DateTime endTime)
        {
            startTime = startTime.AddSeconds(-startTime.Second); //zero out seconds
            endTime = endTime.AddSeconds(-endTime.Second);
            startTime = startTime.AddMilliseconds(-startTime.Millisecond); //zero out milliseconds
            endTime = endTime.AddMilliseconds(-endTime.Millisecond);

            DateTime startMins = RoundTime(startTime);
            DateTime endMins = RoundTime(endTime);

            double minutesWorked = endMins.Subtract(startMins).TotalMinutes;

            return Convert.ToDouble(TimeSpan.FromMinutes(minutesWorked).TotalHours);
        }

        private DateTime RoundTime(DateTime time)
        {
            int threshold = 15 - (time.Minute % 15);
            DateTime minsWorked;
            if (threshold > 7)
                return minsWorked = time.Add(TimeSpan.FromMinutes(threshold - 15));
            else
                return minsWorked = time.Add(TimeSpan.FromMinutes(threshold));
        }


        private DateTime RoundUp(DateTime dt, TimeSpan d)
        {
            return new DateTime(((dt.Ticks + d.Ticks - 1) / d.Ticks) * d.Ticks);
        }

        private void CurrentExplorer_Event()
        {
            Outlook.Application myapp = Globals.ThisAddIn.Application;

            Outlook.Explorer localcalendar =
                myapp.ActiveExplorer(); //For local calender appointments.
            
            Outlook.Explorer sharedcalendar =
                myapp.Session.GetSharedDefaultFolder(
                    myapp.Session.CreateRecipient(myapp.Session.CurrentUser.Address), //our own shared folder
                    Outlook.OlDefaultFolders.olFolderCalendar
                    ).Application.ActiveExplorer(); 

            if (localcalendar.CurrentFolder.Name == "Calendar" ||
                sharedcalendar.CurrentFolder.Name != "")
            {
                if (localcalendar.Selection.Count > 0 &&
                    localcalendar.Selection[1].Subject != "")
                {
                    object localobject = Globals.ThisAddIn.Application.ActiveExplorer().Selection[1];
                    if (localobject is Outlook.AppointmentItem)
                    {
                        Outlook.AppointmentItem localapptitem =
                            (localobject as Outlook.AppointmentItem);
                        myapptitem_ = localapptitem;
                    }
                }
                else if (sharedcalendar.Selection.Count > 0 &&
                    sharedcalendar.Selection[1].Subject != "")
                {
                    object sharedobject = sharedcalendar.Selection[1];
                    if (sharedobject is Outlook.AppointmentItem)
                    {
                        Outlook.AppointmentItem sharedapptitem =
                            (sharedobject as Outlook.AppointmentItem);
                        myapptitem_ = sharedapptitem;
                    }
                }
                else
                {
                    myapptitem_ = null;
                }
            }
        }

    }
}
