using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using System.ComponentModel;
using System.Threading;
using System.IO;

namespace OutlookAddIn1
{
    public partial class ThisAddIn
    {
        //create a new inspector for memory-handling
        Outlook.Inspectors inspectors;

        public static Outlook.AppointmentItem appointmentItem;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {

            //attach an event handler to the inspector
            //the 'Application' class represents the current instance of Outlook
            //'inspector' represents the inspector window of the new mail message
            inspectors = this.Application.Inspectors;
            inspectors.NewInspector += Inspectors_NewInspector;

        }

        private void Inspectors_NewInspector(Outlook.Inspector Inspector)
        {
            //example: when the user creates a new meeting item, the subject and body is populated by code
            appointmentItem = Inspector.CurrentItem as Outlook.AppointmentItem;
            if(appointmentItem != null)
            {
                if(appointmentItem.EntryID==null)
                {
                    appointmentItem.Subject = "This subject was added via code";
                }
            }

        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see http://go.microsoft.com/fwlink/?LinkId=506785
        }


        public static void button1_Click(object sender, EventArgs e)
        {
            var btn = (System.Windows.Forms.Button)sender;
            if (btn.Tag.ToString() != "")
            {
                appointmentItem.Recipients.Add(btn.Tag.ToString());
                btn.BackColor = System.Drawing.Color.DarkOliveGreen;
            }
        }

        public static void decideButtonColor(object sender, EventArgs e)
        {

            var btn = (System.Windows.Forms.Button)sender;
            string buttonTag = btn.Tag.ToString();
            if (btn.Tag.ToString() != "")
            {
                appointmentItem.Recipients.Add(btn.Tag.ToString());

                int startIndex;
                string freeBusy;

                int startHour = appointmentItem.StartInStartTimeZone.Hour;
                int startMinute = appointmentItem.StartInStartTimeZone.Minute;
                if (startMinute < 30)
                    startIndex = startHour * 2;
                else
                    startIndex = startHour * 2 + 1;

                freeBusy = appointmentItem.Recipients[appointmentItem.Recipients.Count].FreeBusy(appointmentItem.StartInStartTimeZone.Date, 30, false);
                
                if (freeBusy != null)
                {
                    if (freeBusy[startIndex] == '0')
                        btn.BackColor = System.Drawing.Color.LightGreen;//free
                    else
                        btn.BackColor = System.Drawing.Color.OrangeRed;//busy
                }

                appointmentItem.Recipients.Remove(appointmentItem.Recipients.Count);
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
    
    //define the list model
    public class listModel
    {
        public List<Site> site { get; set; }
    }

    public class Site
    {
        public string siteName { get; set; }
        public List<Floor> floor { get; set; }
    }

    public class Floor
    {
        public string floorName { get; set; }
        public List<Room> room { get; set; }
    }

    public class Room
    {
        public string roomId { get; set; }
        public string roomName { get; set; }
        public int locationX { get; set; }
        public int locationY { get; set; }
        public int sizeX { get; set; }
        public int sizeY { get; set; }
        public string tag { get; set; }
    }


}
