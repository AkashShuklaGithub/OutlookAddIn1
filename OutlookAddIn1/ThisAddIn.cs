using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using Outlook = Microsoft.Office.Interop.Outlook;

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
                    appointmentItem.Subject = "`";
                }
            }

        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see http://go.microsoft.com/fwlink/?LinkId=506785
        }

        public static void decideButtonColor(object sender, EventArgs e)
        {
            //change all button color to green
            for(int i=0;i<FormRegion.listButton.Count;i++)
            {
                if(FormRegion.listButton[i].Tag.ToString()!="")
                {
                    FormRegion.listButton[i].BackColor = System.Drawing.Color.LightGreen;
                }
            }

            //check the busy rooms
            //change the color to red
            ExchangeService service = new ExchangeService();
            service.UseDefaultCredentials = true;
            service.Url = new Uri("https://email.netapp.com/EWS/Exchange.asmx");
            AvailabilityOptions myOptions = new AvailabilityOptions();
            myOptions.MeetingDuration = 30;
            myOptions.RequestedFreeBusyView = FreeBusyViewType.Detailed;
            GetUserAvailabilityResults freeBusyResults = service.GetUserAvailability(FormRegion.attendees, new TimeWindow(appointmentItem.StartInStartTimeZone, appointmentItem.StartInStartTimeZone.AddDays(1)), AvailabilityData.FreeBusy, myOptions);

            foreach (AttendeeAvailability availability in freeBusyResults.AttendeesAvailability)
            {
                foreach (CalendarEvent calendarItem in availability.CalendarEvents)
                {
                    if (DateTime.Compare(calendarItem.StartTime, appointmentItem.Start) < 0 && DateTime.Compare(calendarItem.EndTime, appointmentItem.End) > 0)
                    {
                        for (int q = 0; q < FormRegion.listButton.Count; q++)
                        {
                            if (FormRegion.listButton[q].Name == calendarItem.Details.Location)
                            {
                                FormRegion.listButton[q].BackColor = System.Drawing.Color.OrangeRed;
                            }
                        }
                    }
                }
            }
        }

        public static void button1_Click(object sender, EventArgs e)
        {
            //when a meeting room is clicked
            var btn = (System.Windows.Forms.Button)sender;
            if (btn.Tag.ToString() != "")
            {
                appointmentItem.Recipients.Add(btn.Tag.ToString());
                btn.BackColor = System.Drawing.Color.DarkOliveGreen;
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
        public string longName { get; set; }
        public int locationX { get; set; }
        public int locationY { get; set; }
        public int sizeX { get; set; }
        public int sizeY { get; set; }
        public string tag { get; set; }
    }


}
