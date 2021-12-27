using System;
using System.Collections.Generic;
using System.Reflection;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace FW_TestTask
{
    public partial class Form1 : Form
    {
        public new List<string> Events = new List<string>();
        private const string SeriesName = "Minutes Spent";

        public Form1()
        {
            InitializeComponent();
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            
            chart1.Titles.Add("Time spent on each event");

            startDateTime.Value = startDateTime.Value.AddDays(-7);
        }

        private void btnGetEvents_Click(object sender, EventArgs e)
        {
            ((Button)sender).Enabled = false;

            // clear existing data
            dgvEvents.Rows.Clear();
            chart1.Series[SeriesName].Points.Clear();

            try
            {
                var outlookApp = new Outlook.Application();
                var oNameSpace = outlookApp.GetNamespace("mapi");

                oNameSpace.Logon(Missing.Value, Missing.Value, true, true);

                // Get the Calendar folder.
                var oCalendar = oNameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);

                var dateFilter = "[Start] >= '" + startDateTime.Value.ToString("g") + "'";
                dateFilter += " AND [End] <= '" + endDateTime.Value.ToString("g") + "'";

                var oItems = oCalendar.Items.Restrict(dateFilter);

                foreach (Outlook.AppointmentItem oItem in oItems)
                {
                    // set the data on the table
                    object[] eventRow =
                        { oItem.Subject, oItem.Start, new TimeSpan(0, oItem.Duration, 0).GetFriendlyName() };
                    dgvEvents.Rows.Add(eventRow);
                    
                    // set the data on the chart
                    chart1.Series[SeriesName].Points.AddXY($"{oItem.Subject}\n{oItem.Start:D}", oItem.Duration);
                }

                oNameSpace.Logoff();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"{ex} Exception caught.");
            }
            finally
            {
                ((Button)sender).Enabled = true;
            }
        }
    }
}