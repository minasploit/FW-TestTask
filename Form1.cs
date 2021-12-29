using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace FW_TestTask
{
    public partial class Form1 : Form
    {
        public new List<string> Events = new List<string>();
        private const string Chart1SeriesName = "Minutes Spent";
        private const string Chart2SeriesName = "Hours Spent";

        public Form1()
        {
            InitializeComponent();
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);

            chart1.Titles.Add("Time spent on each Task");
            chart2.Titles.Add("Hours worked per Day");

            startDateTime.Value = startDateTime.Value.AddDays(-7);
        }

        private void btnGetEvents_Click(object sender, EventArgs e)
        {
            ((Button)sender).Enabled = false;

            // clear existing data
            dgvEvents.Rows.Clear();
            
            chart1.Series[Chart1SeriesName].Points.Clear();
            chart2.Series[Chart2SeriesName].Points.Clear();

            try
            {
                var outlookApp = new Outlook.Application();
                var oNameSpace = outlookApp.GetNamespace("mapi");

                oNameSpace.Logon(Missing.Value, Missing.Value, true, true);

                // Get the Calendar folder.
                var oCalendar = oNameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);

                var searchStartDate = startDateTime.Value;
                var searchEndDate = endDateTime.Value;

                // setup filters for getting items from outlook
                var dateFilter = "[Start] >= '" + searchStartDate.ToString("g") + "'";
                dateFilter += " AND [End] <= '" + searchEndDate.ToString("g") + "'";

                var oItems = oCalendar.Items.Restrict(dateFilter);

                var hoursWorkedPerDay = searchStartDate.DatesBetween(searchEndDate)
                    .Select(d => new { Date = d, Hours = TimeSpan.Zero }).ToDictionary(x => x.Date, x => x.Hours);

                foreach (Outlook.AppointmentItem oItem in oItems)
                {
                    var eventStart = oItem.Start;
                    var eventEnd = oItem.End;
                    var eventDuration = new TimeSpan(0, oItem.Duration, 0);

                    // set the data on the table
                    object[] eventRow =
                        { oItem.Subject, eventStart, new TimeSpan(0, oItem.Duration, 0).GetFriendlyName() };
                    dgvEvents.Rows.Add(eventRow);

                    // set the data on the first chart
                    chart1.Series[Chart1SeriesName].Points.AddXY($"{oItem.Subject}\n{eventStart:D}", oItem.Duration);

                    // append the time worked to the dictionary
                    if (eventStart.Date == eventEnd.Date)
                    {
                        hoursWorkedPerDay[eventStart.Date] =
                            hoursWorkedPerDay[eventStart.Date].Add(eventDuration);
                        continue;
                    }
                    
                    //// the event crosses a day boundary

                    //// trim the end date so it doesn't cross the user-defined margin
                    if (eventEnd.Date > searchEndDate.Date)
                        eventEnd = searchEndDate;

                    //// set the time worked for the first day of the event
                    var timeWorkedOnStartDay = TimeSpan.FromHours(24) - eventStart.TimeOfDay;
                    hoursWorkedPerDay[eventStart.Date] =
                        hoursWorkedPerDay[eventStart.Date].Add(timeWorkedOnStartDay);

                    //// loop and set the full days worked throughout the event
                    var eventDurationCopy = eventDuration;
                    var fullDays = 1;
                    while (eventDurationCopy >= TimeSpan.FromHours(24))
                    {
                        hoursWorkedPerDay[eventStart.AddDays(fullDays).Date] =
                            hoursWorkedPerDay[eventStart.AddDays(fullDays).Date].Add(TimeSpan.FromHours(24));
                        
                        eventDurationCopy = eventDurationCopy.Subtract(TimeSpan.FromHours(24));
                        
                        fullDays++;
                    }
                    
                    //// set the time worked for the last day of the event
                    var timeWorkedOnEndDay = eventEnd.TimeOfDay;
                    hoursWorkedPerDay[eventEnd.Date] =
                        hoursWorkedPerDay[eventEnd.Date].Add(timeWorkedOnEndDay);
                }
                
                // set the data on the second chart
                foreach (var keyValuePair in hoursWorkedPerDay)
                {
                    var timeWorked = keyValuePair.Value;
                    
                    // make sure the time doesn't exceed 24 hours
                    if (timeWorked > TimeSpan.FromHours(24))
                        timeWorked = timeWorked.Subtract(TimeSpan.FromHours(24));
                    
                    chart2.Series[Chart2SeriesName].Points.AddXY(keyValuePair.Key.Date.ToString(), timeWorked.TotalHours);
                }

                oNameSpace.Logoff();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"{ex.Message} Exception caught.");
            }
            finally
            {
                ((Button)sender).Enabled = true;
            }
        }
    }
}