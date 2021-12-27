using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace FW_TestTask
{
    public partial class Form1 : Form
    {
        public new List<string> Events = new List<string>();

        public Form1()
        {
            InitializeComponent();
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
        }

        private void btnGetEvents_Click(object sender, EventArgs e)
        {
            ((Button)sender).Enabled = false;

            dgvEvents.Rows.Clear();

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
                    object[] eventRow =
                        { oItem.Subject, oItem.Start, new TimeSpan(0, oItem.Duration, 0).GetFriendlyName() };
                    dgvEvents.Rows.Add(eventRow);
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