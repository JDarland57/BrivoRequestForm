using System;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using OutlookApp = Microsoft.Office.Interop.Outlook.Application;

namespace Brivo_Access
{
    public partial class UserControl3 : UserControl
    {
        public string realemails = "JDarland@Frutarom.com" + ";" + "tstanley@Frutarom.com";
        // public string realemails = "TStanley@Frutarom.com" + ";" +"USA-IT@Frutarom.com";
        public string selectedDate = "";
        public string useremail = "";
        public string eob;
        public string term;
        public string empName;
        

        public UserControl3()
        {
            InitializeComponent();
            resetCombo();
        }

        private void monthCalendar1_DateSelected(object sender, DateRangeEventArgs e)
        {
            dateBox.Text = e.Start.ToShortDateString();

            selectedDate = dateBox.Text;
            
        }

        /*
         * Button click 1 is the main button click, and will submit all the work
         * it reaches out to other functions, but on click event will complete all work
         * returns no values
         * */
        private void button1_Click(object sender, EventArgs e)
        {
            empName = textBox1.Text;
            term = comboBox1.Text;
            eob = comboBox2.Text;
            useremail = textBox2.Text;
            doEmail();
        }

        /*
         * CheckBlank is checking if useremail is blank
         * If it is, assign useremail to usa-it@frutarom.com
         * this is to ensure that no errors are thrown by a blank string being attachedc
         * returns no value
         * */
        private void checkEBlank()
        {
            if (useremail == "")
            {
                useremail = "USA-IT@Frutarom.com";
            }
        }
        /*
         * The htmlBody function is to attach the html body, and to clean up the code for visiblitity
         * There is no reason it cannot be attached to the doEmail function
         * but is placed here to read easier
         * returns html. which is the body of the html
         * */
        public string htmlBody ()
        {
            string html;

            html = "<html>" +
                "<body><strong>Termination | Suspension Form</strong>" + " " +
                "<br>" + "This employee is being: " + term +
             "<br>" + "Termination Date: " + selectedDate +
             "<br>" + "Time For Termination: " + eob +
                "</body>" +
                "</html>";
            return html;
        }
        /*
         * Send the email here
         * this has all the information from 
         * Emails to send information to
         * subject
         * body
         * label for successful email
         * returns no values
         * */
        private void doEmail()
        {
            OutlookApp outlookApp = new OutlookApp();
            MailItem mailItem = outlookApp.CreateItem(OlItemType.olMailItem);
            // Goto the CheckBlank function to see if userEmail is blank or not
            checkEBlank();

            mailItem.BCC = realemails;
            mailItem.CC = useremail;
            //    mailItem.BCC = myemail;
            mailItem.Subject = "Employee: " + empName;
            // This is the body of the email in HTML. GoTo htmlBody Function
            mailItem.HTMLBody = htmlBody();
            mailItem.Send();
            label6.Text = "Success!";
            resetCombo();
        }

        private void resetCombo ()
        {
            comboBox1.SelectedIndex = 0;
            comboBox2.SelectedIndex = 0;
        }

        private void dateBox_DateChanged(object sender, DateRangeEventArgs e)
        {

        }

        private void UserControl3_Load(object sender, EventArgs e)
        {

        }
    }
}
