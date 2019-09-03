using System;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using OutlookApp = Microsoft.Office.Interop.Outlook.Application;

namespace Brivo_Access
{
    public partial class UserControl1 : UserControl
    {
        public string empName;
        public string reqName;
        public string manName;
        public string other;
        public string building;
        public string dept;
        public string location;
        public string tempF;
        public string dea;
        public string realemails = "JDarland@Frutarom.com" + ";" + "tstanley@Frutarom.com";
    //  public string realemails = "TStanley@Frutarom.com" +";" + "USA-IT@Frutarom.com";
     public string CAmail = "JDarland@Frutarom.com" + ";" + "tstanley@Frutarom.com";
        public string OHdea = "USA-IT@Frutarom.com";
         //  public string CAmail = "TSTanley@frutarom.com" + ";" + "usa-it@frutarom.com" + ";" + "dkaranzias@frutarom.com" + ";" + "santunes@Frutarom.com";
         //  public string OHdea = "TSTanley@frutarom.com" + ";" + "usa-it@frutarom.com" + ";" + "Santunes@frutarom.com";
        public string emailChecked = "";
        public string subject = "";
        Form2 error = new Form2();

        public UserControl1()
        {
            InitializeComponent();
            comboBack();
            label11.Text = "";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            dea = comboBox5.Text;
            empName = textBox1.Text;
             reqName = textBox2.Text;
             manName = textBox3.Text;
             other = textBox4.Text;
             building = comboBox1.Text;
             dept = comboBox2.Text;
             location = comboBox3.Text;
             tempF = comboBox4.Text;
             dea = comboBox5.Text;

         

            // Goto checkToSend Method to do the work
            checkToSend();
    }

        // This will display the HTML Body for each of the above.
        public string htmlBody ()
        {
            OutlookApp outlookApp = new OutlookApp();
            MailItem mailItem = outlookApp.CreateItem(OlItemType.olMailItem);
            string html;
            html = "<html>" +
                       "<body><strong>User Access Requested For:</strong>" + " " + empName +
                       "<br>" + "This is a " + tempF + " employee" +
                    "<br>" + "Does this Employee Need DEA Access: " + dea +
                    " <br>" + "The Manager is: " + manName +
                    "<br>" + "They will need Access to the Following Buildings " + building + "<br>" + "If Other: " + other +
                    "<br> " + "They will Be Located at: " + location +
                    "<br>" + "Department: " + dept +
                    "<br>" + "Person Requesting: " + reqName +
                    "<br>" +
                       "</body>" +
                       "</html>";
            return html;
        }
        // Clear the Form with this button.
        private void button2_Click(object sender, EventArgs e)
        {
            label11.Text = "";
            // Goto the method ComboBack(), which will reset the comboBox's back to their original selection
            comboBack();
            textBox1.Clear();
            textBox2.Clear();
            textBox3.Clear();
            textBox4.Text = "N/A (Default)";
        }
        private void comboBack ()
        {
            comboBox1.SelectedIndex = 0;
            comboBox2.SelectedIndex = 0;
            comboBox3.SelectedIndex = 0;
            comboBox4.SelectedIndex = 0;
            comboBox5.SelectedIndex = 0;
        }

       /*
        * This is the checkToSend Function
        * It's main goal is to check if the email should be sent successfully
        * or if there is an error, in which case it will show Form2, the error message
        * Returns no value.
        * */
       
        private void checkToSend()
        {

            var emptyTextboxes = from tb in this.Controls.OfType<TextBox>()
                                 where string.IsNullOrEmpty(tb.Text)
                                 select tb;

            OutlookApp outlookApp = new OutlookApp();
                MailItem mailItem = outlookApp.CreateItem(OlItemType.olMailItem);

            emailChecker();

            if (building != "Other" && building != "Corona")
            {
                mailItem.BCC = realemails;
                //    mailItem.BCC = myemail;
                mailItem.Subject = "Employee: " + empName;
                // This is the body of the email in HTML
                mailItem.HTMLBody = htmlBody();
                // oh the validation is so sweet
                if (emptyTextboxes.Any() != true && dea != "Yes")
                {
                    mailItem.Send();
                    label11.Text = "Success!";
                }
                else if (building == "Other" && emptyTextboxes.Any() != true && dea == "Yes")
                {
                    mailItem.BCC = OHdea;
                    //  mailItem.BCC = myemail;
                    mailItem.Subject = "User & DEA Access Requested For: " + empName;
                    // This is the body of the email in HTML. Goes to method htmlBody()
                    mailItem.HTMLBody = htmlBody();
                    mailItem.Send();
                    label11.Text = "Success!";
                }
                else if (empName == "" || manName == "" || reqName == "")
                {
                    error.Show();
                }
            }
            else if (building == "Corona")
            {
                mailItem.BCC = emailChecked;
                mailItem.Subject = subject;
                mailItem.HTMLBody = htmlBody();
                mailItem.Send();
                label11.Text = "Success!";
            }
            else if (building == "Other")
            {
                mailItem.BCC = emailChecked;
                //  mailItem.BCC = myemail;
                mailItem.Subject = subject;
                // This is the body of the email in HTML. Goes to method htmlBody()
                mailItem.HTMLBody = htmlBody();
                mailItem.Send();
                label11.Text = "Success!";
            }
            else
            {
                error.Show();
            }
            
        }
     
        /*
         * This function does 2 important things
         * First:
         * It will set the email that it needs to send to depending on a few attributes
         * If it is Corona or Other = building string
         * if it is dea or not
         * 
         * This will then change the Subject of the email and the desired recipients based on the information given.
         * */
        private void emailChecker ()
        {
            if (building == "Corona")
            {
                if (dea.ToUpper() == "YES")
                {
                    emailChecked = CAmail;
                    subject = "User & DEA Access Requested For: " + empName;
                }
                else if (dea.ToUpper() == "NO")
                {
                    emailChecked = realemails;
                    subject = "User Access Requested For: " + empName;
                }
            }
            else if (building == "Other")
            {
                if (dea.ToUpper() == "YES")
                {
                    emailChecked = OHdea;
                    subject = "User & DEA Access Requested For: " + empName;
                }
                else if (dea.ToUpper() == "NO")
                {
                    emailChecked = realemails;
                    subject = "User Access Requested For: " + empName;
                }
            }
        }
        private void deaEmail ()
        {
         
            
        }

        private void UserControl1_Load(object sender, EventArgs e)
        {

        }
    }

    // End of Program
}
