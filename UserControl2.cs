using System;
using System.Windows.Forms;
using System.DirectoryServices.AccountManagement;

namespace Brivo_Access
{
    public partial class UserControl2 : UserControl
    {


        public UserControl2()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            using (PrincipalContext pc = new PrincipalContext(ContextType.Domain, "FRUTAROM"))
            {
                Form1 wForm;
                wForm = (Form1)this.FindForm();
                UserControl3 term = new UserControl3();
                // validate the credentials
                bool isValid = pc.ValidateCredentials(textBox1.Text, textBox2.Text);

                if (isValid == true && textBox1.Text.ToUpper() == "JDARLAND" | textBox1.Text.ToUpper() == "CHURST" | textBox1.Text.ToUpper() == "LSTORTZ" | textBox1.Text.ToUpper() == "TSTANLEY" | textBox1.Text.ToUpper() == "MGOMES")
                {
                    wForm.panel1.Controls.Clear();
                    wForm.panel1.Controls.Add(term);
                }
                else
                {
                    Form2 error = new Form2();

                    error.Show();
                    textBox1.Clear();
                    textBox2.Clear();
                }
            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            UserControl1 login = new UserControl1();
            Form1 wForm;
            wForm = (Form1)this.FindForm();

            wForm.panel1.Controls.Clear();
            wForm.panel1.Controls.Add(login);
        }

        private void UserControl2_Load(object sender, EventArgs e)
        {

        }
    }
}
