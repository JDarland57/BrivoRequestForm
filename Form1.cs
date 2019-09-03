using System;
using System.Windows.Forms;
using HtmlAgilityPack;

namespace Brivo_Access
{
    public partial class Form1 : Form
    {
        UserControl1 form = new UserControl1();
       
        UserControl2 hr = new UserControl2();
       
        public string RELEASE = "1.0.0";

        public Form1()
        {
            InitializeComponent();
            this.Location = new System.Drawing.Point(750, 50);
            panel1.Controls.Add(form);
           
            string version = "";

          
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {
           
        }

        private void terminateToolStripMenuItem_Click(object sender, EventArgs e)
        {
            panel1.Controls.Clear();
            panel1.Controls.Add(hr);
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        int mouseX = 0, mouseY = 0;
        bool mouseDown;

        private void panel2_MouseDown(object sender, MouseEventArgs e)
        {
            mouseDown = true;
        }

        private void panel2_MouseUp(object sender, MouseEventArgs e)
        {
            mouseDown = false;
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void iTUpdatesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
        }

        private void panel2_MouseMove(object sender, MouseEventArgs e)
        {
            if (mouseDown)
            {
                mouseX = MousePosition.X - 200;
                mouseY = MousePosition.Y - 40;

                this.SetDesktopLocation(mouseX, mouseY);
            }
        }
    }
}
