using System;
using System.Windows.Forms;
using System.Threading;


namespace SendKey
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        
        private void btnSend_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < 5; i++)
            {
                
                lblStatus.Text = (5 - i).ToString() + "초 후 입력 시작";
                lblStatus.Update();
                Thread.Sleep(1000);
            }
            
            lblStatus.Text = "입력을 시작합니다.";
            lblStatus.Update();
            Thread.Sleep(1000);
            SendKeys.Send(richTextBox1.Text);
            lblStatus.Text = "입력 완료.";
            lblStatus.Update();
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            DateTime dt = DateTime.Now;
            this.lblStatus.Text = dt.ToString();
        }
    }
}
