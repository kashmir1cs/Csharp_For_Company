using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.Windows.Forms;
using System.Runtime.InteropServices;

// 마우스 이동 : System.Windows.Forms에 있는 Cursor 을 이용하여 사용가능
// EX : Cursor.Position = new Point(242, 83);
// 출처1 : https://hubbleconstant.tistory.com/18
// 출처2 : https://diy-dev-design.tistory.com/13

namespace MouseClick_Test
{
    public partial class Form1 : Form
    {
        //DLL 파일 Load
        [DllImport("user32.dll")]
        static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint dwData, int dwExtraInfo);
        //내장 상수 정의 
        private const uint MOVE = 0x0001;          // Mouse 이동 
        private const uint LEFTDOWN = 0x0002;      // The left button is down.
        private const uint LEFTUP = 0x0004;        // The left button is up.
        private const uint RIGHTDOWN = 0x0008;      // The right butting is down.
        private const uint RIGHTUP = 0x0010;        // The right butting is up.
        private const uint ABSOLUTEMOVE = 0x0080;        // The right butting is up.

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void btnClick_Click(object sender, EventArgs e)
        {
            int PtX,PtY;

            if (!string.IsNullOrEmpty(txtXcoord1.Text)&& !string.IsNullOrEmpty(txtYcoord1.Text))
            {
                PtX = Convert.ToInt32(txtXcoord1.Text);
                PtY = Convert.ToInt32(txtYcoord1.Text);
                Cursor.Position = new Point(PtX, PtY);
                mouse_event(LEFTDOWN, 0, 0, 0, 0);
                mouse_event(LEFTUP, 0, 0, 0, 0);
                Thread.Sleep(500);
            }
            lblStatus.Text = "1st Click 완료";
            lblStatus.Update();

            if (!string.IsNullOrEmpty(txtXcoord2.Text) && !string.IsNullOrEmpty(txtYcoord2.Text))
            {
                PtX = Convert.ToInt32(txtXcoord2.Text);
                PtY = Convert.ToInt32(txtYcoord2.Text);
                Cursor.Position = new Point(PtX, PtY);
                mouse_event(LEFTDOWN, 0, 0, 0, 0);
                mouse_event(LEFTUP, 0, 0, 0, 0);
                Thread.Sleep(500);

            }
            lblStatus.Text = "2st Click 완료";
            lblStatus.Update();

            if (!string.IsNullOrEmpty(txtXcoord3.Text) && !string.IsNullOrEmpty(txtYcoord3.Text))
            {
                PtX = Convert.ToInt32(txtXcoord3.Text);
                PtY = Convert.ToInt32(txtYcoord3.Text);
                Cursor.Position = new Point(PtX, PtY);
                mouse_event(LEFTDOWN, 0, 0, 0, 0);
                mouse_event(LEFTUP, 0, 0, 0, 0);

            }
            lblStatus.Text = "작업완료";
            lblStatus.Update();
              
        }
    }
}
