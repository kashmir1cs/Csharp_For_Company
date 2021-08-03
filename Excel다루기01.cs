using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace SPEU
{
    
    public partial class FormMain : Form
    {
        public List<string> FileList = new List<string>();
        public FormMain()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void tableLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {

        }

        private void fileOpenToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if(openFileDialog1.ShowDialog()==DialogResult.OK)
            {
                foreach (string FilePath in openFileDialog1.FileNames)
                {
                    FileList.Add(FilePath); // filelist에 파일명 추가

                }
                foreach (string FileName in FileList)
                {
                    rtxtStatus.Text += FileName.ToString() + Environment.NewLine;


                }
            }
        }

        private void Setup_Click(object sender, EventArgs e)
        {

        }

        private void btnExec_Click(object sender, EventArgs e)
        {
            if (FileList != null)
            {
                Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
                foreach (string f in FileList)
                {
                    Workbook workbook = app.Workbooks.Open(f);
                    app.Visible = true;
                }


            }
        }
    }
}
