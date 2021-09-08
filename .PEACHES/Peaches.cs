using System;
using System.Collections.Generic;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.IO;

namespace Peaches
{

    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            
            tbCheckStatus.Clear();
            lblCheckStatus.Text = "xml 파일 미선택";



        }
        
        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void btnWeldTypeFileSelect_Click(object sender, EventArgs e)
        {
            
        }

        private void tableLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void lblWeldTypeSelect_Click(object sender, EventArgs e)
        {

        }

        private void lblCheckStatus_Click(object sender, EventArgs e)
        {

        }

        private void btnWeldTypeFileSelect_Click_1(object sender, EventArgs e)
        {
            openFileDialogWeldTypeFileSelect.Filter ="XML File (*.xml)|*.xml; *.XML;"; // File Filter

            openFileDialogWeldTypeFileSelect.Title = "Weld Type (XML) Select";

            openFileDialogWeldTypeFileSelect.InitialDirectory = System.IO.Directory.GetCurrentDirectory(); // set initial folder to current directory
            
            openFileDialogWeldTypeFileSelect.ShowDialog(); // show dialog


            DialogResult dialogResult = openFileDialogWeldTypeFileSelect.ShowDialog();
            
            if (dialogResult==DialogResult.OK && openFileDialogWeldTypeFileSelect.FileName.Length>0)
            {
                tbCheckStatus.Clear();
                lblCheckStatus.Text = "xml 파일 선택";
                tbCheckStatus.Text = ("XML 파일이 선택 되었습니다.") + Environment.NewLine;
                tbCheckStatus.Text += "XML Parsing 시작" + Environment.NewLine;
                WeldType WList = new WeldType();
                string WeldTypeXmlFileNameFull = openFileDialogWeldTypeFileSelect.FileName; // xml파일 (전체 경로 포함)
                
                string WeldTypeXmlFileName = Path.GetFileName(WeldTypeXmlFileNameFull); // 파일명만 추출
                
                XmlDocument WeldTypeXml = new XmlDocument(); // xml 객체 선언
                
                WeldTypeXml.Load(WeldTypeXmlFileNameFull);
                
                tbCheckStatus.Text += "XML Parsing 시작" + Environment.NewLine;
                tbCheckStatus.Text += "파일명 : " + WeldTypeXmlFileName + Environment.NewLine;
                
                XmlNodeList WeldTypeList = WeldTypeXml.SelectNodes("/WELDTYPES/WELD");
                //Console.WriteLine(WeldTypeXml.SelectSingleNode("/WELDTYPES/NAME").InnerText); //console에 text 표시
                tbCheckStatus.Text += "NAME : " + WeldTypeXml.SelectSingleNode("/WELDTYPES/NAME").InnerText + Environment.NewLine;
                
                
                
                foreach (XmlNode WeldType in WeldTypeList)
                {
                    tbCheckStatus.Text += "WELD SKEY : " + WeldType["SKEY"].InnerText + "/ Type : " + WeldType["TYPE"].InnerText+Environment.NewLine;
                
                    WList.WeldSkeyList.Add(WeldType["SKEY"].InnerText);
                    
                    tbCheckStatus.Text += "WELD SKEY 추가 완료" + Environment.NewLine;
                    

                }
                // Weld Skey 수량 확인
                
                tbCheckStatus.Text += "WELD SKEY 수량 : " + WList.WeldSkeyList.Count() + "개 "+ Environment.NewLine;
                lblCheckStatus.Text = "WELD SKEY 수량 : " + WList.WeldSkeyList.Count() + "개 ";

                // Class Weld Type에 저장된 SKEY 확인 (나중에 주석처리)

               /* for (int i=0; i<WList.WeldSkeyList.Count;i++)
                {
                    tbCheckStatus.Text += "WELD SKEY 확인 : " + (i+1)+"번 : " + WList.WeldSkeyList[i] + Environment.NewLine;
                }*/

            }
            else if (dialogResult==DialogResult.Cancel)
            {
                return;
            }
            
        }
        public class WeldType
        {
            public List<string> WeldSkeyList = new List<string>();
            // Weld Type 관리위한 Class 선언

        }

        private void lblIntro_Click(object sender, EventArgs e)
        {
            
        }
    }
}
