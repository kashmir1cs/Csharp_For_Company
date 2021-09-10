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
using EX = Microsoft.Office.Interop.Excel;

namespace Peaches
{

    public partial class Form1 : Form
    {
        public static List<string> WeldSkeyList = new List<string>(); //WeldSkeyList 전역변수 선언
        public static List<string> ExcelFiles = new List<string>(); // Excel Files List 전역변수 선언
        public static List<string> FilteredExcelFiles = new List<string>(); // Error확인된 Excel Files List 전역변수 선언
        public static List<string> FilteringResult = new List<string>(); // 검사 결과 txt파일로 저장하기 위한 List 전역변수 선언


        public Form1()
        {
            InitializeComponent();
            // UI TextBox, Label 초기화
            tbCheckStatus.Clear();
            lblFinalStatus.Text = "Weld Type 미설정/ Folder 미선택";
            pBarCheckStatus.Maximum = 100;
            pBarCheckStatus.Minimum = 0;
            pBarCheckStatus.Value = 0;
            tbAbout.Text = "사용방법 : \r\n 1. Weld Type File선택 (xml) \r\n 2. SSU에서 ISSUE한 파일이 있는 Folder 선택 \r\n 3. Weld Plan Check Click";

        }
        public void LabelCheck()
        {
            // Label
            if (WeldSkeyList.Count == 0 && ExcelFiles.Count == 0)
            {
                lblFinalStatus.Text = "Weld Type 미설정/ Folder 미선택";
            }
            else if (WeldSkeyList.Count != 0 && ExcelFiles.Count == 0)
            {
                lblFinalStatus.Text = "Weld Type 설정/ Folder 미선택";
            }
            else if (WeldSkeyList.Count != 0 && ExcelFiles.Count != 0)
            {
                lblFinalStatus.Text = "Weld Type 설정/ Folder 선택";
            }
            else if (WeldSkeyList.Count == 0 && ExcelFiles.Count != 0)
            {
                lblFinalStatus.Text = "Weld Type 미설정/ Folder 선택";
            }
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
            LabelCheck();
            openFileDialogWeldTypeFileSelect.Filter = "XML File (*.xml)|*.xml; *.XML;"; // File Filter

            openFileDialogWeldTypeFileSelect.Title = "Weld Type (XML) Select";

            openFileDialogWeldTypeFileSelect.InitialDirectory = System.IO.Directory.GetCurrentDirectory(); // set initial folder to current directory

            if (openFileDialogWeldTypeFileSelect.ShowDialog() == DialogResult.OK && WeldSkeyList.Count == 0)
            {
                lblWeldTypeSelect.Text += openFileDialogWeldTypeFileSelect.FileName;
                tbCheckStatus.Text += ("XML 파일이 선택 되었습니다.") + Environment.NewLine;
                tbCheckStatus.Text += "XML Parsing 시작" + Environment.NewLine;

                string WeldTypeXmlFileNameFull = openFileDialogWeldTypeFileSelect.FileName; // xml파일 (전체 경로 포함)

                string WeldTypeXmlFileName = Path.GetFileName(WeldTypeXmlFileNameFull); // 파일명만 추출

                XmlDocument WeldTypeXml = new XmlDocument(); // xml 객체 선언

                WeldTypeXml.Load(WeldTypeXmlFileNameFull);

                tbCheckStatus.Text += "XML Parsing 시작" + Environment.NewLine;
                tbCheckStatus.Text += "파일명 : " + WeldTypeXmlFileName + Environment.NewLine;

                XmlNodeList WeldTypeList = WeldTypeXml.SelectNodes("/WELDTYPES/WELD");

                //Console.WriteLine(WeldTypeXml.SelectSingleNode("/WELDTYPES/NAME").InnerText); //console에 text 표시
                tbCheckStatus.Text += "NAME : " + WeldTypeXml.SelectSingleNode("/WELDTYPES/NAME").InnerText + Environment.NewLine;
                LabelCheck();


                foreach (XmlNode WeldType in WeldTypeList)
                {
                    tbCheckStatus.Text += "WELD SKEY : " + WeldType["SKEY"].InnerText + "/ Type : " + WeldType["TYPE"].InnerText + " - 추가 완료" + Environment.NewLine;
                    WeldSkeyList.Add(WeldType["SKEY"].InnerText);

                }
                // Weld Skey 수량 확인

                tbCheckStatus.Text += "WELD SKEY 수량 : " + WeldSkeyList.Count() + "개 " + Environment.NewLine;
                LabelCheck();
                // Class Weld Type에 저장된 SKEY 확인 (나중에 주석처리)

                /* for (int i=0; i<WList.WeldSkeyList.Count;i++)
                 {
                     tbCheckStatus.Text += "WELD SKEY 확인 : " + (i+1)+"번 : " + WList.WeldSkeyList[i] + Environment.NewLine;
                 }*/

            }
            else if (openFileDialogWeldTypeFileSelect.ShowDialog() == DialogResult.OK && WeldSkeyList.Count != 0)
            {
                lblWeldTypeSelect.Text = "파일을 선택하세요 : ";
                tbCheckStatus.Text += ("XML 파일이 선택 되었습니다.") + Environment.NewLine;
                tbCheckStatus.Text += "XML Parsing 시작" + Environment.NewLine;

                string WeldTypeXmlFileNameFull = openFileDialogWeldTypeFileSelect.FileName; // xml파일 (전체 경로 포함)

                string WeldTypeXmlFileName = Path.GetFileName(WeldTypeXmlFileNameFull); // 파일명만 추출

                XmlDocument WeldTypeXml = new XmlDocument(); // xml 객체 선언

                WeldTypeXml.Load(WeldTypeXmlFileNameFull);

                tbCheckStatus.Text += "XML Parsing 시작" + Environment.NewLine;
                tbCheckStatus.Text += "파일명 : " + WeldTypeXmlFileName + Environment.NewLine;

                XmlNodeList WeldTypeList = WeldTypeXml.SelectNodes("/WELDTYPES/WELD");

                //Console.WriteLine(WeldTypeXml.SelectSingleNode("/WELDTYPES/NAME").InnerText); //console에 text 표시
                tbCheckStatus.Text += "NAME : " + WeldTypeXml.SelectSingleNode("/WELDTYPES/NAME").InnerText + Environment.NewLine;
                LabelCheck();


                foreach (XmlNode WeldType in WeldTypeList)
                {
                    tbCheckStatus.Text += "WELD SKEY : " + WeldType["SKEY"].InnerText + "/ Type : " + WeldType["TYPE"].InnerText + " - 추가 완료" + Environment.NewLine;
                    WeldSkeyList.Add(WeldType["SKEY"].InnerText);
                }
                // Weld Skey 수량 확인

                tbCheckStatus.Text += "WELD SKEY 수량 : " + WeldSkeyList.Count() + "개 " + Environment.NewLine;
                LabelCheck();
                // Class Weld Type에 저장된 SKEY 확인 (나중에 주석처리)

                /* for (int i=0; i<WList.WeldSkeyList.Count;i++)
                 {
                     tbCheckStatus.Text += "WELD SKEY 확인 : " + (i+1)+"번 : " + WList.WeldSkeyList[i] + Environment.NewLine;
                 }*/

            }
            else if (openFileDialogWeldTypeFileSelect.ShowDialog() == DialogResult.Cancel)
            {
                return;
            }

        }

        private void btnWeldPlanFolderSelect_Click(object sender, EventArgs e)
        {
            LabelCheck();
            // "Weld Plan만 선택하도록 반복문 작성"
            // Linq Lambda 식 활용
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            if (fbd.ShowDialog() == DialogResult.OK)
            {
                lblWeldPlanFolderSelect.Text = "폴더를 선택하세요 : ";
                lblWeldPlanFolderSelect.Text += fbd.SelectedPath;
                ExcelFiles = Directory.GetFiles(fbd.SelectedPath, "*.xlsx").Where(s=>s.Contains("__WP__")).ToList(); // Linq 활용
                for (int i =0;i<ExcelFiles.Count;i++)
                {
                    tbCheckStatus.Text += ExcelFiles[i] + Environment.NewLine;
                }

                LabelCheck();
            }
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            tbCheckStatus.Clear();
        }

        private void btnExecute_Click(object sender, EventArgs e)
        {
            // Weld Type과 Weld Plan이 설정된 경우에만 작동
            if (WeldSkeyList.Count == 0 && ExcelFiles.Count == 0)
            {
                MessageBox.Show("Weld Skey Type이 Load되지 않았습니다.\n Weld Plan File이 Load 되지 않았습니다.","설정 미완료");
            }

            else if(WeldSkeyList.Count != 0 && ExcelFiles.Count == 0)
            {
                MessageBox.Show("Weld Plan File이 Load 되지 않았습니다.", "설정 미완료");
            }
            else if (WeldSkeyList.Count == 0 && ExcelFiles.Count != 0)
            {
                MessageBox.Show("Weld Skey Type이 Load되지 않았습니다.", "설정 미완료");
            }
            else
            { 
                EX.Application ExFile = new EX.Application();
                double total = Convert.ToDouble(ExcelFiles.Count);
                pBarCheckStatus.Value = 0; // Progress Bar 초기화
                // Excel 파일 열어서 Human Error 확인
                /* 확인사항 
                 * 1. Weld Plan Joint No. 중복 입력 
                 * 
                 * 2. Shop용접 Point Spool No. 미입력
                 * 
                 * 3. Weld Type 확인 : Xml에서 읽어온 Weld Type과 비교
                 * 
                */
                for (int i = 0; i < ExcelFiles.Count; i++)
                {
                    lblProgress.Text = "진행율 : " + "(" + (i + 1).ToString() + "of" + ExcelFiles.Count + ")";
                    double p = Math.Round((i + 1) * 100 / total);
                    pBarCheckStatus.Value = Convert.ToInt32(p);
                    pBarCheckStatus.Update();
                    //tbCheckStatus.Text += p.ToString() + Environment.NewLine; //진행율 확인
                    tbCheckStatus.Text += "Weld Plan :" + ExcelFiles[i] + " - 확인 시작" + Environment.NewLine;

                    // Excel Class 선언
                    // ExcelFiles에 있는 파일 하나씩 실행
                    EX.Workbook WeldPlanExcelFile = null;
                    EX.Worksheet WeldPlanExcelSheet = null;
                    WeldPlanExcelFile = ExFile.Workbooks.Open(Filename: ExcelFiles[i], ReadOnly: true);
                    WeldPlanExcelSheet = WeldPlanExcelFile.Worksheets.Item[1];
                    int RowEnd = WeldPlanExcelSheet.Range["A1"].End[EX.XlDirection.xlDown].Row; //마지막 행번호 변수 할당*/
                                                                                                // Weld Plan값을 입력받을 List 선언

                    /*                FileStream ExStream = File.Open(ExcelFiles[i], FileMode.Open, FileAccess.Read);
                                    IExcelDataReader excelReader;
                                    excelReader = ExcelReaderFactory.CreateBinaryReader(ExStream);
                                    DataSet WeldPlanData = excelReader.AsDataSet();*/
                    List<string> JointList = new List<string>(); //A열
                    List<string> SpoolNoList = new List<string>(); //B열
                    List<string> RawMaterialList = new List<string>(); //D열
                    List<string> SchList = new List<string>(); //F열
                    List<string> WeldTypeList = new List<string>(); //G열
                    List<string> ShopFieldList = new List<string>(); //H열
                    List<string> ShopSpoolList = new List<string>(); //B열,D열 함께입력

                    // 중복 검사
                    for (int j = 2; j <= RowEnd; j++)
                    {
                        //각 List에 Weld Plan Data 입력
                        JointList.Add(WeldPlanExcelSheet.Range["A" + j].Value);
                        SpoolNoList.Add(WeldPlanExcelSheet.Range["B" + j].Value);
                        RawMaterialList.Add(WeldPlanExcelSheet.Range["D" + j].Value);
                        SchList.Add(WeldPlanExcelSheet.Range["F" + j].Value);
                        WeldTypeList.Add(WeldPlanExcelSheet.Range["G" + j].Value);
                        ShopFieldList.Add(WeldPlanExcelSheet.Range["H" + j].Value);
                        ShopSpoolList.Add(WeldPlanExcelSheet.Range["B" + j].Value + "_" + WeldPlanExcelSheet.Range["H" + j].Value);
                    }
                    // HUman Error 확인 
                    // 1. Joint 중복 확인
                    if (JointList.Count == JointList.Distinct().Count())
                    {
                        tbCheckStatus.Text += " 1) Joint no. 중복 확인 : Joint No. 중복 없음" + Environment.NewLine;
                    }
                    else
                    {
                        tbCheckStatus.Text += " 1) Joint no. 중복 확인 : Joint No. 중복 입력" + Environment.NewLine;
                        // Error 확인한 Weld Plan 파일 List에 추가
                        FilteredExcelFiles.Add(ExcelFiles[i]);
                    }
                    // 2. Spool No. 미표기 확인
                    if (ShopSpoolList.Contains("_SHOP"))
                    {
                        tbCheckStatus.Text += " 2) Shop Joint Spool no. 미입력 : 있음" + Environment.NewLine;
                        if (FilteredExcelFiles.Contains(ExcelFiles[i]) == false)
                        {
                            FilteredExcelFiles.Add(ExcelFiles[i]);
                        }

                    }
                    else
                    {
                        tbCheckStatus.Text += " 2) Shop Joint Spool no. 미입력 : 없음" + Environment.NewLine;
                    }
                    // 3. 미등록된 Weld Type 입력
                    // Linq Except 메소드 이용
                    IEnumerable<string> diffWeldType = WeldTypeList.Except(WeldSkeyList);
                    if (diffWeldType.ToList().Count > 0)
                    {
                        tbCheckStatus.Text += " 3) Weld Type 입력 : 미등록 Weld Type 입력 (공백포함)" + Environment.NewLine;
                        if (FilteredExcelFiles.Contains(ExcelFiles[i]) == false)
                        {
                            FilteredExcelFiles.Add(ExcelFiles[i]);
                        }
                    }
                    else
                    {
                        tbCheckStatus.Text += " 3) Weld Type 입력 : OK" + Environment.NewLine;
                    }
                    WeldPlanExcelFile.Close();
                }
            }

        }

        private void btnXmlReset_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Weld Type이 삭제됩니다 계속 하시겠습니까?", "Data Reset", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {

                if (WeldSkeyList.Count == 0)
                {
                    MessageBox.Show("Reset할 Data가 없습니다.", "Reset 완료");
                    LabelCheck();
                }
                else
                {
                    WeldSkeyList.Clear();
                    MessageBox.Show("Weld Type Reset완료.", "Reset 완료");
                    lblWeldTypeSelect.Text = "파일을 선택하세요 : ";
                    LabelCheck();
                }
            }
            else
            {
                MessageBox.Show("작업 취소되었습니다.", "작업 취소");
                LabelCheck();
            }

            /*private void btnReset_Click(object sender, EventArgs e)
            {
                if(MessageBox.Show("Weld Type과 Weld Plan 목록이 삭제됩니다 계속 하시겠습니까?","Data Reset",MessageBoxButtons.YesNo)==DialogResult.Yes)
                {

                    if (WeldSkeyList.Count == 0 && ExcelFiles.Count == 0)
                    {
                        MessageBox.Show("Reset할 Data가 없습니다.","Reset 완료");
                        LableCheck();
                    }
                    else if (WeldSkeyList.Count != 0 && ExcelFiles.Count == 0)
                    {
                        WeldSkeyList.Clear();
                        MessageBox.Show("Weld Type Reset완료.", "Reset 완료");
                        LableCheck();
                    }
                    else if (WeldSkeyList.Count != 0 && ExcelFiles.Count != 0)
                    {
                        WeldSkeyList.Clear();
                        ExcelFiles.Clear();
                        MessageBox.Show("Weld Type/ Weld Plan Reset 완료", "Reset 완료");
                        LableCheck();
                    }
                    else if (WeldSkeyList.Count == 0 && ExcelFiles.Count != 0)
                    {
                        ExcelFiles.Clear();
                        MessageBox.Show("Weld Plan Reset 완료", "Reset 완료");
                        LableCheck();
                    }

                }
                else
                {
                    MessageBox.Show("작업 취소되었습니다.", "작업 취소");
                    LableCheck();
                }
            }*/
        }

        private void btnWeldPlanFolderSelectReset_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Weld Plan이 삭제됩니다 계속 하시겠습니까?", "Data Reset", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {

                if (ExcelFiles.Count == 0)
                {
                    MessageBox.Show("Reset할 Data가 없습니다.", "Reset 완료");
                    LabelCheck();
                }
                else
                {
                    ExcelFiles.Clear();
                    MessageBox.Show("Weld Type Reset완료.", "Reset 완료");
                    lblWeldPlanFolderSelect.Text = "폴더를 선택하세요 : ";
                    LabelCheck();
                }
            }
            else
            {
                MessageBox.Show("작업 취소되었습니다.", "작업 취소");
                LabelCheck();
            }
        }
    }
}
