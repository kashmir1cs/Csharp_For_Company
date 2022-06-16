using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.IO; // File Open할 때 이용
using ExcelDataReader;
using System.Data;
using System.Diagnostics; // Stopwatch 활용 (처리 시간 측정)
// 1. xml 활용 
// 2. Datareader활용 - Data 읽는 속도 획기적으로 단축 (Interop.Excel 대비)
// 참고 자료 1: https://pgh268400.tistory.com/247
// 참고 자료 2 : https://m.blog.naver.com/kimmingul/221866465945



namespace XmlDataReader
{
    class Program
    {
        static void Main(string[] args)
        {
            string xmlpath = @"C:\SdataInfo\Setting.xml";
            XmlDocument xml = new XmlDocument(); // xml 객체 선언
            xml.Load(xmlpath);
            XmlNodeList xmltaglist = xml.SelectNodes("/Setup"); //Setup 노드를 탐색하며 저장
            // tag안에 들어 있는 "Text"추출
            // 숫자도 기본적으로 Text로 인식
            Console.WriteLine(xml.SelectSingleNode("/Setup/Directory").InnerText);
            Console.WriteLine("\n");
            Console.WriteLine(xml.SelectSingleNode("/Setup/File").InnerText);
            Console.WriteLine("\n");
            Console.WriteLine(xml.SelectSingleNode("/Setup/Month").InnerText);
            Console.WriteLine("\n");
            Console.WriteLine(xml.SelectSingleNode("/Setup/Weight").InnerText);
            Console.WriteLine("\n");
            Console.WriteLine("Tag 파싱한 값은 기본적으로 문자열 취급 : 숫자 변환시 int.Pars()등 활용 필요");
            Console.WriteLine("\n");
            Console.WriteLine(int.Parse( xml.SelectSingleNode("/Setup/Weight").InnerText)+1);
            Console.WriteLine("\n");
            Console.WriteLine(xml.SelectSingleNode("/Setup/Directory").InnerText + "\\"+ xml.SelectSingleNode("/Setup/File").InnerText);
            Console.WriteLine("\n");
            
            string exFileName = xml.SelectSingleNode("/Setup/Directory").InnerText + "\\" + xml.SelectSingleNode("/Setup/File").InnerText;
            // Data Reader Test하기 

            var stopwatch = new Stopwatch();// 처리 시간 측정
            stopwatch.Start();//시간 측정 시작

            // File STream 생성  하기 
            using (var stream = File.Open(exFileName, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                    {
                        ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                        {
                            EmptyColumnNamePrefix = "Column",
                            UseHeaderRow = true
                        }
                    });

                    var datatable = result.Tables[0];
                    Console.WriteLine("{0} seconds elapsed.(ExcelDataReader이용하여 Excel 파일 Read에 걸린 시간)", stopwatch.Elapsed.TotalSeconds);
                    Console.ReadLine();
                    for (var i = 1; i < datatable.Rows.Count; i++)
                    {
                        string s = Convert.ToDateTime(datatable.Rows[i][0]).ToString("yyyy/MM/dd");
                        Console.WriteLine(s);
                        
                    }
                    Console.ReadLine();
                }

            }

        }
    }
}
