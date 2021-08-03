using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EX = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Diagnostics;
// Excel 다루기
// 엑셀파일 생성 C#에서 접근 및 제어하기
/// <summary>
/// 0. 엑셀 기초 연습
/// 1. 특정 셀에 값 넣기
/// 2. 반복문
/// 3. 열수/ 행수 구하기
/// 관련 링크 : https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel.range?redirectedfrom=MSDN&view=excel-pia
/// </summary>
namespace Excel활용
{
    class Program
    {

        static void Main(string[] args)
        {
            // Cells 속성 + 반복문 활용 
            Stopwatch stopwatch1 = new Stopwatch(); // 객체 선언
            stopwatch1.Start();
            string xlpath1 = @"C:\Users\kashm\OneDrive\Desktop\PYJ\Excel_Sample_for Csharp\02_Excel다루기_기초";
            string xlfile1 = "sample1.xlsx";
            string samplefile1 = System.IO.Path.Combine(xlpath1, xlfile1);
            // 0. 엑셀 다루기 기초 문법 
            EX.Application app = new EX.Application(); // 엑셀 개체 생성하기
            EX.Workbook wb = null; // 워크북 개체 생성
            EX.Worksheet ws = null; // 워크시트 생성

            // 엑셀파일 생성
            wb = app.Workbooks.Add();
            ws = wb.Worksheets.get_Item(1) as EX.Worksheet;
            // 엑셀 열 크기 조정
            ws.Columns.ColumnWidth += 10;
            // 가운데 정렬로 바꿈
            ws.Cells.Style.HorizontalAlignment = EX.XlHAlign.xlHAlignCenter;
            // 반복문을 활용하여 엑셀 셀에 값 입력하기 
            // cells 개체 활용
            for (int i = 1; i < 201; i++)
            {
                for (int j = 1; j < 151; j++)
                {
                    Console.WriteLine("Cell[{0},{1}]입력 중", i, j);
                    string a = string.Format("{0:D4}", i);
                    a += "__";
                    a += string.Format("{0:D6}", j);
                    ws.Cells[i, j] = a;

                }
            }
            // 파일 저장하기 
            Console.WriteLine("파일저장");
            stopwatch1.Stop();//시간측정 끝
            Console.WriteLine("소요시간 {0} ms", stopwatch1.ElapsedMilliseconds);
            wb.SaveAs(samplefile1);

            wb.Close();
            wb = null;// 변수 초기화 

            Stopwatch stopwatch = new Stopwatch(); // 객체 선언
            stopwatch.Start();
            string xlpath2 = @"C:\Users\kashm\OneDrive\Desktop\PYJ\Excel_Sample_for Csharp\02_Excel다루기_기초";
            string xlfile2 = "sample2.xlsx";
            string samplefile2 = System.IO.Path.Combine(xlpath2, xlfile2);
            // 0. 엑셀 다루기 기초 문법 
            EX.Application app1 = new EX.Application(); // 엑셀 개체 생성하기
            EX.Workbook wb1 = null; // 워크북 개체 생성
            EX.Worksheet ws1 = null; // 워크시트 생성

            // 엑셀파일 생성
            wb1 = app.Workbooks.Add();
            ws1 = wb1.Worksheets.get_Item(1) as EX.Worksheet;
            // 엑셀 열 크기 조정
            ws1.Columns.ColumnWidth += 10;
            // 가운데 정렬로 바꿈
            ws1.Cells.Style.HorizontalAlignment = EX.XlHAlign.xlHAlignCenter;
            // 반복문을 활용하여 엑셀 셀에 값 입력하기 
            // Range 개체 활용
            // 행/열 관련된 Property 활용
            for (int i = 1; i < 201; i++)
            {
                Console.WriteLine("{0}번째 행 입력중 ", i);
                string c = i.ToString();
                string n = string.Format("{0:D4}", i);
                n += "-A";
                ws1.Range["A"+c].Value2 = n;
                n += "-B";
                ws1.Range["E" + c].Value2 = n;

            }
            // 파일 저장하기 
            Console.WriteLine("파일저장");
            stopwatch.Stop();//시간측정 끝
            Console.WriteLine("소요시간 {0} ms", stopwatch.ElapsedMilliseconds);
            wb1.SaveAs(samplefile2);

            wb1.Close();
        }
    }
}
