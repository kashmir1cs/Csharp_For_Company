using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EX = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Diagnostics;
// 엑셀파일 Parsing하기
/// <summary>
/// 1. 특정 셀에 있는 데이터 읽기
/// 2. 반복문 활용하여 셀 값 출력
/// 2. UsedRange 등 특수한 셀 선택 방법
/// 3. Rows / Columns 속성
/// 4. 파싱 결과 Console 창에 출력하기
/// 참고사이트 : https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel.range?view=excel-pia
/// </summary>
namespace ExcelParsing
{
    class ExParser
    {
        // Sheet1 = Basic
        // Sheet2 = Advanced

        static void Main(string[] args)
        {
            // Cells 속성 + 반복문 활용 
            Stopwatch stopwatch = new Stopwatch(); // 객체 선언
            stopwatch.Start();
            // Excel 파일 경로 정보 
            string xlpath1 = @"C:\Excel_Sample_for Csharp\03_Excel파일 Parsing";
            string xlfile1 = "sample2.xlsx";
            string samplefile1 = System.IO.Path.Combine(xlpath1, xlfile1);

            // Excel 객체  생성
            EX.Application app1 = new EX.Application(); // 엑셀 개체 생성하기
            EX.Workbook wb1 = null; // 워크북 개체 생성
            EX.Worksheet ws1 = null; // 워크시트 생성
            EX.Worksheet ws2 = null; // 워크시트 생성
            wb1 = app1.Workbooks.Open(Filename:samplefile1,ReadOnly:true); ; //파일 열어서 객체 초기화 하기 

            int cntSheet = wb1.Worksheets.Count; //시트 숫자 확인
            Console.WriteLine("{0}개 시트 있음",cntSheet);
            // 워크 시트 객체에 워크북의 객체 할당
            // 번호/ 이름 둘다 지정 가능  
            ws1 = wb1.Worksheets.Item[1];
            ws2 = wb1.Worksheets.Item["Basic"];
            string Worksheetname1 = ws1.Name;
            string Worksheetname2 = ws2.Name;
            
            Console.WriteLine("Sheet 이름 확인하기 ");
            Console.WriteLine("1번시트 : {0}", Worksheetname1);
            Console.WriteLine("2번시트 : {0}", Worksheetname2);
            // 반복문 활용하여 시트이름 출력하기 
            Console.WriteLine("반복문을 활용하여 Sheet이름 꺼내기");
            for (int i=1; i<cntSheet+1;i++)
            {
                Console.WriteLine("{0}번시트 : {1}",i, wb1.Worksheets.Item[i].Name);
                
            }
            // UsedRange
            
            Console.WriteLine("1번시트 이름  {0}",  Worksheetname1);
            Console.WriteLine("UsedRange {0}행 {1}열", ws1.UsedRange.Rows.Count, ws1.UsedRange.Columns.Count);

            Console.WriteLine("1번시트 이름  {0}", Worksheetname1);
            Console.WriteLine("A1셀의 CurrentRegion {0}행 {1}열", ws1.Range["A1"].CurrentRegion.Rows.Count, ws1.Range["A1"].CurrentRegion.Columns.Count);

            // 특수한 셀 선택
            // 예시 : .End(xlUp)
            // Offset(*,*) -> Offset[*,*]
            Console.WriteLine("A1셀에서 xlDown을 하면 {0}행으로 이동", ws1.Range["A1"].End[EX.XlDirection.xlDown].Offset[1,0].Row-1);
            Console.WriteLine("B1셀에서 xlDown을 하면 {0}행으로 이동", ws1.Range["B1"].End[EX.XlDirection.xlDown].Offset[1, 0].Row-1);
            Console.WriteLine("C1셀에서 xlDown을 하면 {0}행으로 이동", ws1.Range["C1"].End[EX.XlDirection.xlDown].Offset[1, 0].Row-1);
            Console.WriteLine("A1셀에서 xlRight를 하면 {0}번째 열 ", ws1.Range["A1"].End[EX.XlDirection.xlToRight].Column);
            Console.WriteLine("A1셀에서 xlRight로 이동한 셀의 값은 \"{0}\"", ws1.Range["A1"].End[EX.XlDirection.xlToRight].Value);
            stopwatch.Stop();
            Console.WriteLine("소요시간 {0} ms", stopwatch.ElapsedMilliseconds);
            wb1.Close();
            app1.Quit();

        }
    }
}
