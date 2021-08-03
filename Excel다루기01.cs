using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading.Tasks;
using System.Diagnostics;

// 엑셀 파일 여러개 생성하기 
// 파일 복사하기 연습
/// <summary>
/// 한폴더에 일련번호를 붙여서 여러개의 파일 생성(복사)
/// 경로는 고정
/// format 함수
/// Stopwatch class 활용하여 시간 측정하기 
/// </summary>

namespace ExcelCopy
{
    class Excopy01

    {
        
        static void Main(string[] args)
        {
            string sourcePath = @"C:\Users\kashm\OneDrive\Desktop\PYJ\Excel_Sample_for Csharp\01_Excel파일 복사하기"; // 원본 엑셀 파일 경로
            string File = "Sample.xlsx"; // 원본 엑셀 파일 명 
            string targetPath = @"C:\Users\kashm\OneDrive\Desktop\PYJ\Excel_Sample_for Csharp\01_Excel파일 복사하기\생성결과"; // 실행 결과 저장 경로
            // IO 함수를 이용하여 파일 경로 완성
            string rawFilePath = System.IO.Path.Combine(sourcePath, File); // 파일명과 파일 경로 합치기 
            int num = 500; // 복사할 파일 개수
            
            string suffix = "__WP__00.xlsx";
            if (System.IO.Directory.Exists(sourcePath) && System.IO.Directory.Exists(targetPath)) //원본 파일와 복사할 파일 경로가 존재할 경우 
            {
                Stopwatch stopwatch = new Stopwatch(); // 객체 선언
                stopwatch.Start();//시간측정 시작
                for (int i =0; i<num;i++)
                {
                    string newFileName;
                    Console.WriteLine("{0} 번째 파일 생성 시작", i + 1);
                    //숫자 포맷팅
                    string s = string.Format("{0:D4}", i+1);
                    s += suffix; // 일련번호와 suffix합치기 -> 0001__WP__00.xlsx
                    newFileName = System.IO.Path.Combine(targetPath, s);
                    // 파일 복사 하기 
                    System.IO.File.Copy(rawFilePath, newFileName);
                    Console.WriteLine("{0} 번째 파일 생성 완료", i + 1);
                }
                Console.WriteLine("{0}개 파일 생성 완료", num);
                stopwatch.Stop();//시간측정 끝
                Console.WriteLine("소요시간 {0} ms", stopwatch.ElapsedMilliseconds);
            }
        }
    }
}
