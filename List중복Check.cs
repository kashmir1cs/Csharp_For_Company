using System;
using System.Collections.Generic;
using System.Linq;


namespace Duplicate확인
{
    class Program
    {
        static void Main(string[] args)
        {
            List<string> test = new List<string>();
            test.Add("A");
            test.Add("B");
            test.Add("B");
            test.Add("C");
            test.Add("C");
            test.Add("C");
            test.Add("D");
            test.Add("D");
            test.Add("E");
            test.Add("E");
            test.Add("F");
            test.Add("F");
            test.Add("F");
            test.Add("F");
            test.Add("F");
            test.Add("F");
            test.Add("G");
            var dictTestCount = test.GroupBy(x => x).Where(g => g.Count() > 1).ToDictionary(y => y.Key, x => x.Count());
            foreach(KeyValuePair<string,int> items in dictTestCount)
            {
                Console.WriteLine("{0} : {1}회 중복입력", items.Key, items.Value);
            }

            Console.WriteLine("출력완료");
            Console.ReadKey();

        }
    }
}
