using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GetMonitorWidthAndHeight
{
    class Program
    {
        // 모니터 해상도 구하기
        // 참조 추가 : System.Drawing, System.Drawing.Design
        
        static void Main(string[] args)
        {
            int x = Screen.PrimaryScreen.Bounds.Width; // 주모니터의 가로 해상도 
            int y = Screen.PrimaryScreen.Bounds.Height; // 주모니터의 세로 해상도 
            int e = Screen.AllScreens.Length; // 모니터 개수 구하기

            Console.WriteLine("Screen.PrimaryScreen.Bounds.Width / Screen.PrimaryScreen.Bounds.Height /Screen.AllScreens.Length 사용 ");
            Console.WriteLine("모니터의 해상도는 가로 : {0} / 세로 : {1} 그리고 모니터 대수는 : {2}", x,y,e);

            x = SystemInformation.VirtualScreen.Width;
            y = SystemInformation.VirtualScreen.Height;

            Console.WriteLine("SystemInformation.VirtualScreen.Width / SystemInformation.VirtualScreen.Height 사용");
            Console.WriteLine("모니터의 해상도는 가로 : {0} / 세로 : {1} 그리고 모니터 대수는 : {2}", x, y, e);

        }
    }
}
