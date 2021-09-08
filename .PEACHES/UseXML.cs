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
using System.Xml; // xml 읽어오는 기본 어셈블리
using System.IO;
XmlDocument Xml = new XmlDocument(); // xml 객체 선언
Xml.Load(xmlpath); //xml 파일 불러오기
XmlNodeList xmlList = Xml.SelectNodes("/Path1/Path2");

Console.WriteLine(Xml.SelectSingleNode("/Path1/Path2").InnerText); //console에 text 표시

// 반복문 활용하여 xml tag 읽어오기
foreach (XmlNode xmlElem in xmlList)
  {
    // winform의 Text Box등에 표시
    tbCheckStatus.Text += "파싱결과" + xmlElem["Tag1"].InnerText + " / " + xmlElem["Tag1"]["Tag2"].InnerText+Environment.NewLine;
    List.xmlTagList.Add(xmlElem["Tag1"].InnerText); //List에 추가
  }
