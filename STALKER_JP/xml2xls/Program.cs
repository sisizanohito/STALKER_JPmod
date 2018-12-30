using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;
using ClosedXML.Excel;

namespace xml2xls
{
    class Program
    {
        static void Main(string[] args)
        {
            var workbook = new XLWorkbook();
            var topWorksheet = workbook.Worksheets.Add(@"日本語化");

            //保存元のファイル名
            string folderName = @"./eng/";

            System.IO.DirectoryInfo di = new System.IO.DirectoryInfo(folderName);
            System.IO.FileInfo[] files =
                di.GetFiles("*.xml", System.IO.SearchOption.AllDirectories);

            //ListBox1に結果を表示する
            foreach (System.IO.FileInfo f in files)
            {
                uiText data = ReadXML(f.FullName);
                string sheetName = NormalizeLength(f.Name.Replace(f.Extension, ""),31);
                var worksheet = workbook.Worksheets.Add(sheetName);
                worksheet.Cell("A1").Value = f.Name.Replace(f.Extension, "");

                //トップにシートのリンクを制作
                var topCell = topWorksheet.Cell(workbook.Worksheets.Count, 1);
                topCell.Value= f.Name.Replace(f.Extension, "");
                topCell.Hyperlink = new XLHyperlink($"'{sheetName}'!A1");

            }
            workbook.SaveAs(@"Stalker日本語化.xlsx");
            
        }

        static string NormalizeLength(string value, int maxLength)
        {
            if(value.Length > maxLength)
            {
                return value.Substring(0, maxLength);
            }
            return value;

        }

        static public uiText ReadXML(string fileName)
        {
            //XmlSerializerオブジェクトを作成
            XmlSerializer serializer =
                new XmlSerializer(typeof(uiText));

            //読み込むファイルを開く
            StreamReader sr = new StreamReader(fileName, Encoding.GetEncoding("windows-1251"));

            string str = sr.ReadToEnd();
            //Console.WriteLine(str);
            str = str.Replace("&", "[#]");
            //XMLファイルから読み込み、逆シリアル化する
            uiText obj = (uiText)serializer.Deserialize(new StringReader(str));
            //ファイルを閉じる
            sr.Close();

            return obj;
        }

    }

    [XmlRoot("string_table")]
    public class uiText
    {
        [XmlElement("string")]
        public List<Texts> UITexts { get; set; }
    }

    public class Texts
    {
        [XmlAttribute("id")]
        public string ID { get; set; }
        [XmlElement("text")]
        public string text { get; set; }
    }


}
