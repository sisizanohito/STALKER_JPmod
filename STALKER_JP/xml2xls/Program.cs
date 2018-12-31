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
            var topWorksheet = workbook.Worksheets.Add("トップ");
            topWorksheet.Cell("A1").Value = "ファイル";
            topWorksheet.Cell("B1").Value = "進捗";

            //保存元のファイル名
            string folderName = @"./eng/";

            DirectoryInfo di = new DirectoryInfo(folderName);
           FileInfo[] files =
                di.GetFiles("*.xml", SearchOption.AllDirectories);

            foreach (FileInfo f in files)
            {
                uiText data = ReadXML(f.FullName);
                string sheetName = NormalizeLength(f.Name.Replace(f.Extension, ""),31);
                var worksheet = workbook.Worksheets.Add(sheetName);
                worksheet.Cell("A1").Value = f.Name.Replace(f.Extension, "");//ファイル名
                worksheet.Range("A1:F1").Merge();
                //ヘッダ
                worksheet.Cell("A2").Value = "ID";
                worksheet.Cell("B2").Value = "原文";
                worksheet.Cell("C2").Value = "表示";
                worksheet.Cell("D2").Value = "翻訳(こちらに記述してください)";
                worksheet.Cell("E2").Value = "コメント";
                worksheet.Cell("F2").Value = "機械翻訳";

                int cellRow = 3;
                //本文
                foreach(Texts texts in data.UITexts)
                {
                    worksheet.Cell(cellRow, 1).Value = texts.ID;
                    worksheet.Cell(cellRow, 2).Value = texts.text;
                    worksheet.Cell(cellRow, 3).FormulaA1 = $"IF(EXACT(\"\",D{cellRow}),IF(EXACT(\"\",F{cellRow}),B{cellRow},F{cellRow}),D{cellRow})";
                    cellRow++;
                }
               

                //スタイル
                worksheet.SheetView.FreezeRows(1);
                worksheet.SheetView.FreezeRows(2);
                worksheet.RangeUsed().Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                worksheet.RangeUsed().Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                worksheet.RangeUsed().Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
                worksheet.RangeUsed().Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.RangeUsed().Style.Alignment.WrapText = true;
                worksheet.Column(1).Width = 20;//ID
                worksheet.Column(2).Width = 30;//原文
                worksheet.Column(3).Width = 40;//表示
                worksheet.Column(4).Width = 40;//翻訳
                worksheet.Column(5).Width = 10;//コメント
                worksheet.Column(6).Width = 5;//機械翻訳

                //トップにシートのリンクを制作
                var topCell = topWorksheet.Cell(workbook.Worksheets.Count, 1);
                topCell.Value= f.Name.Replace(f.Extension, "");
                topCell.Hyperlink = new XLHyperlink($"'{sheetName}'!A1");
                topWorksheet.Cell(workbook.Worksheets.Count, 2).FormulaA1 = $"1-(COUNTBLANK('{sheetName}'!D3:D)/ROWS('{sheetName}'!D3:D))";
            }
            workbook.SaveAs("Stalker Anomaly日本語化.xlsx");
            
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
