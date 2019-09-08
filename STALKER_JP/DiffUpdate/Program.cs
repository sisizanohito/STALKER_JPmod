using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;
using System.IO;
using ClosedXML.Excel;

namespace DiffUpdate
{
    public class Program
    {
        static void Main(string[] args)
        {
            XDB.XMLDB mDB = new XDB.XMLDB("./Stalker Anomaly日本語化.xlsx");

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
                string sheetName = NormalizeLength(f.Name.Replace(f.Extension, ""), 31);
                var worksheet = workbook.Worksheets.Add(sheetName);
                worksheet.Cell("A1").Value = f.Name.Replace(f.Extension, "");//ファイル名
                worksheet.Range("A1:G1").Merge();
                //ヘッダ
                worksheet.Cell("A2").Value = "ID";
                worksheet.Cell("B2").Value = "原文";
                worksheet.Cell("C2").Value = "表示";
                worksheet.Cell("D2").Value = "翻訳(こちらに記述してください)";
                worksheet.Cell("E2").Value = "コメント";
                worksheet.Cell("F2").Value = "前ver原文";
                worksheet.Cell("G2").Value = "前ver日本語";

                int cellRow = 3;
                //本文
                foreach (Texts texts in data.UITexts)
                {
                    worksheet.Cell(cellRow, 1).Value = texts.ID;
                    worksheet.Cell(cellRow, 2).Value = texts.TEXT;
                    worksheet.Cell(cellRow, 3).FormulaA1 = $"IF(EXACT(\"\",D{cellRow}),B{cellRow},D{cellRow})";
                    XDB.Texts textDB = mDB.db_text.UITexts.Find(n => n.ID == texts.ID);
                    if (textDB == null)
                    {
                        cellRow++;
                        continue;
                    }
                    else
                    {
                        worksheet.Cell(cellRow, 5).Value = textDB.COMMENT;
                        if (texts.TEXT== textDB.ENTEXT)
                        {
                            worksheet.Cell(cellRow, 4).Value = textDB.JPTEXT;
                        }
                        else
                        {
                            worksheet.Cell(cellRow, 6).Value = textDB.ENTEXT;
                            worksheet.Cell(cellRow, 7).Value = textDB.JPTEXT;
                        }
                        
                    }
                    
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
                worksheet.Column(6).Width = 30;//前原文
                worksheet.Column(7).Width = 40;//前日本語

                //トップにシートのリンクを制作
                var topCell = topWorksheet.Cell(workbook.Worksheets.Count, 1);
                topCell.Value = f.Name.Replace(f.Extension, "");
                topCell.Hyperlink = new XLHyperlink($"'{sheetName}'!A1");
                topWorksheet.Cell(workbook.Worksheets.Count, 2).FormulaA1 = $"1-(COUNTBLANK('{sheetName}'!D3:D{cellRow-1})/ROWS('{sheetName}'!D3:D{cellRow-1}))";
                topWorksheet.Cell(workbook.Worksheets.Count, 2).Style.NumberFormat.NumberFormatId = 10;//%
                topWorksheet.Cell(workbook.Worksheets.Count, 2).AddConditionalFormat().ColorScale()
                .LowestValue(XLColor.FromArgb(244,102,102))
                .Midpoint(XLCFContentType.Percent, 50, XLColor.FromArgb(255, 229, 153))
                .HighestValue(XLColor.FromArgb(182, 215, 168));
            }
            workbook.SaveAs("Stalker Anomaly日本語化2.xlsx");
        }

        static string NormalizeLength(string value, int maxLength)
        {
            if (value.Length > maxLength)
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
            public string TEXT { get; set; }
        }
    }

    
}
