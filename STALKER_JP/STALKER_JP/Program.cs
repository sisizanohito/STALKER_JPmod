using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace STALKER_JP
{
    class Program
    {
        static void Main(string[] args)
        {
            using (var workbook = new XLWorkbook("./Stalker Anomaly日本語化.xlsx"))
            {
                foreach (var worksheet in workbook.Worksheets)
                {
                    uiText ui_text = new uiText();
                    ui_text.UITexts = new List<Texts>();
                    if (!worksheet.Name.Equals("トップ"))
                    {
                        var rowCount = worksheet.RangeUsed().RowCount();
                        for (int i = 3; i <= rowCount; i++)
                        {
                            var id = worksheet.Cell(i, 1).Value.ToString();
                            var text = worksheet.Cell(i, 3).Value.ToString();
                            ui_text.UITexts.Add(new Texts {
                                ID = id,
                                TEXT = text
                            });
                            //Console.WriteLine("{0}:{1}",id,text);
                        }
                        Console.WriteLine(worksheet.Name);
                    }
                    else
                    {
                        continue;
                    }

                    XmlSerializerNamespaces ns = new XmlSerializerNamespaces();
                    ns.Add(String.Empty, String.Empty);
                    var writer = new StringWriter(); // 出力先のWriterを定義
                    var serializer = new XmlSerializer(typeof(uiText)); // Bookクラスのシリアライザを定義
                    serializer.Serialize(writer, ui_text, ns);
                    var xml = writer.ToString();
                    xml = xml.Replace("[#]", "&");

                    string filename = worksheet.Cell("A1").Value.ToString();
                    File.WriteAllText($@"./gamedata/configs/text/jpn/{filename}.xml", xml);
                }
                Console.WriteLine("終了-何かキーを押してください");
                Console.ReadKey();
            }
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
        public string TEXT { get; set; }
    }
}
