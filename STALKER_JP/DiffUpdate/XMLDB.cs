using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;
using System.Xml.Serialization;
using System.IO;

namespace XDB
{
    class XMLDB
    {
        public TEXTDB db_text = new TEXTDB();
        public XMLDB(string filename)
        {
            db_text.UITexts = new List<Texts>();
            using (var workbook = new XLWorkbook(filename))
            {
                foreach (var worksheet in workbook.Worksheets)
                {
                    
                    
                    if (!worksheet.Name.Equals("トップ"))//トップを除く
                    {
                        var rowCount = worksheet.RangeUsed().RowCount();
                        for (int i = 3; i <= rowCount; i++)
                        {
                            var id = worksheet.Cell(i, 1).Value.ToString();
                            var en = worksheet.Cell(i, 2).Value.ToString();
                            var jp = worksheet.Cell(i, 4).Value.ToString();
                            var comment = worksheet.Cell(i, 5).Value.ToString();
                            db_text.UITexts.Add(new Texts
                            {
                                ID = id,
                                ENTEXT = en,
                                JPTEXT = jp,
                                COMMENT = comment
                            });
                            //Console.WriteLine("{0}:{1}",id,text);
                        }
                        Console.WriteLine(worksheet.Name);
                    }
                    else
                    {
                        continue;
                    }
                }
            }
        }
    }

    public class TEXTDB
    {
        public List<Texts> UITexts { get; set; }
    }

    public class Texts
    {  
        public string ID { get; set; }
        public string ENTEXT { get; set; }
        public string JPTEXT { get; set; }
        public string COMMENT { get; set; }
    }
}
