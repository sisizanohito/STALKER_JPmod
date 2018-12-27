using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace xml2xls
{
    class Program
    {
        static void Main(string[] args)
        {

            //保存元のファイル名
            string fileName = @"";

            //XmlSerializerオブジェクトを作成
            System.Xml.Serialization.XmlSerializer serializer =
                new System.Xml.Serialization.XmlSerializer(typeof(uiText));

            //読み込むファイルを開く
            System.IO.StreamReader sr = new System.IO.StreamReader(fileName, new System.Text.UTF8Encoding(false));
            //XMLファイルから読み込み、逆シリアル化する
            uiText obj = (uiText)serializer.Deserialize(sr);
            //ファイルを閉じる
            sr.Close();
        }
    }

    [XmlRoot("string_table")]
    public class uiText
    {
        [XmlElement("string")]
        public List<Texts> uiTexts { get; set; }
    }

    public class Texts
    {
        [XmlAttribute("id")]
        public string ID { get; set; }
        [XmlElement("text")]
        public string text { get; set; }
    }


}
