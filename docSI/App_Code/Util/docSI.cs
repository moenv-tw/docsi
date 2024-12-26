using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Xml;
using System.Xml.Serialization;

namespace Util.doc
{

    /// <summary>
    /// 公文流程Si檔案物件
    /// </summary>
    public class docSI
    {
        public docSI()
        {
            //
            // TODO: 在這裡新增建構函式邏輯
            //
        }

        public Util.doc.線上簽核 getClassData(string as_xmlPath)
        {
            線上簽核 lo_onlineSi = null;
            if (System.IO.File.Exists(as_xmlPath))
            {
                System.IO.StreamReader lo_streamReader = null;
                try
                {
                    XmlSerializer lo_xmlSerializer = new XmlSerializer(typeof(線上簽核));

                    lo_streamReader = new System.IO.StreamReader(as_xmlPath);

                    object lo_deserializationObj = lo_xmlSerializer.Deserialize(lo_streamReader);

                    lo_onlineSi = lo_deserializationObj as 線上簽核;
                }
                finally
                {
                    if (lo_streamReader != null)
                    {
                        lo_streamReader.Close();
                        lo_streamReader.Dispose();
                    }
                }
            }
            return lo_onlineSi;
         }
    }

    [XmlRoot(ElementName = "簽核流程")]
    public class 簽核流程
    {
        [XmlAttribute(AttributeName = "Id")]
        public string Id { get; set; }
        [XmlAttribute(AttributeName = "異動別")]
        public string 異動別 { get; set; }
    }

    [XmlRoot(ElementName = "線上簽核流程")]
    public class 線上簽核流程
    {
        [XmlElement(ElementName = "簽核流程")]
        public List<簽核流程> 簽核流程 { get; set; }
        [XmlAttribute(AttributeName = "Id")]
        public string Id { get; set; }
    }

    [XmlRoot(ElementName = "CanonicalizationMethod")]
    public class CanonicalizationMethod
    {
        [XmlAttribute(AttributeName = "Algorithm")]
        public string Algorithm { get; set; }
    }

    [XmlRoot(ElementName = "SignatureMethod")]
    public class SignatureMethod
    {
        [XmlAttribute(AttributeName = "Algorithm")]
        public string Algorithm { get; set; }
    }

    [XmlRoot(ElementName = "DigestMethod")]
    public class DigestMethod
    {
        [XmlAttribute(AttributeName = "Algorithm")]
        public string Algorithm { get; set; }
    }

    [XmlRoot(ElementName = "Reference")]
    public class Reference
    {
        [XmlElement(ElementName = "DigestMethod")]
        public DigestMethod DigestMethod { get; set; }
        [XmlElement(ElementName = "DigestValue")]
        public string DigestValue { get; set; }
        [XmlAttribute(AttributeName = "URI")]
        public string URI { get; set; }
    }

    [XmlRoot(ElementName = "SignedInfo")]
    public class SignedInfo
    {
        [XmlElement(ElementName = "CanonicalizationMethod")]
        public CanonicalizationMethod CanonicalizationMethod { get; set; }
        [XmlElement(ElementName = "SignatureMethod")]
        public SignatureMethod SignatureMethod { get; set; }
        [XmlElement(ElementName = "Reference")]
        public List<Reference> Reference { get; set; }
    }

    [XmlRoot(ElementName = "X509Data")]
    public class X509Data
    {
        [XmlElement(ElementName = "X509Certificate")]
        public string X509Certificate { get; set; }
    }

    [XmlRoot(ElementName = "KeyInfo")]
    public class KeyInfo
    {
        [XmlElement(ElementName = "X509Data")]
        public X509Data X509Data { get; set; }
    }

    [XmlRoot(ElementName = "Signature")]
    public class Signature
    {
        [XmlElement(ElementName = "SignedInfo")]
        public SignedInfo SignedInfo { get; set; }
        [XmlElement(ElementName = "SignatureValue")]
        public string SignatureValue { get; set; }
        [XmlElement(ElementName = "KeyInfo")]
        public KeyInfo KeyInfo { get; set; }
        [XmlAttribute(AttributeName = "Id")]
        public string Id { get; set; }
    }

    [XmlRoot(ElementName = "機關")]
    public class 機關
    {
        [XmlElement(ElementName = "全銜")]
        public string 全銜 { get; set; }
    }

    [XmlRoot(ElementName = "簽核人員")]
    public class 簽核人員
    {
        [XmlElement(ElementName = "機關")]
        public 機關 機關 { get; set; }
        [XmlElement(ElementName = "單位")]
        public string 單位 { get; set; }
        [XmlElement(ElementName = "職稱")]
        public string 職稱 { get; set; }
        [XmlElement(ElementName = "姓名")]
        public string 姓名 { get; set; }
        [XmlElement(ElementName = "帳號")]
        public string 帳號 { get; set; }
        [XmlElement(ElementName = "角色")]
        public string 角色 { get; set; }
    }

    [XmlRoot(ElementName = "次位簽核人員")]
    public class 次位簽核人員
    {
        [XmlElement(ElementName = "簽核人員")]
        public 簽核人員 簽核人員 { get; set; }
    }

    [XmlRoot(ElementName = "異動資訊")]
    public class 異動資訊
    {
        [XmlElement(ElementName = "簽核人員")]
        public 簽核人員 簽核人員 { get; set; }
        [XmlElement(ElementName = "異動別")]
        public string 異動別 { get; set; }
        [XmlElement(ElementName = "簽章時間")]
        public string 簽章時間 { get; set; }
        [XmlElement(ElementName = "次位簽核人員")]
        public 次位簽核人員 次位簽核人員 { get; set; }
        [XmlElement(ElementName = "簽核意見")]
        public 簽核意見 簽核意見 { get; set; }
    }

    [XmlRoot(ElementName = "電子檔格式附件")]
    public class 電子檔格式附件
    {
        [XmlAttribute(AttributeName = "原始檔序號")]
        public string 原始檔序號 { get; set; }
        [XmlAttribute(AttributeName = "會辦附件")]
        public string 會辦附件 { get; set; }
        [XmlAttribute(AttributeName = "參考附件")]
        public string 參考附件 { get; set; }
    }

    [XmlRoot(ElementName = "附件")]
    public class 附件
    {
        [XmlElement(ElementName = "名稱")]
        public string 名稱 { get; set; }
        [XmlElement(ElementName = "附件類型")]
        public string 附件類型 { get; set; }
        [XmlElement(ElementName = "電子檔格式附件")]
        public 電子檔格式附件 電子檔格式附件 { get; set; }
        [XmlAttribute(AttributeName = "文件夾識別碼")]
        public string 文件夾識別碼 { get; set; }
        [XmlAttribute(AttributeName = "物件識別碼")]
        public string 物件識別碼 { get; set; }
        [XmlAttribute(AttributeName = "格式")]
        public string 格式 { get; set; }
        [XmlAttribute(AttributeName = "序號")]
        public string 序號 { get; set; }
        [XmlAttribute(AttributeName = "產生時間")]
        public string 產生時間 { get; set; }
    }

    [XmlRoot(ElementName = "附件清單")]
    public class 附件清單
    {
        [XmlElement(ElementName = "附件")]
        public List<附件> 附件 { get; set; }
        [XmlAttribute(AttributeName = "附件數")]
        public string 附件數 { get; set; }
    }

    [XmlRoot(ElementName = "文稿")]
    public class 文稿
    {
        [XmlElement(ElementName = "名稱")]
        public string 名稱 { get; set; }
        [XmlElement(ElementName = "文稿類型")]
        public string 文稿類型 { get; set; }
        [XmlElement(ElementName = "附件清單")]
        public 附件清單 附件清單 { get; set; }
        [XmlAttribute(AttributeName = "文件夾識別碼")]
        public string 文件夾識別碼 { get; set; }
        [XmlAttribute(AttributeName = "物件識別碼")]
        public string 物件識別碼 { get; set; }
        [XmlAttribute(AttributeName = "類型")]
        public string 類型 { get; set; }
        [XmlAttribute(AttributeName = "產生點資訊")]
        public string 產生點資訊 { get; set; }
        [XmlAttribute(AttributeName = "原始檔序號")]
        public string 原始檔序號 { get; set; }
        [XmlAttribute(AttributeName = "序號")]
        public string 序號 { get; set; }
        [XmlAttribute(AttributeName = "產生時間")]
        public string 產生時間 { get; set; }
        [XmlElement(ElementName = "文稿頁面檔")]
        public 文稿頁面檔 文稿頁面檔 { get; set; }
    }

    [XmlRoot(ElementName = "簽核文稿清單")]
    public class 簽核文稿清單
    {
        [XmlElement(ElementName = "文稿")]
        public List<文稿> 文稿 { get; set; }
        [XmlAttribute(AttributeName = "文稿數")]
        public string 文稿數 { get; set; }
    }

    [XmlRoot(ElementName = "電子檔案資訊")]
    public class 電子檔案資訊
    {
        [XmlElement(ElementName = "檔案名稱")]
        public string 檔案名稱 { get; set; }
        [XmlElement(ElementName = "檔案大小")]
        public string 檔案大小 { get; set; }
        [XmlElement(ElementName = "檔案格式")]
        public string 檔案格式 { get; set; }
        [XmlAttribute(AttributeName = "文件夾識別碼")]
        public string 文件夾識別碼 { get; set; }
        [XmlAttribute(AttributeName = "物件識別碼")]
        public string 物件識別碼 { get; set; }
        [XmlAttribute(AttributeName = "原始檔序號")]
        public string 原始檔序號 { get; set; }
        [XmlAttribute(AttributeName = "序號")]
        public string 序號 { get; set; }
        [XmlAttribute(AttributeName = "產生時間")]
        public string 產生時間 { get; set; }
    }

    [XmlRoot(ElementName = "檔案清單")]
    public class 檔案清單
    {
        [XmlElement(ElementName = "電子檔案資訊")]
        public List<電子檔案資訊> 電子檔案資訊 { get; set; }
        [XmlAttribute(AttributeName = "檔案數")]
        public string 檔案數 { get; set; }
    }

    [XmlRoot(ElementName = "簽核文件夾")]
    public class 簽核文件夾
    {
        [XmlElement(ElementName = "簽核文稿清單")]
        public 簽核文稿清單 簽核文稿清單 { get; set; }
        [XmlElement(ElementName = "檔案清單")]
        public 檔案清單 檔案清單 { get; set; }
        [XmlAttribute(AttributeName = "Id")]
        public string Id { get; set; }
        [XmlAttribute(AttributeName = "產生時間")]
        public string 產生時間 { get; set; }
    }

    [XmlRoot(ElementName = "簽核資訊")]
    public class 簽核資訊
    {
        [XmlElement(ElementName = "簽核文件夾")]
        public 簽核文件夾 簽核文件夾 { get; set; }
    }

    [XmlRoot(ElementName = "Object")]
    public class Object
    {
        [XmlElement(ElementName = "異動資訊")]
        public 異動資訊 異動資訊 { get; set; }
        [XmlElement(ElementName = "簽核資訊")]
        public 簽核資訊 簽核資訊 { get; set; }
        [XmlAttribute(AttributeName = "Id")]
        public string Id { get; set; }
    }

    [XmlRoot(ElementName = "簽核點定義")]
    public class 簽核點定義
    {
        [XmlElement(ElementName = "Signature")]
        public Signature Signature { get; set; }
        [XmlElement(ElementName = "Object")]
        public Object Object { get; set; }
        [XmlAttribute(AttributeName = "URI")]
        public string URI { get; set; }
        [XmlAttribute(AttributeName = "Id")]
        public string Id { get; set; }
    }

    [XmlRoot(ElementName = "簽核意見")]
    public class 簽核意見
    {
        [XmlAttribute(AttributeName = "URI")]
        public string URI { get; set; }
        [XmlText]
        public string Text { get; set; }
    }

    [XmlRoot(ElementName = "頁面")]
    public class 頁面
    {
        [XmlAttribute(AttributeName = "文件夾識別碼")]
        public string 文件夾識別碼 { get; set; }
        [XmlAttribute(AttributeName = "物件識別碼")]
        public string 物件識別碼 { get; set; }
        [XmlAttribute(AttributeName = "產生點資訊")]
        public string 產生點資訊 { get; set; }
        [XmlAttribute(AttributeName = "原始檔序號")]
        public string 原始檔序號 { get; set; }
        [XmlAttribute(AttributeName = "序號")]
        public string 序號 { get; set; }
        [XmlAttribute(AttributeName = "產生時間")]
        public string 產生時間 { get; set; }
    }

    [XmlRoot(ElementName = "文稿頁面清單")]
    public class 文稿頁面清單
    {
        [XmlElement(ElementName = "頁面")]
        public 頁面 頁面 { get; set; }
        [XmlAttribute(AttributeName = "頁面數")]
        public string 頁面數 { get; set; }
    }

    [XmlRoot(ElementName = "文稿頁面檔")]
    public class 文稿頁面檔
    {
        [XmlElement(ElementName = "文稿頁面清單")]
        public 文稿頁面清單 文稿頁面清單 { get; set; }
        [XmlAttribute(AttributeName = "文件夾識別碼")]
        public string 文件夾識別碼 { get; set; }
        [XmlAttribute(AttributeName = "物件識別碼")]
        public string 物件識別碼 { get; set; }
        [XmlAttribute(AttributeName = "記錄方式")]
        public string 記錄方式 { get; set; }
        [XmlAttribute(AttributeName = "產生時間")]
        public string 產生時間 { get; set; }
    }

    [XmlRoot(ElementName = "線上簽核資訊")]
    public class 線上簽核資訊
    {
        [XmlElement(ElementName = "簽核點定義")]
        public List<簽核點定義> 簽核點定義 { get; set; }
        [XmlAttribute(AttributeName = "Id")]
        public string Id { get; set; }
    }

    [Serializable()]
    [XmlRoot(ElementName = "線上簽核")]
    public class 線上簽核
    {
        [XmlElement(ElementName = "線上簽核流程")]
        public 線上簽核流程 線上簽核流程 { get; set; }
        [XmlElement(ElementName = "線上簽核資訊")]
        public 線上簽核資訊 線上簽核資訊 { get; set; }
        [XmlAttribute(AttributeName = "Id")]
        public string Id { get; set; }
    }
}