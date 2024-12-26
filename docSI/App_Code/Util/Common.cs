using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Web;
using System.Xml;
using System.Xml.Serialization;

using Newtonsoft.Json;


namespace Util
{
    /// <summary>
    /// 共用
    /// </summary>
    public class Common
    {
        public static string get(object s)
        {
            string ls_retValue = string.Empty;
            try
            {
                ls_retValue = Convert.ToString(s).Trim();
            }
            catch
            {
                ls_retValue = "";
            }
            return ls_retValue;
        }

        public static int getInt(object s)
        {
            int li_retValue = 0;
            if (s != null)
            {
                try
                {
                    string ls_value = s.ToString().Replace(",", "");
                    if (string.IsNullOrEmpty(ls_value) == false)
                    {
                        li_retValue = Convert.ToInt32(ls_value);
                    }
                }
                catch 
                {
                    li_retValue = 0;
                }
            }
            else
            {
                li_retValue = 0;
            }
            return li_retValue;
        }

        public static string EncodeBase64(string as_text)
        {
            byte[] lo_textBytes = System.Text.Encoding.UTF8.GetBytes(as_text);
            return System.Convert.ToBase64String(lo_textBytes);
        }

        public static string EncodeBase64Uri(string as_text)
        {
            string ls_text = EncodeBase64(as_text);

            //Uri.EscapeDataString = javaScript => encodeURIComponent     
            return System.Uri.EscapeDataString(ls_text);
        }

        public static string DecodeBase64(string as_base64EncodedData)
        {
            byte[] lo_base64EncodedBytes = System.Convert.FromBase64String(as_base64EncodedData);
            return System.Text.Encoding.UTF8.GetString(lo_base64EncodedBytes);
        }

        public static string DecodeBase64Uri(string as_base64EncodedData)
        {
            string ls_decodeText = as_base64EncodedData;
            try
            {
                string ls_text = System.Uri.UnescapeDataString(ls_decodeText);
                ls_decodeText = DecodeBase64(ls_text);
            }
            catch
            {
                ls_decodeText = string.Empty;
            }
            // Uri.UnescapeDataString == javaScript => decodeURIComponent
            return ls_decodeText;
        }


        public static DataTable JsonConvertToDataTable(string as_jsonString, string as_tableName = "")
        {
            DataTable lo_dt_table = null;
            if (string.IsNullOrEmpty(as_jsonString))
            {
                return lo_dt_table;
            }
            try
            {
                lo_dt_table = Newtonsoft.Json.JsonConvert.DeserializeObject<DataTable>(as_jsonString);
                if (lo_dt_table != null && string.IsNullOrEmpty(as_tableName) == false)
                {
                    lo_dt_table.TableName = as_tableName;
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            return lo_dt_table;
        }

        public static string hashTableToJson(Hashtable ao_ht)
        {
            if (ao_ht == null || ao_ht.Count == 0)
            {
                return null;
            }
            ArrayList lo_arrayList = new ArrayList();

            lo_arrayList.AddRange(ao_ht);

            string ls_jsonString = JsonConvert.SerializeObject(lo_arrayList);

            return ls_jsonString;
        }
    }
}