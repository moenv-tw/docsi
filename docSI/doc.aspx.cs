using System;
using System.Collections;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO.Compression;
using System.Runtime.CompilerServices;
using System.Text;
using System.Web;
using System.Web.DynamicData;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;

public partial class doc : System.Web.UI.Page
{
    #region 檔案資料
    /// <summary>
    /// 暫存檔案路徑
    /// </summary>
    string is_tempDefault = "";
    /// <summary>
    /// 公文檔案預設路徑
    /// </summary>
    string is_docLocation = "";
    /// <summary>
    /// 壓縮檔案的文件目錄
    /// </summary>
    string is_dirPath = "";
    /// <summary>
    /// 壓縮檔案的完整路徑
    /// </summary>
    string is_zipfilePath = "";

    /// <summary>
    /// doc_SI_附件資訊
    /// </summary>
    private System.Collections.Hashtable lo_ht_attachment = new Hashtable();
    /// <summary>
    /// doc_SI_電子檔案資訊
    /// </summary>
    private System.Collections.Hashtable lo_ht_electronicFile = new Hashtable();
    /// <summary>
    /// 電子檔案與附件資訊對應
    /// </summary>
    private System.Collections.Hashtable lo_ht_electronicList = new Hashtable();
    /// <summary>
    /// 顯示副檔名
    /// </summary>
    private System.Collections.Hashtable lo_ht_Extension = new Hashtable();
    #endregion

    protected void Page_Load(object sender, EventArgs e)
    {
        init();

        #region repeaterItems
        CollectionPager1.Visible = true;
        CollectionPager1.DataSource = queryAll();
        CollectionPager1.BindToControl = repeaterItems;
        repeaterItems.DataSource = CollectionPager1.DataSourcePaged;
        repeaterItems.DataBind();
        #endregion

        #region download
        switch (state.Value)
        {
            case "popup_pdf":
                state.Value = "";
                doJavaScript_download();
                break;
            case "download":
                state.Value = "";
                download(fileEntry.Value);
                break;
        }
        #endregion

        deleteTempDir();
    }

    /// <summary>
    /// 設定初始值資料
    /// </summary>
    private void init()
    {
        string ls_applicationPath = System.Web.HttpContext.Current.Request.ApplicationPath;
        string ls_fullPath = System.Web.HttpContext.Current.Request.MapPath(ls_applicationPath);
        //string is_docLocation = Util.Common.get(ConfigurationManager.AppSettings["DOCLocation"]);
        string ls_docLocation = System.IO.Path.Combine(ls_fullPath, "temp", "doc");

        string ls_tempDefault = System.IO.Path.Combine(ls_fullPath, "temp");
        string ls_guid = System.Guid.NewGuid().ToString("N");
        //string is_tempDefault = Util.Common.get(ConfigurationManager.AppSettings["TempFilePath"]);
        tempDir.Value = Util.Common.EncodeBase64Uri(System.IO.Path.Combine(ls_tempDefault, ls_guid));

        #region 讀取SI檔案的副檔名
        lo_ht_Extension.Add(".pdf", ".pdf");
        lo_ht_Extension.Add(".xls", ".xls");
        lo_ht_Extension.Add(".xlsx", ".xlsx");
        lo_ht_Extension.Add(".doc", ".doc");
        lo_ht_Extension.Add(".docx", ".docx");
        #endregion

        //設定zip的路徑
        string ls_docNo = docNo.Value;
        if (string.IsNullOrEmpty(ls_docNo) == false)
        {
            is_dirPath = System.IO.Path.Combine(ls_docLocation, ls_docNo);

            is_zipfilePath = System.IO.Path.Combine(is_dirPath, ls_docNo + ".zip");
        }
    }
    /// <summary>
    /// 目的：依查詢欄位查詢多筆資料
    /// </summary>
    /// <returns></returns>
    protected System.Data.DataView queryAll()
    {
        #region dt
        System.Data.DataTable lo_dt = new DataTable("doc");
        System.Data.DataColumn lo_dc_rid = new DataColumn("rid", typeof(string));
        lo_dc_rid.Caption = "系統編號";
        System.Data.DataColumn lo_dc_fileName = new DataColumn("filaName", typeof(string));
        lo_dc_fileName.Caption = "檔案名稱";
        System.Data.DataColumn lo_dc_fileMemo = new DataColumn("fileMemo", typeof(string));
        lo_dc_fileMemo.Caption = "原始附件檔案名稱";
        System.Data.DataColumn lo_dc_filelength = new DataColumn("filelength", typeof(string));
        lo_dc_fileName.Caption = "檔案長度";
        System.Data.DataColumn lo_dc_lastWriteTime = new DataColumn("lastWriteTime", typeof(string));
        lo_dc_fileName.Caption = "最後異動日期";

        lo_dt.Columns.AddRange(new DataColumn[] 
                                  {
                                       lo_dc_rid
                                      ,lo_dc_fileName
                                      ,lo_dc_fileMemo
                                      ,lo_dc_filelength
                                      ,lo_dc_lastWriteTime
                                  }
                              );
        #endregion
        lo_dt.Rows.Clear();

        string ls_docNo = docNo.Value;
        
        if (string.IsNullOrEmpty(ls_docNo) == false)
        {
            string ls_fileMemo = "";
            
            if (System.IO.File.Exists(is_zipfilePath))
            {
                #region fileProcess
                lb_text.InnerText = "";
                System.IO.FileInfo lo_fileInfo = new System.IO.FileInfo(is_zipfilePath);
                try
                {
                    #region 讀取公文SI檔案,如有資料則直接讀取json資料以避免效能負擔
                    if (string.IsNullOrEmpty(si_json.Value))
                    {
                        read_doc_SI();
                    }
                    else
                    {
                        string ls_jsonString = Util.Common.DecodeBase64Uri(si_json.Value);
                        if (string.IsNullOrEmpty(ls_jsonString) == false)
                        {
                            System.Data.DataTable lo_dt_si = Util.Common.JsonConvertToDataTable(ls_jsonString);
                            if (lo_dt_si != null && lo_dt_si.Rows.Count > 0)
                            {
                                lo_ht_electronicList = new Hashtable();
                                foreach (System.Data.DataRow o_row in lo_dt_si.Rows)
                                {
                                    lo_ht_electronicList.Add(o_row[0], o_row[1]);
                                }
                            }
                        }
                    }
                    #endregion

                    #region 讀取zip
                    ZipArchive lo_zipArchive = System.IO.Compression.ZipFile.OpenRead(lo_fileInfo.FullName);
                    foreach (ZipArchiveEntry lo_entry in lo_zipArchive.Entries)
                    {
                        string ls_fileName = Util.Common.get(lo_entry.FullName);

                        ls_fileMemo = "";

                        bool lb_fund = false;
                        string ls_search_name = ls_fileName;
                        do
                        {
                            if (lo_ht_electronicList.ContainsKey(ls_search_name))
                            {
                                ls_fileMemo = Util.Common.get(lo_ht_electronicList[ls_search_name]);
                                ls_search_name = ls_fileMemo;
                                lb_fund = true;
                            }
                            else
                            {
                                lb_fund = false;
                            }
                        } while (lb_fund);

                        
                        string ls_fileExtension = System.IO.Path.GetExtension(ls_fileName).ToLower();
                        if (lo_ht_Extension.ContainsKey(ls_fileExtension))
                        {
                            System.Data.DataRow lo_row = lo_dt.NewRow();
                            lo_row["rid"] = ls_fileName;
                            lo_row["filaName"] = ls_fileName;
                            lo_row["fileMemo"] = ls_fileMemo;
                            lo_row["filelength"] = lo_entry.Length.ToString("###,##0");
                            lo_row["lastWriteTime"] = Util.Common.get(lo_entry.LastWriteTime.DateTime);

                            lo_dt.Rows.Add(lo_row);
                        }
                    }
                    #endregion
                }
                catch (Exception ex)
                {
                    lb_text.InnerText = "錯誤訊息:" + ex.Message;
                }
                #endregion
            }
            else
            {
                lb_text.InnerText = "檔案不存在，請確認！";
            }
        }

        if (string.IsNullOrEmpty(lb_text.InnerText))
        {
            tr_showMsg.Style["display"] = "none";
        }
        else
        {
            tr_showMsg.Style["display"] = "";
        }

        #region ls_attachment mark
        //string ls_attachment = attachment.Value;
        //if (string.IsNullOrEmpty(ls_attachment) == false)
        //{
        //    #region attachment

        //    //抓取公文文號目錄
        //    string ls_dirPath = System.IO.Path.Combine(is_docLocation, ls_docNo, ls_attachment);

        //    if (System.IO.Directory.Exists(ls_dirPath))
        //    {
        //        System.IO.DirectoryInfo lo_dirInfo = new System.IO.DirectoryInfo(ls_dirPath);
        //        System.IO.FileInfo[] lo_fileInfo = lo_dirInfo.GetFiles();
        //        foreach (System.IO.FileInfo o_file in lo_fileInfo)
        //        {
        //            System.Data.DataRow lo_row = lo_dt.NewRow();
        //            lo_row["filaName"] = Util.Common.get(o_file.Name);
        //            lo_row["fileMemo"] = "";
        //            lo_row["filelength"] = o_file.Length.ToString("###,##0");
        //            lo_row["lastWriteTime"] = Util.Common.get(o_file.LastWriteTime);

        //            lo_dt.Rows.Add(lo_row);
        //        }
        //    }
        //    #endregion
        //}
        #endregion
        

        return lo_dt.DefaultView;
    }

    /// <summary>
    /// 下載檔案
    /// </summary>
    private void zipDownloadFile()
    {
        try
        {
            string ls_docNo = docNo.Value;
            if (string.IsNullOrEmpty(ls_docNo) == false)
            {
                if (System.IO.File.Exists(is_zipfilePath))
                {
                    //Util.Common.doJavaScript_download(this, is_zipfilePath, "N");
                }
                else
                {
                    lb_text.InnerText = "檔案不存在，請確認！";
                }
            }
            else
            {
                lb_text.InnerText = "檔案資料為空白，請確認！";
            }
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
    }
    
    /// <summary>
    /// 讀取公文系統SI檔案
    /// </summary>
    /// <returns></returns>
    private bool read_doc_SI()
    {
        bool lb_bool = false;
        if (System.IO.File.Exists(is_zipfilePath))
        {
            string ls_TempfileFullPath = "";

            #region 1.將SI解壓縮到temp
            System.IO.FileInfo lo_zip_fileInfo = new System.IO.FileInfo(is_zipfilePath);
            if (lo_zip_fileInfo.Exists)
            {
                ZipArchive lo_zipArchive = System.IO.Compression.ZipFile.OpenRead(lo_zip_fileInfo.FullName);
                foreach (ZipArchiveEntry lo_entry in lo_zipArchive.Entries)
                {
                    System.IO.FileInfo lo_entryFileInfo = new System.IO.FileInfo(lo_entry.FullName);
                    if (lo_entryFileInfo.Extension.ToLower() == ".si")
                    {
                        string ls_tempDirPath = Util.Common.DecodeBase64Uri(tempDir.Value);
                        System.IO.DirectoryInfo lo_tempDirInfo = new System.IO.DirectoryInfo(ls_tempDirPath);
                        if (lo_tempDirInfo.Exists == false)
                        {
                            lo_tempDirInfo.Create();
                        }
                        ls_TempfileFullPath = System.IO.Path.Combine(lo_tempDirInfo.FullName, lo_entry.FullName);
                        lo_entry.ExtractToFile(ls_TempfileFullPath);
                        
                        break;
                    }
                }
                lo_zipArchive.Dispose();
            }
            #endregion

            #region 2.讀取SI的檔案到HashTable
            System.IO.FileInfo lo_tempfileInfo = new System.IO.FileInfo(ls_TempfileFullPath);
            if (lo_tempfileInfo.Exists)
            {
                #region read process
                Util.doc.docSI lo_doc_si = new Util.doc.docSI();
                Util.doc.線上簽核 lo_onlineSign = lo_doc_si.getClassData(lo_tempfileInfo.FullName);
                if (lo_onlineSign != null)
                {
                    string ls_objectId = "";
                    string ls_modifykind = "";

                    foreach (Util.doc.簽核點定義 lo_data in lo_onlineSign.線上簽核資訊.簽核點定義)
                    {
                        if (lo_data != null &&
                            lo_data.Object != null &&
                            lo_data.Object.簽核資訊 != null &&
                            lo_data.Object.簽核資訊.簽核文件夾 != null &&
                            lo_data.Object.簽核資訊.簽核文件夾.簽核文稿清單 != null)
                        {
                            lb_bool = true;
                            ls_objectId = lo_data.Object.Id;
                            ls_modifykind = lo_data.Object.異動資訊.異動別;

                            #region 文稿
                            //foreach (Util.doc.文稿 lo_signDoc in lo_data.Object.簽核資訊.簽核文件夾.簽核文稿清單.文稿)
                            //{
                            //    if (lo_signDoc != null &&
                            //        lo_signDoc.附件清單 != null &&
                            //        Util.Common.getInt(lo_signDoc.附件清單.附件數) > 0)
                            //    {
                            //        foreach (Util.doc.附件 lo_signDoc_attachment in lo_signDoc.附件清單.附件)
                            //        {
                            //            if (lo_ht_attachment.ContainsKey(lo_signDoc_attachment.電子檔格式附件.原始檔序號) == false)
                            //            {
                            //                lo_ht_attachment.Add(lo_signDoc_attachment.電子檔格式附件.原始檔序號, lo_signDoc_attachment.名稱);
                            //            }
                            //        }
                            //    }
                            //}
                            foreach (Util.doc.文稿 lo_signDoc in lo_data.Object.簽核資訊.簽核文件夾.簽核文稿清單.文稿)
                            {
                                #region 文稿
                                if (Util.Common.getInt(lo_signDoc.原始檔序號) > 0)
                                {
                                    if (lo_ht_attachment.ContainsKey(lo_signDoc.原始檔序號) == false)
                                    {
                                        lo_ht_attachment.Add(lo_signDoc.原始檔序號, lo_signDoc.類型);
                                    }
                                }

                                if (lo_signDoc.文稿頁面檔 != null &&
                                    lo_signDoc.文稿頁面檔.文稿頁面清單 != null &&
                                    lo_signDoc.文稿頁面檔.文稿頁面清單.頁面 != null &&
                                    Util.Common.getInt(lo_signDoc.文稿頁面檔.文稿頁面清單.頁面.原始檔序號) > 0)
                                {
                                    if (lo_ht_attachment.ContainsKey(lo_signDoc.文稿頁面檔.文稿頁面清單.頁面.原始檔序號) == false)
                                    {
                                        lo_ht_attachment.Add(lo_signDoc.文稿頁面檔.文稿頁面清單.頁面.原始檔序號, lo_signDoc.類型 + "_公文稿");
                                    }
                                }
                                #endregion


                                if (lo_signDoc.附件清單 != null &&
                                    Util.Common.getInt(lo_signDoc.附件清單.附件數) > 0)
                                {
                                    //附件清單
                                    foreach (Util.doc.附件 lo_signDoc_attachment in lo_signDoc.附件清單.附件)
                                    {
                                        switch (lo_signDoc.文稿類型)
                                        {
                                            case "簽":
                                                if (lo_ht_attachment.ContainsKey(lo_signDoc_attachment.電子檔格式附件.原始檔序號) == false)
                                                {
                                                    lo_ht_attachment.Add(lo_signDoc_attachment.電子檔格式附件.原始檔序號, lo_signDoc_attachment.名稱);
                                                }
                                                break;
                                            default:
                                                if (lo_ht_attachment.ContainsKey(lo_signDoc_attachment.電子檔格式附件.原始檔序號) == false)
                                                {
                                                    //lo_ht_attachment.Add(lo_signDoc_attachment.電子檔格式附件.原始檔序號, lo_signDoc.文稿類型 + "_" + lo_signDoc_attachment.名稱);
                                                    //會辦意見上傳附件檔
                                                    lo_ht_attachment.Add(lo_signDoc_attachment.電子檔格式附件.原始檔序號, "會辦意見上傳附件檔");
                                                }
                                                break;
                                        }
                                    }
                                }
                            }
                            #endregion

                            #region 電子檔案資訊
                            if (lo_data.Object.簽核資訊.簽核文件夾.檔案清單 != null &&
                                Util.Common.getInt(lo_data.Object.簽核資訊.簽核文件夾.檔案清單.檔案數) > 0)
                            {
                                //電子檔案資訊 value 對應文稿key
                                foreach (Util.doc.電子檔案資訊 lo_fileDoc in lo_data.Object.簽核資訊.簽核文件夾.檔案清單.電子檔案資訊)
                                {
                                    if (lo_ht_electronicFile.ContainsKey(lo_fileDoc.原始檔序號) == false)
                                    {
                                        lo_ht_electronicFile.Add(lo_fileDoc.原始檔序號, lo_fileDoc.檔案名稱);
                                    }
                                }
                            }
                            #endregion

                            #region 由附件清單的原始檔序號抓取電子檔案檔名並由電子檔名為Key,附件名稱為值
                            foreach (System.Collections.DictionaryEntry o_de_attach in lo_ht_attachment)
                            {
                                if (lo_ht_electronicFile.ContainsKey(o_de_attach.Key))
                                {
                                    if (lo_ht_electronicList.ContainsKey(lo_ht_electronicFile[o_de_attach.Key]) == false)
                                    {
                                        lo_ht_electronicList.Add(lo_ht_electronicFile[o_de_attach.Key], o_de_attach.Value);
                                    }
                                }
                            }
                            
                            #endregion
                        }
                    }
                }
                #endregion

                if (lo_ht_electronicList != null && lo_ht_electronicList.Count > 0)
                {
                    si_json.Value = Util.Common.EncodeBase64Uri(Util.Common.hashTableToJson(lo_ht_electronicList));
                }
            }
            #endregion
        }
        return lb_bool;
    }
    /// <summary>
    /// 下載PDF檔案
    /// </summary>
    /// <param name="as_fileEntry"></param>
    private void download(string as_fileEntry)
    {
        lb_text.InnerText = "";
        if (System.IO.File.Exists(is_zipfilePath) && string.IsNullOrEmpty(as_fileEntry) == false)
        {
            try
            {
                string ls_tempDirPath = Util.Common.DecodeBase64Uri(tempDir.Value);
                if (string.IsNullOrEmpty(ls_tempDirPath) == false)
                {
                    System.IO.DirectoryInfo lo_tempDirInfo = new System.IO.DirectoryInfo(ls_tempDirPath);
                    if (lo_tempDirInfo.Exists == false)
                    {
                        lo_tempDirInfo.Create();
                    }

                    string ls_filePath = System.IO.Path.Combine(lo_tempDirInfo.FullName, as_fileEntry);

                    ZipArchive lo_zipArchive = System.IO.Compression.ZipFile.OpenRead(is_zipfilePath);
                    ZipArchiveEntry lo_zipArchiveEntry = lo_zipArchive.GetEntry(as_fileEntry);
                    lo_zipArchiveEntry.ExtractToFile(ls_filePath, true);

                    System.IO.FileInfo lo_fileInfo = new System.IO.FileInfo(ls_filePath);
                    if (lo_fileInfo.Exists)
                    {
                        fileToDownload_byTransmitFile(lo_fileInfo.FullName, true);
                    }
                }
            }
            catch (Exception ex)
            {
                lb_text.InnerText = "錯誤訊息：" + ex.Message;
            }
        }
        else
        {
            lb_text.InnerText = "公文zip檔案不存在，請確認！";
        }
    }


    /// <summary>
    /// 刪除暫存檔案的資料
    /// </summary>
    private void deleteTempDir()
    {
        if (string.IsNullOrEmpty(tempDir.Value) == false)
        {
            string ls_tempDirPath = Util.Common.DecodeBase64Uri(tempDir.Value);
            System.IO.DirectoryInfo lo_tempDirInfo = new System.IO.DirectoryInfo(ls_tempDirPath);
            if (lo_tempDirInfo.Exists)
            {
                lo_tempDirInfo.Delete(true);
                System.Threading.Thread.Sleep(100);
                System.Threading.Thread.Yield();
            }
        }
    }
    

    private void doJavaScript_download()
    {
        //filePath.Value = Util.Common.EncodeBase64Uri(lo_fileInfo.FullName);
        //fileYN.Value = Util.Common.EncodeBase64Uri(as_delete);

        string ls_javaScript = "<script type=\"text/javascript\">       \n"
                             + "\t form1.state.value = \"download\";    \n"
                             + "\t form1.submit(); \n"
                             + "</script> \n"
                             + "";
        this.ClientScript.RegisterStartupScript(this.GetType(), "fn", ls_javaScript);
    }

    public void fileToDownload_byTransmitFile(string as_filePath, bool deleteDirectory = false)
    {
        System.IO.FileInfo lo_fileInfo = new System.IO.FileInfo(as_filePath);
        if (lo_fileInfo.Exists)
        {
            System.IO.DirectoryInfo lo_dirinfo = lo_fileInfo.Directory;

            try
            {
                #region try
                string ls_outfileName = HttpUtility.UrlEncode(lo_fileInfo.Name, System.Text.Encoding.UTF8);

                if (System.Web.HttpContext.Current.Response != null &&
                    System.Web.HttpContext.Current.Response.IsClientConnected)
                {
                    System.Web.HttpContext.Current.Response.Clear();        //清除buffer
                    System.Web.HttpContext.Current.Response.ClearHeaders(); //清除buffer 表頭
                    System.Web.HttpContext.Current.Response.Buffer = false;
                    System.Web.HttpContext.Current.Response.BufferOutput = false;

                    //設定MIME類型為二進位檔案
                    HttpContext.Current.Response.ContentType = "application/octet-stream";
                    System.Web.HttpContext.Current.Response.AddHeader("Content-Disposition", string.Format("attachment; filename=\"{0}\"", ls_outfileName));
                    System.Web.HttpContext.Current.Response.AddHeader("Transfer-Encoding", "identity");
                    //表頭加入檔案大小
                    System.Web.HttpContext.Current.Response.AppendHeader("Content-Length", lo_fileInfo.Length.ToString());
                    System.Web.HttpContext.Current.Response.TransmitFile(lo_fileInfo.FullName);

                    //Response.End會引發錯誤訊息。文件參考:微軟的KBQ312629，Response.End, Response.Redirect, Server.Transfer 
                    //由下列程式取代End
                    //System.Web.HttpContext.Current.Response.End();
                    System.Web.HttpContext.Current.Response.SuppressContent = true;
                    if (System.Web.HttpContext.Current.ApplicationInstance != null)
                    {
                        System.Web.HttpContext.Current.ApplicationInstance.CompleteRequest();
                    }
                    System.Threading.Thread.Sleep(100);
                    //System.Web.HttpContext.Current.Response.Close();
                }
                #endregion
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                try
                {
                    if (deleteDirectory)
                    {
                        lo_dirinfo.Delete(true);
                    }
                }
                catch { };
            }
        }
    }

    


}


