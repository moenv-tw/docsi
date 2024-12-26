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
    #region �ɮ׸��
    /// <summary>
    /// �Ȧs�ɮ׸��|
    /// </summary>
    string is_tempDefault = "";
    /// <summary>
    /// �����ɮ׹w�]���|
    /// </summary>
    string is_docLocation = "";
    /// <summary>
    /// ���Y�ɮת����ؿ�
    /// </summary>
    string is_dirPath = "";
    /// <summary>
    /// ���Y�ɮת�������|
    /// </summary>
    string is_zipfilePath = "";

    /// <summary>
    /// doc_SI_�����T
    /// </summary>
    private System.Collections.Hashtable lo_ht_attachment = new Hashtable();
    /// <summary>
    /// doc_SI_�q�l�ɮ׸�T
    /// </summary>
    private System.Collections.Hashtable lo_ht_electronicFile = new Hashtable();
    /// <summary>
    /// �q�l�ɮ׻P�����T����
    /// </summary>
    private System.Collections.Hashtable lo_ht_electronicList = new Hashtable();
    /// <summary>
    /// ��ܰ��ɦW
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
    /// �]�w��l�ȸ��
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

        #region Ū��SI�ɮת����ɦW
        lo_ht_Extension.Add(".pdf", ".pdf");
        lo_ht_Extension.Add(".xls", ".xls");
        lo_ht_Extension.Add(".xlsx", ".xlsx");
        lo_ht_Extension.Add(".doc", ".doc");
        lo_ht_Extension.Add(".docx", ".docx");
        #endregion

        //�]�wzip�����|
        string ls_docNo = docNo.Value;
        if (string.IsNullOrEmpty(ls_docNo) == false)
        {
            is_dirPath = System.IO.Path.Combine(ls_docLocation, ls_docNo);

            is_zipfilePath = System.IO.Path.Combine(is_dirPath, ls_docNo + ".zip");
        }
    }
    /// <summary>
    /// �ت��G�̬d�����d�ߦh�����
    /// </summary>
    /// <returns></returns>
    protected System.Data.DataView queryAll()
    {
        #region dt
        System.Data.DataTable lo_dt = new DataTable("doc");
        System.Data.DataColumn lo_dc_rid = new DataColumn("rid", typeof(string));
        lo_dc_rid.Caption = "�t�νs��";
        System.Data.DataColumn lo_dc_fileName = new DataColumn("filaName", typeof(string));
        lo_dc_fileName.Caption = "�ɮצW��";
        System.Data.DataColumn lo_dc_fileMemo = new DataColumn("fileMemo", typeof(string));
        lo_dc_fileMemo.Caption = "��l�����ɮצW��";
        System.Data.DataColumn lo_dc_filelength = new DataColumn("filelength", typeof(string));
        lo_dc_fileName.Caption = "�ɮת���";
        System.Data.DataColumn lo_dc_lastWriteTime = new DataColumn("lastWriteTime", typeof(string));
        lo_dc_fileName.Caption = "�̫Ყ�ʤ��";

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
                    #region Ū������SI�ɮ�,�p����ƫh����Ū��json��ƥH�קK�į�t��
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

                    #region Ū��zip
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
                    lb_text.InnerText = "���~�T��:" + ex.Message;
                }
                #endregion
            }
            else
            {
                lb_text.InnerText = "�ɮפ��s�b�A�нT�{�I";
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

        //    //�������帹�ؿ�
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
    /// �U���ɮ�
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
                    lb_text.InnerText = "�ɮפ��s�b�A�нT�{�I";
                }
            }
            else
            {
                lb_text.InnerText = "�ɮ׸�Ƭ��ťաA�нT�{�I";
            }
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
    }
    
    /// <summary>
    /// Ū������t��SI�ɮ�
    /// </summary>
    /// <returns></returns>
    private bool read_doc_SI()
    {
        bool lb_bool = false;
        if (System.IO.File.Exists(is_zipfilePath))
        {
            string ls_TempfileFullPath = "";

            #region 1.�NSI�����Y��temp
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

            #region 2.Ū��SI���ɮר�HashTable
            System.IO.FileInfo lo_tempfileInfo = new System.IO.FileInfo(ls_TempfileFullPath);
            if (lo_tempfileInfo.Exists)
            {
                #region read process
                Util.doc.docSI lo_doc_si = new Util.doc.docSI();
                Util.doc.�u�Wñ�� lo_onlineSign = lo_doc_si.getClassData(lo_tempfileInfo.FullName);
                if (lo_onlineSign != null)
                {
                    string ls_objectId = "";
                    string ls_modifykind = "";

                    foreach (Util.doc.ñ���I�w�q lo_data in lo_onlineSign.�u�Wñ�ָ�T.ñ���I�w�q)
                    {
                        if (lo_data != null &&
                            lo_data.Object != null &&
                            lo_data.Object.ñ�ָ�T != null &&
                            lo_data.Object.ñ�ָ�T.ñ�֤�� != null &&
                            lo_data.Object.ñ�ָ�T.ñ�֤��.ñ�֤�Z�M�� != null)
                        {
                            lb_bool = true;
                            ls_objectId = lo_data.Object.Id;
                            ls_modifykind = lo_data.Object.���ʸ�T.���ʧO;

                            #region ��Z
                            //foreach (Util.doc.��Z lo_signDoc in lo_data.Object.ñ�ָ�T.ñ�֤��.ñ�֤�Z�M��.��Z)
                            //{
                            //    if (lo_signDoc != null &&
                            //        lo_signDoc.����M�� != null &&
                            //        Util.Common.getInt(lo_signDoc.����M��.�����) > 0)
                            //    {
                            //        foreach (Util.doc.���� lo_signDoc_attachment in lo_signDoc.����M��.����)
                            //        {
                            //            if (lo_ht_attachment.ContainsKey(lo_signDoc_attachment.�q�l�ɮ榡����.��l�ɧǸ�) == false)
                            //            {
                            //                lo_ht_attachment.Add(lo_signDoc_attachment.�q�l�ɮ榡����.��l�ɧǸ�, lo_signDoc_attachment.�W��);
                            //            }
                            //        }
                            //    }
                            //}
                            foreach (Util.doc.��Z lo_signDoc in lo_data.Object.ñ�ָ�T.ñ�֤��.ñ�֤�Z�M��.��Z)
                            {
                                #region ��Z
                                if (Util.Common.getInt(lo_signDoc.��l�ɧǸ�) > 0)
                                {
                                    if (lo_ht_attachment.ContainsKey(lo_signDoc.��l�ɧǸ�) == false)
                                    {
                                        lo_ht_attachment.Add(lo_signDoc.��l�ɧǸ�, lo_signDoc.����);
                                    }
                                }

                                if (lo_signDoc.��Z������ != null &&
                                    lo_signDoc.��Z������.��Z�����M�� != null &&
                                    lo_signDoc.��Z������.��Z�����M��.���� != null &&
                                    Util.Common.getInt(lo_signDoc.��Z������.��Z�����M��.����.��l�ɧǸ�) > 0)
                                {
                                    if (lo_ht_attachment.ContainsKey(lo_signDoc.��Z������.��Z�����M��.����.��l�ɧǸ�) == false)
                                    {
                                        lo_ht_attachment.Add(lo_signDoc.��Z������.��Z�����M��.����.��l�ɧǸ�, lo_signDoc.���� + "_����Z");
                                    }
                                }
                                #endregion


                                if (lo_signDoc.����M�� != null &&
                                    Util.Common.getInt(lo_signDoc.����M��.�����) > 0)
                                {
                                    //����M��
                                    foreach (Util.doc.���� lo_signDoc_attachment in lo_signDoc.����M��.����)
                                    {
                                        switch (lo_signDoc.��Z����)
                                        {
                                            case "ñ":
                                                if (lo_ht_attachment.ContainsKey(lo_signDoc_attachment.�q�l�ɮ榡����.��l�ɧǸ�) == false)
                                                {
                                                    lo_ht_attachment.Add(lo_signDoc_attachment.�q�l�ɮ榡����.��l�ɧǸ�, lo_signDoc_attachment.�W��);
                                                }
                                                break;
                                            default:
                                                if (lo_ht_attachment.ContainsKey(lo_signDoc_attachment.�q�l�ɮ榡����.��l�ɧǸ�) == false)
                                                {
                                                    //lo_ht_attachment.Add(lo_signDoc_attachment.�q�l�ɮ榡����.��l�ɧǸ�, lo_signDoc.��Z���� + "_" + lo_signDoc_attachment.�W��);
                                                    //�|��N���W�Ǫ�����
                                                    lo_ht_attachment.Add(lo_signDoc_attachment.�q�l�ɮ榡����.��l�ɧǸ�, "�|��N���W�Ǫ�����");
                                                }
                                                break;
                                        }
                                    }
                                }
                            }
                            #endregion

                            #region �q�l�ɮ׸�T
                            if (lo_data.Object.ñ�ָ�T.ñ�֤��.�ɮײM�� != null &&
                                Util.Common.getInt(lo_data.Object.ñ�ָ�T.ñ�֤��.�ɮײM��.�ɮ׼�) > 0)
                            {
                                //�q�l�ɮ׸�T value ������Zkey
                                foreach (Util.doc.�q�l�ɮ׸�T lo_fileDoc in lo_data.Object.ñ�ָ�T.ñ�֤��.�ɮײM��.�q�l�ɮ׸�T)
                                {
                                    if (lo_ht_electronicFile.ContainsKey(lo_fileDoc.��l�ɧǸ�) == false)
                                    {
                                        lo_ht_electronicFile.Add(lo_fileDoc.��l�ɧǸ�, lo_fileDoc.�ɮצW��);
                                    }
                                }
                            }
                            #endregion

                            #region �Ѫ���M�檺��l�ɧǸ�����q�l�ɮ��ɦW�åѹq�l�ɦW��Key,����W�٬���
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
    /// �U��PDF�ɮ�
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
                lb_text.InnerText = "���~�T���G" + ex.Message;
            }
        }
        else
        {
            lb_text.InnerText = "����zip�ɮפ��s�b�A�нT�{�I";
        }
    }


    /// <summary>
    /// �R���Ȧs�ɮת����
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
                    System.Web.HttpContext.Current.Response.Clear();        //�M��buffer
                    System.Web.HttpContext.Current.Response.ClearHeaders(); //�M��buffer ���Y
                    System.Web.HttpContext.Current.Response.Buffer = false;
                    System.Web.HttpContext.Current.Response.BufferOutput = false;

                    //�]�wMIME�������G�i���ɮ�
                    HttpContext.Current.Response.ContentType = "application/octet-stream";
                    System.Web.HttpContext.Current.Response.AddHeader("Content-Disposition", string.Format("attachment; filename=\"{0}\"", ls_outfileName));
                    System.Web.HttpContext.Current.Response.AddHeader("Transfer-Encoding", "identity");
                    //���Y�[�J�ɮפj�p
                    System.Web.HttpContext.Current.Response.AppendHeader("Content-Length", lo_fileInfo.Length.ToString());
                    System.Web.HttpContext.Current.Response.TransmitFile(lo_fileInfo.FullName);

                    //Response.End�|�޵o���~�T���C���Ѧ�:�L�n��KBQ312629�AResponse.End, Response.Redirect, Server.Transfer 
                    //�ѤU�C�{�����NEnd
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


