<%@ Application Language="C#" %>

<script runat="server">
    /// <summary>
    /// 在應用程式開始執行的程式碼
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    void Application_Start(object sender, EventArgs e)
    {
        System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.SystemDefault
                                                        | System.Net.SecurityProtocolType.Tls13
                                                        | System.Net.SecurityProtocolType.Tls12;
    }

    /// <summary>
    /// 在應用程式關閉時執行的程式碼
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    void Application_End(object sender, EventArgs e)
    {


    }
    /// <summary>
    /// 在應用程式發生錯誤
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    void Application_Error(object sender, EventArgs e)
    {
        try
        {

            string ls_yn = getString(System.Configuration.ConfigurationManager.AppSettings["writeErrorLog"]);
            if (ls_yn.ToUpper().Equals("Y"))
            {
                Exception lo_exception = this.Server.GetLastError();
                if (lo_exception != null)
                {
                    doWriteErrorLog(lo_exception.InnerException);
                }
            }
        }
        catch { }
    }

    /// <summary>
    /// 在新的工作階段啟動時執行的程式碼
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    void Session_Start(object sender, EventArgs e)
    {

    }
    /// <summary>
    /// 在工作階段結束時執行的程式碼
    /// <para>注意: 只有在  Web.config 檔案中將 sessionstate 模式設定為 InProc 時，</para>
    /// <para>才會引起 Session_End 事件。如果將 session 模式設定為 StateServer </para>
    /// <para>或 SQLServer，則不會引起該事件。</para>
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    void Session_End(object sender, EventArgs e)
    {
    }

    /// <summary>
    /// 取得文字
    /// </summary>
    /// <param name="ao_obj"></param>
    /// <returns></returns>
    private string getString(object ao_obj)
    {
        if (ao_obj != null)
            return ao_obj.ToString().Trim();
        else
            return "";
    }

    /// <summary>
    /// 建立錯誤記錄檔案
    /// </summary>
    /// <param name="as_filePath"></param>
    /// <returns></returns>
    private bool CreateErrorLogFile(string as_filePath)
    {
        bool lb_retValue = false;
        try
        {
            if (!System.IO.File.Exists(as_filePath))
            {
                System.IO.FileInfo lo_fileInfo = new System.IO.FileInfo(as_filePath);
                if (lo_fileInfo.Directory.Exists == false)
                {
                    System.IO.Directory.CreateDirectory(lo_fileInfo.Directory.FullName);
                }

                using (System.IO.FileStream lo_fileStream = System.IO.File.Create(as_filePath))
                {
                    lo_fileStream.Close();
                    lo_fileStream.Dispose();
                }
            }
            lb_retValue = true;
        }
        catch
        {
            return lb_retValue;
        }

        return lb_retValue;
    }

    /// <summary>
    /// 寫入錯誤訊息到檔案
    /// </summary>
    /// <param name="ao_exception"></param>
    /// <param name="as_msg"></param>
    private void doWriteErrorLog(Exception ao_exception, string as_msg = "")
    {
        bool lb_flag = false;

        string ls_logFilePath = HttpContext.Current.Server.MapPath(@"~\ErrorLog\ErrorLog_" + DateTime.Now.ToString("yyyyMMdd") + ".txt");
        if (!System.IO.File.Exists(ls_logFilePath))
        {
            lb_flag = CreateErrorLogFile(ls_logFilePath);
        }
        else
        {
            lb_flag = true;
        }

        if (lb_flag)
        {
            System.IO.StreamWriter lo_streamWriter = new System.IO.StreamWriter(ls_logFilePath, true, System.Text.Encoding.UTF8);
            try
            {
                string ls_now = DateTime.Now.ToString("yyyyMMddHHmmss");
                if (ao_exception != null)
                {
                    lo_streamWriter.WriteLine();
                    lo_streamWriter.WriteLine("** " + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss.ffff") + ", 錯誤編號：" + ls_now + ", 錯誤程式：" + HttpContext.Current.Request.FilePath + " **");
                    lo_streamWriter.WriteLine();
                    lo_streamWriter.WriteLine(ao_exception.StackTrace);
                    lo_streamWriter.WriteLine();
                    lo_streamWriter.WriteLine(ao_exception.Message);
                    lo_streamWriter.WriteLine();
                    lo_streamWriter.WriteLine("============================================================================================================================================");
                }
                else if (string.IsNullOrEmpty(as_msg) == false)
                {
                    lo_streamWriter.WriteLine();
                    lo_streamWriter.WriteLine("** " + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss.ffff") + ", 錯誤編號：" + ls_now + ", 錯誤程式：" + HttpContext.Current.Request.FilePath + " **");
                    lo_streamWriter.WriteLine("錯誤訊息：" + as_msg);
                    lo_streamWriter.WriteLine("============================================================================================================================================");
                }

            }
            finally
            {
                lo_streamWriter.Close();
            }
        }
    }
</script>
