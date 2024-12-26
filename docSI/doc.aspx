<!--
程式目的：公文解析SI檔
程式代號：doc
程式日期：20241106
程式作者：sheng
--------------------------------------------------------
修改作者　　修改日期　　　修改目的
--------------------------------------------------------
-->

<%@ Page Language="C#" AutoEventWireup="true" CodeFile="doc.aspx.cs" Inherits="doc" %>
<%@ Register TagPrefix="ctl" Namespace="Util.WebControls" Assembly="CollectionPager" %>

<!DOCTYPE html>
<html>
<head runat="server">
    <title>電子公文封裝檔附件名稱解析</title>
    <meta http-equiv="pragma" content="no-cache" />
    <meta http-equiv="Cache-control" content="no-cache" />
    <link rel="stylesheet" href="js/default.css" type="text/css" />
    <script type="text/javascript" src="js/jquery.min.js"></script>
    <script type="text/javascript" src="js/json.min.js"></script>
    <script type="text/javascript" src="js/jquery.blockUI.js"></script>
    <script type="text/javascript" src="js/function.js"></script>
    <script type="text/javascript">
        function checkField()
        {
            var alertStr = "";

            if (form1.docNo.value == "")
            {
                alertStr += "公文文號不得為空白";
            }

            if (alertStr.length != 0)
            {
                alert(alertStr);
                return false;
            }
            else
            {
                if (form1.state.value == "download")
                {
                    form1.state.value = "queryAll";
                }
                showBlockUI("查詢中...");
                return true;
            }
        }

        function queryOne(as_rid)
        {
            if (form1.state.value == "download")
            {
                form1.state.value = "queryAll";
                form1.submit();
            }
            else
            {
                var ls_ext = as_rid.substr(as_rid.length - 3, 3).toLowerCase();
                if (ls_ext == "pdf" || ls_ext == "xls" || ls_ext == "doc")
                {
                    showBlockUI("下載檔案處理中...");
                    form1.fileEntry.value = as_rid;
                    form1.state.value = "popup_pdf";
                    form1.submit();
                }
            }
        }

        function init()
        {
            hiddenBlockUI();
        }
    </script>
</head>

<body onload="init();">
    <form id="form1" name="form1" method="post" autocomplete="off" runat="server" onsubmit="return checkField()">
        <table style="width: 100%; border-collapse: collapse; border-spacing: 0;">
            <!--Form區============================================================-->
            <tr>
                <td class="bg" style="text-align: center">
                    <div id="formContainer1">
                        <table class="table_form" style="width: 100%; height: auto;">
                            <tr>
                                <td class="queryTDLable" style="width: 50%; text-align:right;"><span style="color: red">*</span>電子公文傳簽號：</td>
                                <td class="queryTDInput" style="width: 50%; text-align:left;">
                                    [<input type="text" id="docNo" name="docNo" size="21" maxlength="20" value ="1131000908" runat="server" />]
                                    <input type="submit" id="bt_docNo" name="bt_docNo" value="查詢" runat="server" />
                               </td>
                            </tr>
                            <tr id="tr_showMsg" runat="server" style="display: none">
                                <td class="queryTDLable" colspan="2" style="text-align:center">
                                    <label class="labelRequire" id="lb_text" runat="server"></label>
                                </td>
                            </tr>
                        </table>
                    </div>
                </td>
            </tr>
            <!--Toolbar區============================================================-->
            <tr>
                <td class="bg" style="text-align: center">
                    <input type="hidden" id="rid" name="rid" value="" runat="server"/>
                    <input type="hidden" id="state" name="state" runat="server" value="queryAll"/>
                    <input type="hidden" id="queryAllFlag" name="queryAllFlag" runat="server" value=""/>
                    <input type="hidden" id="attachment" name="attachment" value="" runat="server" />
                    <input type="hidden" id="fileEntry" name="fileEntry" value="" runat="server" />
                    <!-- 暫存完整路徑 -->
                    <input type="hidden" id="tempDir" name="tempDir" value="" runat="server" />
                    <input type="hidden" id="filePath" name="filePath" value="" runat="server" />
                    <input type="hidden" id="fileYN" name="fileYN" value="" runat="server" />
                    <!-- si_json -->
                    <input type="hidden" id="si_json" name="si_json" value="" runat="server" />
                </td>
            </tr>
            <!--List區============================================================-->
            <tr>
                <td>
                    <ctl:CollectionPager ID="CollectionPager1"  runat="server" 
                        BackNextDisplay="HyperLinks" 
                        BackNextLocation="Right" 
                        BackText="前一頁" 
                        ShowFirstLast="True" 
                        ResultsLocation="Parallel" 
                        PagingMode="PostBack" 
                        MaxPages="5000" 
                        ResultsStyle="" 
                        ResultsFormat="共{2}筆, 第{0}-{1}筆" 
                        PageSize="10" 
                        ControlCssClass="CollectionPager" 
                        ShowLabel="False"
                        Visible="true">
                    </ctl:CollectionPager>
                </td>
            </tr>

            <tr>
                <td class="bg">
                    <div id="listContainer" class="listContainer" style="height: auto;">
                        <asp:Repeater ID="repeaterItems" runat="server">
                            <HeaderTemplate>
                                <table class="table_form" style="width: 100%; border-collapse: collapse; border-spacing: 0;">
                                    <thead id="listTHEAD" class="listTHEAD">
                                        <tr>
                                            <th class="listTH"><a class="text_link_w" href="#">檔案名稱</a></th>
                                            <th class="listTH"><a class="text_link_w" href="#">原始附件檔案名稱</a></th>
                                            <th class="listTH"><a class="text_link_w" href="#">檔案長度</a></th>
                                            <th class="listTH"><a class="text_link_w" href="#">最後異動日期</a></th>
                                        </tr>
                                    </thead>
                                    <tbody id="listTBODY">
                            </HeaderTemplate>
                            <ItemTemplate>
                                        <tr class="listTR" id="ID<%#Eval("rid")%>" onmouseover="this.className='listTRMouseover'" onmouseout="this.className='listTR'" onclick="queryOne('<%#Eval("rid")%>')" >
                                            <td class="listTDL"><%#Eval("filaName")%></td>
                                            <td class="listTDL"><%#Eval("fileMemo")%></td>
                                            <td class="listTDR"><%#Eval("filelength")%></td>
                                            <td class="listTD"><%#Eval("lastWriteTime")%></td>
                                        </tr>
                            </ItemTemplate>
                            <FooterTemplate>
                                    </tbody>
                                </table>
                            </FooterTemplate>
                        </asp:Repeater>
                    </div>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>



