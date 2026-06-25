using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;

namespace Zotero_linker
{
    public partial class ThisAddIn
    {
        internal ZoteroLinkerService LinkerService { get; private set; }

        internal ZoteroLinkerOptions CurrentOptions { get; private set; }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            LinkerService = new ZoteroLinkerService();
            CurrentOptions = ZoteroLinkerOptions.Load();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        internal void RefreshOptions()
        {
            CurrentOptions = ZoteroLinkerOptions.Load();
        }

        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
