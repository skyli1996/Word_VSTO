using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using Microsoft.Office.Tools;

namespace MyWordAddIn
{
    public partial class ThisAddIn
    {
        public CustomTaskPane _MyCustomTaskPane = null;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            UCForRichText richTextBox = new UCForRichText();
            _MyCustomTaskPane = this.CustomTaskPanes.Add(richTextBox, "解析结果");
            _MyCustomTaskPane.Width = 800;
            _MyCustomTaskPane.Visible = false;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
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
