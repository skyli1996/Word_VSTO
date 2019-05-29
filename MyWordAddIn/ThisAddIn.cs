using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using Microsoft.Office.Tools;
using System.Drawing;
using System.Windows.Forms;

namespace MyWordAddIn
{
    public partial class ThisAddIn
    {
        LineQuestion lineQuestion;
        int clickNumber = 0; //用于防止多次添加点击事件
        string answer = null; //答案
        public CustomTaskPane _MyCustomTaskPane = null;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // 初始化右侧解析框
            UCForRichText richTextBox = new UCForRichText();
            _MyCustomTaskPane = this.CustomTaskPanes.Add(richTextBox, "解析结果");
            _MyCustomTaskPane.Width = 800;
            _MyCustomTaskPane.Visible = false;

            //获取右键菜单对象mzBar
            Office.CommandBar mzBar = Application.CommandBars["Text"];
            //重置菜单
            mzBar.Reset();
            //获取右键菜单的控制对象
            Office.CommandBarControls bars = mzBar.Controls;

            /*已被mzBar.Reset();替代，用于除去其余标签为"autoAnswer"的菜单项
            foreach (Office.CommandBarControl temp_contrl in bars)
            {
                string t = temp_contrl.Tag.Trim();
                if (t.Equals("autoAnswer"))
                {
                    temp_contrl.Delete();
                }
            }
            */

            //在右键菜单列表添加新的菜单项
            Office.CommandBarControl comControl = bars.Add(Office.MsoControlType.msoControlButton, missing, missing, missing, true); //添加自己的菜单项
            //将该菜单项转为按钮可点击模式
            Office.CommandBarButton comButton = comControl as Office.CommandBarButton;
            if (comControl != null)
            {
                comButton.Tag = "autoAnswer";//添加标签"autoAnswer"
                comButton.Caption = "选择答案";//设置文本
                comButton.Style = Office.MsoButtonStyle.msoButtonIconAndCaption;
                comButton.Enabled = false;//设置不可点击
                //name = comButton.accName;
                //comButton.Click += new Office._CommandBarButtonEvents_ClickEventHandler(_RightBtn_Click);

            }

            Document vstoDoc = Globals.Factory.GetVstoObject(this.Application.ActiveDocument);
            vstoDoc.BeforeDoubleClick += new Microsoft.Office.Tools.Word.ClickEventHandler(ThisDocument_BeforeDoubleClick);

            this.Application.WindowBeforeRightClick += new Word.ApplicationEvents4_WindowBeforeRightClickEventHandler(Application_WindowBeforeRightClick);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        void Application_WindowBeforeRightClick(Word.Selection Sel, ref bool Cancel)
        {
            clickNumber = 0;
            Word.Document document = this.Application.ActiveDocument; // 全文文档对象
            lineQuestion = new LineQuestion();
            lineQuestion.leftAnswerList = new List<LineLeftAnswer>();
            lineQuestion.rightAnswerList = new List<LineRightAnswer>();
            //获取标签为"autoAnswer"的菜单项
            Office.CommandBarButton addBtn = (Office.CommandBarButton)Application.CommandBars["Text"].FindControl(Office.MsoControlType.msoControlButton, missing, "autoAnswer", false);
            addBtn.Enabled = false;

            //addBtn.Click -= new Office._CommandBarButtonEvents_ClickEventHandler(_RightBtn_Click);


            if (!string.IsNullOrWhiteSpace(Sel.Range.Text) && Sel.Range.Text.Trim().Length >= 2 && Sel.Range.Text.Trim().Length <= 5 && Sel.Range.Text.Trim().Contains("连线"))
            {
                addBtn.Enabled = true;

                //object charUnit = Word.WdUnits.wdCharacter; //字符移动
                object lineUnit = Word.WdUnits.wdLine; //行移动
                //object paragraphUnit = Word.WdUnits.wdParagraph; //段落移动
                //object storyUnit = Word.WdUnits.wdStory; // 当前文档
                //object count = 3;  // 移动次数
                object extend = Word.WdMovementType.wdExtend; //extend对光标移动区域进行扩展选择

                Sel.HomeKey(lineUnit);
                Sel.EndKey(lineUnit, extend);
                //选择光标所在段落的内容
                //Sel.MoveUp(paragraphUnit);
                //Sel.MoveDown(paragraphUnit, extend); 

                string currentParagraphText = Sel.Range.Text; // 获取该行文本内容
                lineQuestion.currentParagraph = 0;
                for (int i = 1; i <= document.Paragraphs.Count; i++)
                {
                    if (currentParagraphText.Trim().Equals(document.Paragraphs[i].Range.Text.Trim().ToString()))
                    {
                        // 获取该行在文档中的段落数
                        lineQuestion.currentParagraph = i;
                    }
                }
                if (lineQuestion.currentParagraph != 0)
                {
                    // 获取每一道连线题的连线行总数
                    lineQuestion.answerCount = validParagraphs(lineQuestion.currentParagraph);
                }
                if (lineQuestion.answerCount != 0)
                {

                    for (int i = 1; i <= lineQuestion.answerCount; i++)
                    {
                        // 获取连线行内容
                        string paragraphText = document.Paragraphs[lineQuestion.currentParagraph + i].Range.Text.ToString();

                        // 操作连线行左侧
                        string leftAnswerStr = paragraphText.Substring(0, paragraphText.IndexOf(" "));
                        LineLeftAnswer lla = new LineLeftAnswer();
                        lla.leftIndex = leftAnswerStr.Substring(0, leftAnswerStr.IndexOf("."));
                        lla.leftAnswer = leftAnswerStr.Substring(leftAnswerStr.IndexOf(".") + 1);
                        lineQuestion.leftAnswerList.Add(lla);

                        // 操作连线行左侧
                        string rightAnswerStr = paragraphText.Substring(paragraphText.LastIndexOf(" ") + 1);
                        LineRightAnswer lra = new LineRightAnswer();
                        lra.rightIndex = rightAnswerStr.Substring(0, rightAnswerStr.IndexOf("."));
                        lra.rightAnswer = rightAnswerStr.Substring(rightAnswerStr.IndexOf(".") + 1);
                        lineQuestion.rightAnswerList.Add(lra);
                    }
                }
                addBtn.Click += new Office._CommandBarButtonEvents_ClickEventHandler(_RightBtn_Click);
            }
        }

        /// <summary>
        /// 返回每一道连线题的连线行总数
        /// </summary>
        /// <param name="currentParagraph"></param>
        /// <returns></returns>
        private int validParagraphs(int currentParagraph)
        {
            int i = 0;
            Word.Document document = this.Application.ActiveDocument;
            while (!string.IsNullOrWhiteSpace(document.Paragraphs[currentParagraph].Range.Text.ToString()))
            {
                i++;
                if (currentParagraph == document.Paragraphs.Count)
                    break;
                currentParagraph++;
            }
            return i - 1;
        }

        /// <summary>
        /// 让窗体停靠在光标附近
        /// </summary>
        /// <param name="Sel"></param>
        /// <returns></returns>
        private static Point GetPositionForShowing(Word.Selection Sel)
        {
            // get range postion
            int left = 0;
            int top = 0;
            int width = 0;
            int height = 0;
            Globals.ThisAddIn.Application.ActiveDocument.ActiveWindow.GetPoint(out left, out top, out width, out height, Sel.Range);

            Point currentPos = new Point(left, top);
            if (Screen.PrimaryScreen.Bounds.Height - top > 340)
            {
                currentPos.Y += 20;
            }
            else
            {
                currentPos.Y -= 320;
            }
            return currentPos;
        }

        /// <summary>
        /// 点击事件
        /// </summary>
        /// <param name="Ctrl"></param>
        /// <param name="CancelDefault"></param>
        void _RightBtn_Click(Office.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            if (clickNumber == 0)
            {  
                //Point currentPos = GetPositionForShowing(this.Application.Selection);
                LineSelectAnswerForm answerForm = new LineSelectAnswerForm(lineQuestion) { StartPosition = FormStartPosition.CenterParent };
                answerForm.AnswerEvent += (str) => { answer = str; }; // 获得答案
                //answerForm.Location = currentPos;
                answerForm.ShowDialog();

                Word.Document document = this.Application.ActiveDocument;
                string originalTxt = document.Paragraphs[lineQuestion.currentParagraph].Range.Text.ToString();
                originalTxt = originalTxt.Substring(0, originalTxt.IndexOf("]", 5) + 1);
                string currentTxt = originalTxt + answer + "\n";
                answer = null;
                document.Paragraphs[lineQuestion.currentParagraph].Range.Text = currentTxt; // 替换原段落的内容

                clickNumber++;
            }
        }

        void ThisDocument_BeforeDoubleClick(object sender, Microsoft.Office.Tools.Word.ClickEventArgs e)
        {
            Document vstoDoc = Globals.Factory.GetVstoObject(this.Application.ActiveDocument);
            lineQuestion = new LineQuestion();
            lineQuestion.leftAnswerList = new List<LineLeftAnswer>();
            lineQuestion.rightAnswerList = new List<LineRightAnswer>();
            Word.Selection Sel = this.Application.Selection;
            object charUnit = Word.WdUnits.wdCharacter; //字符移动
            object lineUnit = Word.WdUnits.wdLine; //行移动
            object count = 5;  // 移动次数
            object extend = Word.WdMovementType.wdExtend; //extend对光标移动区域进行扩展选择
            Sel.MoveEnd(charUnit, count);
            Sel.HomeKey(lineUnit, extend);
            if (!string.IsNullOrWhiteSpace(Sel.Range.Text) && Sel.Range.Text.Trim().ToString().Contains("连线"))
            {
                Sel.HomeKey(lineUnit);
                Sel.EndKey(lineUnit, extend);
                string currentParagraphText = Sel.Range.Text; // 获取该行文本内容
                lineQuestion.currentParagraph = 0;
                for (int i = 1; i <= vstoDoc.Paragraphs.Count; i++)
                {
                    if (currentParagraphText.Trim().Equals(vstoDoc.Paragraphs[i].Range.Text.Trim().ToString()))
                    {
                        // 获取该行在文档中的段落数
                        lineQuestion.currentParagraph = i;
                        break;
                    }
                }
                if (lineQuestion.currentParagraph != 0)
                {
                    // 获取每一道连线题的连线行总数
                    lineQuestion.answerCount = validParagraphs(lineQuestion.currentParagraph);
                }
                if (lineQuestion.answerCount != 0)
                {

                    for (int i = 1; i <= lineQuestion.answerCount; i++)
                    {
                        // 获取连线行内容
                        string paragraphText = vstoDoc.Paragraphs[lineQuestion.currentParagraph + i].Range.Text.ToString();

                        // 操作连线行左侧
                        string leftAnswerStr = paragraphText.Substring(0, paragraphText.IndexOf(" "));
                        LineLeftAnswer lla = new LineLeftAnswer();
                        lla.leftIndex = leftAnswerStr.Substring(0, leftAnswerStr.IndexOf("."));
                        lla.leftAnswer = leftAnswerStr.Substring(leftAnswerStr.IndexOf(".") + 1);
                        lineQuestion.leftAnswerList.Add(lla);

                        // 操作连线行左侧
                        string rightAnswerStr = paragraphText.Substring(paragraphText.LastIndexOf(" ") + 1);
                        LineRightAnswer lra = new LineRightAnswer();
                        lra.rightIndex = rightAnswerStr.Substring(0, rightAnswerStr.IndexOf("."));
                        lra.rightAnswer = rightAnswerStr.Substring(rightAnswerStr.IndexOf(".") + 1);
                        lineQuestion.rightAnswerList.Add(lra);
                    }
                }

                if(lineQuestion != null)
                {
                    LineSelectAnswerForm answerForm = new LineSelectAnswerForm(lineQuestion) { StartPosition = FormStartPosition.CenterParent };
                    answerForm.AnswerEvent += (str) => { answer = str; }; // 获得答案
                    answerForm.ShowDialog();

                    Word.Document document = this.Application.ActiveDocument;
                    string originalTxt = document.Paragraphs[lineQuestion.currentParagraph].Range.Text.ToString();
                    originalTxt = originalTxt.Substring(0, originalTxt.IndexOf("]", 5) + 1);
                    string currentTxt = originalTxt + answer + "\n";
                    answer = null;
                    document.Paragraphs[lineQuestion.currentParagraph].Range.Text = currentTxt; // 替换原段落的内容
                }

            }
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
