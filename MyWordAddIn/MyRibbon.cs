using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Word = Microsoft.Office.Interop.Word;

namespace MyWordAddIn
{
    public partial class MyRibbon
    {
        Question q;
        private void MyRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            if (Globals.ThisAddIn._MyCustomTaskPane != null)
            {
                Globals.ThisAddIn._MyCustomTaskPane.Visible = true;
            }
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            if (Globals.ThisAddIn._MyCustomTaskPane != null)
            {
                Globals.ThisAddIn._MyCustomTaskPane.Visible = false;
            }
        }

        #region 一键组卷
        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Document document = Globals.ThisAddIn.Application.ActiveDocument;
            Word.Range rng = document.Content; // 获取全部文本范围
            rng.Delete(); // 清空原先文本
            addHeaderAndFooter();
            creatTitle();
            creatSingleSelection();
            creatMultipleSelection();
            creatExperimentTest();
            creatCalculateTest();
        }



        /// <summary>
        /// 标题
        /// </summary>
        private void creatTitle()
        {
            Word._Document document = Globals.ThisAddIn.Application.ActiveDocument;
            Word.Selection focusSelect = Globals.ThisAddIn.Application.Selection; //获取光标位置     
            focusSelect.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter; // 居中
            focusSelect.Font.Bold = 1; // 1加粗,0不加粗
            focusSelect.Font.Size = float.Parse("15"); // 15为小三字体
            focusSelect.Font.Name = "黑体"; // 黑体
            DateTime currentTime = new System.DateTime();
            currentTime = DateTime.Now; // 获取当前时间
            string strYMD = currentTime.ToString("yyyy-MM-dd");
            focusSelect.TypeText(strYMD + " 高中年级-物理考试");

            focusSelect.TypeParagraph(); // 添加一个段落
            focusSelect.Font.Bold = 0;
            focusSelect.Font.Size = float.Parse("11");
            focusSelect.Font.Name = "宋体"; // 宋体
            focusSelect.TypeText("\n总分: 100分\n");
        }

        /// <summary>
        /// 单选题
        /// </summary>
        private void creatSingleSelection()
        {
            Word.Selection focusSelect = Globals.ThisAddIn.Application.Selection; 
            focusSelect.TypeParagraph();
            focusSelect.TypeParagraph();
            focusSelect.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft; // 居左
            focusSelect.Font.Bold = 1;
            focusSelect.TypeText("一、单选题");
            focusSelect.Font.Bold = 0;
            focusSelect.TypeText("（共6题；共18分）\n");

            int[] singleSelectionList = new int[6];
            singleSelectionList = myRandom(6, MyGlobal.SingleSelection.Rows.Count); // 获取要打印的题目集合
            
            for (int i = 0; i < 6; i++)
            {
                focusSelect.TypeParagraph();
                // 打印题目
                focusSelect.TypeText((i+1).ToString() + ".(3分)" + 
                    MyGlobal.SingleSelection.Rows[singleSelectionList[i]-1]["Squestion"].ToString() + "\n");

                // 打印图片
                if(MyGlobal.SingleSelection.Rows[singleSelectionList[i] - 1]["Spicture"].ToString() != string.Empty)
                {
                    focusSelect.InlineShapes.AddPicture(SqlAndPath.PicturePath + "\\" +
                        MyGlobal.SingleSelection.Rows[singleSelectionList[i] - 1]["Spicture"].ToString());
                    focusSelect.TypeParagraph();
                }

                // 打印选择题
                if(MyGlobal.SingleSelection.Rows[singleSelectionList[i] - 1]["SoptionA"].ToString().Length <= 6)
                {
                    focusSelect.TypeText("A." + MyGlobal.SingleSelection.Rows[singleSelectionList[i] - 1]["SoptionA"].ToString() + "    ");
                    focusSelect.TypeText("B." + MyGlobal.SingleSelection.Rows[singleSelectionList[i] - 1]["SoptionB"].ToString() + "    ");
                    focusSelect.TypeText("C." + MyGlobal.SingleSelection.Rows[singleSelectionList[i] - 1]["SoptionC"].ToString() + "    ");
                    focusSelect.TypeText("D." + MyGlobal.SingleSelection.Rows[singleSelectionList[i] - 1]["SoptionD"].ToString() + "\n");
                }
                else
                {
                    focusSelect.TypeText("A." + MyGlobal.SingleSelection.Rows[singleSelectionList[i] - 1]["SoptionA"].ToString() + "\n");
                    focusSelect.TypeText("B." + MyGlobal.SingleSelection.Rows[singleSelectionList[i] - 1]["SoptionB"].ToString() + "\n");
                    focusSelect.TypeText("C." + MyGlobal.SingleSelection.Rows[singleSelectionList[i] - 1]["SoptionC"].ToString() + "\n");
                    focusSelect.TypeText("D." + MyGlobal.SingleSelection.Rows[singleSelectionList[i] - 1]["SoptionD"].ToString() + "\n");
                }

            }
                       
        }

        /// <summary>
        /// 多选题
        /// </summary>
        private void creatMultipleSelection()
        {
            Word.Selection focusSelect = Globals.ThisAddIn.Application.Selection;
            focusSelect.TypeParagraph();
            focusSelect.TypeParagraph();
            focusSelect.Font.Bold = 1;
            focusSelect.TypeText("二、多选题");
            focusSelect.Font.Bold = 0;
            focusSelect.TypeText("（共6题；共24分）\n");

            int[] multipleSelectionList = new int[6];
            multipleSelectionList = myRandom(6, MyGlobal.MultipleSelection.Rows.Count); // 获取要打印的题目集合
            for (int i = 0; i < 6; i++)
            {
                focusSelect.TypeParagraph();
                // 打印题目
                focusSelect.TypeText((i + 7).ToString() + ".(4分)" +
                    MyGlobal.MultipleSelection.Rows[multipleSelectionList[i] - 1]["Squestion"].ToString() + "\n");

                // 打印图片
                if (MyGlobal.MultipleSelection.Rows[multipleSelectionList[i] - 1]["Spicture"].ToString() != string.Empty)
                {
                    focusSelect.InlineShapes.AddPicture(SqlAndPath.PicturePath + "\\" +
                        MyGlobal.MultipleSelection.Rows[multipleSelectionList[i] - 1]["Spicture"].ToString());
                    focusSelect.TypeParagraph();
                }

                // 打印选择题
                focusSelect.TypeText("A." + MyGlobal.MultipleSelection.Rows[multipleSelectionList[i] - 1]["SoptionA"].ToString() + "\n");
                focusSelect.TypeText("B." + MyGlobal.MultipleSelection.Rows[multipleSelectionList[i] - 1]["SoptionB"].ToString() + "\n");
                focusSelect.TypeText("C." + MyGlobal.MultipleSelection.Rows[multipleSelectionList[i] - 1]["SoptionC"].ToString() + "\n");
                focusSelect.TypeText("D." + MyGlobal.MultipleSelection.Rows[multipleSelectionList[i] - 1]["SoptionD"].ToString() + "\n");
            }

        }

        /// <summary>
        /// 实验题
        /// </summary>
        private void creatExperimentTest()
        {
            Word.Selection focusSelect = Globals.ThisAddIn.Application.Selection;
            focusSelect.TypeParagraph();
            focusSelect.TypeParagraph();
            focusSelect.Font.Bold = 1;
            focusSelect.TypeText("三、实验题");
            focusSelect.Font.Bold = 0;
            focusSelect.TypeText("（共3题；共18分）\n");

            int[] experimentTestList = new int[3];
            experimentTestList = myRandom(3, MyGlobal.ExperimentTest.Rows.Count); // 获取要打印的题目集合
            for (int i = 0; i < 3; i++)
            {
                focusSelect.TypeParagraph();

                // 打印题目
                focusSelect.TypeText((i + 13).ToString() + ".(6分)" +
                    MyGlobal.ExperimentTest.Rows[experimentTestList[i] - 1]["Squestion"].ToString() + "\n");

                // 打印图片
                if (MyGlobal.ExperimentTest.Rows[experimentTestList[i] - 1]["Spicture"].ToString() != string.Empty)
                {
                    focusSelect.InlineShapes.AddPicture(SqlAndPath.PicturePath + "\\" +
                        MyGlobal.ExperimentTest.Rows[experimentTestList[i] - 1]["Spicture"].ToString());
                    focusSelect.TypeParagraph();
                }

                // 打印实验题的每小问
                if (MyGlobal.ExperimentTest.Rows[experimentTestList[i] - 1]["SoptionA"].ToString() != string.Empty)
                {
                    focusSelect.TypeText(MyGlobal.ExperimentTest.Rows[experimentTestList[i] - 1]["SoptionA"].ToString() + "\n");
                }
                else continue;
                if(MyGlobal.ExperimentTest.Rows[experimentTestList[i] - 1]["SoptionB"].ToString() != string.Empty)
                {
                    focusSelect.TypeText(MyGlobal.ExperimentTest.Rows[experimentTestList[i] - 1]["SoptionB"].ToString() + "\n");
                }
                else continue;
                if (MyGlobal.ExperimentTest.Rows[experimentTestList[i] - 1]["SoptionC"].ToString() != string.Empty)
                {
                    focusSelect.TypeText(MyGlobal.ExperimentTest.Rows[experimentTestList[i] - 1]["SoptionC"].ToString() + "\n");
                }
                else continue;
                if (MyGlobal.ExperimentTest.Rows[experimentTestList[i] - 1]["SoptionD"].ToString() != string.Empty)
                {
                    focusSelect.TypeText(MyGlobal.ExperimentTest.Rows[experimentTestList[i] - 1]["SoptionD"].ToString() + "\n");
                }
            }
        }

        /// <summary>
        /// 计算题
        /// </summary>
        private void creatCalculateTest()
        {
            Word.Selection focusSelect = Globals.ThisAddIn.Application.Selection;
            focusSelect.TypeParagraph();
            focusSelect.TypeParagraph();
            focusSelect.Font.Bold = 1;
            focusSelect.TypeText("四、计算题");
            focusSelect.Font.Bold = 0;
            focusSelect.TypeText("（共5题；共40分）\n");

            int[] calculateTestList = new int[5];
            calculateTestList = myRandom(5, MyGlobal.CalculateTest.Rows.Count); // 获取要打印的题目集合
            for (int i = 0; i < 5; i++)
            {
                focusSelect.TypeParagraph();

                // 打印题目
                focusSelect.TypeText((i + 16).ToString() + ".(" + (i+6).ToString() + "分)" +
                    MyGlobal.CalculateTest.Rows[calculateTestList[i] - 1]["Squestion"].ToString() + "\n");

                // 打印图片
                if (MyGlobal.CalculateTest.Rows[calculateTestList[i] - 1]["Spicture"].ToString() != string.Empty)
                {
                    focusSelect.InlineShapes.AddPicture(SqlAndPath.PicturePath + "\\" +
                        MyGlobal.CalculateTest.Rows[calculateTestList[i] - 1]["Spicture"].ToString());
                    focusSelect.TypeParagraph();
                }

                // 打印计算题的每小问
                if (MyGlobal.CalculateTest.Rows[calculateTestList[i] - 1]["SoptionA"].ToString() != string.Empty)
                {
                    focusSelect.TypeText(MyGlobal.CalculateTest.Rows[calculateTestList[i] - 1]["SoptionA"].ToString() + "\n");
                }
                if (MyGlobal.CalculateTest.Rows[calculateTestList[i] - 1]["SoptionB"].ToString() != string.Empty)
                {
                    focusSelect.TypeText(MyGlobal.CalculateTest.Rows[calculateTestList[i] - 1]["SoptionB"].ToString() + "\n");
                }
                if (MyGlobal.CalculateTest.Rows[calculateTestList[i] - 1]["SoptionC"].ToString() != string.Empty)
                {
                    focusSelect.TypeText(MyGlobal.CalculateTest.Rows[calculateTestList[i] - 1]["SoptionC"].ToString() + "\n");
                }
                if (MyGlobal.CalculateTest.Rows[calculateTestList[i] - 1]["SoptionD"].ToString() != string.Empty)
                {
                    focusSelect.TypeText(MyGlobal.CalculateTest.Rows[calculateTestList[i] - 1]["SoptionD"].ToString() + "\n");
                }
                focusSelect.TypeParagraph();
                /*
                int j = i * 2 + 12;
                while (j-- >= 0)
                {
                    focusSelect.TypeParagraph();
                }
                */

            }
        }

        /// <summary>
        /// 添加页眉页脚
        /// </summary>
        private void addHeaderAndFooter()
        {
            
            foreach (Word.Section wordSection in Globals.ThisAddIn.Application.ActiveDocument.Sections)
            {
                Word.Range footerRange = wordSection.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                Word.Range headerRange = wordSection.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                headerRange.Text = "北京师范大学珠海分校 1518060028_李佳颖";
                headerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                footerRange.Font.ColorIndex = Word.WdColorIndex.wdBlack;
                //footerRange.Font.Size = 5;
                footerRange.Fields.Add(footerRange, Word.WdFieldType.wdFieldPage);
                footerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            }
        }

        /// <summary>
        /// 随机获取要在数据库抽取的题目序号数组
        /// </summary>
        /// <param name="needNum">所需要的题目总数</param>
        /// <param name="totalNum">数据库中所有符合的题目总数</param>
        /// <returns></returns>
        private int[] myRandom(int needNum, int totalNum)
        {
            int[] a = new int[needNum];
            // flag用来标记取到的随机数和之前的有没有重复,1代表有重复,0代表没有重复
            int i,j,temp, flag;
            Random rand = new Random();
            for (i = 0; i < needNum; i++)
            {
                flag = 0;
                temp = rand.Next(1, totalNum+1);
                for (j = 0; j < i; j++)
                {
                    // 如果有重复
                    if (temp == a[j])
                    {
                        flag = 1;
                        break;
                    }
                }
                // i--为了抵消for循环括号里的i++
                if (flag == 1)         
                {
                    i--;
                    continue; // 重新取数
                }
                else { a[i] = temp; }

            }

            //使用冒泡法对数组进行升序排序
            for(i = 0; i < needNum; i++)
            {
                for(j = 0; j < needNum-i-1; j++)
                {
                    if(a[j] > a[j+1])
                    {
                        temp = a[j];
                        a[j] = a[j + 1];
                        a[j + 1] = temp;
                    }
                }
            }
            return a; // 返回数组
        }
        #endregion

        private void button4_Click(object sender, RibbonControlEventArgs e)
        {
            AddTestForm atf = new AddTestForm();
            atf.Show();

        }

        #region 扫描解析
        private void button5_Click(object sender, RibbonControlEventArgs e)
        {
            MyGlobal.QuestionList = new List<Question>();
            Word.Document document = Globals.ThisAddIn.Application.ActiveDocument;
            int currentParagraph = 1; // Word文档的段落序号是从1开始的
            while(currentParagraph <= document.Paragraphs.Count)
            {
                q = new Question();
                while(string.IsNullOrWhiteSpace(document.Paragraphs[currentParagraph].Range.Text.ToString())) // 循环直到段落不为空
                {
                    if (currentParagraph == document.Paragraphs.Count) break;
                    currentParagraph++;
                }
                if (currentParagraph == document.Paragraphs.Count) break;
                int paragraphsNum = validParagraphs(currentParagraph); // 当前题的所有段落总数
                int currentMark = 0; // 当前题正在解析的段落数-1
                if(analyzeFirstParagraph(document.Paragraphs[currentParagraph].Range.Text, q)) // 解析题目的第一段 
                {
                    currentMark++;
                }  
                if(!q.type.Equals("连线"))
                {
                    if (currentMark < paragraphsNum)
                    {
                        analyzeTitle(document.Paragraphs[currentParagraph + currentMark].Range.Text, q); //解析题目的标题
                        currentMark++;
                    }
                    analyzeOptionsOrTests(currentParagraph, currentMark, paragraphsNum, q); //解析题目的所有选项/小问
                }
                else if(q.type.Equals("连线"))
                {
                    analyzeLinkTests(currentParagraph, currentMark, paragraphsNum, q);
                }                     
                MyGlobal.QuestionList.Add(q);
                currentParagraph += paragraphsNum;
            }
            executeTaskPane(MyGlobal.QuestionList); // 打开TaskPane
        }

        /// <summary>
        /// 返回每一道题的段落总数
        /// </summary>
        /// <param name="currentParagraph"></param>
        /// <returns></returns>
        private int validParagraphs(int currentParagraph)
        {
            int i = 0;
            Word.Document document = Globals.ThisAddIn.Application.ActiveDocument;
            while (!string.IsNullOrWhiteSpace(document.Paragraphs[currentParagraph].Range.Text.ToString()))
            {
                i++;
                if (currentParagraph == document.Paragraphs.Count)
                    break;
                currentParagraph++;               
            }
            return i;
        }

        /// <summary>
        /// 解析题目的第一段，解析出类型、分值、正确答案
        /// </summary>
        /// <param name="str"></param>
        private bool analyzeFirstParagraph(string str, Question q)
        {
            // 类型
            try
            {
                if(!string.IsNullOrEmpty(str.Substring(str.IndexOf('[') + 1, str.IndexOf(']') - str.IndexOf('[') - 1)))
                {
                    q.type = str.Substring(str.IndexOf('[') + 1, str.IndexOf(']') - str.IndexOf('[') - 1); // 获得第一个[]内的内容
                }
                else
                {
                    q.type = "无类型";
                }                
                str = str.Substring(str.IndexOf(']') + 1); // 除去第一个[]
            }
            catch(ArgumentOutOfRangeException)
            {
                q.type = null;
                return false;
            }

            // 分值
            try
            {
                if (!string.IsNullOrEmpty(str.Substring(str.IndexOf('[') + 1, str.IndexOf(']') - str.IndexOf('[') - 1)))
                {
                    q.value = str.Substring(str.IndexOf('[') + 1, str.IndexOf(']') - str.IndexOf('[') - 1); // 获得第二个[]内的内容
                }
                else
                {
                    q.value = null;
                }
                str = str.Substring(str.IndexOf(']') + 1); // 除去第二个[]
            }
            catch (ArgumentOutOfRangeException)
            {
                q.value = null;
            }

            // 正确答案
            try
            {
                if (!string.IsNullOrEmpty(str.Substring(str.IndexOf('[') + 1, str.IndexOf(']') - str.IndexOf('[') - 1)))
                {
                    q.correct = str.Substring(str.IndexOf('[') + 1, str.IndexOf(']') - str.IndexOf('[') - 1); // 获得第三个[]内的内容
                }
                else
                {
                    q.correct = null;
                }
            }
            catch (ArgumentOutOfRangeException)
            {
                q.correct = null;
            }
            return true;
        }

        /// <summary>
        /// 解析题目的标题，解析出题目序号、问题
        /// </summary>
        /// <param name="str"></param>
        private void analyzeTitle(string str, Question q)
        {
            try
            {
                q.id = Convert.ToInt32(str.Substring(0, str.IndexOf('.')));
                q.title = str.Substring(str.IndexOf('.') + 1);
            }
            catch(ArgumentOutOfRangeException)
            {
                q.id = -1;
                q.title = str.Trim();

            }
            catch(FormatException)
            {
                q.id = -1;
                q.title = str.Trim() ;
            }
            
        }

        /// <summary>
        /// 解析题目的所有选项/小问
        /// </summary>
        /// <param name="currentParagraph"></param>
        /// <param name="currentMark"></param>
        /// <param name="paragraphsNum"></param>
        /// <param name="q"></param>
        private void analyzeOptionsOrTests(int currentParagraph, int currentMark, int paragraphsNum,  Question q)
        {
            List<string> tempList = new List<string>();
            Word.Document document = Globals.ThisAddIn.Application.ActiveDocument;
            for (; currentMark < paragraphsNum; currentMark++)
            {
                tempList.Add(document.Paragraphs[currentParagraph + currentMark].Range.Text);
            }
            q.optionsOrTestsList = tempList;
        }

        /// <summary>
        /// 解析特殊题型--连线题
        /// </summary>
        /// <param name="currentParagraph"></param>
        /// <param name="currentMark"></param>
        /// <param name="paragraphsNum"></param>
        /// <param name="q"></param>
        private void analyzeLinkTests(int currentParagraph, int currentMark, int paragraphsNum, Question q)
        {
            Word.Document document = Globals.ThisAddIn.Application.ActiveDocument;
            List<string> tempLeftList = new List<string>();
            List<string> tempRightList = new List<string>();
            for (; currentMark < paragraphsNum; currentMark++)
            {
                string str = document.Paragraphs[currentParagraph + currentMark].Range.Text;
                str = str.Trim(); // 去掉字符串前后的空格
                //int blankNum = str.Length - str.Replace(" ", "").Length; // 获取字符串中间空格的数量
                try
                {
                    string leftStr = str.Substring(0, str.IndexOf(" "));
                    tempLeftList.Add(leftStr);
                }
                catch (ArgumentOutOfRangeException)
                {
                    string leftStr = str + "<注意>格式不正确<注意>";
                    tempLeftList.Add(leftStr);
                }

                try
                {
                    string rightStr = str.Substring(str.LastIndexOf(" "));
                    tempRightList.Add(rightStr);
                }
                catch (ArgumentOutOfRangeException) { }

            }
            q.linkLeftList = tempLeftList;
            q.linkRightList = tempRightList;
        }

        /// <summary>
        /// 打开TaskPane
        /// </summary>
        /// <param name="QuestionList"></param>
        private void executeTaskPane(List<Question> QuestionList)
        {
            if (Globals.ThisAddIn._MyCustomTaskPane != null)
            {
                Globals.ThisAddIn._MyCustomTaskPane.Dispose(); // 释放旧TaskPane
                UCForRichText richTextBox = new UCForRichText(QuestionList);
                Globals.ThisAddIn._MyCustomTaskPane = Globals.ThisAddIn.CustomTaskPanes.Add(richTextBox, "解析结果");
                Globals.ThisAddIn._MyCustomTaskPane.Width = 800;
                Globals.ThisAddIn._MyCustomTaskPane.Visible = true;
            }
        }

        #endregion
    }
}
