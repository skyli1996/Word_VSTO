using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MyWordAddIn
{
    public partial class UCForRichText : UserControl
    {
        public UCForRichText(List<Question> QuestionList = null)
        {
            InitializeComponent();
            if(QuestionList != null)
            {
                int i = 1; // 题号
                richTextBox1.Text += "已添加题目：" + QuestionList.Count + "道\n";
                richTextBox1.Text += "------------------------------------------------------------\n";
                foreach (Question q in QuestionList)
                {
                    richTextBox1.Text += q.correct != null ? string.Format("第{0}题：正确答案为[{1}]\n\n", i, q.correct) : string.Format("第{0}题：该题无答案\n\n", i);
                    richTextBox1.Text += q.type != null ? (q.type.Equals("连线") ? string.Format("{0}题({1})\n\n", q.type, q.value) : string.Format("{0}题\n\n", q.type)) : "<注意>请添加题目类型<注意>\n\n";

                    if(!q.type.Equals("连线"))
                    {
                        // 适用单选、多选、填空、判断
                        if (q.id != -1 && q.id != 0)
                        {
                            richTextBox1.Text += q.value != null ? string.Format("{0}.({1}){2}\n\n", q.id, q.value, q.title) : string.Format("{0}.{1}\n\n", q.id, q.title);
                        }
                        else
                        {
                            richTextBox1.Text += q.value != null ? string.Format("{0}.({1}){2}\n\n", "<注意>题目/题号格式有误<注意>", q.value, q.title) : string.Format("{0}.{1}\n\n", "<注意>题目/题号格式有误<注意>", q.title);
                        }
                        foreach (string str in q.optionsOrTestsList)
                        {
                            richTextBox1.Text += str + "\n";
                        }
                    }
                    else if(q.type.Equals("连线"))
                    {
                        int j = 0;
                        while(j < q.linkLeftList.Count || j < q.linkRightList.Count)
                        {
                            if (j < q.linkLeftList.Count) richTextBox1.Text += q.linkLeftList[j];
                            richTextBox1.Text += "\t\t";
                            if (j < q.linkRightList.Count) richTextBox1.Text += q.linkRightList[j];
                            richTextBox1.Text += "\n";
                            j++;
                        }
                    }               
                    richTextBox1.Text += "------------------------------------------------------------\n";
                    i++;
                }
            }

        }

    }

}
