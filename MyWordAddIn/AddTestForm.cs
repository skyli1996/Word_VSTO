using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MyWordAddIn
{
    public partial class AddTestForm : Form
    {
        SqlDao s1; // sql方法类
        string loadType; // 记录图片格式
        public AddTestForm()
        {
            InitializeComponent();
            s1 = new SqlDao();
        }

        private void AddTestForm_Load(object sender, EventArgs e)
        {
            comboBox1.SelectedIndex = 0;
            this.Left = 500;
            this.Top = 500;
        }

        /// <summary>
        /// 添加图片事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();//显示选择文件对话框
            ofd.ShowDialog();
            //得到上传文件的完整名
            string loadFullName = ofd.FileName.ToString();
            this.pictureBox1.ImageLocation = loadFullName;

            //上传文件的文件名
            string loadName = loadFullName.Substring(loadFullName.LastIndexOf("\\") + 1);
            System.Diagnostics.Debug.WriteLine(loadName);

            //上传文件的类型
            loadType = loadFullName.Substring(loadFullName.LastIndexOf(".") + 1).ToLower();

            //判断文件类型
            if (!loadType.Equals("jpg") && !loadType.Equals("gif") && !loadType.Equals("png") && (loadType != string.Empty))
            {
                MessageBox.Show("文件不合法!仅限于JPG,GIF,PNG格式!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            //将文件名显示到标签上
            this.textBox5.Text = loadFullName;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            // 逻辑判断
            if(richTextBox1.Text == string.Empty)
            {
                MessageBox.Show("题目不能为空!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if(comboBox1.Text.Equals("单选") || comboBox1.Text.Equals("多选"))
            {
                if(textBox1.Text == string.Empty || textBox2.Text == string.Empty || textBox3.Text == string.Empty || textBox4.Text == string.Empty)
                {
                    MessageBox.Show("选择题的选项不能缺省!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
            }
            if(textBox1.Text == string.Empty && (textBox2.Text != string.Empty || textBox3.Text != string.Empty || textBox4.Text != string.Empty))
            {
                MessageBox.Show("选项/问题出题顺序错误!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if ((textBox1.Text == string.Empty || textBox2.Text == string.Empty) && (textBox3.Text != string.Empty || textBox4.Text != string.Empty))
            {
                MessageBox.Show("选项/问题出题顺序错误!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if ((textBox1.Text == string.Empty || textBox2.Text == string.Empty || textBox3.Text == string.Empty) && textBox4.Text != string.Empty)
            {
                MessageBox.Show("选项/问题出题顺序错误!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 导入试题主要代码块
            if (MessageBox.Show("请再次确认题目类型", "注意", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
            {
                string type = comboBox1.Text;
                string question = richTextBox1.Text;
                string optionA = textBox1.Text;
                string optionB = textBox2.Text;
                string optionC = textBox3.Text;
                string optionD = textBox4.Text;
                string sqlStr = "INSERT INTO Table_Sky (Stype,Squestion,SoptionA,SoptionB,SoptionC,SoptionD) VALUES('" +
                   type + "','" + question + "','" + optionA + "','" + optionB + "','" + optionC + "','" + optionD + "')";
                int mark = s1.ExecuteUpdate(sqlStr); // 执行成功返回受影响行数，执行失败返回0；
                if (mark != 0)
                {
                    if(textBox5.Text.Equals("图片路径") || textBox5.Text == string.Empty)
                    {
                        MessageBox.Show("添加成功！", "添加结果", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        Close();
                    }
                    else
                    {
                        sqlStr = "SELECT TOP 1 * FROM Table_Sky order by Sid desc";
                        DataTable d1 = s1.ExecuteQuery(sqlStr);
                        int id = Convert.ToInt32(d1.Rows[0]["Sid"].ToString());
                        string pictureName = "sky_" + id.ToString() + "." + loadType;
                        if (addPicture(textBox5.Text, pictureName)) // 成功将图片添加进Picture文件夹
                        {
                            sqlStr = "update Table_Sky set Spicture='" + pictureName + "' where Sid=" + id;
                            mark = s1.ExecuteUpdate(sqlStr); // 将图片信息添加到数据库
                            if (mark != 0)
                            {
                                MessageBox.Show("添加成功！", "添加结果", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                Close();

                            }
                        }
                        else // 如果没有成功导入图片，则删除数据库刚添加的最后一行记录
                        {
                            sqlStr = "delete from Table_Sky where Sid like (SELECT TOP 1 * FROM Table_Sky order by Sid desc)";
                            mark = s1.ExecuteUpdate(sqlStr);
                            MessageBox.Show("添加失败！导入图片发生错误！", "添加结果", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            return;
                        }
                    }                    
                }
                else
                {
                    MessageBox.Show("添加失败！数据导入数据库发生错误！", "添加结果", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
            }            
        }

        /// <summary>
        /// 将图片导入Picture文件夹
        /// </summary>
        /// <param name="originalPath">图片来源路径</param>
        /// <param name="pictureName">图片命名</param>
        /// <returns>true Or false</returns>
        private bool addPicture(string originalPath, string pictureName)
        {
            string filePath = SqlAndPath.PicturePath + "\\" + pictureName;
            byte[] btFile = FileHelper.FileToBinary(originalPath); // 调用读取方法保存图片
            if(FileHelper.BinaryToFile(filePath, btFile))
            {
                return true;
            }
            return false;

        }
    }
}
