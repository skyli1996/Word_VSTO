using MyWordAddIn.LineAnswerHelper;
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
    public partial class LineSelectAnswerForm : Form
    {
        private LineQuestion _lineQuestion;
        public event Action<string> AnswerEvent;
        //实时记录鼠标坐标
        Point po = Point.Empty;
        //鼠标按下时的坐标
        Point mouseDownPoint = Point.Empty;
        //记录所拖动的panel
        PanelExLine panel = null;
        //记录靠近的panelEx
        PanelEx nearPanel = null;
        //是否正在拖拽 
        bool isDrag = false;

        //panelEx集合
        List<PanelEx> panelExList = new List<PanelEx>();
        //panelExLine集合
        List<PanelExLine> panelExLineList = new List<PanelExLine>();

        public LineSelectAnswerForm(LineQuestion lineQuestion)
        {
            InitializeComponent();
            //设置Style支持透明背景色
            this.SetStyle(ControlStyles.SupportsTransparentBackColor, true);
            this.BackColor = Color.FromArgb(0, 0, 0, 0);

            _lineQuestion = lineQuestion;
        }

        private void LineSelectAnswerForm_Load(object sender, EventArgs e)
        {
            if (_lineQuestion != null)
            {
                this.ClientSize = new System.Drawing.Size(600, _lineQuestion.answerCount * 110 + 40);
                addTextAndPanel();
                AddButton();
            }
        }

        public void addTextAndPanel()
        {
            for(int i = 1; i <= _lineQuestion.answerCount; i++)
            {
                Label leftTxt = new Label();
                leftTxt.Name = "leftTxt" + (char)(i + 64);
                leftTxt.Text = _lineQuestion.leftAnswerList[i - 1].leftAnswer;
                leftTxt.Location = new Point(30, 40 + (i - 1) * 80);
                leftTxt.AutoSize = true;
                this.Controls.Add(leftTxt);

                PanelExLine panelExLine = new PanelExLine("panel" + (char)(i + 64), new Point(80, 35 + (i - 1) * 80), new Size(25, 25));
                panelExLine.MouseDown += new System.Windows.Forms.MouseEventHandler(this.panel_MouseDown);
                panelExLine.MouseMove += new System.Windows.Forms.MouseEventHandler(this.panel_MouseMove);
                panelExLine.MouseUp += new System.Windows.Forms.MouseEventHandler(this.panel_MouseUp);
                this.Controls.Add(panelExLine);

                Label rightTxt = new Label();
                rightTxt.Name = "rightLabel" + i;
                rightTxt.Text = _lineQuestion.rightAnswerList[i - 1].rightAnswer;
                rightTxt.Location = new Point(300, 40 + (i - 1) * 80);
                rightTxt.AutoSize = true;
                this.Controls.Add(rightTxt);

                PanelEx panelEx = new PanelEx("panel" + i, new Point(265, 35 + (i - 1) * 80), new Size(25, 25));
                this.Controls.Add(panelEx);

                //遍历所有的PanelEx，存入集合中
                foreach (Control control in this.Controls)
                {
                    if (control is PanelEx)
                    {
                        PanelEx pe = (PanelEx)control;
                        panelExList.Add(pe);
                    }
                }
                //遍历所有的PanelExLine，存入集合中
                foreach (Control control in this.Controls)
                {
                    if (control is PanelExLine)
                    {
                        PanelExLine pe = (PanelExLine)control;
                        panelExLineList.Add(pe);
                    }
                }
            }
        }

        public void AddButton()
        {
            Button certainButton = new Button();
            certainButton.Name = "certainButton";
            certainButton.Text = "确定";
            certainButton.Size = new Size(135, 55);
            certainButton.Location = new Point(150, _lineQuestion.answerCount * 90);
            certainButton.Click += new EventHandler(certainButton_Click);
            this.Controls.Add(certainButton);

            Button cancelButton = new Button();
            cancelButton.Name = "cancelButton";
            cancelButton.Text = "取消";
            cancelButton.Size = new Size(135, 55);
            cancelButton.Location = new Point(320, _lineQuestion.answerCount * 90);
            cancelButton.Click += new EventHandler(cancelButton_Click);
            this.Controls.Add(cancelButton);
        }

        /// <summary>
        /// “确定”按钮点击事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void certainButton_Click(object sender, EventArgs e)
        {
            string str = "[";

            foreach (Control control in this.Controls)
            {
                if (control is PanelExLine)
                {
                    PanelExLine pe = (PanelExLine)control;
                    if (pe.panelEx == null)
                    {
                        MessageBox.Show("请完成所有连线");
                        return;
                    }
                    else
                    {
                        str += pe.index + "(" + pe.panelEx.index + "),";
                    }
                }
            }
            str = str.Substring(0, str.Length - 1);
            str += "]";
            AnswerEvent?.Invoke(str);
            Close();
            //System.Diagnostics.Debug.WriteLine(str);
        }

        /// <summary>
        /// “取消”按钮点击事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cancelButton_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void panel_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                panel = (PanelExLine)sender;
                if (panel.panelEx != null)
                {
                    if (panel.panelEx.panelExLine != null) panel.panelEx.panelExLine = null;
                    panel.panelEx = null;
                }
                //鼠标点击时相对于控件的位置进行记录
                mouseDownPoint = new Point(MousePosition.X - this.Location.X, MousePosition.Y - this.Location.Y);
            }
        }

        private void panel_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                isDrag = true;
                //鼠标位置
                po = getPointToForm(new Point(MousePosition.X - Location.X - mouseDownPoint.X, MousePosition.Y - Location.Y - mouseDownPoint.Y));
                this.Invalidate();
            }
        }

        //把相对与control控件的坐标，转换成相对于窗体的坐标。
        private Point getPointToForm(Point p)
        {
            return this.PointToClient(panel.PointToScreen(p));
        }

        private void panel_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                if (isDrag)
                {
                    isDrag = false;
                    if (isNearPanelEx(po))
                    {
                        if (nearPanel != null)
                        {
                            panel.panelEx = nearPanel;
                            if (nearPanel.panelExLine != null)
                            {
                                nearPanel.panelExLine.panelEx = null;
                                nearPanel.panelExLine = panel;
                            }
                            else
                                nearPanel.panelExLine = panel;
                        }
                        this.Invalidate();
                    }
                    this.Invalidate();
                }
                reset();
            }
        }

        //重置变量
        private void reset()
        {
            isDrag = false;
            //panel = null;
        }

        private void LineSelectAnswerForm_Paint(object sender, PaintEventArgs e)
        {
            if (panel != null)
            {
                if (isDrag)
                {
                    //绘制线段
                    e.Graphics.DrawLine(new Pen(Color.FromArgb(255, Color.Black), 1), getLineStartPoint(po, panel).X, getLineStartPoint(po, panel).Y, getLineEndPoint(po, panel).X, getLineEndPoint(po, panel).Y);
                    //绘制文字
                    e.Graphics.DrawString(panel.index, new Font("宋体", 10.5F), Brushes.Black, po.X + panel.Size.Width / 4.0F, po.Y + panel.Size.Height / 4.0F);
                    //绘制圆
                    e.Graphics.DrawEllipse(new Pen(Color.FromArgb(255, Color.Black), 1), po.X, po.Y, panel.Size.Width, panel.Size.Height);

                    if (isNearPanelEx(po))
                        //填充圆,半透明显示文字
                        e.Graphics.FillEllipse(new SolidBrush(Color.FromArgb(100, Color.Red)), po.X, po.Y, panel.Size.Width, panel.Size.Height);
                    else
                        e.Graphics.FillEllipse(new SolidBrush(Color.FromArgb(100, Color.DarkBlue)), po.X, po.Y, panel.Size.Width, panel.Size.Height);

                }
            }
        }

        //线段结尾处坐标
        private Point getLineEndPoint(Point p, PanelExLine tempP)
        {
            float bTriangleX = p.X + tempP.Size.Width / 2.0F - tempP.center.X;
            float bTriangleY = p.Y + tempP.Size.Height / 2.0F - tempP.center.Y;
            float length = (float)Math.Sqrt(Math.Pow(bTriangleX, 2) + Math.Pow(bTriangleY, 2));
            float sTriangleX = bTriangleX / length * (tempP.Size.Width / 2.0F);
            float sTriangleY = bTriangleY / length * (tempP.Size.Height / 2.0F);
            return new Point((int)(p.X + tempP.Size.Width / 2.0F - sTriangleX), (int)(p.Y + tempP.Size.Height / 2.0F - sTriangleY));

        }

        //线段开始处坐标
        private Point getLineStartPoint(Point p, PanelExLine tempP)
        {
            float bTriangleX = p.X + tempP.Size.Width / 2.0F - tempP.center.X;
            float bTriangleY = p.Y + tempP.Size.Height / 2.0F - tempP.center.Y;
            float length = (float)Math.Sqrt(Math.Pow(bTriangleX, 2) + Math.Pow(bTriangleY, 2));
            float sTriangleX = bTriangleX / length * (tempP.Size.Width / 2.0F);
            float sTriangleY = bTriangleY / length * (tempP.Size.Height / 2.0F);
            return new Point((int)(tempP.center.X + sTriangleX), (int)(tempP.center.Y + sTriangleY));
        }

        //获取两点间距离
        private float getPointsDistance(Point p1, Point p2)
        {
            float triangleX = (float)Math.Pow(p2.X - p1.X, 2);
            float triangleY = (float)Math.Pow(p2.Y - p1.Y, 2);
            return (float)Math.Sqrt(triangleX + triangleY);
        }

        //是否靠近PanelEx
        private bool isNearPanelEx(Point p1)
        {

            foreach (PanelEx panelEx in panelExList)
            {
                float distance = getPointsDistance(p1, panelEx.Location);
                if (distance <= 30)
                {
                    nearPanel = panelEx;
                    return true;
                }
                else
                {
                    nearPanel = null;
                }
            }
            return false;
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            foreach (PanelExLine panelExLine in panelExLineList)
            {
                if (panelExLine.panelEx != null)
                {
                    e.Graphics.DrawLine(new Pen(Color.FromArgb(255, Color.Black), 1), getLineStartPoint(panelExLine.panelEx.Location, panelExLine).X, getLineStartPoint(panelExLine.panelEx.Location, panelExLine).Y, getLineEndPoint(panelExLine.panelEx.Location, panelExLine).X, getLineEndPoint(panelExLine.panelEx.Location, panelExLine).Y);
                }
            }
        }
    }
}
