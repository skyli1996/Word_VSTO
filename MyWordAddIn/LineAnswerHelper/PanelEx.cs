using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MyWordAddIn.LineAnswerHelper
{
    public class PanelEx : Panel
    {
        public Point center { get; private set; } //圆心
        public PanelExLine panelExLine { get; set; }
        public string index { get; private set; }
        public PanelEx(string name, Point location, Size size)
        {
            this.Name = name;
            this.Location = location;
            this.Size = size;
            this.Paint += new PaintEventHandler(myPaint);
            center = new Point(Location.X + Size.Width / 2, Location.Y + Size.Height / 2);
            index = name.Substring(name.LastIndexOf('l') + 1);
        }

        private void myPaint(object sender, PaintEventArgs e)
        {
            //绘制文字
            StringFormat sf = new StringFormat(); //文字格式
            sf.Alignment = StringAlignment.Center; //使文字居中
            string text = this.Name.Substring(Name.LastIndexOf('l') + 1);
            Font font = new Font("宋体", 10.5F);
            Brush brush = Brushes.Black;
            e.Graphics.DrawString(text, font, brush, (float)Size.Width / 2, (float)Size.Height / 2 - 9F, sf);

            //绘制圆
            int radius = Size.Width / 2; //半径
            Point circle = new Point(0, 0);
            int border = 1;
            int d = radius * 2 - border; //直径
            //e.Graphics.FillEllipse(new SolidBrush(Color.FromArgb(255, Color.Yellow)), centre.X, centre.Y, d, d);
            e.Graphics.DrawEllipse(new Pen(Color.FromArgb(255, Color.Black), border), circle.X, circle.Y, d, d);
        }

    }
}
