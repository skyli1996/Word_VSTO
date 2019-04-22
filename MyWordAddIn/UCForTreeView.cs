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
    public partial class UCForTreeView : UserControl
    {
        private TreeNode book; //教材
        public UCForTreeView()
        {
            InitializeComponent();
        }

        private void UCForTreeView_Load(object sender, EventArgs e)
        {
            CreateTree();
        }

        private void CreateTree()

        {
            this.treeView1.CheckBoxes = true;
            book = new TreeNode("高中物理必修1");

            var chapter1 = book.Nodes.Add("第一章 运动的描述");
            var section = chapter1.Nodes.Add("1 质点 参考系和坐标系");
            section = chapter1.Nodes.Add("2 时间和位移");
            section = chapter1.Nodes.Add("3 运动快慢的描述-速度");
            section = chapter1.Nodes.Add("4 实验：用打点计时器测速度");
            section = chapter1.Nodes.Add("5 速度变化快慢的描述-加速度");

            var chapter2 = book.Nodes.Add("第二章 匀变速直线运动的研究");
            section = chapter2.Nodes.Add("1 匀变速直线运动的速度与时间的关系");
            section = chapter2.Nodes.Add("2 自由落体运动");

            var chapter3 = book.Nodes.Add("第三章 相互作用");
            section = chapter3.Nodes.Add("1 重力 基本相互作用");
            section = chapter3.Nodes.Add("2 弹力");
            section = chapter3.Nodes.Add("3 摩擦力");
            section = chapter3.Nodes.Add("4 力的合成");
            section = chapter3.Nodes.Add("5 力的分解");

            var chapter4 = book.Nodes.Add("第四章 牛顿运动定律");
            section = chapter4.Nodes.Add("1 牛顿第一定律");
            section = chapter4.Nodes.Add("2 牛顿第二定律");
            section = chapter4.Nodes.Add("3 牛顿第三定律");
            section = chapter4.Nodes.Add("4 用牛顿运动定律解决问题");

            //pointer.Tag = "123";            

            this.treeView1.Nodes.Add(book);
            book.Checked = true;
            //更新子节点状态
            UpdateChildNodes(book);
            //展开子节点
            ExpandChildNodes(book);
        }

        private void treeView1_AfterCheck(object sender, TreeViewEventArgs e)
        {
            //只处理鼠标点击引起的状态变化
            if (e.Action == TreeViewAction.ByMouse)
            {
                if (e.Node.Checked)
                {
                    //更新子节点状态
                    UpdateChildNodes(e.Node);
                    //展开子节点
                    ExpandChildNodes(e.Node);
                }
                else
                {
                    //更新子节点状态
                    UpdateChildNodes(e.Node);
                    //折叠子节点
                    CollapseChildNodes(e.Node);
                }

            }
        }

        private void UpdateChildNodes(TreeNode node)
        {
            foreach (TreeNode child in node.Nodes)
            {
                child.Checked = node.Checked;
                UpdateChildNodes(child);
            }
        }

        private void ExpandChildNodes(TreeNode node)
        {
            node.ExpandAll();
        }

        private void CollapseChildNodes(TreeNode node)
        {
            node.Collapse();
        }
    }
}
