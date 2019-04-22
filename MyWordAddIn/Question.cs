using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyWordAddIn
{
    public class Question
    {
        public int id { get; set; } // 题目序号
        public string type { get; set; } // 题目类型
        public string value { get; set; } // 题目分值
        public string correct { get; set; } // 题目答案
        public string title { get; set; } // 题目问题
        public List<string> optionsOrTestsList { get; set; } // 选项/小问集合
        public List<string> linkLeftList { get; set; } // 连线题左侧集合
        public List<string> linkRightList { get; set; } // 连线题右侧集合
    }
}
