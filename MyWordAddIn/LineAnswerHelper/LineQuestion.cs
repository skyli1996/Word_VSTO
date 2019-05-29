using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyWordAddIn
{
    public class LineQuestion
    {
        public int currentParagraph { get; set; }

        public int answerCount { get; set; }

        public List<LineLeftAnswer> leftAnswerList { get; set; }

        public List<LineRightAnswer> rightAnswerList { get; set; }
    }
}
