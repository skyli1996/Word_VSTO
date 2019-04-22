using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyWordAddIn
{
    /// <summary>
    /// sql语句以及图片路径存放类
    /// </summary>
    class SqlAndPath
    {
        public static readonly string SqlForSingleSelection = "select * FROM Table_Sky WHERE Stype='单选'";

        public static readonly string SqlForMultipleSelection = "select * FROM Table_Sky WHERE Stype='多选'";

        public static readonly string SqlForExperimentTest = "select * FROM Table_Sky WHERE Stype='实验'";

        public static readonly string SqlForCalculateTest = "select * FROM Table_Sky WHERE Stype='计算'";

        public static readonly string PicturePath = AppDomain.CurrentDomain.BaseDirectory + "Picture";
    }
}
