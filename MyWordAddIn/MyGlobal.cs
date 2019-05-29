using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyWordAddIn
{
    class MyGlobal
    {
        public static DataTable SingleSelection { get; set; }
        public static DataTable MultipleSelection { get; set; }
        public static DataTable ExperimentTest { get; set; }
        public static DataTable CalculateTest { get; set; }
        public static List<Question> QuestionList { get; set; }

        static MyGlobal()
        {
            QuestionList = null;
            SqlDao s1 = new SqlDao();            
            //SingleSelection = s1.ExecuteQuery(SqlAndPath.SqlForSingleSelection);
            //MultipleSelection = s1.ExecuteQuery(SqlAndPath.SqlForMultipleSelection);
            //ExperimentTest = s1.ExecuteQuery(SqlAndPath.SqlForExperimentTest);
            //CalculateTest = s1.ExecuteQuery(SqlAndPath.SqlForCalculateTest);
        }
    }
}
