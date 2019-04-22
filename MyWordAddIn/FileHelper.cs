using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MyWordAddIn
{
    public class FileHelper
    {
        /// <summary>
        /// 将文件转换为二进制流进行读取
        /// </summary>
        /// <param name="fileName">文件完整名</param>
        /// <returns>二进制流</returns>
        public static byte[] FileToBinary(string fileName)
        {
            try
            {
                using (FileStream fsRead = new FileStream(fileName, FileMode.Open, FileAccess.Read))
                {
                    if (fsRead.CanRead)
                    {
                        int fsSize = Convert.ToInt32(fsRead.Length);

                        byte[] btRead = new byte[fsSize];

                        fsRead.Read(btRead, 0, fsSize);

                        return btRead;
                    }
                    else
                    {
                        MessageBox.Show("文件读取错误！");
                        return null;
                    }
                }            
            }
            catch (Exception ce)
            {
                MessageBox.Show(ce.Message);

                return null;
            }
        }

        /// <summary>
        /// 将二进制流转换为对应的文件进行存储
        /// </summary>
        /// <param name="filePath">接收的文件</param>
        /// <param name="btBinary">二进制流</param>
        /// <returns>转换结果</returns>
        public static bool BinaryToFile(string fileName, byte[] btBinary)
        {
            bool result = false;

            try
            {
                using (FileStream fsWrite = new FileStream(fileName, FileMode.Create, FileAccess.Write))
                {
                    if (fsWrite.CanWrite)
                    {
                        fsWrite.Write(btBinary, 0, btBinary.Length);
                        result = true;
                    }
                    else
                    {
                        result = false;
                    }
                }           
            }
            catch (Exception ce)
            {
                MessageBox.Show(ce.Message);
                result = false;
            }

            return result;
        }
    }
}
