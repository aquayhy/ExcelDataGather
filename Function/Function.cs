using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelDataGather
{
    public static class Function
    {
        /// <summary>
        /// 检查一个字符串是否为纯数字
        /// </summary>
        /// <param name="message">对应字符串</param>
        /// <returns>纯数字返回true，否则返回假</returns>
        public static bool IsNumberic(string message)
        {
            System.Text.RegularExpressions.Regex rex = new System.Text.RegularExpressions.Regex(@"^\d+(\.\d+)?$");
            if (rex.IsMatch(message))
            {
                return true;
            }
            else
                return false;
        }
    }
}
