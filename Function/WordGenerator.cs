using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Reflection;
using Microsoft.Office.Interop.Word;
using System.Globalization;
using System.ComponentModel;
using System.Data;

namespace ExcelDataGather
{

    public class WordGenerator : WordOperate
    {
        /// <summary>
        /// 批量生成word文件
        /// </summary>
        /// <param name="data">datatable的一行</param>
        /// <param name="alldatafield">需要替换字段</param>
        /// <param name="templatefile">模板文件</param>
        /// <param name="targetdirectory">输出文件目录</param>
        /// <param name="filenamefield">输出文件名所在字段名称</param>
        /// <param name="genPDF">是否输出PDF</param>
        /// <param name="insertPictureOption">插入图片选项</param>
        /// <param name="isSuccess">是否成功生成</param>
        /// <param name="remark">输出备注信息</param>
        /// <returns>输出word文件的地址</returns>
        public string GenWordFileFromDatatable(DataRow data, IList<String>alldatafield,string templatefile, string targetdirectory, string filenamefield, bool genPDF, InsertPictureOption insertPictureOption, out bool isSuccess, out string remark)
        {
            string targetfilename, targetfilefullname, targetPDFfilefullname;
            targetfilename = data[filenamefield].ToString();

            targetfilefullname = string.Format("{0}\\{1}.docx", targetdirectory, targetfilename);
            targetPDFfilefullname = string.Format("{0}\\{1}.pdf", targetdirectory, targetfilename);
            string photofilefullpath="";
            if (insertPictureOption.InsertPicture)
                photofilefullpath = string.Format("{0}\\{1}{2}", insertPictureOption.PhotoFilePath, data[insertPictureOption.PhotoField], insertPictureOption.PhotoExt);
            if (File.Exists(targetfilefullname))
            {
                try
                {
                    File.Delete(targetfilefullname);
                }
                catch
                {
                    isSuccess = false;
                    remark = String.Format("目录中存在 {0} 文件，且无法被删除，请检查", targetfilefullname);
                    return "";
                }
            }
            if (!File.Exists(templatefile))
            {
                isSuccess = false;
                remark = String.Format("未找到原始模板文件 {0} ，请检查", templatefile);
                return "";
            }
            if (insertPictureOption.InsertPicture&& !File.Exists(photofilefullpath))
            {
                isSuccess = false;
                remark = String.Format("未找到照片文件 {0} ，请检查", photofilefullpath);
                return "";
            }
            File.Copy(templatefile, targetfilefullname);
            Open(targetfilefullname);
            for (int i = 0; i < alldatafield.Count; i++)
            {
                string field = alldatafield[i].ToString();
                Replace(field, data[field].ToString());
            }
            if(insertPictureOption.InsertPicture)
                InsertPicture("picture", photofilefullpath,insertPictureOption.PictureLength,insertPictureOption.PictureWidth);
            Save(false);
            if (genPDF)
            {
                if (File.Exists(targetPDFfilefullname))
                {
                    try
                    {
                        File.Delete(targetPDFfilefullname);
                    }
                    catch
                    {
                        isSuccess = false;
                        remark = String.Format("目录中存在 {0} 文件，且无法被删除，请检查", targetPDFfilefullname);
                        return "";
                    }
                }
                SaveAsPdf(targetPDFfilefullname);
            }
            Close();
            isSuccess = true;
            remark = "";
            return (targetfilefullname);
        }
    }
}