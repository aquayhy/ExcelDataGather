using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Windows.Forms.Design;

namespace ExcelDataGather
{
    public class LocalDirDlg : FolderNameEditor
    {
        FolderNameEditor.FolderBrowser fDialog = new System.Windows.Forms.Design.FolderNameEditor.FolderBrowser();
        public LocalDirDlg()
        {
            //   
            //   TODO:   在此处添加构造函数逻辑 
            // 
        }

        public DialogResult DisplaySourceDialog()
        {
            return DisplayDialog("请选择源目录");
        }


        public DialogResult DisplayTargetDialog()
        {
            return DisplayDialog("请选择目标目录 ");
        }

        public DialogResult DisplayDialog(string description)
        {
            fDialog.Description = description;
            return fDialog.ShowDialog();
        }

        ///   <summary> 
        ///   获得当前路径 
        ///   </summary> 
        public string Path
        {
            get
            {
                return fDialog.DirectoryPath;
            }
        }
    }

    public struct InsertPictureOption
    {
        public bool InsertPicture;//指示是否插入图片
        public string PhotoField;//图片文件名字段名称
        public string PhotoExt;//图片后缀
        public float PictureLength;//图片长度，以毫米输入
        public float PictureWidth;//图片宽度，以毫米输入
        public string PhotoFilePath;//图片所在目录
    }
}
