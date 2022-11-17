using System;
using System.Diagnostics;
using System.Linq;
using Microsoft.Office.Interop.Word;

namespace ExcelDataGather
{
    public class WordOperate : IDisposable
    {
        private Microsoft.Office.Interop.Word.Application _app;
        private Microsoft.Office.Interop.Word._Document _doc;
        object _nullobj = System.Reflection.Missing.Value;
        string _filepath="";

        /// <summary>
        /// 关闭Word进程
        /// </summary>
        public void KillWinword()
        {
            var p = Process.GetProcessesByName("WINWORD");
            if (p.Any()) p[0].Kill();
        }

        /// <summary>
        /// 打开word文档
        /// </summary>
        /// <param name="filePath"></param>
        public void Open(string filePath)
        {

            _app = new Microsoft.Office.Interop.Word.Application();
            _filepath = filePath;
            _doc = _app.Documents.Add(_filepath);
            //_doc = _app.Documents.Open(filePath);
            /*ref file, ref _nullobj, ref _nullobj,
            ref _nullobj, ref _nullobj, ref _nullobj,
            ref _nullobj, ref _nullobj, ref _nullobj,
            ref _nullobj, ref _nullobj, ref _nullobj,
            ref _nullobj, ref _nullobj, ref _nullobj, ref _nullobj);*/
        }


        /// <summary>
        /// 替换word中的文字
        /// </summary>
        /// <param name="strOld">查找的文字</param>
        /// <param name="strNew">替换的文字</param>
        public void Replace(string strOld, string strNew)
        {
            _app.Selection.Find.ClearFormatting();
            _app.Selection.Find.Replacement.ClearFormatting();
            _app.Selection.Find.Text = strOld;
            _app.Selection.Find.Replacement.Text = strNew;
            object objReplace = Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll;
            _app.Selection.Find.Execute(ref _nullobj, ref _nullobj, ref _nullobj,
                                       ref _nullobj, ref _nullobj, ref _nullobj,
                                       ref _nullobj, ref _nullobj, ref _nullobj,
                                       ref _nullobj, ref objReplace, ref _nullobj,
                                       ref _nullobj, ref _nullobj, ref _nullobj);
        }

        /// <summary>
        /// 在word文档中插入图片
        /// </summary>
        /// <param name="bookmarkname">需要插入图片的标签名</param>
        /// <param name="photofile">图片地址</param>
        public void InsertPicture(string bookmarkname,string photofile,float picturelength,float picturewidth)
        {
            foreach(Bookmark bk in _doc.Bookmarks)
            {
                if(bk.Name== bookmarkname)
                {
                    bk.Select();
                    Selection sel = _app.Selection;
                    Microsoft.Office.Interop.Word.InlineShape inshape = sel.InlineShapes.AddPicture(photofile);
                    if (picturelength > 1e-6)
                        inshape.Width = Convert.ToSingle(picturelength / 0.3527);
                    if (picturewidth > 1e-6)
                        inshape.Height = Convert.ToSingle(picturewidth / 0.3527); 
                }
            }
        }

        /// <summary>
        /// 保存
        /// </summary>
        public void Save()
        {
            _doc.Save();
        }

        public void Save(bool visible)
        {
            _app.Visible = visible;
            _doc.SaveAs2(_filepath);
        }

        public void SaveAsPdf(string pdffilepath)
        {
            _doc.ExportAsFixedFormat(pdffilepath, WdExportFormat.wdExportFormatPDF);
        }

        public void Close()
        {
            _doc.Close(ref _nullobj, ref _nullobj, ref _nullobj);
            _app.Quit(ref _nullobj, ref _nullobj, ref _nullobj);
        }

        /// <summary>
        /// 退出
        /// </summary>
        public void Dispose()
        {
            _doc.Close(ref _nullobj, ref _nullobj, ref _nullobj);
            _app.Quit(ref _nullobj, ref _nullobj, ref _nullobj);
        }
    }
}