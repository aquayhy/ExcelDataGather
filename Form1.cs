using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Windows.Forms.Design;

namespace ExcelDataGather
{
    public partial class Form1 : Form
    {
        IList<string> sheetNameList = new List<string>();
        DataTable datatable;
        OpenFileDialog ofd;
        public Form1()
        {
            InitializeComponent();
            ofd = new OpenFileDialog
            {
                FilterIndex = 1,
                Multiselect = false,
                RestoreDirectory = true
            };
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ofd.Filter = "Excel files (*.xls,*.xlsx)|*.xls;*.xlsx";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                textBoxfilename.Text = ofd.FileName;
                sheetNameList = ExcelHelper.GetAllSheetNameOfExcel(ofd.FileName);
                comboBoxSheetName.DataSource = sheetNameList;
                comboBoxSheetName.SelectedIndex = 0;
                if (sheetNameList.Count == 1)
                {
                    LoadTable();
                }
            }
        }

        private void LoadTable()
        {
            datatable = ExcelHelper.TransExcelToDataTable(textBoxfilename.Text, comboBoxSheetName.Text);
            dataGridView1.DataSource = datatable;
            if (datatable != null)
            {

                sheetNameList.Clear();
                int colcount = datatable.Columns.Count;
                for (int i = 0; i < colcount; i++)
                {
                    sheetNameList.Add(datatable.Columns[i].ColumnName);
                }
                if (checkedListBoxfield.Items != null && checkedListBoxfield.Items.Count > 0)
                    checkedListBoxfield.Items.Clear();
                comboBoxPhotoField.Items.Clear();
                comboBoxFilename.Items.Clear();
                comboBoxGroupName.Items.Clear();
                foreach (string data in sheetNameList)
                {
                    checkedListBoxfield.Items.Add(data);
                    comboBoxPhotoField.Items.Add(data);
                    comboBoxFilename.Items.Add(data);
                    comboBoxGroupName.Items.Add(data);
                }
                if (sheetNameList.Count > 0)
                {
                    comboBoxPhotoField.SelectedIndex = 0;
                    comboBoxFilename.SelectedIndex = 0;
                    comboBoxGroupName.SelectedIndex = 0;
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            {
                for (int i = 0; i < checkedListBoxfield.Items.Count; i++)
                {
                    checkedListBoxfield.SetItemChecked(i, true);
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            {
                for (int i = 0; i < checkedListBoxfield.Items.Count; i++)
                {
                    checkedListBoxfield.SetItemChecked(i, false);
                }
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < checkedListBoxfield.Items.Count; i++)
            {
                checkedListBoxfield.SetItemChecked(i, !checkedListBoxfield.GetItemChecked(i));
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            ofd.Filter = "Word files (*.doc,*.docx)|*.doc;*.docx";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                textBoxtemplatefile.Text = ofd.FileName;
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            LocalDirDlg DirDlg = new LocalDirDlg();
            if (DirDlg.DisplayDialog("选择输出文件夹") == DialogResult.OK)
                textBoxoutputpath.Text = DirDlg.Path;
            else
                return;
        }

        private void checkBoxEnablePhoto_CheckedChanged(object sender, EventArgs e)
        {
            textBoxphotopath.Enabled = checkBoxEnablePhoto.Checked;
            button8.Enabled = checkBoxEnablePhoto.Checked;
            comboBoxPhotoExt.Enabled = checkBoxEnablePhoto.Checked;
            comboBoxPhotoField.Enabled = checkBoxEnablePhoto.Checked;
            checkBoxSpeciayImageSize.Enabled=checkBoxEnablePhoto.Checked;
            textBoxPicLength.Enabled = textBoxPicWidth.Enabled = checkBoxSpeciayImageSize.Checked && checkBoxEnablePhoto.Checked;

        }

        private void checkBoxSpeciayImageSize_CheckedChanged(object sender, EventArgs e)
        {
            textBoxPicLength.Enabled = textBoxPicWidth.Enabled = checkBoxSpeciayImageSize.Checked && checkBoxEnablePhoto.Checked;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if(checkBoxSpeciayImageSize.Enabled&& checkBoxSpeciayImageSize.Checked)
            {
                if (!Function.IsNumberic(textBoxPicLength.Text) || !Function.IsNumberic(textBoxPicWidth.Text))
                {
                    MessageBox.Show("照片尺寸必须为数字");
                    return;
                }
            }
            string composefilepath, oneoutputfile;
            List<String> allfile = new List<String>();
            if (datatable == null || datatable.Rows.Count == 0)
            {
                textBoxlog.Text = "内存中无任何数据";
                return;
            }
            textBoxlog.Clear();
            bool isSuccess;
            string remark;
            int successNum = 0, failNum = 0, skipNum = 0;

            textBoxphotopath.Text = removeLastSlash(textBoxphotopath.Text);
            textBoxoutputpath.Text = removeLastSlash(textBoxoutputpath.Text);
            int count = datatable.Rows.Count;
            string lastGroupName = "";
            InsertPictureOption insertPictureOption = new InsertPictureOption();
            insertPictureOption.InsertPicture = checkBoxEnablePhoto.Checked;
            if (insertPictureOption.InsertPicture)
            {
                insertPictureOption.PhotoField = comboBoxPhotoField.Text;
                insertPictureOption.PhotoExt = comboBoxPhotoExt.Text;
                if (checkBoxSpeciayImageSize.Checked)
                {
                    insertPictureOption.PictureLength = Convert.ToSingle(textBoxPicLength.Text);
                    insertPictureOption.PictureWidth = Convert.ToSingle(textBoxPicWidth.Text);
                }
                insertPictureOption.PhotoFilePath = textBoxphotopath.Text;
            }
            for (int i = 0; i < count; i++)
            {
                if (datatable.Rows[i][comboBoxFilename.Text].ToString().Trim().Length == 0)
                {
                    skipNum++;
                    continue;
                }
                WordGenerator wg = new WordGenerator();
                oneoutputfile = wg.GenWordFileFromDatatable(datatable.Rows[i], GetCheckedListBoxStringList(checkedListBoxfield), textBoxtemplatefile.Text, textBoxoutputpath.Text, comboBoxFilename.Text, checkBoxSingleFileOutputPdf.Checked, insertPictureOption, out isSuccess, out remark);
                if (oneoutputfile != null && oneoutputfile.Length > 0 && isSuccess)
                {
                    if (checkBoxGenInGroup.Checked)
                    {
                        if (lastGroupName == "")
                        {
                            lastGroupName = datatable.Rows[i][comboBoxGroupName.Text].ToString().Trim();
                        }
                        else if (lastGroupName != datatable.Rows[i][comboBoxGroupName.Text].ToString().Trim())
                        {
                            WordDocumentMerger wdm = new WordDocumentMerger();
                            composefilepath = string.Format("{0}\\{1}{2}.docx", textBoxoutputpath.Text, lastGroupName, DateTime.Now.ToString("yyyyMMddHHmmss"));
                            wdm.InsertMerge(allfile, composefilepath, checkBoxComposeFileOutPutPDF.Checked);
                            lastGroupName = datatable.Rows[i][comboBoxGroupName.Text].ToString().Trim();
                            allfile.Clear();
                        }
                    }
                    allfile.Add(oneoutputfile);
                }
                if (isSuccess)
                {
                    successNum++;
                }
                else
                {
                    failNum++;
                    textBoxlog.Text = textBoxlog.Text + remark + Environment.NewLine;
                }
            }
            if (checkBoxComposeFile.Checked)
            {
                if (checkBoxGenInGroup.Checked)
                    composefilepath = string.Format("{0}\\{1}{2}.docx", textBoxoutputpath.Text, lastGroupName, DateTime.Now.ToString("yyyyMMddHHmmss"));
                else
                    composefilepath = string.Format("{0}\\{1}.docx", textBoxoutputpath.Text, DateTime.Now.ToString("yyyyMMddHHmmss"));
                WordDocumentMerger wdm = new WordDocumentMerger();
                wdm.InsertMerge(allfile, composefilepath, checkBoxComposeFileOutPutPDF.Checked);
            }
            textBoxlog.Text = textBoxlog.Text + String.Format("成功生成{0}条，失败生成{1}条，跳过数据{2}条{3}", successNum, failNum, skipNum, Environment.NewLine);
        }

        private string removeLastSlash(string str)
        {
            if (str.Length == 0)
                return str;
            while (str.Substring(str.Length - 1, 1) == "\\" || str.Substring(str.Length - 1, 1) == "/")
            {
                str = str.Substring(0, str.Length - 1);
            }
            return str;
        }

        private void button8_Click(object sender, EventArgs e)
        {
            LocalDirDlg DirDlg = new LocalDirDlg();
            if (DirDlg.DisplayDialog("选择图片所在文件夹") == DialogResult.OK)
                textBoxphotopath.Text = DirDlg.Path;
            else
                return;
        }

        private IList<String> GetCheckedListBoxStringList(CheckedListBox clb)
        {
            List<String> alldata = new List<string>();
            for (int i = 0; i < clb.Items.Count; i++)
            {
                if (clb.GetItemChecked(i))
                {
                    alldata.Add(clb.Items[i].ToString());
                }
            }
            return alldata;
        }

        private void checkBoxComposeFile_CheckStateChanged(object sender, EventArgs e)
        {
            checkBoxComposeFileOutPutPDF.Enabled = checkBoxComposeFile.Checked;
            checkBoxGenInGroup.Enabled=checkBoxComposeFile.Checked;
            comboBoxGroupName.Enabled = checkBoxComposeFile.Checked;
        }
    }
}