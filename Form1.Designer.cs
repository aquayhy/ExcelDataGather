namespace ExcelDataGather
{
    partial class Form1
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.comboBoxSheetName = new System.Windows.Forms.ComboBox();
            this.button1 = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.textBoxfilename = new System.Windows.Forms.TextBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox5 = new System.Windows.Forms.GroupBox();
            this.label11 = new System.Windows.Forms.Label();
            this.comboBoxFilename = new System.Windows.Forms.ComboBox();
            this.checkBoxSingleFileOutputPdf = new System.Windows.Forms.CheckBox();
            this.label9 = new System.Windows.Forms.Label();
            this.button9 = new System.Windows.Forms.Button();
            this.button7 = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.textBoxoutputpath = new System.Windows.Forms.TextBox();
            this.textBoxtemplatefile = new System.Windows.Forms.TextBox();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.comboBoxGroupName = new System.Windows.Forms.ComboBox();
            this.checkBoxGenInGroup = new System.Windows.Forms.CheckBox();
            this.checkBoxComposeFile = new System.Windows.Forms.CheckBox();
            this.label10 = new System.Windows.Forms.Label();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.textBoxlog = new System.Windows.Forms.TextBox();
            this.button6 = new System.Windows.Forms.Button();
            this.button5 = new System.Windows.Forms.Button();
            this.button4 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.checkBoxEnablePhoto = new System.Windows.Forms.CheckBox();
            this.button8 = new System.Windows.Forms.Button();
            this.comboBoxPhotoExt = new System.Windows.Forms.ComboBox();
            this.comboBoxPhotoField = new System.Windows.Forms.ComboBox();
            this.textBoxphotopath = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.checkedListBoxfield = new System.Windows.Forms.CheckedListBox();
            this.label3 = new System.Windows.Forms.Label();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.label12 = new System.Windows.Forms.Label();
            this.checkBoxSpeciayImageSize = new System.Windows.Forms.CheckBox();
            this.textBoxPicLength = new System.Windows.Forms.TextBox();
            this.label13 = new System.Windows.Forms.Label();
            this.textBoxPicWidth = new System.Windows.Forms.TextBox();
            this.label14 = new System.Windows.Forms.Label();
            this.label15 = new System.Windows.Forms.Label();
            this.checkBoxComposeFileOutPutPDF = new System.Windows.Forms.CheckBox();
            this.groupBox1.SuspendLayout();
            this.groupBox5.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // comboBoxSheetName
            // 
            this.comboBoxSheetName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBoxSheetName.FormattingEnabled = true;
            this.comboBoxSheetName.Location = new System.Drawing.Point(660, 13);
            this.comboBoxSheetName.Name = "comboBoxSheetName";
            this.comboBoxSheetName.Size = new System.Drawing.Size(209, 20);
            this.comboBoxSheetName.TabIndex = 8;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(468, 11);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 7;
            this.button1.Text = "选择文件";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(613, 16);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(41, 12);
            this.label2.TabIndex = 5;
            this.label2.Text = "选择表";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(10, 16);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(83, 12);
            this.label1.TabIndex = 6;
            this.label1.Text = "选择Excel文件";
            // 
            // textBoxfilename
            // 
            this.textBoxfilename.Location = new System.Drawing.Point(99, 12);
            this.textBoxfilename.Name = "textBoxfilename";
            this.textBoxfilename.ReadOnly = true;
            this.textBoxfilename.Size = new System.Drawing.Size(360, 21);
            this.textBoxfilename.TabIndex = 4;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.groupBox5);
            this.groupBox1.Controls.Add(this.groupBox4);
            this.groupBox1.Controls.Add(this.groupBox3);
            this.groupBox1.Controls.Add(this.button6);
            this.groupBox1.Controls.Add(this.button5);
            this.groupBox1.Controls.Add(this.button4);
            this.groupBox1.Controls.Add(this.button3);
            this.groupBox1.Controls.Add(this.button2);
            this.groupBox1.Controls.Add(this.groupBox2);
            this.groupBox1.Controls.Add(this.checkedListBoxfield);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Location = new System.Drawing.Point(14, 308);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(1096, 348);
            this.groupBox1.TabIndex = 10;
            this.groupBox1.TabStop = false;
            // 
            // groupBox5
            // 
            this.groupBox5.Controls.Add(this.label11);
            this.groupBox5.Controls.Add(this.comboBoxFilename);
            this.groupBox5.Controls.Add(this.checkBoxSingleFileOutputPdf);
            this.groupBox5.Controls.Add(this.label9);
            this.groupBox5.Controls.Add(this.button9);
            this.groupBox5.Controls.Add(this.button7);
            this.groupBox5.Controls.Add(this.label4);
            this.groupBox5.Controls.Add(this.label6);
            this.groupBox5.Controls.Add(this.textBoxoutputpath);
            this.groupBox5.Controls.Add(this.textBoxtemplatefile);
            this.groupBox5.Location = new System.Drawing.Point(366, 20);
            this.groupBox5.Name = "groupBox5";
            this.groupBox5.Size = new System.Drawing.Size(354, 135);
            this.groupBox5.TabIndex = 23;
            this.groupBox5.TabStop = false;
            this.groupBox5.Text = "输出文件选项";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(115, 114);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(221, 12);
            this.label11.TabIndex = 22;
            this.label11.Text = "<--默认生成word文件，勾选后会输出PDF";
            // 
            // comboBoxFilename
            // 
            this.comboBoxFilename.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBoxFilename.FormattingEnabled = true;
            this.comboBoxFilename.Location = new System.Drawing.Point(63, 84);
            this.comboBoxFilename.Name = "comboBoxFilename";
            this.comboBoxFilename.Size = new System.Drawing.Size(281, 20);
            this.comboBoxFilename.TabIndex = 3;
            // 
            // checkBoxSingleFileOutputPdf
            // 
            this.checkBoxSingleFileOutputPdf.AutoSize = true;
            this.checkBoxSingleFileOutputPdf.Location = new System.Drawing.Point(7, 114);
            this.checkBoxSingleFileOutputPdf.Name = "checkBoxSingleFileOutputPdf";
            this.checkBoxSingleFileOutputPdf.Size = new System.Drawing.Size(102, 16);
            this.checkBoxSingleFileOutputPdf.TabIndex = 18;
            this.checkBoxSingleFileOutputPdf.Text = "单文件输出PDF";
            this.checkBoxSingleFileOutputPdf.UseVisualStyleBackColor = true;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(4, 87);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(53, 12);
            this.label9.TabIndex = 11;
            this.label9.Text = "命名字段";
            // 
            // button9
            // 
            this.button9.Location = new System.Drawing.Point(289, 49);
            this.button9.Name = "button9";
            this.button9.Size = new System.Drawing.Size(55, 23);
            this.button9.TabIndex = 13;
            this.button9.Text = "选择";
            this.button9.UseVisualStyleBackColor = true;
            this.button9.Click += new System.EventHandler(this.button9_Click);
            // 
            // button7
            // 
            this.button7.Location = new System.Drawing.Point(289, 16);
            this.button7.Name = "button7";
            this.button7.Size = new System.Drawing.Size(55, 23);
            this.button7.TabIndex = 15;
            this.button7.Text = "选择";
            this.button7.UseVisualStyleBackColor = true;
            this.button7.Click += new System.EventHandler(this.button7_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(3, 54);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(53, 12);
            this.label4.TabIndex = 10;
            this.label4.Text = "输出目录";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(4, 21);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(53, 12);
            this.label6.TabIndex = 12;
            this.label6.Text = "模板文件";
            // 
            // textBoxoutputpath
            // 
            this.textBoxoutputpath.Location = new System.Drawing.Point(63, 51);
            this.textBoxoutputpath.Name = "textBoxoutputpath";
            this.textBoxoutputpath.Size = new System.Drawing.Size(220, 21);
            this.textBoxoutputpath.TabIndex = 7;
            // 
            // textBoxtemplatefile
            // 
            this.textBoxtemplatefile.Location = new System.Drawing.Point(63, 18);
            this.textBoxtemplatefile.Name = "textBoxtemplatefile";
            this.textBoxtemplatefile.Size = new System.Drawing.Size(220, 21);
            this.textBoxtemplatefile.TabIndex = 9;
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.comboBoxGroupName);
            this.groupBox4.Controls.Add(this.checkBoxGenInGroup);
            this.groupBox4.Controls.Add(this.checkBoxComposeFileOutPutPDF);
            this.groupBox4.Controls.Add(this.checkBoxComposeFile);
            this.groupBox4.Controls.Add(this.label10);
            this.groupBox4.Location = new System.Drawing.Point(365, 158);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(355, 71);
            this.groupBox4.TabIndex = 21;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "文件合并选项";
            // 
            // comboBoxGroupName
            // 
            this.comboBoxGroupName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBoxGroupName.FormattingEnabled = true;
            this.comboBoxGroupName.Location = new System.Drawing.Point(143, 41);
            this.comboBoxGroupName.Name = "comboBoxGroupName";
            this.comboBoxGroupName.Size = new System.Drawing.Size(198, 20);
            this.comboBoxGroupName.TabIndex = 3;
            // 
            // checkBoxGenInGroup
            // 
            this.checkBoxGenInGroup.AutoSize = true;
            this.checkBoxGenInGroup.Location = new System.Drawing.Point(10, 43);
            this.checkBoxGenInGroup.Name = "checkBoxGenInGroup";
            this.checkBoxGenInGroup.Size = new System.Drawing.Size(72, 16);
            this.checkBoxGenInGroup.TabIndex = 18;
            this.checkBoxGenInGroup.Text = "分组输出";
            this.checkBoxGenInGroup.UseVisualStyleBackColor = true;
            // 
            // checkBoxComposeFile
            // 
            this.checkBoxComposeFile.AutoSize = true;
            this.checkBoxComposeFile.Checked = true;
            this.checkBoxComposeFile.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxComposeFile.Location = new System.Drawing.Point(10, 18);
            this.checkBoxComposeFile.Name = "checkBoxComposeFile";
            this.checkBoxComposeFile.Size = new System.Drawing.Size(72, 16);
            this.checkBoxComposeFile.TabIndex = 18;
            this.checkBoxComposeFile.Text = "合成文件";
            this.checkBoxComposeFile.UseVisualStyleBackColor = true;
            this.checkBoxComposeFile.CheckStateChanged += new System.EventHandler(this.checkBoxComposeFile_CheckStateChanged);
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(82, 45);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(53, 12);
            this.label10.TabIndex = 11;
            this.label10.Text = "分组依据";
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.textBoxlog);
            this.groupBox3.Location = new System.Drawing.Point(369, 238);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(718, 92);
            this.groupBox3.TabIndex = 20;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "日记";
            // 
            // textBoxlog
            // 
            this.textBoxlog.Location = new System.Drawing.Point(6, 17);
            this.textBoxlog.Multiline = true;
            this.textBoxlog.Name = "textBoxlog";
            this.textBoxlog.ReadOnly = true;
            this.textBoxlog.Size = new System.Drawing.Size(706, 68);
            this.textBoxlog.TabIndex = 0;
            // 
            // button6
            // 
            this.button6.Location = new System.Drawing.Point(966, 209);
            this.button6.Name = "button6";
            this.button6.Size = new System.Drawing.Size(75, 23);
            this.button6.TabIndex = 19;
            this.button6.Text = "清空日记";
            this.button6.UseVisualStyleBackColor = true;
            // 
            // button5
            // 
            this.button5.Location = new System.Drawing.Point(810, 209);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(75, 23);
            this.button5.TabIndex = 19;
            this.button5.Text = "开始工作";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(285, 310);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(75, 23);
            this.button4.TabIndex = 17;
            this.button4.Text = "反选";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(149, 310);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(75, 23);
            this.button3.TabIndex = 17;
            this.button3.Text = "全取消";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(8, 310);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(75, 23);
            this.button2.TabIndex = 17;
            this.button2.Text = "全选";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.checkBoxSpeciayImageSize);
            this.groupBox2.Controls.Add(this.checkBoxEnablePhoto);
            this.groupBox2.Controls.Add(this.button8);
            this.groupBox2.Controls.Add(this.comboBoxPhotoExt);
            this.groupBox2.Controls.Add(this.comboBoxPhotoField);
            this.groupBox2.Controls.Add(this.textBoxphotopath);
            this.groupBox2.Controls.Add(this.label15);
            this.groupBox2.Controls.Add(this.label13);
            this.groupBox2.Controls.Add(this.label14);
            this.groupBox2.Controls.Add(this.label12);
            this.groupBox2.Controls.Add(this.label8);
            this.groupBox2.Controls.Add(this.textBoxPicWidth);
            this.groupBox2.Controls.Add(this.textBoxPicLength);
            this.groupBox2.Controls.Add(this.label7);
            this.groupBox2.Controls.Add(this.label5);
            this.groupBox2.Location = new System.Drawing.Point(727, 20);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(354, 183);
            this.groupBox2.TabIndex = 16;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "照片选项";
            // 
            // checkBoxEnablePhoto
            // 
            this.checkBoxEnablePhoto.AutoSize = true;
            this.checkBoxEnablePhoto.Checked = true;
            this.checkBoxEnablePhoto.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxEnablePhoto.Location = new System.Drawing.Point(13, 22);
            this.checkBoxEnablePhoto.Name = "checkBoxEnablePhoto";
            this.checkBoxEnablePhoto.Size = new System.Drawing.Size(312, 16);
            this.checkBoxEnablePhoto.TabIndex = 15;
            this.checkBoxEnablePhoto.Text = "启动照片替换（照片使用书签和文本框在word中实现）";
            this.checkBoxEnablePhoto.UseVisualStyleBackColor = true;
            this.checkBoxEnablePhoto.CheckedChanged += new System.EventHandler(this.checkBoxEnablePhoto_CheckedChanged);
            // 
            // button8
            // 
            this.button8.Location = new System.Drawing.Point(296, 41);
            this.button8.Name = "button8";
            this.button8.Size = new System.Drawing.Size(55, 23);
            this.button8.TabIndex = 14;
            this.button8.Text = "选择";
            this.button8.UseVisualStyleBackColor = true;
            this.button8.Click += new System.EventHandler(this.button8_Click);
            // 
            // comboBoxPhotoExt
            // 
            this.comboBoxPhotoExt.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBoxPhotoExt.FormattingEnabled = true;
            this.comboBoxPhotoExt.Items.AddRange(new object[] {
            ".jpg",
            ".jpeg",
            ".bmp",
            ".png"});
            this.comboBoxPhotoExt.Location = new System.Drawing.Point(71, 99);
            this.comboBoxPhotoExt.Name = "comboBoxPhotoExt";
            this.comboBoxPhotoExt.Size = new System.Drawing.Size(280, 20);
            this.comboBoxPhotoExt.TabIndex = 3;
            // 
            // comboBoxPhotoField
            // 
            this.comboBoxPhotoField.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBoxPhotoField.FormattingEnabled = true;
            this.comboBoxPhotoField.Location = new System.Drawing.Point(70, 70);
            this.comboBoxPhotoField.Name = "comboBoxPhotoField";
            this.comboBoxPhotoField.Size = new System.Drawing.Size(281, 20);
            this.comboBoxPhotoField.TabIndex = 3;
            // 
            // textBoxphotopath
            // 
            this.textBoxphotopath.Location = new System.Drawing.Point(70, 43);
            this.textBoxphotopath.Name = "textBoxphotopath";
            this.textBoxphotopath.Size = new System.Drawing.Size(220, 21);
            this.textBoxphotopath.TabIndex = 8;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(10, 102);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(53, 12);
            this.label8.TabIndex = 11;
            this.label8.Text = "照片后缀";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(11, 73);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(53, 12);
            this.label7.TabIndex = 11;
            this.label7.Text = "名称字段";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(11, 46);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(53, 12);
            this.label5.TabIndex = 11;
            this.label5.Text = "照片目录";
            // 
            // checkedListBoxfield
            // 
            this.checkedListBoxfield.CheckOnClick = true;
            this.checkedListBoxfield.FormattingEnabled = true;
            this.checkedListBoxfield.Location = new System.Drawing.Point(6, 29);
            this.checkedListBoxfield.Name = "checkedListBoxfield";
            this.checkedListBoxfield.Size = new System.Drawing.Size(354, 276);
            this.checkedListBoxfield.TabIndex = 1;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(9, 14);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(77, 12);
            this.label3.TabIndex = 0;
            this.label3.Text = "选择替换字段";
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.AllowUserToOrderColumns = true;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(12, 40);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.ReadOnly = true;
            this.dataGridView1.RowTemplate.Height = 23;
            this.dataGridView1.Size = new System.Drawing.Size(1099, 262);
            this.dataGridView1.TabIndex = 9;
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(10, 156);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(17, 12);
            this.label12.TabIndex = 11;
            this.label12.Text = "长";
            // 
            // checkBoxSpeciayImageSize
            // 
            this.checkBoxSpeciayImageSize.AutoSize = true;
            this.checkBoxSpeciayImageSize.Checked = true;
            this.checkBoxSpeciayImageSize.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxSpeciayImageSize.Location = new System.Drawing.Point(12, 129);
            this.checkBoxSpeciayImageSize.Name = "checkBoxSpeciayImageSize";
            this.checkBoxSpeciayImageSize.Size = new System.Drawing.Size(96, 16);
            this.checkBoxSpeciayImageSize.TabIndex = 15;
            this.checkBoxSpeciayImageSize.Text = "指定图片大小";
            this.checkBoxSpeciayImageSize.UseVisualStyleBackColor = true;
            this.checkBoxSpeciayImageSize.CheckedChanged += new System.EventHandler(this.checkBoxSpeciayImageSize_CheckedChanged);
            // 
            // textBoxPicLength
            // 
            this.textBoxPicLength.Location = new System.Drawing.Point(33, 151);
            this.textBoxPicLength.Name = "textBoxPicLength";
            this.textBoxPicLength.Size = new System.Drawing.Size(64, 21);
            this.textBoxPicLength.TabIndex = 7;
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(103, 155);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(29, 12);
            this.label13.TabIndex = 11;
            this.label13.Text = "毫米";
            // 
            // textBoxPicWidth
            // 
            this.textBoxPicWidth.Location = new System.Drawing.Point(195, 151);
            this.textBoxPicWidth.Name = "textBoxPicWidth";
            this.textBoxPicWidth.Size = new System.Drawing.Size(64, 21);
            this.textBoxPicWidth.TabIndex = 7;
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Location = new System.Drawing.Point(172, 156);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(17, 12);
            this.label14.TabIndex = 11;
            this.label14.Text = "宽";
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Location = new System.Drawing.Point(265, 155);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(29, 12);
            this.label15.TabIndex = 11;
            this.label15.Text = "毫米";
            // 
            // checkBoxComposeFileOutPutPDF
            // 
            this.checkBoxComposeFileOutPutPDF.AutoSize = true;
            this.checkBoxComposeFileOutPutPDF.Location = new System.Drawing.Point(143, 17);
            this.checkBoxComposeFileOutPutPDF.Name = "checkBoxComposeFileOutPutPDF";
            this.checkBoxComposeFileOutPutPDF.Size = new System.Drawing.Size(114, 16);
            this.checkBoxComposeFileOutPutPDF.TabIndex = 18;
            this.checkBoxComposeFileOutPutPDF.Text = "合成文件输出PDF";
            this.checkBoxComposeFileOutPutPDF.UseVisualStyleBackColor = true;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1125, 671);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.comboBoxSheetName);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.textBoxfilename);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Form1";
            this.Text = "Word文件批量生成器";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox5.ResumeLayout(false);
            this.groupBox5.PerformLayout();
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox comboBoxSheetName;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textBoxfilename;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.TextBox textBoxlog;
        private System.Windows.Forms.Button button6;
        private System.Windows.Forms.Button button5;
        private System.Windows.Forms.ComboBox comboBoxGroupName;
        private System.Windows.Forms.ComboBox comboBoxFilename;
        private System.Windows.Forms.CheckBox checkBoxSingleFileOutputPdf;
        private System.Windows.Forms.CheckBox checkBoxGenInGroup;
        private System.Windows.Forms.CheckBox checkBoxComposeFile;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.CheckBox checkBoxEnablePhoto;
        private System.Windows.Forms.Button button8;
        private System.Windows.Forms.ComboBox comboBoxPhotoExt;
        private System.Windows.Forms.ComboBox comboBoxPhotoField;
        private System.Windows.Forms.TextBox textBoxphotopath;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button button9;
        private System.Windows.Forms.Button button7;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox textBoxoutputpath;
        private System.Windows.Forms.TextBox textBoxtemplatefile;
        private System.Windows.Forms.CheckedListBox checkedListBoxfield;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.GroupBox groupBox5;
        private System.Windows.Forms.CheckBox checkBoxSpeciayImageSize;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.Label label15;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.TextBox textBoxPicWidth;
        private System.Windows.Forms.TextBox textBoxPicLength;
        private System.Windows.Forms.CheckBox checkBoxComposeFileOutPutPDF;
    }
}

