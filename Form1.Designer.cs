using static Generator.Program;
using System.IO;
using System.Diagnostics.Eventing.Reader;

namespace Generator
{
    partial class Form1
    {
        private string mode = string.Empty;

        private string part0File = Directory.GetCurrentDirectory() + "\\Templates\\Part0-Template.DOC";

        private string part1File = string.Empty;

        private string part1mode1File = Directory.GetCurrentDirectory() + "\\Templates\\Part1-Template1.DOC";
        private string part1mode2File = Directory.GetCurrentDirectory() + "\\Templates\\Part1-Template2.DOC";
        private string part1mode3File = Directory.GetCurrentDirectory() + "\\Templates\\Part1-Template3.DOC";
        private string part1mode4File = Directory.GetCurrentDirectory() + "\\Templates\\Part1-Template4.DOC";
        private string part1mode5File = Directory.GetCurrentDirectory() + "\\Templates\\Part1-Template5.DOC";

        private string part2File = Directory.GetCurrentDirectory() + "\\Templates\\Part2-Template.DOC";

        private string destFile = string.Empty;

        private DateTime endDate;
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void Form1_Load1(object sender, EventArgs e)
        {
            DriveInfo[] drives = DriveInfo.GetDrives();
            foreach (DriveInfo drive in drives)
            {
                if (drive.DriveType == DriveType.Fixed)
                {
                    comboBox2.Items.Add(drive.Name);
                }
            }

            for (int i = 1; i <= 20; i++)
            {
                TextBox textbox;
                string textBoxStr = "textBox" + (i * 2 + 1).ToString();
                textbox = this.Controls.Find(textBoxStr, true).FirstOrDefault() as TextBox;
                textbox.Enabled = false;
            }
        }

        private void generate_Click(object sender, EventArgs e)
        {
            if (check_Values())
            {
                button1.Enabled = false;
                button2.Enabled = false;
                button1.Text = "生成中...";
                generate_Contract();
            }
        }
        private void close_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        private bool check_Values()
        {
            endDate = dateTimePicker1.Value;

            if (textBox1.Text == "")
            {
                MessageBox.Show("请输入新合同名称！");
                return false;
            }
            else if (textBox2.Text == "")
            {
                MessageBox.Show("请输入基金名称！");
                return false;
            }
            else if (comboBox1.Text == "")
            {
                MessageBox.Show("请选择合同模板！");
                return false;
            }
            else if (comboBox2.Text == "")
            {
                MessageBox.Show("请选择存储路径！");
                return false;
            }
            else if (DateTime.Compare(endDate, DateTime.Now) <= 0)
            {
                MessageBox.Show("请选择有效的投资运作到期日！");
                return false;
            }
            else
            {
                destFile = comboBox2.Text + textBox1.Text.Trim() + ".doc";
                part1File = comboBox2.Text + textBox1.Text.Trim() + ".part1.doc";
                return true; 
            }
        }

        private void generate_Contract()
        {
            string aShareTime;
            string cShareTime;
            string aShareInvest;
            string cShareInvest;
            string interesTime;
            string invEndDate = endDate.Year.ToString() + "年" + endDate.Month.ToString() + "月" + endDate.Day.ToString() + "日";

            if (mode == "mode1")
            {
                aShareTime = textBox24.Text.Trim();
                cShareTime = textBox26.Text.Trim();
                aShareInvest = textBox28.Text.Trim();
                cShareInvest = textBox30.Text.Trim();
                interesTime = textBox32.Text.Trim();

                if (check_shareTime(aShareTime) & check_shareTime(cShareTime) & check_returnRate(aShareInvest) & check_returnRate(cShareInvest) & check_interestTime(interesTime))
                {
                    WordOperate wordOperate = new WordOperate();
                    wordOperate.Open_Tem(part1mode1File);

                    wordOperate.Replace("Asharetime", aShareTime);
                    wordOperate.Replace("Csharetime", cShareTime);
                    wordOperate.Replace("Ashareinvest", aShareInvest);
                    wordOperate.Replace("Cshareinvest", cShareInvest);
                    wordOperate.Replace("interesttime", interesTime);
                    wordOperate.Replace("Enddate", invEndDate);

                    wordOperate.Save(part1File);

                    wordOperate.Open_Standard();

                    wordOperate.Combine(part0File, part1File, part2File);

                    wordOperate.Replace("fundname", textBox2.Text.Trim());
                    wordOperate.Replace("trustcorp", textBox6.Text.Trim());
                    wordOperate.Replace("trustplan", textBox8.Text.Trim());
                    wordOperate.Replace("accountname", textBox10.Text.Trim());
                    wordOperate.Replace("accountno", textBox12.Text.Trim());
                    wordOperate.Replace("bank", textBox14.Text.Trim());
                    wordOperate.Replace("payno", textBox16.Text.Trim());
                    wordOperate.Replace("managername", textBox18.Text.Trim());
                    wordOperate.Replace("managerphoneno", textBox20.Text.Trim());
                    wordOperate.Replace("managermail", textBox22.Text.Trim());

                    wordOperate.addPageNo();

                    wordOperate.update_page();

                    wordOperate.Save(destFile);

                    button1.Text = "生成";
                    MessageBox.Show("生成成功!    文件保存在 " + destFile);
                    button1.Enabled = true;
                    button2.Enabled = true;
                }
                else
                {
                    button1.Text = "生成";
                    button1.Enabled = true;
                    button2.Enabled = true;
                    return;
                }
            }

            else if (mode == "mode2")
            {
                aShareTime = textBox24.Text.Trim();
                cShareTime = textBox26.Text.Trim();
                aShareInvest = textBox28.Text.Trim();
                cShareInvest = textBox30.Text.Trim();
                interesTime = textBox32.Text.Trim();

                if (check_shareTime(aShareTime) & check_shareTime(cShareTime) & check_returnRate(aShareInvest) & check_returnRate(cShareInvest) & check_interestTime(interesTime))
                {
                    WordOperate wordOperate = new WordOperate();
                    wordOperate.Open_Tem(part1mode2File);

                    wordOperate.Replace("Asharetime", aShareTime);
                    wordOperate.Replace("Csharetime", cShareTime);
                    wordOperate.Replace("Ashareinvest", aShareInvest);
                    wordOperate.Replace("Cshareinvest", cShareInvest);
                    wordOperate.Replace("interesttime", interesTime);
                    wordOperate.Replace("Enddate", invEndDate);
                    wordOperate.Replace("shareholder", textBox36.Text.Trim());

                    wordOperate.Save(part1File);

                    wordOperate.Open_Standard();

                    wordOperate.Combine(part0File, part1File, part2File);

                    wordOperate.Replace("fundname", textBox2.Text.Trim());
                    wordOperate.Replace("trustcorp", textBox6.Text.Trim());
                    wordOperate.Replace("trustplan", textBox8.Text.Trim());
                    wordOperate.Replace("accountname", textBox10.Text.Trim());
                    wordOperate.Replace("accountno", textBox12.Text.Trim());
                    wordOperate.Replace("bank", textBox14.Text.Trim());
                    wordOperate.Replace("payno", textBox16.Text.Trim());
                    wordOperate.Replace("managername", textBox18.Text.Trim());
                    wordOperate.Replace("managerphoneno", textBox20.Text.Trim());
                    wordOperate.Replace("managermail", textBox22.Text.Trim());

                    wordOperate.addPageNo();

                    wordOperate.update_page();

                    wordOperate.Save(destFile);

                    button1.Text = "生成";
                    MessageBox.Show("生成成功!    文件保存在 " + destFile);
                    button1.Enabled = true;
                    button2.Enabled = true;
                }
                else
                {
                    button1.Text = "生成";
                    button1.Enabled = true;
                    button2.Enabled = true;
                    return;
                }
            }

            else if (mode == "mode3")
            {
                aShareTime = textBox24.Text.Trim();
                aShareInvest = textBox28.Text.Trim();
                interesTime = textBox32.Text.Trim();

                if (check_shareTime(aShareTime) & check_returnRate(aShareInvest) & check_interestTime(interesTime))
                {
                    WordOperate wordOperate = new WordOperate();
                    wordOperate.Open_Tem(part1mode3File);

                    wordOperate.Replace("Asharetime", aShareTime);
                    wordOperate.Replace("Ashareinvest", aShareInvest);
                    wordOperate.Replace("interesttime", interesTime);
                    wordOperate.Replace("Enddate", invEndDate);
                    wordOperate.Replace("assetholder", textBox38.Text.Trim());
                    wordOperate.Replace("debtor", textBox40.Text.Trim());

                    wordOperate.Save(part1File);

                    wordOperate.Open_Standard();

                    wordOperate.Combine(part0File, part1File, part2File);

                    wordOperate.Replace("fundname", textBox2.Text.Trim());
                    wordOperate.Replace("accountname", textBox10.Text.Trim());
                    wordOperate.Replace("accountno", textBox12.Text.Trim());
                    wordOperate.Replace("bank", textBox14.Text.Trim());
                    wordOperate.Replace("payno", textBox16.Text.Trim());
                    wordOperate.Replace("managername", textBox18.Text.Trim());
                    wordOperate.Replace("managerphoneno", textBox20.Text.Trim());
                    wordOperate.Replace("managermail", textBox22.Text.Trim());

                    wordOperate.addPageNo();

                    wordOperate.update_page();

                    wordOperate.Save(destFile);

                    button1.Text = "生成";
                    MessageBox.Show("生成成功!    文件保存在 " + destFile);
                    button1.Enabled = true;
                    button2.Enabled = true;
                }
                else
                {
                    button1.Text = "生成";
                    button1.Enabled = true;
                    button2.Enabled = true;
                    return;
                }
            }

            else if (mode == "mode4")
            {
                aShareTime = textBox24.Text.Trim();
                cShareTime = textBox26.Text.Trim();
                aShareInvest = textBox28.Text.Trim();
                cShareInvest = textBox30.Text.Trim();
                interesTime = textBox32.Text.Trim();

                if (check_shareTime(aShareTime) & check_shareTime(cShareTime) & check_returnRate(aShareInvest) & check_returnRate(cShareInvest) & check_interestTime(interesTime))
                {
                    WordOperate wordOperate = new WordOperate();
                    wordOperate.Open_Tem(part1mode4File);

                    wordOperate.Replace("Asharetime", aShareTime);
                    wordOperate.Replace("Csharetime", cShareTime);
                    wordOperate.Replace("Ashareinvest", aShareInvest);
                    wordOperate.Replace("Cshareinvest", cShareInvest);
                    wordOperate.Replace("interesttime", interesTime);
                    wordOperate.Replace("Enddate", invEndDate);

                    wordOperate.Save(part1File);

                    wordOperate.Open_Standard();

                    wordOperate.Combine(part0File, part1File, part2File);

                    wordOperate.Replace("fundname", textBox2.Text.Trim());
                    wordOperate.Replace("trustcorp", textBox6.Text.Trim());
                    wordOperate.Replace("trustplan", textBox8.Text.Trim());
                    wordOperate.Replace("accountname", textBox10.Text.Trim());
                    wordOperate.Replace("accountno", textBox12.Text.Trim());
                    wordOperate.Replace("bank", textBox14.Text.Trim());
                    wordOperate.Replace("payno", textBox16.Text.Trim());
                    wordOperate.Replace("managername", textBox18.Text.Trim());
                    wordOperate.Replace("managerphoneno", textBox20.Text.Trim());
                    wordOperate.Replace("managermail", textBox22.Text.Trim());

                    wordOperate.addPageNo();

                    wordOperate.update_page();

                    wordOperate.Save(destFile);

                    button1.Text = "生成";
                    MessageBox.Show("生成成功!    文件保存在 " + destFile);
                    button1.Enabled = true;
                    button2.Enabled = true;
                }
                else
                {
                    button1.Text = "生成";
                    button1.Enabled = true;
                    button2.Enabled = true;
                    return;
                }
            }
            
            else if (mode == "mode5")
            {
                aShareTime = textBox24.Text.Trim();
                cShareTime = textBox26.Text.Trim();
                aShareInvest = textBox28.Text.Trim();
                cShareInvest = textBox30.Text.Trim();
                interesTime = textBox32.Text.Trim();

                if (check_shareTime(aShareTime) & check_shareTime(cShareTime) & check_returnRate(aShareInvest) & check_returnRate(cShareInvest) & check_interestTime(interesTime))
                {
                    WordOperate wordOperate = new WordOperate();
                    wordOperate.Open_Tem(part1mode5File);

                    wordOperate.Replace("Asharetime", aShareTime);
                    wordOperate.Replace("Csharetime", cShareTime);
                    wordOperate.Replace("Ashareinvest", aShareInvest);
                    wordOperate.Replace("Cshareinvest", cShareInvest);
                    wordOperate.Replace("interesttime", interesTime);
                    wordOperate.Replace("interesttime", interesTime);
                    wordOperate.Replace("Enddate", invEndDate);
                    wordOperate.Replace("borrower", textBox42.Text.Trim());

                    wordOperate.Save(part1File);

                    wordOperate.Open_Standard();

                    wordOperate.Combine(part0File, part1File, part2File);

                    wordOperate.Replace("fundname", textBox2.Text.Trim());
                    wordOperate.Replace("accountname", textBox10.Text.Trim());
                    wordOperate.Replace("accountno", textBox12.Text.Trim());
                    wordOperate.Replace("bank", textBox14.Text.Trim());
                    wordOperate.Replace("payno", textBox16.Text.Trim());
                    wordOperate.Replace("managername", textBox18.Text.Trim());
                    wordOperate.Replace("managerphoneno", textBox20.Text.Trim());
                    wordOperate.Replace("managermail", textBox22.Text.Trim());

                    wordOperate.addPageNo();

                    wordOperate.update_page();

                    wordOperate.Save(destFile);

                    button1.Text = "生成";
                    MessageBox.Show("生成成功!    文件保存在 " + destFile);
                    button1.Enabled = true;
                    button2.Enabled = true;
                }
                else
                {
                    button1.Text = "生成";
                    button1.Enabled = true;
                    button2.Enabled = true;
                    return;
                }
            }
        }

            private bool check_shareTime(string shareTime)
            {
                int i = 0;
                if (shareTime != "" && int.TryParse(shareTime, out i))
                {
                    int month = int.Parse(shareTime);
                    if (month <= 0 || month > 36)
                    {
                        MessageBox.Show("请输入1-36范围内的投资期限!");
                        return false;
                    }
                    else
                    {
                        return true;
                    }
                }
                else
                {
                    MessageBox.Show("请输入份额投资期限!");
                    return false;
                }
            }

            private bool check_returnRate(string returnRate)
            {
                int i = 0;
                if (returnRate != "" && int.TryParse(returnRate, out i))
                {
                    int rate = int.Parse(returnRate);
                    if (rate <= 0 || rate > 15)
                    {
                        MessageBox.Show("请输入1-15范围内的年化收益！");
                        return false;
                    }
                    else
                    {
                        return true;
                    }
                }
                else
                {
                    MessageBox.Show("请输入预计年化收益!");
                    return false;
                }
            }

            private bool check_interestTime(string interestTime)
            {
                int i = 0;
                if (interestTime != "" && int.TryParse(interestTime, out i))
                {
                    int period = int.Parse(interestTime);
                    if (period <= 0 || period > 3)
                    {
                        MessageBox.Show("请输入1-3范围内的收益分配周期");
                        return false;
                    }
                    else
                    {
                        return true;
                    }
                }
                else
                {
                    MessageBox.Show("请输入收益分配周期!");
                    return false;
                }
            }

            private void comboBox1_SelectedMode(object sender, EventArgs e)
            {
                if (comboBox1.Text == "信托单位")
                {
                    mode = "mode1";
                    textBox6.Enabled = true;
                    textBox8.Enabled = true;
                    textBox24.Enabled = true;
                    textBox26.Enabled = true;
                    textBox28.Enabled = true;
                    textBox30.Enabled = true;
                    textBox32.Enabled = true;
                    textBox36.Enabled = false;
                    textBox38.Enabled = false;
                    textBox40.Enabled = false;
                    textBox42.Enabled = false;

            }
                else if (comboBox1.Text == "信托受益权")
                {
                    mode = "mode2";
                    textBox6.Enabled = true;
                    textBox8.Enabled = true;
                    textBox24.Enabled = true;
                    textBox26.Enabled = true;
                    textBox28.Enabled = true;
                    textBox30.Enabled = true;
                    textBox32.Enabled = true;
                    textBox36.Enabled = true;
                    textBox38.Enabled = false;
                    textBox40.Enabled = false;
                    textBox42.Enabled = false;
                }
                else if (comboBox1.Text == "资产包")
                {
                    mode = "mode3";
                    textBox6.Enabled = false;
                    textBox8.Enabled = false;
                    textBox24.Enabled = true;
                    textBox26.Enabled = false;
                    textBox28.Enabled = true;
                    textBox30.Enabled = false;
                    textBox32.Enabled = true;
                    textBox36.Enabled = false;
                    textBox38.Enabled = true;
                    textBox40.Enabled = true;
                    textBox42.Enabled = false;
                }
                else if (comboBox1.Text == "资管计划")
                {
                    mode = "mode4";
                    textBox6.Enabled = true;
                    textBox8.Enabled = true;
                    textBox24.Enabled = true;
                    textBox26.Enabled = true;
                    textBox28.Enabled = true;
                    textBox30.Enabled = true;
                    textBox32.Enabled = true;
                    textBox36.Enabled = false;
                    textBox38.Enabled = false;
                    textBox40.Enabled = false;
                    textBox42.Enabled = false;
                }
                else if (comboBox1.Text == "委托贷款")
                {
                    mode = "mode5";
                    textBox6.Enabled = false;
                    textBox8.Enabled = false;
                    textBox24.Enabled = true;
                    textBox26.Enabled = true;
                    textBox28.Enabled = true;
                    textBox30.Enabled = true;
                    textBox32.Enabled = true;
                    textBox36.Enabled = false;
                    textBox38.Enabled = false;
                    textBox40.Enabled = false;
                    textBox42.Enabled = true;
                }
            }

        #region Windows Form Designer generated code
        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>

        private Button button1;
        private Button button2;
        private Label label1;
        private TextBox textBox1;
        private Label label2;
        private Label label3;
        private ColorDialog colorDialog1;
        private ComboBox comboBox1;
        private Label label4;
        private ComboBox comboBox2;
        private Label label5;
        private Label label6;
        private TextBox textBox2;
        private Label label7;
        private TextBox textBox5;
        private TextBox textBox6;
        private TextBox textBox7;
        private TextBox textBox8;
        private TextBox textBox9;
        private TextBox textBox10;
        private TextBox textBox11;
        private TextBox textBox12;
        private TextBox textBox13;
        private TextBox textBox14;
        private TextBox textBox15;
        private TextBox textBox16;
        private TextBox textBox17;
        private TextBox textBox18;
        private TextBox textBox19;
        private TextBox textBox20;
        private TextBox textBox21;
        private TextBox textBox22;
        private TextBox textBox23;
        private TextBox textBox24;
        private TextBox textBox25;
        private TextBox textBox26;
        private TextBox textBox27;
        private TextBox textBox28;
        private TextBox textBox29;
        private TextBox textBox30;
        private TextBox textBox31;
        private TextBox textBox33;
        private TextBox textBox32;
        private TextBox textBox35;
        private TextBox textBox36;
        private TextBox textBox37;
        private TextBox textBox38;
        private TextBox textBox39;
        private TextBox textBox40;
        private TextBox textBox41;
        private TextBox textBox42;
        private DateTimePicker dateTimePicker1;
        private void InitializeComponent()
        {
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.colorDialog1 = new System.Windows.Forms.ColorDialog();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.label4 = new System.Windows.Forms.Label();
            this.comboBox2 = new System.Windows.Forms.ComboBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.textBox5 = new System.Windows.Forms.TextBox();
            this.textBox6 = new System.Windows.Forms.TextBox();
            this.textBox7 = new System.Windows.Forms.TextBox();
            this.textBox8 = new System.Windows.Forms.TextBox();
            this.textBox9 = new System.Windows.Forms.TextBox();
            this.textBox10 = new System.Windows.Forms.TextBox();
            this.textBox11 = new System.Windows.Forms.TextBox();
            this.textBox12 = new System.Windows.Forms.TextBox();
            this.textBox13 = new System.Windows.Forms.TextBox();
            this.textBox14 = new System.Windows.Forms.TextBox();
            this.textBox15 = new System.Windows.Forms.TextBox();
            this.textBox16 = new System.Windows.Forms.TextBox();
            this.textBox17 = new System.Windows.Forms.TextBox();
            this.textBox18 = new System.Windows.Forms.TextBox();
            this.textBox19 = new System.Windows.Forms.TextBox();
            this.textBox20 = new System.Windows.Forms.TextBox();
            this.textBox21 = new System.Windows.Forms.TextBox();
            this.textBox22 = new System.Windows.Forms.TextBox();
            this.textBox23 = new System.Windows.Forms.TextBox();
            this.textBox24 = new System.Windows.Forms.TextBox();
            this.textBox25 = new System.Windows.Forms.TextBox();
            this.textBox26 = new System.Windows.Forms.TextBox();
            this.textBox27 = new System.Windows.Forms.TextBox();
            this.textBox28 = new System.Windows.Forms.TextBox();
            this.textBox29 = new System.Windows.Forms.TextBox();
            this.textBox30 = new System.Windows.Forms.TextBox();
            this.textBox31 = new System.Windows.Forms.TextBox();
            this.textBox33 = new System.Windows.Forms.TextBox();
            this.textBox32 = new System.Windows.Forms.TextBox();
            this.textBox35 = new System.Windows.Forms.TextBox();
            this.textBox36 = new System.Windows.Forms.TextBox();
            this.textBox37 = new System.Windows.Forms.TextBox();
            this.textBox38 = new System.Windows.Forms.TextBox();
            this.textBox39 = new System.Windows.Forms.TextBox();
            this.textBox40 = new System.Windows.Forms.TextBox();
            this.textBox41 = new System.Windows.Forms.TextBox();
            this.textBox42 = new System.Windows.Forms.TextBox();
            this.dateTimePicker1 = new System.Windows.Forms.DateTimePicker();
            this.textBox3 = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Font = new System.Drawing.Font("Microsoft YaHei UI", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.button1.Location = new System.Drawing.Point(317, 742);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(145, 38);
            this.button1.TabIndex = 0;
            this.button1.Text = "生成";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.generate_Click);
            // 
            // button2
            // 
            this.button2.Font = new System.Drawing.Font("Microsoft YaHei UI", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.button2.Location = new System.Drawing.Point(730, 742);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(145, 38);
            this.button2.TabIndex = 2;
            this.button2.Text = "关闭";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.close_Click);
            // 
            // label1
            // 
            this.label1.Font = new System.Drawing.Font("Microsoft YaHei UI", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.label1.Location = new System.Drawing.Point(12, 22);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(206, 38);
            this.label1.TabIndex = 3;
            this.label1.Text = "生成新合同名称：";
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(211, 22);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(277, 38);
            this.textBox1.TabIndex = 4;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(494, 25);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(74, 31);
            this.label2.TabIndex = 5;
            this.label2.Text = ".docx";
            // 
            // label3
            // 
            this.label3.Font = new System.Drawing.Font("Microsoft YaHei UI", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.label3.Location = new System.Drawing.Point(660, 25);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(230, 38);
            this.label3.TabIndex = 6;
            this.label3.Text = "选择合同模板类型：";
            // 
            // comboBox1
            // 
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Items.AddRange(new object[] {
            "信托单位",
            "信托受益权",
            "资产包",
            "资管计划",
            "委托贷款"});
            this.comboBox1.Location = new System.Drawing.Point(887, 22);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(333, 39);
            this.comboBox1.TabIndex = 7;
            this.comboBox1.SelectedIndexChanged += new System.EventHandler(this.comboBox1_SelectedMode);
            // 
            // label4
            // 
            this.label4.Font = new System.Drawing.Font("Microsoft YaHei UI", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.label4.Location = new System.Drawing.Point(12, 82);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(182, 38);
            this.label4.TabIndex = 8;
            this.label4.Text = "选择存储路径：";
            // 
            // comboBox2
            // 
            this.comboBox2.FormattingEnabled = true;
            this.comboBox2.Location = new System.Drawing.Point(211, 79);
            this.comboBox2.Name = "comboBox2";
            this.comboBox2.Size = new System.Drawing.Size(333, 39);
            this.comboBox2.TabIndex = 9;
            // 
            // label5
            // 
            this.label5.Font = new System.Drawing.Font("Microsoft YaHei UI", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.label5.Location = new System.Drawing.Point(12, 145);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(174, 38);
            this.label5.TabIndex = 10;
            // 
            // label6
            // 
            this.label6.Location = new System.Drawing.Point(306, 139);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(79, 38);
            this.label6.TabIndex = 11;
            this.label6.Text = "ABC - ";
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(396, 139);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(101, 38);
            this.textBox2.TabIndex = 12;
            // 
            // label7
            // 
            this.label7.Location = new System.Drawing.Point(503, 139);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(65, 38);
            this.label7.TabIndex = 13;
            this.label7.Text = "基金";
            // 
            // textBox5
            // 
            this.textBox5.Font = new System.Drawing.Font("Microsoft YaHei UI", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.textBox5.Location = new System.Drawing.Point(12, 202);
            this.textBox5.Name = "textBox5";
            this.textBox5.Size = new System.Drawing.Size(262, 38);
            this.textBox5.TabIndex = 14;
            this.textBox5.Text = "信托/资管计划公司全称";
            // 
            // textBox6
            // 
            this.textBox6.Location = new System.Drawing.Point(306, 202);
            this.textBox6.Name = "textBox6";
            this.textBox6.Size = new System.Drawing.Size(262, 38);
            this.textBox6.TabIndex = 15;
            // 
            // textBox7
            // 
            this.textBox7.Font = new System.Drawing.Font("Microsoft YaHei UI", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.textBox7.Location = new System.Drawing.Point(12, 262);
            this.textBox7.Name = "textBox7";
            this.textBox7.Size = new System.Drawing.Size(262, 38);
            this.textBox7.TabIndex = 16;
            this.textBox7.Text = "信托/资管计划名称";
            // 
            // textBox8
            // 
            this.textBox8.Location = new System.Drawing.Point(306, 262);
            this.textBox8.Name = "textBox8";
            this.textBox8.Size = new System.Drawing.Size(262, 38);
            this.textBox8.TabIndex = 17;
            // 
            // textBox9
            // 
            this.textBox9.Font = new System.Drawing.Font("Microsoft YaHei UI", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.textBox9.Location = new System.Drawing.Point(12, 322);
            this.textBox9.Name = "textBox9";
            this.textBox9.Size = new System.Drawing.Size(262, 38);
            this.textBox9.TabIndex = 18;
            this.textBox9.Text = "募集账户账户名";
            // 
            // textBox10
            // 
            this.textBox10.Location = new System.Drawing.Point(306, 322);
            this.textBox10.Name = "textBox10";
            this.textBox10.Size = new System.Drawing.Size(262, 38);
            this.textBox10.TabIndex = 19;
            // 
            // textBox11
            // 
            this.textBox11.Font = new System.Drawing.Font("Microsoft YaHei UI", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.textBox11.Location = new System.Drawing.Point(12, 382);
            this.textBox11.Name = "textBox11";
            this.textBox11.Size = new System.Drawing.Size(262, 38);
            this.textBox11.TabIndex = 20;
            this.textBox11.Text = "募集账户账户号";
            // 
            // textBox12
            // 
            this.textBox12.Location = new System.Drawing.Point(306, 382);
            this.textBox12.Name = "textBox12";
            this.textBox12.Size = new System.Drawing.Size(262, 38);
            this.textBox12.TabIndex = 21;
            // 
            // textBox13
            // 
            this.textBox13.Font = new System.Drawing.Font("Microsoft YaHei UI", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.textBox13.Location = new System.Drawing.Point(12, 442);
            this.textBox13.Name = "textBox13";
            this.textBox13.Size = new System.Drawing.Size(262, 38);
            this.textBox13.TabIndex = 22;
            this.textBox13.Text = "募集账户开户行";
            // 
            // textBox14
            // 
            this.textBox14.Location = new System.Drawing.Point(306, 442);
            this.textBox14.Name = "textBox14";
            this.textBox14.Size = new System.Drawing.Size(262, 38);
            this.textBox14.TabIndex = 23;
            // 
            // textBox15
            // 
            this.textBox15.Font = new System.Drawing.Font("Microsoft YaHei UI", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.textBox15.Location = new System.Drawing.Point(12, 502);
            this.textBox15.Name = "textBox15";
            this.textBox15.Size = new System.Drawing.Size(262, 38);
            this.textBox15.TabIndex = 24;
            this.textBox15.Text = "大额支付系统行号";
            this.textBox15.TextChanged += new System.EventHandler(this.textBox13_TextChanged);
            // 
            // textBox16
            // 
            this.textBox16.Location = new System.Drawing.Point(306, 502);
            this.textBox16.Name = "textBox16";
            this.textBox16.Size = new System.Drawing.Size(262, 38);
            this.textBox16.TabIndex = 25;
            // 
            // textBox17
            // 
            this.textBox17.Font = new System.Drawing.Font("Microsoft YaHei UI", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.textBox17.Location = new System.Drawing.Point(12, 562);
            this.textBox17.Name = "textBox17";
            this.textBox17.Size = new System.Drawing.Size(262, 38);
            this.textBox17.TabIndex = 26;
            this.textBox17.Text = "产品经理姓名";
            // 
            // textBox18
            // 
            this.textBox18.Location = new System.Drawing.Point(306, 562);
            this.textBox18.Name = "textBox18";
            this.textBox18.Size = new System.Drawing.Size(262, 38);
            this.textBox18.TabIndex = 27;
            // 
            // textBox19
            // 
            this.textBox19.Font = new System.Drawing.Font("Microsoft YaHei UI", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.textBox19.Location = new System.Drawing.Point(12, 622);
            this.textBox19.Name = "textBox19";
            this.textBox19.Size = new System.Drawing.Size(262, 38);
            this.textBox19.TabIndex = 28;
            this.textBox19.Text = "产品经理固话-无区号";
            // 
            // textBox20
            // 
            this.textBox20.Location = new System.Drawing.Point(306, 622);
            this.textBox20.Name = "textBox20";
            this.textBox20.Size = new System.Drawing.Size(262, 38);
            this.textBox20.TabIndex = 29;
            // 
            // textBox21
            // 
            this.textBox21.Font = new System.Drawing.Font("Microsoft YaHei UI", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.textBox21.Location = new System.Drawing.Point(12, 682);
            this.textBox21.Name = "textBox21";
            this.textBox21.Size = new System.Drawing.Size(262, 38);
            this.textBox21.TabIndex = 30;
            this.textBox21.Text = "产品经理邮箱";
            // 
            // textBox22
            // 
            this.textBox22.Location = new System.Drawing.Point(306, 682);
            this.textBox22.Name = "textBox22";
            this.textBox22.Size = new System.Drawing.Size(262, 38);
            this.textBox22.TabIndex = 31;
            // 
            // textBox23
            // 
            this.textBox23.Font = new System.Drawing.Font("Microsoft YaHei UI", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.textBox23.Location = new System.Drawing.Point(660, 142);
            this.textBox23.Name = "textBox23";
            this.textBox23.Size = new System.Drawing.Size(262, 38);
            this.textBox23.TabIndex = 34;
            this.textBox23.Text = "A类份额投资期限（月）";
            // 
            // textBox24
            // 
            this.textBox24.Location = new System.Drawing.Point(958, 142);
            this.textBox24.Name = "textBox24";
            this.textBox24.Size = new System.Drawing.Size(262, 38);
            this.textBox24.TabIndex = 35;
            // 
            // textBox25
            // 
            this.textBox25.Font = new System.Drawing.Font("Microsoft YaHei UI", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.textBox25.Location = new System.Drawing.Point(660, 202);
            this.textBox25.Name = "textBox25";
            this.textBox25.Size = new System.Drawing.Size(262, 38);
            this.textBox25.TabIndex = 36;
            this.textBox25.Text = "C类份额投资期限（月）";
            // 
            // textBox26
            // 
            this.textBox26.Location = new System.Drawing.Point(958, 202);
            this.textBox26.Name = "textBox26";
            this.textBox26.Size = new System.Drawing.Size(262, 38);
            this.textBox26.TabIndex = 37;
            // 
            // textBox27
            // 
            this.textBox27.Font = new System.Drawing.Font("Microsoft YaHei UI", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.textBox27.Location = new System.Drawing.Point(660, 262);
            this.textBox27.Name = "textBox27";
            this.textBox27.Size = new System.Drawing.Size(262, 38);
            this.textBox27.TabIndex = 38;
            this.textBox27.Text = "A预计年化收益（%）";
            // 
            // textBox28
            // 
            this.textBox28.Location = new System.Drawing.Point(958, 262);
            this.textBox28.Name = "textBox28";
            this.textBox28.Size = new System.Drawing.Size(262, 38);
            this.textBox28.TabIndex = 39;
            // 
            // textBox29
            // 
            this.textBox29.Font = new System.Drawing.Font("Microsoft YaHei UI", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.textBox29.Location = new System.Drawing.Point(660, 322);
            this.textBox29.Name = "textBox29";
            this.textBox29.Size = new System.Drawing.Size(262, 38);
            this.textBox29.TabIndex = 40;
            this.textBox29.Text = "C预计年化收益（%）";
            // 
            // textBox30
            // 
            this.textBox30.Location = new System.Drawing.Point(958, 322);
            this.textBox30.Name = "textBox30";
            this.textBox30.Size = new System.Drawing.Size(262, 38);
            this.textBox30.TabIndex = 41;
            // 
            // textBox31
            // 
            this.textBox31.Font = new System.Drawing.Font("Microsoft YaHei UI", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.textBox31.Location = new System.Drawing.Point(660, 382);
            this.textBox31.Name = "textBox31";
            this.textBox31.Size = new System.Drawing.Size(262, 38);
            this.textBox31.TabIndex = 42;
            this.textBox31.Text = "收益分配周期（月）";
            // 
            // textBox33
            // 
            this.textBox33.Font = new System.Drawing.Font("Microsoft YaHei UI", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.textBox33.Location = new System.Drawing.Point(660, 442);
            this.textBox33.Name = "textBox33";
            this.textBox33.Size = new System.Drawing.Size(262, 38);
            this.textBox33.TabIndex = 44;
            this.textBox33.Text = "投资运作到期日";
            // 
            // textBox32
            // 
            this.textBox32.Location = new System.Drawing.Point(958, 382);
            this.textBox32.Name = "textBox32";
            this.textBox32.Size = new System.Drawing.Size(262, 38);
            this.textBox32.TabIndex = 43;
            // 
            // textBox35
            // 
            this.textBox35.Font = new System.Drawing.Font("Microsoft YaHei UI", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.textBox35.Location = new System.Drawing.Point(660, 502);
            this.textBox35.Name = "textBox35";
            this.textBox35.Size = new System.Drawing.Size(262, 38);
            this.textBox35.TabIndex = 46;
            this.textBox35.Text = "信托计划受益权持有者";
            // 
            // textBox36
            // 
            this.textBox36.Location = new System.Drawing.Point(958, 502);
            this.textBox36.Name = "textBox36";
            this.textBox36.Size = new System.Drawing.Size(262, 38);
            this.textBox36.TabIndex = 47;
            // 
            // textBox37
            // 
            this.textBox37.Font = new System.Drawing.Font("Microsoft YaHei UI", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.textBox37.Location = new System.Drawing.Point(660, 562);
            this.textBox37.Name = "textBox37";
            this.textBox37.Size = new System.Drawing.Size(262, 38);
            this.textBox37.TabIndex = 48;
            this.textBox37.Text = "项目公司名称";
            // 
            // textBox38
            // 
            this.textBox38.Location = new System.Drawing.Point(958, 562);
            this.textBox38.Name = "textBox38";
            this.textBox38.Size = new System.Drawing.Size(262, 38);
            this.textBox38.TabIndex = 49;
            // 
            // textBox39
            // 
            this.textBox39.Font = new System.Drawing.Font("Microsoft YaHei UI", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.textBox39.Location = new System.Drawing.Point(660, 622);
            this.textBox39.Name = "textBox39";
            this.textBox39.Size = new System.Drawing.Size(262, 38);
            this.textBox39.TabIndex = 50;
            this.textBox39.Text = "基础义务承担方";
            // 
            // textBox40
            // 
            this.textBox40.Location = new System.Drawing.Point(958, 622);
            this.textBox40.Name = "textBox40";
            this.textBox40.Size = new System.Drawing.Size(262, 38);
            this.textBox40.TabIndex = 51;
            // 
            // textBox41
            // 
            this.textBox41.Font = new System.Drawing.Font("Microsoft YaHei UI", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.textBox41.Location = new System.Drawing.Point(660, 682);
            this.textBox41.Name = "textBox41";
            this.textBox41.Size = new System.Drawing.Size(262, 38);
            this.textBox41.TabIndex = 52;
            this.textBox41.Text = "借款人名称";
            // 
            // textBox42
            // 
            this.textBox42.Location = new System.Drawing.Point(958, 682);
            this.textBox42.Name = "textBox42";
            this.textBox42.Size = new System.Drawing.Size(262, 38);
            this.textBox42.TabIndex = 53;
            // 
            // dateTimePicker1
            // 
            this.dateTimePicker1.Location = new System.Drawing.Point(958, 442);
            this.dateTimePicker1.Name = "dateTimePicker1";
            this.dateTimePicker1.Size = new System.Drawing.Size(262, 38);
            this.dateTimePicker1.TabIndex = 54;
            // 
            // textBox3
            // 
            this.textBox3.Font = new System.Drawing.Font("Microsoft YaHei UI", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.textBox3.Location = new System.Drawing.Point(12, 142);
            this.textBox3.Name = "textBox3";
            this.textBox3.Size = new System.Drawing.Size(262, 38);
            this.textBox3.TabIndex = 55;
            this.textBox3.Text = "基金名称";
            this.textBox3.TextChanged += new System.EventHandler(this.textBox21_TextChanged);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(14F, 31F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1234, 802);
            this.Controls.Add(this.textBox3);
            this.Controls.Add(this.dateTimePicker1);
            this.Controls.Add(this.textBox42);
            this.Controls.Add(this.textBox41);
            this.Controls.Add(this.textBox40);
            this.Controls.Add(this.textBox39);
            this.Controls.Add(this.textBox38);
            this.Controls.Add(this.textBox37);
            this.Controls.Add(this.textBox36);
            this.Controls.Add(this.textBox35);
            this.Controls.Add(this.textBox32);
            this.Controls.Add(this.textBox33);
            this.Controls.Add(this.textBox31);
            this.Controls.Add(this.textBox30);
            this.Controls.Add(this.textBox29);
            this.Controls.Add(this.textBox28);
            this.Controls.Add(this.textBox27);
            this.Controls.Add(this.textBox26);
            this.Controls.Add(this.textBox25);
            this.Controls.Add(this.textBox24);
            this.Controls.Add(this.textBox23);
            this.Controls.Add(this.textBox22);
            this.Controls.Add(this.textBox21);
            this.Controls.Add(this.textBox20);
            this.Controls.Add(this.textBox19);
            this.Controls.Add(this.textBox18);
            this.Controls.Add(this.textBox17);
            this.Controls.Add(this.textBox16);
            this.Controls.Add(this.textBox15);
            this.Controls.Add(this.textBox14);
            this.Controls.Add(this.textBox13);
            this.Controls.Add(this.textBox12);
            this.Controls.Add(this.textBox11);
            this.Controls.Add(this.textBox10);
            this.Controls.Add(this.textBox9);
            this.Controls.Add(this.textBox8);
            this.Controls.Add(this.textBox7);
            this.Controls.Add(this.textBox6);
            this.Controls.Add(this.textBox5);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.textBox2);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.comboBox2);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.comboBox1);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Name = "Form1";
            this.Text = "合同生成器";
            this.Load += new System.EventHandler(this.Form1_Load1);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private TextBox textBox3;
    }
}