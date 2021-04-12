namespace WindowsFormsApplication1
{
    partial class PrintForm8
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.label1 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.birthTextBox = new System.Windows.Forms.TextBox();
            this.districtComboBox = new System.Windows.Forms.ComboBox();
            this.printButton = new System.Windows.Forms.Button();
            this.closeButton = new System.Windows.Forms.Button();
            this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label1.Location = new System.Drawing.Point(12, 20);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(118, 20);
            this.label1.TabIndex = 0;
            this.label1.Text = "Год рождения";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label3.Location = new System.Drawing.Point(97, 54);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(93, 20);
            this.label3.TabIndex = 2;
            this.label3.Text = "Военкомат";
            // 
            // birthTextBox
            // 
            this.birthTextBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.birthTextBox.Location = new System.Drawing.Point(208, 14);
            this.birthTextBox.Name = "birthTextBox";
            this.birthTextBox.Size = new System.Drawing.Size(58, 26);
            this.birthTextBox.TabIndex = 3;
            this.birthTextBox.Text = "1996";
            // 
            // districtComboBox
            // 
            this.districtComboBox.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.districtComboBox.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.districtComboBox.DropDownWidth = 400;
            this.districtComboBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.districtComboBox.FormattingEnabled = true;
            this.districtComboBox.Location = new System.Drawing.Point(16, 80);
            this.districtComboBox.Name = "districtComboBox";
            this.districtComboBox.Size = new System.Drawing.Size(250, 28);
            this.districtComboBox.TabIndex = 5;
            // 
            // printButton
            // 
            this.printButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.printButton.Location = new System.Drawing.Point(16, 123);
            this.printButton.Name = "printButton";
            this.printButton.Size = new System.Drawing.Size(250, 30);
            this.printButton.TabIndex = 6;
            this.printButton.Text = "Печать";
            this.printButton.UseVisualStyleBackColor = true;
            this.printButton.Click += new System.EventHandler(this.printButton_Click);
            // 
            // closeButton
            // 
            this.closeButton.Location = new System.Drawing.Point(197, 161);
            this.closeButton.Name = "closeButton";
            this.closeButton.Size = new System.Drawing.Size(75, 23);
            this.closeButton.TabIndex = 7;
            this.closeButton.Text = "Закрыть";
            this.closeButton.UseVisualStyleBackColor = true;
            this.closeButton.Click += new System.EventHandler(this.closeButton_Click);
            // 
            // backgroundWorker1
            // 
            this.backgroundWorker1.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorker1_DoWork);
            this.backgroundWorker1.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.backgroundWorker1_RunWorkerCompleted);
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::WindowsFormsApplication1.Properties.Resources.preview;
            this.pictureBox1.Location = new System.Drawing.Point(124, 54);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(35, 35);
            this.pictureBox1.TabIndex = 8;
            this.pictureBox1.TabStop = false;
            this.pictureBox1.UseWaitCursor = true;
            this.pictureBox1.Visible = false;
            // 
            // printForm8
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 196);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.closeButton);
            this.Controls.Add(this.printButton);
            this.Controls.Add(this.districtComboBox);
            this.Controls.Add(this.birthTextBox);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = global::WindowsFormsApplication1.Properties.Resources.Print;
            this.MaximizeBox = false;
            this.Name = "printForm8";
            this.Text = "Форма 8";
            this.Load += new System.EventHandler(this.PrintForm8_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox birthTextBox;
        private System.Windows.Forms.ComboBox districtComboBox;
        private System.Windows.Forms.Button printButton;
        private System.Windows.Forms.Button closeButton;
        private System.ComponentModel.BackgroundWorker backgroundWorker1;
        private System.Windows.Forms.PictureBox pictureBox1;
    }
}