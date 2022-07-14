
namespace excel_interop
{
    partial class MainForm
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        #region Windows Form Designer generated code

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.checkBoxExcelVisible = new System.Windows.Forms.CheckBox();
            this.checkBoxWorkbookOpen = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // textBox1
            // 
            this.textBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBox1.Location = new System.Drawing.Point(31, 99);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(518, 261);
            this.textBox1.TabIndex = 1;
            // 
            // checkBoxExcelVisible
            // 
            this.checkBoxExcelVisible.AutoSize = true;
            this.checkBoxExcelVisible.Location = new System.Drawing.Point(36, 40);
            this.checkBoxExcelVisible.Name = "checkBoxExcelVisible";
            this.checkBoxExcelVisible.Size = new System.Drawing.Size(132, 29);
            this.checkBoxExcelVisible.TabIndex = 2;
            this.checkBoxExcelVisible.Text = "Excel Visible";
            this.checkBoxExcelVisible.UseVisualStyleBackColor = true;
            this.checkBoxExcelVisible.CheckedChanged += new System.EventHandler(this.checkBoxExcelVisible_CheckedChanged);
            // 
            // checkBoxWorkbookOpen
            // 
            this.checkBoxWorkbookOpen.AutoSize = true;
            this.checkBoxWorkbookOpen.Location = new System.Drawing.Point(209, 40);
            this.checkBoxWorkbookOpen.Name = "checkBoxWorkbookOpen";
            this.checkBoxWorkbookOpen.Size = new System.Drawing.Size(171, 29);
            this.checkBoxWorkbookOpen.TabIndex = 2;
            this.checkBoxWorkbookOpen.Text = "Workbook Open";
            this.checkBoxWorkbookOpen.UseVisualStyleBackColor = true;
            this.checkBoxWorkbookOpen.Visible = false;
            this.checkBoxWorkbookOpen.CheckedChanged += new System.EventHandler(this.checkBoxWorkbookOpen_CheckedChanged);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(10F, 25F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(609, 384);
            this.Controls.Add(this.checkBoxWorkbookOpen);
            this.Controls.Add(this.checkBoxExcelVisible);
            this.Controls.Add(this.textBox1);
            this.Name = "MainForm";
            this.Text = "Excel Form";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.CheckBox checkBoxExcelVisible;
        private System.Windows.Forms.CheckBox checkBoxWorkbookOpen;
    }
}

