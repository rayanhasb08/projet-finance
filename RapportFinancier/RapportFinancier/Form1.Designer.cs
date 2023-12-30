namespace RapportFinancier
{
    partial class Form1
    {
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

        #region Windows Form Designer generated code

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            buttonGenerate = new Button();
            labelCompanyName = new Label();
            textBoxCompanyName = new TextBox();
            SuspendLayout();
            // 
            // buttonGenerate
            // 
            buttonGenerate.Location = new Point(153, 107);
            buttonGenerate.Name = "buttonGenerate";
            buttonGenerate.Size = new Size(120, 36);
            buttonGenerate.TabIndex = 0;
            buttonGenerate.Text = "Generate";
            buttonGenerate.UseVisualStyleBackColor = true;
            buttonGenerate.Click += buttonGenerate_Click;
            // 
            // labelCompanyName
            // 
            labelCompanyName.AutoSize = true;
            labelCompanyName.Location = new Point(16, 17);
            labelCompanyName.Name = "labelCompanyName";
            labelCompanyName.Size = new Size(131, 15);
            labelCompanyName.TabIndex = 1;
            labelCompanyName.Text = "Nom de la compagnie :";
            // 
            // textBoxCompanyName
            // 
            textBoxCompanyName.Location = new Point(148, 14);
            textBoxCompanyName.Name = "textBoxCompanyName";
            textBoxCompanyName.Size = new Size(282, 23);
            textBoxCompanyName.TabIndex = 2;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(442, 247);
            Controls.Add(textBoxCompanyName);
            Controls.Add(labelCompanyName);
            Controls.Add(buttonGenerate);
            Name = "Form1";
            Text = "Form1";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button buttonGenerate;
        private Label labelCompanyName;
        private TextBox textBoxCompanyName;
    }
}
