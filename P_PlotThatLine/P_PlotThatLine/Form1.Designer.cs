namespace P_PlotThatLine
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
            formsPlot1 = new ScottPlot.WinForms.FormsPlot();
            dateTimePicker1 = new DateTimePicker();
            dateTimePicker2 = new DateTimePicker();
            checkBox1 = new CheckBox();
            textBox1 = new TextBox();
            textBox2 = new TextBox();
            SuspendLayout();
            // 
            // formsPlot1
            // 
            formsPlot1.DisplayScale = 1F;
            formsPlot1.Location = new Point(12, 34);
            formsPlot1.Name = "formsPlot1";
            formsPlot1.Size = new Size(574, 381);
            formsPlot1.TabIndex = 0;
            formsPlot1.Load += formsPlot1_Load;
            // 
            // dateTimePicker1
            // 
            dateTimePicker1.Location = new Point(627, 51);
            dateTimePicker1.Name = "dateTimePicker1";
            dateTimePicker1.Size = new Size(161, 23);
            dateTimePicker1.TabIndex = 1;
            // 
            // dateTimePicker2
            // 
            dateTimePicker2.Location = new Point(627, 84);
            dateTimePicker2.Name = "dateTimePicker2";
            dateTimePicker2.Size = new Size(162, 23);
            dateTimePicker2.TabIndex = 2;
            // 
            // checkBox1
            // 
            checkBox1.AccessibleDescription = "";
            checkBox1.AutoSize = true;
            checkBox1.Location = new Point(592, 124);
            checkBox1.Name = "checkBox1";
            checkBox1.Size = new Size(68, 19);
            checkBox1.TabIndex = 3;
            checkBox1.Text = "Bougies";
            checkBox1.UseVisualStyleBackColor = true;
            checkBox1.CheckedChanged += checkBox1_CheckedChanged;
            // 
            // textBox1
            // 
            textBox1.Location = new Point(592, 51);
            textBox1.Name = "textBox1";
            textBox1.Size = new Size(29, 23);
            textBox1.TabIndex = 4;
            textBox1.Text = "De";
            // 
            // textBox2
            // 
            textBox2.Location = new Point(592, 84);
            textBox2.Name = "textBox2";
            textBox2.Size = new Size(29, 23);
            textBox2.TabIndex = 5;
            textBox2.Text = "à";
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(800, 450);
            Controls.Add(textBox2);
            Controls.Add(textBox1);
            Controls.Add(checkBox1);
            Controls.Add(dateTimePicker2);
            Controls.Add(dateTimePicker1);
            Controls.Add(formsPlot1);
            Name = "Form1";
            Text = "Form1";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private ScottPlot.WinForms.FormsPlot formsPlot1;
        private DateTimePicker dateTimePicker1;
        private DateTimePicker dateTimePicker2;
        private CheckBox checkBox1;
        private TextBox textBox1;
        private TextBox textBox2;
    }
}
