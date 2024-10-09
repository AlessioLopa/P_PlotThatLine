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
            label1 = new Label();
            label2 = new Label();
            button1 = new Button();
            checkBox1 = new CheckBox();
            SuspendLayout();
            // 
            // formsPlot1
            // 
            formsPlot1.AutoSize = true;
            formsPlot1.DisplayScale = 1F;
            formsPlot1.Location = new Point(11, 2);
            formsPlot1.Name = "formsPlot1";
            formsPlot1.Size = new Size(633, 434);
            formsPlot1.TabIndex = 0;
            formsPlot1.Load += formsPlot1_Load;
            // 
            // dateTimePicker1
            // 
            dateTimePicker1.Location = new Point(699, 34);
            dateTimePicker1.Name = "dateTimePicker1";
            dateTimePicker1.Size = new Size(161, 23);
            dateTimePicker1.TabIndex = 1;
            dateTimePicker1.ValueChanged += dateTimePicker1_ValueChanged;
            // 
            // dateTimePicker2
            // 
            dateTimePicker2.Location = new Point(699, 67);
            dateTimePicker2.Name = "dateTimePicker2";
            dateTimePicker2.Size = new Size(162, 23);
            dateTimePicker2.TabIndex = 2;
            dateTimePicker2.ValueChanged += dateTimePicker2_ValueChanged;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(664, 40);
            label1.Name = "label1";
            label1.Size = new Size(21, 15);
            label1.TabIndex = 6;
            label1.Text = "De";
            label1.Click += label1_Click;
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new Point(664, 73);
            label2.Name = "label2";
            label2.Size = new Size(13, 15);
            label2.TabIndex = 7;
            label2.Text = "à";
            label2.Click += label2_Click;
            // 
            // button1
            // 
            button1.Location = new Point(738, 96);
            button1.Name = "button1";
            button1.Size = new Size(75, 23);
            button1.TabIndex = 8;
            button1.Text = "Filtrer";
            button1.UseVisualStyleBackColor = true;
            button1.Click += button1_Click_1;
            // 
            // checkBox1
            // 
            checkBox1.AccessibleDescription = "";
            checkBox1.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            checkBox1.AutoSize = true;
            checkBox1.Location = new Point(664, 134);
            checkBox1.Name = "checkBox1";
            checkBox1.Size = new Size(68, 19);
            checkBox1.TabIndex = 3;
            checkBox1.Text = "Bougies";
            checkBox1.UseVisualStyleBackColor = true;
            checkBox1.CheckedChanged += checkBox1_CheckedChanged;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            AutoSize = true;
            ClientSize = new Size(932, 476);
            Controls.Add(button1);
            Controls.Add(label2);
            Controls.Add(label1);
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
        private Label label1;
        private Label label2;
        private Button button1;
        private CheckBox checkBox1;
    }
}
