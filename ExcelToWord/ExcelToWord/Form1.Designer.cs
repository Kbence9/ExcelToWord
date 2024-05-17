namespace ExcelToWord;

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
        button1 = new Button();
        button2 = new Button();
        textBox1 = new TextBox();
        textBox2 = new TextBox();
        button3 = new Button();
        label1 = new Label();
        label2 = new Label();
        button4 = new Button();
        textBox3 = new TextBox();
        label3 = new Label();
        SuspendLayout();
        // 
        // button1
        // 
        button1.Location = new Point(485, 142);
        button1.Name = "button1";
        button1.Size = new Size(141, 37);
        button1.TabIndex = 0;
        button1.Text = "Válassz Excel  fájl";
        button1.UseVisualStyleBackColor = true;
        button1.Click += button1_Click;
        // 
        // button2
        // 
        button2.Location = new Point(485, 197);
        button2.Name = "button2";
        button2.Size = new Size(141, 35);
        button2.TabIndex = 1;
        button2.Text = "Válassz Word fájlt";
        button2.UseVisualStyleBackColor = true;
        button2.Click += button2_Click;
        // 
        // textBox1
        // 
        textBox1.Location = new Point(199, 150);
        textBox1.Name = "textBox1";
        textBox1.Size = new Size(258, 23);
        textBox1.TabIndex = 2;
        // 
        // textBox2
        // 
        textBox2.Location = new Point(199, 204);
        textBox2.Name = "textBox2";
        textBox2.Size = new Size(258, 23);
        textBox2.TabIndex = 3;
        // 
        // button3
        // 
        button3.Location = new Point(314, 329);
        button3.Name = "button3";
        button3.Size = new Size(238, 79);
        button3.TabIndex = 4;
        button3.Text = "Konvertálás";
        button3.UseVisualStyleBackColor = true;
        button3.Click += button3_Click;
        // 
        // label1
        // 
        label1.AutoSize = true;
        label1.Location = new Point(199, 132);
        label1.Name = "label1";
        label1.Size = new Size(103, 15);
        label1.TabIndex = 5;
        label1.Text = "Excel fájl elérési út";
        // 
        // label2
        // 
        label2.AutoSize = true;
        label2.Location = new Point(199, 186);
        label2.Name = "label2";
        label2.Size = new Size(105, 15);
        label2.TabIndex = 6;
        label2.Text = "Word fájl elérési út";
        // 
        // button4
        // 
        button4.Location = new Point(485, 251);
        button4.Name = "button4";
        button4.Size = new Size(141, 38);
        button4.TabIndex = 7;
        button4.Text = "Válassz cél mappát";
        button4.UseVisualStyleBackColor = true;
        button4.Click += button4_Click;
        // 
        // textBox3
        // 
        textBox3.Location = new Point(199, 260);
        textBox3.Name = "textBox3";
        textBox3.Size = new Size(258, 23);
        textBox3.TabIndex = 8;
        // 
        // label3
        // 
        label3.AutoSize = true;
        label3.Location = new Point(199, 242);
        label3.Name = "label3";
        label3.Size = new Size(124, 15);
        label3.TabIndex = 9;
        label3.Text = "Konvertált fájlok helye";
        // 
        // Form1
        // 
        AutoScaleDimensions = new SizeF(7F, 15F);
        AutoScaleMode = AutoScaleMode.Font;
        BackColor = SystemColors.Window;
        ClientSize = new Size(845, 560);
        Controls.Add(label3);
        Controls.Add(textBox3);
        Controls.Add(button4);
        Controls.Add(label2);
        Controls.Add(label1);
        Controls.Add(button3);
        Controls.Add(textBox2);
        Controls.Add(textBox1);
        Controls.Add(button2);
        Controls.Add(button1);
        Name = "Form1";
        Text = "ExcelToWord";
        ResumeLayout(false);
        PerformLayout();
    }

    #endregion

    private Button button1;
    private Button button2;
    private TextBox textBox1;
    private TextBox textBox2;
    private Button button3;
    private Label label1;
    private Label label2;
    private Button button4;
    private TextBox textBox3;
    private Label label3;
}