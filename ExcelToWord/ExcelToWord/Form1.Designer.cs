﻿namespace ExcelToWord;

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
        SuspendLayout();
        // 
        // button1
        // 
        button1.Location = new Point(485, 150);
        button1.Name = "button1";
        button1.Size = new Size(141, 23);
        button1.TabIndex = 0;
        button1.Text = "button1";
        button1.UseVisualStyleBackColor = true;
        button1.Click += button1_Click;
        // 
        // button2
        // 
        button2.Location = new Point(485, 205);
        button2.Name = "button2";
        button2.Size = new Size(141, 22);
        button2.TabIndex = 1;
        button2.Text = "button2";
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
        button3.Location = new Point(360, 285);
        button3.Name = "button3";
        button3.Size = new Size(148, 47);
        button3.TabIndex = 4;
        button3.Text = "button3";
        button3.UseVisualStyleBackColor = true;
        button3.Click += button3_Click;
        // 
        // label1
        // 
        label1.AutoSize = true;
        label1.Location = new Point(199, 132);
        label1.Name = "label1";
        label1.Size = new Size(38, 15);
        label1.TabIndex = 5;
        label1.Text = "label1";
        // 
        // label2
        // 
        label2.AutoSize = true;
        label2.Location = new Point(199, 186);
        label2.Name = "label2";
        label2.Size = new Size(38, 15);
        label2.TabIndex = 6;
        label2.Text = "label2";
        // 
        // Form1
        // 
        AutoScaleDimensions = new SizeF(7F, 15F);
        AutoScaleMode = AutoScaleMode.Font;
        ClientSize = new Size(845, 560);
        Controls.Add(label2);
        Controls.Add(label1);
        Controls.Add(button3);
        Controls.Add(textBox2);
        Controls.Add(textBox1);
        Controls.Add(button2);
        Controls.Add(button1);
        Name = "Form1";
        Text = "Form1";
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
}