using ExcelToWord.Services;
using Microsoft.WindowsAPICodePack.Dialogs;

namespace ExcelToWord;

public partial class Form1 : Form
{
    public string ExcelFile = "";
    public string WordFile = "";
    public string Folder = "";

    public Form1()
    {
        InitializeComponent();
    }

    private void button1_Click(object sender, EventArgs e)
    {
        int size = -1;
        OpenFileDialog openFileDialog1 = new();
        openFileDialog1.Filter = "Office Files|*.xls;*.xlsx";
        DialogResult result = openFileDialog1.ShowDialog(); // Show the dialog.
        if (result == DialogResult.OK) // Test result.
        {
            ExcelFile = openFileDialog1.FileName;
            textBox1.Text = ExcelFile;
            try
            {
                string text = File.ReadAllText(ExcelFile);
                size = text.Length;
            }
            catch (IOException)
            {
            }
        }
        Console.WriteLine(size); // <-- Shows file size in debugging mode.
        Console.WriteLine(result); // <-- For debugging use.
    }

    private void button2_Click(object sender, EventArgs e)
    {
        int size = -1;
        OpenFileDialog openFileDialog1 = new();
        openFileDialog1.Filter = "Office Files|*.doc;*.docx";
        DialogResult result = openFileDialog1.ShowDialog(); // Show the dialog.
        if (result == DialogResult.OK) // Test result.
        {
            WordFile = openFileDialog1.FileName;
            textBox2.Text = WordFile;
            try
            {
                string text = File.ReadAllText(WordFile);
                size = text.Length;
            }
            catch (IOException)
            {
            }
        }
        Console.WriteLine(size); // <-- Shows file size in debugging mode.
        Console.WriteLine(result); // <-- For debugging use.
    }

    private void button3_Click(object sender, EventArgs e)
    {
        ConvertExcelToWord.ConvertFile(ExcelFile, WordFile, Folder);
    }

    private void button4_Click(object sender, EventArgs e)
    {
        CommonOpenFileDialog dialog = new CommonOpenFileDialog();
        dialog.InitialDirectory = "C:\\Users";
        dialog.IsFolderPicker = true;
        if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
        {
            Folder = dialog.FileName;
            textBox3.Text = Folder;
        }

    }
}