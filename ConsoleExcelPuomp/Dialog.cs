using System;
using System.Windows.Forms;

namespace ConsoleExcelPuomp
{
    class Dialog
    {
        [STAThread]
        public static dynamic FolderBrowser()
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog
            {
                Description = " 1) Example text \n 2) Example text \n 3) Example text ",
                RootFolder = System.Environment.SpecialFolder.Desktop,
                SelectedPath = "C:\\Windows\\",
                ShowNewFolderButton = true
            };
            if (fbd.ShowDialog(new Form() { TopMost = true, TopLevel = true, WindowState = FormWindowState.Minimized }) == DialogResult.OK)
            {
                return fbd.SelectedPath;
            }
            return false;
        }

        [STAThread]
        public static dynamic FileBrowser(string ext = "Excel|*.xls;*.xlsx")
        {
            OpenFileDialog ofd = new OpenFileDialog
            {
                Title = "Выбор файла Excel",
                Filter = (ext)
            };
            if (ofd.ShowDialog(new Form() { TopMost = true, TopLevel = true, WindowState = FormWindowState.Minimized }) == DialogResult.OK)
            {
                return ofd.FileName;
            }
            return false;
        }
    }
}
