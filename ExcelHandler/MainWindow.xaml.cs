using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;
using SaveFileDialog = Microsoft.Win32.SaveFileDialog;

namespace ExcelHandler
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        private string _preProcessFilePath;

        private string _reProcessFilePath;

        public MainWindow()
        {
            InitializeComponent();
        }

        public StringBuilder PreProcess(string strCon)
        {
            OleDbConnection connection = new OleDbConnection(strCon);
            connection.Open();
            string strExcel = "";
            OleDbDataAdapter adapter = null;
            DataSet ds = null;
            strExcel = "select * from [Sheet1$]";
            adapter = new OleDbDataAdapter(strExcel, strCon);
            ds = new DataSet();
            adapter.Fill(ds, "table1");
            DataTable table = ds.Tables[0];

            StringBuilder result = new StringBuilder();
            Console.WriteLine(table.Rows.Count);

            //每一行
            ArrayList list = new ArrayList();

            int i = 0;
            int j = 0;
            for (i = 0; i < table.Rows.Count; i++)
            {
                list.Add(table.Rows[i][1].ToString());
                //result.Append(table.Rows[i][1].ToString());
                //如果这一行有多项，则把此行的全部项都添加到 list 中
                for (j = i; j < table.Rows.Count - 1; j++)
                {
                    if (table.Rows[j][0].ToString() == table.Rows[j + 1][0].ToString())
                    {
                        list.Add(table.Rows[j + 1][1].ToString());
                        //result.Append("," + table.Rows[j + 1][1].ToString());
                    }
                    else
                    {
                        break;
                    }
                }
                i = j;
                //如果此行只有一项，则直接跳下一行
                if (list.Count == 1)
                {
                    list.Clear();
                    continue;
                }
                list.Sort();
                foreach (string item in list)
                {
                    result.Append(item + " ");
                }
                list.Clear();
                result.Append(Environment.NewLine);
                //Console.WriteLine(i);
            }

            //Console.Write(result.ToString());

            //String[] tempResult_1 = result.ToString().Split(Environment.NewLine.ToCharArray());

            return result;
        }

        /// <summary>
        /// 选择需要预处理的文件 xlsx 文件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TxtPreProcess_GotFocus(object sender, RoutedEventArgs e)
        {
            OpenFileDialog open = new OpenFileDialog();
            open.DefaultExt = ".xlsx";
            open.Filter = "Excel 工作簿 (.xlsx)|*.xlsx";
            Nullable<bool> result = open.ShowDialog();
            if (result == true)
            {
                this._preProcessFilePath = open.FileName;
                this.TxtPreProcess.Text = open.FileName;
                this.TxtResult.AppendText("已选择预处理文件：" + _preProcessFilePath + Environment.NewLine);
            }
        }

        /// <summary>
        /// 预处理程序
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnPreProcess_Click(object sender, RoutedEventArgs e)
        {
            string filePath = this._preProcessFilePath;
            string strCon = "Provider=Microsoft.Ace.OLEDB.12.0;" + "Data Source=" + filePath + ";" + "Extended Properties=Excel 12.0;";

            TxtResult.AppendText("开始预处理，请稍后" + Environment.NewLine);

            StringBuilder result = PreProcess(strCon);

            SaveFileDialog save = new SaveFileDialog
            {
                DefaultExt = ".txt",
                Filter = "文本文件 (.txt)|*.txt"
            };


            if (save.ShowDialog() == true)
            {
                FileStream output = new FileStream(save.FileName, FileMode.Create);
                StreamWriter writer = new StreamWriter(output, Encoding.UTF8);
                writer.Write(result);
                writer.Close();
                output.Close();
            }
         
            TxtResult.AppendText("预处理完成，结果为：" + save.FileName + Environment.NewLine);
        }

        /// <summary>
        /// 选择需要再处理的文件 txt 文件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TxtReProcess_GotFocus(object sender, RoutedEventArgs e)
        {
            OpenFileDialog open = new OpenFileDialog();
            open.DefaultExt = ".txt";
            open.Filter = "文本文件 (.txt)|*.txt";
            Nullable<bool> result = open.ShowDialog();
            if (result == true)
            {
                this._reProcessFilePath = open.FileName;
                this.TxtReProcess.Text = open.FileName;
                this.TxtResult.AppendText("已选择再处理文件：" + _reProcessFilePath + Environment.NewLine);
            }
        }


        //TODO: 逆序、合并、去重、输出

        /// <summary>
        /// 再处理程序
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnReProcess_Click(object sender, RoutedEventArgs e)
        {
            TxtResult.AppendText("开始再处理，请稍后" + Environment.NewLine);

            //去除 FP-Growth 所得结果中的出现次数
            List<List<string>> reverse = RemovePrefix();

            //使用删减版流通记录制作图书名称字典
            OpenFileDialog open = new OpenFileDialog();

            open.DefaultExt = ".xlsx";
            open.Filter = "Excel 工作簿 (.xlsx)|*.xlsx";

            string filePath = open.FileName;
            string strCon = "";

            Nullable<bool> openedFile = open.ShowDialog();
            if (openedFile == true)
            {
                //this._preProcessFilePath = open.FileName;
                //this.TxtPreProcess.Text = open.FileName;
                this.TxtResult.AppendText("已选择图书集文件：" + open.FileName + Environment.NewLine);
                strCon = "Provider=Microsoft.Ace.OLEDB.12.0;" + "Data Source=" + open.FileName + ";" + "Extended Properties=Excel 12.0;";
            }

            OleDbConnection connection = new OleDbConnection(strCon);
            connection.Open();
            string strExcel = "";
            OleDbDataAdapter adapter = null;
            DataSet ds = null;
            strExcel = "select * from [Sheet1$]";
            adapter = new OleDbDataAdapter(strExcel, strCon);
            ds = new DataSet();
            adapter.Fill(ds, "table1");
            DataTable table = ds.Tables[0];

            //HashSet<string> bookName = new HashSet<string>();

            Dictionary<string, string> bookName = new Dictionary<string, string>();

            for (int i = 0; i < table.Rows.Count; i++)
            {
                try
                {
                    bookName.Add(table.Rows[i][0].ToString(), table.Rows[i][1].ToString());
                }
                catch (Exception exception)
                {
                    
                }
                
            }

            //逆序
            for (int i = 0; i < reverse.Count; i++)
            {
                reverse[i].Reverse();
            }

            //合并与去重
            //每一行
            List<string> row = new List<string>();
            List<List<string>> result = new List<List<string>>();
            //List<HashSet<string>> hashList = new List<HashSet<string>>();
            HashSet<string> hashSet = new HashSet<string>();

            /*
            for (int i = 0; i < reverse.Count; i++)
            {
                for (int j = i; j < reverse.Count; j++)
                {
                    if (reverse[i][0] == reverse[j][0])
                    {
                        row.AddRange(reverse[j]);
                    }
                    else
                    {
                        i = j;
                        hashSet = new HashSet<string>(row);
                        result.Add(hashSet.ToList());
                        row.Clear();
                        hashSet.Clear();
                        break;
                    }
                }
            }
            */
            
            for (int i = 0; i < reverse.Count; i++)
            {
                row.AddRange(reverse[i]);
                for (int j = i + 1; j < reverse.Count; j++)
                {
                    if (reverse[i][0] == reverse[j][0])
                    {
                        row.AddRange(reverse[j]);
                    }
                    else
                    {
                        i = j - 1;
                        hashSet = new HashSet<string>(row);
                        result.Add(hashSet.ToList());
                        row.Clear();
                        hashSet.Clear();
                        break;
                    }
                }
            }

            //替换结果中的图书种类号

            for (int i = 0; i < result.Count; i++)
            {
                for (int j = 0; j < result[i].Count; j++)
                {
                    result[i][j] = bookName[result[i][j]];
                }
                //result[i].Add(Environment.NewLine);
            }
            
            /*
            for (int i = 0; i < result.Count; i++)
            {
                result[i].Add(Environment.NewLine);
            }
            */

            //图书名称保存本次分析结果
            SaveFileDialog save = new SaveFileDialog
            {
                DefaultExt = ".txt",
                Filter = "文本文件 (.txt)|*.txt"
            };


            if (save.ShowDialog() == true)
            {
                FileStream output = new FileStream(save.FileName, FileMode.Create);
                StreamWriter writer = new StreamWriter(output, Encoding.UTF8);
                for (int i = 0; i < result.Count; i++)
                {
                    for (int j = 0; j < result[i].Count; j++)
                    {
                        writer.Write(result[i][j] + "，");
                    }
                    writer.Write(Environment.NewLine + Environment.NewLine );
                }
                //writer.Write(result);
                writer.Close();
                output.Close();
            }

            TxtResult.AppendText("再处理完成，结果为：" + save.FileName + Environment.NewLine);
        }

        /// <summary>
        /// 
        /// </summary>
        private List<List<string>> RemovePrefix()
        {
            List<List<string>> result = new List<List<string>>();

            StreamReader sr = new StreamReader(this._reProcessFilePath, Encoding.Default);
            string line;
            while ((line = sr.ReadLine()) != null)
            {
                //result.Append(line.Substring(line.IndexOf("b")) + Environment.NewLine);
                //Console.WriteLine(line.ToString());
                string currentLine = line.Substring(line.IndexOf("b"));
                result.Add(currentLine.Split(' ').ToList());
            }

            return result;
        }

    }
}
