using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace ExcelHandler
{
    class Program
    {
        static void Main(string[] args)
        {

            string filePath = @"D:\test.xlsx";
            string strCon = "Provider=Microsoft.Ace.OLEDB.12.0;" + "Data Source=" + filePath + ";" + "Extended Properties=Excel 12.0;";

            Program.PreProcess(strCon);

        }

        /// <summary>
        /// Pre 
        /// </summary>
        /// <param name="strCon"></param>
        public static void PreProcess(string strCon)
        {
            OleDbConnection connection = new OleDbConnection(strCon);
            connection.Open();
            string strExcel = "";
            OleDbDataAdapter adapter = null;
            DataSet ds = null;
            strExcel = "select top 1000 * from [Sheet1$]";
            adapter = new OleDbDataAdapter(strExcel, strCon);
            ds = new DataSet();
            adapter.Fill(ds, "table1");
            DataTable table = ds.Tables[0];

            StringBuilder result = new StringBuilder();
            Console.WriteLine(table.Rows.Count);

            ArrayList list = new ArrayList();

            int i = 0;
            int j = 0;
            for (i = 0; i < table.Rows.Count; i++)
            {
                list.Add(table.Rows[i][1].ToString());
                //result.Append(table.Rows[i][1].ToString());
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
                if (list.Count == 1)
                {
                    list.Clear();
                    continue;
                }
                list.Sort();
                foreach (string item in list)
                {
                    result.Append(item + ",");
                }
                list.Clear();
                result.Append(Environment.NewLine);
                //Console.WriteLine(i);
            }

            //Console.Write(result.ToString());

            //String[] tempResult_1 = result.ToString().Split(Environment.NewLine.ToCharArray());



            FileStream output = new FileStream(@"D:\test.txt", FileMode.Create);
            StreamWriter writer = new StreamWriter(output, Encoding.UTF8);
            writer.Write(result);
            writer.Close();
            output.Close();

            Console.ReadKey();
        }

    }
}
