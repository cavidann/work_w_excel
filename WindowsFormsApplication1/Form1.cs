using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;


namespace WindowsFormsApplication1
{

    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
        }
        class CreateExcelDoc
        {
            private Excel.Application app = null;
            private Excel.Workbook workbook = null;
            private Excel.Worksheet worksheet = null;
            private Excel.Range workSheet_range = null;
            public CreateExcelDoc()
            {
                createDoc();
            }
            public void createDoc()
            {
                try
                {
                    app = new Excel.Application();
                    app.Visible = true;
                    workbook = app.Workbooks.Add(1);
                    worksheet = (Excel.Worksheet)workbook.Sheets[1];
                }
                catch (Exception e)
                {
                    Console.Write("Error");
                }
                finally
                {
                }
            }

            public void createHeaders(int row, int col, string htext, string cell1,
            string cell2, int mergeColumns, string b, bool font, int size, string
            fcolor)
            {
                worksheet.Cells[row, col] = htext;
                workSheet_range = worksheet.get_Range(cell1, cell2);
                workSheet_range.Merge(mergeColumns);
                switch (b)
                {
                    case "YELLOW":
                        workSheet_range.Interior.Color = System.Drawing.Color.Yellow.ToArgb();
                        break;
                    case "WHITE":
                        workSheet_range.Interior.Color = System.Drawing.Color.White.ToArgb();
                        break;
                    case "GAINSBORO":
                        workSheet_range.Interior.Color =
                System.Drawing.Color.Gainsboro.ToArgb();
                        break;
                    case "Turquoise":
                        workSheet_range.Interior.Color =
                System.Drawing.Color.Turquoise.ToArgb();
                        break;
                    case "PeachPuff":
                        workSheet_range.Interior.Color =
                System.Drawing.Color.PeachPuff.ToArgb();
                        break;
                    default:
                        //  workSheet_range.Interior.Color = System.Drawing.Color..ToArgb();
                        break;
                }

                //workSheet_range.Borders.Color = System.Drawing.Color.Black.ToArgb();
                worksheet.Cells[row, col].Font.Bold = font;
                workSheet_range.ColumnWidth = size;
                if (fcolor.Equals(""))
                {
                    workSheet_range.Font.Color = System.Drawing.Color.White.ToArgb();
                }
                else
                {
                    workSheet_range.Font.Color = System.Drawing.Color.Black.ToArgb();
                }
            }

            public void addData(int row, int col, string data,
                string cell1, string cell2, string format)
            {
                worksheet.Cells[row, col] = data;
                workSheet_range = worksheet.get_Range(cell1, cell2);
                workSheet_range.Borders.Color = System.Drawing.Color.Black.ToArgb();
                workSheet_range.NumberFormat = format;
            }
        }
        int imtahan_sayi = 0;
        int interval = 0;
        int next = 0;
        public void mainFunc()
        {

        }
        private void upload_Click(object sender, EventArgs e)
        {

            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;
            CreateExcelDoc excell_app = new CreateExcelDoc();

            string str;
            string str1;
            int rCnt;
            int cCnt;
            int rw = 0;
            int cl = 0;
            int gc = 0;
            DateTime now = DateTime.Now;
            List<int> fcounts = new List<int>();
            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(@"d:\Test.xlsx", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            range = xlWorkSheet.UsedRange;
            rw = range.Rows.Count;
            cl = range.Columns.Count;

            for (cCnt = 1; cCnt <= cl; cCnt++)
            {
                str = (string)(range.Cells[1, cCnt] as Excel.Range).Value2;
                excell_app.createHeaders(1, cCnt, str, "Z2", "Z2", 2, "WHITE", true, 10, "n");
            }


            for (cCnt = 1; cCnt <= cl; cCnt++)
            {
                for (rCnt = 2; rCnt <= rw; rCnt++)
                {
                    str = (string)(range.Cells[rCnt, cCnt] as Excel.Range).Value2;
                    if ((string)(range.Cells[rCnt, 3] as Excel.Range).Value2 == (string)(range.Cells[rCnt + 1, 3] as Excel.Range).Value2)
                    {
                        gc++;
                        excell_app.addData(rCnt, cCnt, str, "Z3", "Z3", "#,##0");
                    }
                    else if ((string)(range.Cells[rCnt, 3] as Excel.Range).Value2 == (string)(range.Cells[rCnt - 1, 3] as Excel.Range).Value2)
                    {
                        gc++;
                        excell_app.addData(rCnt, cCnt, str, "Z3", "Z3", "#,##0");
                        if (cCnt == cl)
                        {
                            fcounts.Add(gc);
                        }
                        gc = 0;
                    }
                }
            }
            int group_num_st = 0;
            int group_num_en = 0;
            int start_date = 0;
            for (int dif_group = 0; dif_group < fcounts.Count; dif_group++)
            {
                    start_date++;
                if (start_date == 4)
                {
                    start_date = 1;
                }
                //MessageBox.Show(dif_group.ToString());
                //start_date = dif_group+1;
                if (dif_group == 0)
                {
                    int say = 0;

                    excell_app.createHeaders(1, cl + 1, "Tarix", "Z2", "Z2", 2, "WHITE", true, 10, "n");

                    for (int t = 0; t < fcounts[dif_group]; t++)
                    {
                        say++;
                        //Random start = new Random();
                        //int start_date = start.Next(1, 5);
                        interval = 30 / fcounts[dif_group];
                        if (say == 1)
                        {
                            excell_app.addData(t + 2, cl + 1, (start_date + "." + now.Month + "." + now.Year).ToString(), "Z3", "Z3", "#,##0");
                        }
                        else
                        {
                            next = start_date + interval * (say-1);
                            int rdate = next;
                            int month = 0;
                            if (next > 30)
                            {
                                rdate = next % 30;
                                month = next / 30;
                                int rmonth = now.Month + month;
                                excell_app.addData(t + 2, cl + 1, (rdate + "." + rmonth + "." + now.Year).ToString(), "Z3", "Z3", "#,##0");
                                //MessageBox.Show((rdate + "/" + rmonth + "/" + now.Year).ToString());
                            }
                            else
                            {
                                excell_app.addData(t + 2, cl + 1, (rdate + "." + now.Month + "." + now.Year).ToString(), "Z3", "Z3", "#,##0");
                                ////MessageBox.Show((rdate + "/" + now.Month + "/" + now.Year).ToString());
                            }
                        }
                    }
                    //birinci qrup
                }
                else
                {
                    int say = 0;
                    //yerde qalan qruplar
                    group_num_st += fcounts[dif_group - 1];
                    group_num_en = group_num_st + fcounts[dif_group];
                    for (int t = group_num_st; t < group_num_en; t++)
                    {
                        //excell_app.addData(t + 2, cl + 1, fcounts[dif_group].ToString(), "Z3", "Z3", "#,##0");
                        say++;
                        //Random start = new Random();
                        //int start_date = start.Next(1, 5);
                        interval = 30 / fcounts[dif_group];
                        if (say == 1)
                        {
                            excell_app.addData(t + 2, cl + 1, (start_date + "." + now.Month + "." + now.Year).ToString(), "Z3", "Z3", "#,##0");
                        }
                        else
                        {
                            next = start_date + interval * (say - 1);
                            int rdate = next;
                            int month = 0;
                            if (next > 30)
                            {
                                rdate = next % 30;
                                month = next / 30;
                                int rmonth = now.Month + month;
                                excell_app.addData(t + 2, cl + 1, (rdate + "." + rmonth+"."+now.Year).ToString(), "Z3", "Z3", "#,##0");
                                //MessageBox.Show((rdate + "/" + rmonth + "/" + now.Year).ToString());
                            }
                            else
                            {
                                excell_app.addData(t + 2, cl + 1, (rdate + "." + now.Month + "." + now.Year).ToString(), "Z3", "Z3", "#,##0");
                                ////MessageBox.Show((rdate + "/" + now.Month + "/" + now.Year).ToString());
                            }
                        }
                    }
                }
            }


            xlWorkBook.Close(true, null, null);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {

        }
    }
}
