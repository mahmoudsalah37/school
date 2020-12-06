using Microsoft.Office.Interop.Excel;
using System;
using System.Threading;
using System.Windows.Forms;
namespace School
{
    public partial class Form1 : Form
    {
        int rowIndex = 0;
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
        }
        //get data from excel
        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog d = new OpenFileDialog();
            d.Filter = "Excel Files(.xls)|*.xls| Excel Files(.xlsx)|*.xlsx";
            d.Multiselect = true;
            if (d.ShowDialog() == DialogResult.OK)
            {
                import.Enabled = false;
                save.Enabled = false;
                ThreadPool.QueueUserWorkItem(state =>
                {
                    int number = 0;
                    foreach (string path in d.FileNames)
                    {
                        loadProgressBar.Invoke(new System.Action(() => this.loadProgressBar.Value = 0));
                        loadLabel.Invoke(new System.Action(() => this.loadLabel.Text = "0%"));
                        numberSheets.Invoke(new System.Action(() => this.numberSheets.Text = " جاري التحميل " + "... " + number + "/" + d.FileNames.Length));
                        Excel excel = new Excel(path, 2);
                        int row = 25;
                        int column = 38;
                        try
                        {
                            getData(excel, row, column);
                        }
                        catch (Exception) {
                        }
                        finally
                        {
                            //Close excel file
                            excel.closeFile();
                        }
                        number++;
                        loadProgressBar.Invoke(new System.Action(() => this.loadProgressBar.Value = 100));
                        loadLabel.Invoke(new System.Action(() => this.loadLabel.Text = "100%"));
                        Thread.Sleep(500);
                    }
                    numberSheets.Invoke(new System.Action(() => this.numberSheets.Text = " تم التحميل " + number + "/" + d.FileNames.Length));
                    import.Invoke(new System.Action(() => this.import.Enabled = true));
                    save.Invoke(new System.Action(() => this.save.Enabled = true));
                });
            }
        }

        //save data in excel
        private void save_Click(object sender, EventArgs e)
        {
            import.Enabled = false;
            save.Enabled = false;
            //ThreadPool.QueueUserWorkItem(state =>
            //{
            loadProgressBar.Invoke(new System.Action(() => this.loadProgressBar.Value = 0));
            loadLabel.Invoke(new System.Action(() => this.loadLabel.Text = "0%"));
            numberSheets.Invoke(new System.Action(() => this.numberSheets.Text = " جاري الحفظ " + "... "  + "0/1" ));
            _Application app = new Microsoft.Office.Interop.Excel.Application();
            Workbook workbook = app.Workbooks.Add(Type.Missing);
            Worksheet worksheet = null;
            try
            {
                worksheet = workbook.Sheets["ورقة1"];
            }
            catch (Exception)
            {
                worksheet = workbook.Sheets["sheet1"];

            }
            worksheet = workbook.ActiveSheet;
            worksheet.DisplayRightToLeft = true;
            worksheet.Name = "MAS_SPHierarchyDetailedReport";
            //size columns
            worksheet.Columns[1].ColumnWidth = 25;
            worksheet.Columns[2].ColumnWidth = 15;
            worksheet.Columns[3].ColumnWidth = 15;
            worksheet.Columns[4].ColumnWidth = 15;
            worksheet.Columns[5].ColumnWidth = 15;
            worksheet.Columns[6].ColumnWidth = 15;
            worksheet.Columns[7].ColumnWidth = 15;
            worksheet.Columns[8].ColumnWidth = 15;
            worksheet.Columns[9].ColumnWidth = 15;
            worksheet.Columns[10].ColumnWidth = 15;
            worksheet.Columns[11].ColumnWidth = 15;
            worksheet.Columns[12].ColumnWidth = 15;
            worksheet.Columns[13].ColumnWidth = 15;
            worksheet.Columns[14].ColumnWidth = 15;
            worksheet.Columns[15].ColumnWidth = 15;
            worksheet.Columns[16].ColumnWidth = 15;
            worksheet.Columns[17].ColumnWidth = 15;

            for (int i=1;i < dataGridView1.Columns.Count + 1; i++)
                {
                    worksheet.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;
                }
                double total = dataGridView1.Rows.Count + dataGridView1.Rows.Count;
                for (int i = 0; i< dataGridView1.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridView1.Columns.Count; j++)
                    {
                        loadProgressBar.Invoke(new System.Action(() => this.loadProgressBar.Value = Convert.ToInt32((i+j)/total *100)));
                        loadLabel.Invoke(new System.Action(() => this.loadLabel.Text = Convert.ToInt32((i + j) / total * 100) + "%"));
                        worksheet.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value;
                    }
                }
                var saveFile = new SaveFileDialog();
                saveFile.FileName = "output";
                saveFile.DefaultExt = ".xlsx";
                if(saveFile.ShowDialog() == DialogResult.OK)
                {
                try { 
                    workbook.SaveAs(saveFile.FileName,Type.Missing, Type.Missing, Type.Missing , Type.Missing, Type.Missing,Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                }
                catch (Exception)
                {
                    MessageBox.Show("أغلق ملف الإكسل المراد التعديل عليه");
                }
                finally
                {
                    //workbook.Close(true);
                    app.Quit();
                    //System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
                    import.Invoke(new System.Action(() => this.import.Enabled = true));
                    save.Invoke(new System.Action(() => this.save.Enabled = true));
                    loadProgressBar.Invoke(new System.Action(() => this.loadProgressBar.Value = 100));
                    loadLabel.Invoke(new System.Action(() => this.loadLabel.Text = "100%"));
                    numberSheets.Invoke(new System.Action(() => this.numberSheets.Text = " تم الحفظ " + "... " + "1/1"));
                }

            }
               
             //});
        }

        private void getData(Excel excel, int row, int column)
        {
            //school's name
            var school = excel.readCell(13, 31);
            //Grade's name
            var grade = excel.readCell(2, 3);
            //Section's name
            var section = excel.readCell(5, 3);
            //class's name
            var classIsName = excel.readCell(9, 3);
            //year's from to
            var year = excel.readCell(12, 4);
            double cellsNumber = excel.cellsCount(excel, row,column);
            double counter = 0;
            while (row <= cellsNumber -4) {
                loadProgressBar.Invoke(new System.Action(() => this.loadProgressBar.Value = Convert.ToInt32(counter / cellsNumber * 100)));
                loadLabel.Invoke(new System.Action(() => this.loadLabel.Text = Convert.ToInt32(counter / cellsNumber * 100) + "%"));
                //add row
                dataGridView1.Invoke(new System.Action(() => this.dataGridView1.Rows.Add()));
               
                //insert data in view
                dataGridView1.Rows[rowIndex].Cells[1].Value = school;
                dataGridView1.Rows[rowIndex].Cells[2].Value = grade;
                dataGridView1.Rows[rowIndex].Cells[3].Value = section;
                dataGridView1.Rows[rowIndex].Cells[4].Value = classIsName;
                dataGridView1.Rows[rowIndex].Cells[5].Value = year;
                //Koran's name
                var koran = excel.readCell(row, column);
                //insert data in view
                dataGridView1.Rows[rowIndex].Cells[6].Value = koran;

                //column explanation's degree
                column = 39;
                //explanation's degree
                var type = excel.readCell(row, column);
                //insert data in view
                dataGridView1.Rows[rowIndex].Cells[0].Value = type;

                //column explanation's degree
                column = 37;
                //explanation's degree
                var explanation = excel.readCell(row, column);
                //insert data in view
                dataGridView1.Rows[rowIndex].Cells[7].Value = explanation;

                //column Monotheism's degree
                column = 36;
                //Monotheism's degree
                var monotheism = excel.readCell(row, column);
                //insert data in view
                dataGridView1.Rows[rowIndex].Cells[8].Value = monotheism;

                //column Interpretation's degree
                column = 34;
                //Interpretation's degree
                var interpretation = excel.readCell(row, column);
                //insert data in view
                dataGridView1.Rows[rowIndex].Cells[9].Value = interpretation;

                //column tradition's degree
                column = 33;
                //Interpretation's degree
                var tradition = excel.readCell(row, column);
                //insert data in view
                dataGridView1.Rows[rowIndex].Cells[10].Value = tradition;
                //column language's degree
                column = 32;
                //language's degree
                var language = excel.readCell(row, column);
                //insert data in view
                dataGridView1.Rows[rowIndex].Cells[11].Value = language;
                //column Studies's degree
                column = 29;
                //Studies's degree
                var Studies = excel.readCell(row, column);
                //insert data in view
                dataGridView1.Rows[rowIndex].Cells[12].Value = Studies;
                //column math 's degree
                column = 28;
                //math's degree
                var math = excel.readCell(row, column);
                //insert data in view
                dataGridView1.Rows[rowIndex].Cells[13].Value = math;
                //column science 's degree
                column = 27;
                //science's degree
                var science = excel.readCell(row, column);
                //insert data in view
                dataGridView1.Rows[rowIndex].Cells[14].Value = science;
                //column eng 's degree
                column = 24;
                //eng's degree
                var eng = excel.readCell(row, column);
                //insert data in view
                dataGridView1.Rows[rowIndex].Cells[15].Value = eng;
                //column computer 's degree
                column = 20;
                //computer's degree
                var computer = excel.readCell(row, column);
                //insert data in view
                dataGridView1.Rows[rowIndex].Cells[16].Value = computer;
                //column Technical 's degree
                column = 15;
                //Technical's degree
                var Technical = excel.readCell(row, column);
                //insert data in view
                dataGridView1.Rows[rowIndex].Cells[17].Value = Technical;
                //column Technical 's degree
                column = 8;
                //debt's degree
                var debt = excel.readCell(row, column);
                //insert data in view
                dataGridView1.Rows[rowIndex].Cells[18].Value = debt;
                rowIndex++;
                if (counter % 8 != 0)
                    row++;
                else if(counter >7)
                    row += 4;
                counter++;
                //to check in student is name index
                //column student's name
                column = 38;
            }
        }

        
    }
}
