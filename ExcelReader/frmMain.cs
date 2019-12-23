using System;
using System.Data;
using System.Windows.Forms;
using System.Collections;
using System.Threading;

namespace ExcelReader
{
    public partial class frmMain : Form
    {
        private Thread ReadThread;
        private Thread ProcessThread;

        
        private delegate void ReadProgessCallback(int rowsPassed, int totalCount, DataTable dataTable);
        private delegate void ProcessProgressCallBack();


        private string MainFileName
        {
            get;
            set;
        }

        private void ReadRecords()
        {
            string fileName = MainFileName;

            if (fileName != "")
            {
                Microsoft.Office.Interop.Excel.Application xlApp;
                Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
                Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
                Microsoft.Office.Interop.Excel.Range range;



                xlApp = new Microsoft.Office.Interop.Excel.Application();
                xlWorkBook = xlApp.Workbooks.Open(fileName, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);


                range = xlWorkSheet.UsedRange;

                DataTable dataTable = new DataTable();

                try
                {
                    for (int rowCount = 1; rowCount <= range.Rows.Count; rowCount++)
                    {
                        this.Invoke(new ReadProgessCallback(this.ReadProgress), new Object[] { rowCount, range.Rows.Count, dataTable });


                        string prevColumnName = "";
                        int prevColumnCounter = 1;
                        ArrayList rowValues = new ArrayList();

                        for (int columnCount = 1; columnCount <= range.Columns.Count; columnCount++)
                        {
                            if (rowCount == 1)
                            {
                                var columnHeader = (range.Cells[rowCount, columnCount] as Microsoft.Office.Interop.Excel.Range).Value;

                                if (columnHeader != null)
                                {
                                    dataTable.Columns.Add(new DataColumn(columnHeader.ToString()));
                                    prevColumnName = columnHeader.ToString();
                                    prevColumnCounter = 1;
                                }
                                else
                                {
                                    dataTable.Columns.Add(prevColumnName + prevColumnCounter.ToString());
                                    prevColumnCounter++;
                                }
                            }
                            else
                            {
                                var columnValue = (range.Cells[rowCount, columnCount] as Microsoft.Office.Interop.Excel.Range).Value;

                                if (columnValue != null)
                                    rowValues.Add(columnValue.ToString());
                                else
                                    rowValues.Add("");
                            }
                        }

                        if (rowValues.Count > 0)
                        {
                            dataTable.Rows.Add(rowValues.ToArray());
                        }
                    }
                }
                catch (DuplicateNameException)
                {
                    MessageBox.Show("Duplicate column name in selected xls file.", "", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    dataTable = null;
                }
                catch (Exception exc)
                {
                    MessageBox.Show(exc.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    dataTable = null;
                }


                xlWorkBook.Close(true, null, null);
                xlApp.Quit();


                ReleaseObject(xlWorkSheet);
                ReleaseObject(xlWorkBook);
                ReleaseObject(xlApp);
            }
        }

        private void ReadProgress(int currentIndex, int totalCount, DataTable dataTable)
        {
            progressBar.Maximum = totalCount;
            progressBar.Value = currentIndex;


            lblPercent.Text = ((100 * currentIndex) / totalCount).ToString() + "% completed.";


            if (currentIndex == totalCount)
            {
                grdExcel.DataSource = dataTable;
                grdExcel.Columns[0].Width = 50;


                btnProcess.Enabled = true;
                lblTotal.Text = (totalCount - 1).ToString();
                lblRemaining.Text = (totalCount - 1).ToString();
            }
        }

        private void ProcessRecords()
        {
            this.Invoke(new ProcessProgressCallBack(this.ProcessProgress), new object[] { });
        }

        private void ProcessProgress()
        {
            btnProcess.Enabled = false;


            int processed = 0;
            foreach (DataGridViewRow dataGridViewRow in grdExcel.Rows)
            {
                DataGridViewCheckBoxCell chkSelected = (DataGridViewCheckBoxCell)dataGridViewRow.Cells[0];

                if (chkSelected.Value != null && bool.Parse(chkSelected.Value.ToString()) == true)
                    processed++;

                // Do whatever you want with record.
                // SomeFunction()
            }


            lblProcessed.Text = processed.ToString();
            lblRemaining.Text = (int.Parse(lblTotal.Text) - int.Parse(lblProcessed.Text)).ToString();

            btnProcess.Enabled = true;
        }

        private void ReleaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        public frmMain()
        {
            InitializeComponent();
        }

        private void frmMain_Load(object sender, EventArgs e)
        {
        }
        
        private void btnRead_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            DialogResult dialogResult = openFileDialog.ShowDialog();

            if (dialogResult == DialogResult.OK)
            {
                MainFileName = openFileDialog.FileName;
                lblPercent.Text = "Reading file...";


                ReadThread = new Thread(ReadRecords);
                ReadThread.Start();


                btnRead.Enabled = false;
            }
        }

        private void btnProcess_Click(object sender, EventArgs e)
        {
            ProcessThread = new Thread(ProcessRecords);
            ProcessThread.Start();
        }
    }
}
