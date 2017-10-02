using System;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data;
using System.Windows.Forms;
using System.Data.OleDb;
using Range = Microsoft.Office.Interop.Excel.Range;
using Workbook = Microsoft.Office.Interop.Excel.Workbook;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;
using ExcelTool = Microsoft.Office.Tools.Excel;

namespace Structural_Designer
{
    public partial class Calculate
    {
        private void Calculate_Load(object sender, RibbonUIEventArgs e)
        {
            this.btnThongsodam.Enabled = false;
            this.btnTinhtoandam.Enabled = false;
            this.btnThuyetminhdam.Enabled = false;
            this.btnVeDam.Enabled = false;
        }

        private void btnNewDam_Click(object sender, RibbonControlEventArgs e)
        {
            OpenFileDialog MDB = new OpenFileDialog();
            string sFile;
            MDB.Filter = "Access files|*.mdb";
            MDB.AddExtension = true;
            MDB.CheckPathExists = true;
            MDB.Title = "Chọn file dữ liệu đầu vào";

            if (MDB.ShowDialog() == DialogResult.OK)
            {
                sFile = MDB.FileName;
                try
                {
                    //Copy Beam force
                    #region
                    OleDbConnection MyConnection = new OleDbConnection();
                    OleDbCommand MyCommand = new OleDbCommand();
                    MyConnection = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + sFile);
                    MyConnection.Open();
                    string mySQL = "SELECT [Story],[Beam], [CaseCombo], [Station], P, V2, V3, T, M2, M3 " + "From [Beam Forces]";
                    OleDbCommand cmd = new OleDbCommand(mySQL, MyConnection);
                    OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                    DataTable dtb = new DataTable();
                    da.Fill(dtb);
                    MyConnection.Close();
                    var ds = new DataSet("temp");
                    ds.Tables.Add(dtb);
                    #endregion

                    //Copy Section
                    #region 
                    OleDbConnection MyConnection2 = new OleDbConnection();
                    OleDbCommand MyCommand2 = new OleDbCommand();
                    MyConnection2 = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + sFile);
                    MyConnection2.Open();
                    string mySQL2 = "SELECT [Story],[Label], [AnalysisSect] " + "From [Frame Assignments - Sections]";
                    OleDbCommand cmd2 = new OleDbCommand(mySQL2, MyConnection2);
                    OleDbDataAdapter da2 = new OleDbDataAdapter(cmd2);
                    DataTable dtb2 = new DataTable();
                    da2.Fill(dtb2);
                    MyConnection2.Close();
                    var ds2 = new DataSet("temp2");
                    ds2.Tables.Add(dtb2);
                    #endregion

                    //Copy Frame Sections
                    #region 
                    OleDbConnection MyConnection3 = new OleDbConnection();
                    OleDbCommand MyCommand3 = new OleDbCommand();
                    MyConnection3 = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + sFile);
                    MyConnection3.Open();
                    string mySQL3 = "SELECT [Name],[t2], [t3] " + "From [Frame Sections]";
                    OleDbCommand cmd3 = new OleDbCommand(mySQL3, MyConnection3);
                    OleDbDataAdapter da3 = new OleDbDataAdapter(cmd3);
                    DataTable dtb3 = new DataTable();
                    da3.Fill(dtb3);
                    MyConnection3.Close();
                    var ds3 = new DataSet("temp3");
                    ds3.Tables.Add(dtb3);
                    #endregion

                    //Copy Story Data
                    #region 
                    OleDbConnection MyConnection4 = new OleDbConnection();
                    OleDbCommand MyCommand4 = new OleDbCommand();
                    MyConnection4 = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + sFile);
                    MyConnection4.Open();
                    string mySQL4 = "SELECT [Name],[Height], [Elevation], [SimilarTo] " + "From [Story Data]";
                    OleDbCommand cmd4 = new OleDbCommand(mySQL4, MyConnection4);
                    OleDbDataAdapter da4 = new OleDbDataAdapter(cmd4);
                    DataTable dtb4 = new DataTable();
                    da4.Fill(dtb4);
                    MyConnection4.Close();
                    var ds4 = new DataSet("temp4");
                    ds4.Tables.Add(dtb4);
                    #endregion

                    //Copy Beam Connectivity
                    #region 
                    OleDbConnection MyConnection5 = new OleDbConnection();
                    OleDbCommand MyCommand5 = new OleDbCommand();
                    MyConnection5 = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + sFile);
                    MyConnection5.Open();
                    string mySQL5 = "SELECT [Story],[Label], [UniqueName], [Points] " + "From [Beam Connectivity]";
                    OleDbCommand cmd5 = new OleDbCommand(mySQL5, MyConnection5);
                    OleDbDataAdapter da5 = new OleDbDataAdapter(cmd5);
                    DataTable dtb5 = new DataTable();
                    da5.Fill(dtb5);
                    MyConnection5.Close();
                    var ds5 = new DataSet("temp5");
                    ds5.Tables.Add(dtb5);
                    #endregion

                    //Copy Joint Coordinates
                    #region 
                    OleDbConnection MyConnection6 = new OleDbConnection();
                    OleDbCommand MyCommand6 = new OleDbCommand();
                    MyConnection6 = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + sFile);
                    MyConnection6.Open();
                    string mySQL6 = "SELECT [Story],[Label], [UniqueName], [X], [Y], [Z] " + "From [Joint Coordinates]";
                    OleDbCommand cmd6 = new OleDbCommand(mySQL6, MyConnection6);
                    OleDbDataAdapter da6 = new OleDbDataAdapter(cmd6);
                    DataTable dtb6 = new DataTable();
                    da6.Fill(dtb6);
                    MyConnection6.Close();
                    var ds6 = new DataSet("temp6");
                    ds6.Tables.Add(dtb6);
                    #endregion


                    MessageBox.Show("Chọn vị trí lưu file! ");
                    SaveFileDialog Sfd = new SaveFileDialog();
                    string saveFile;
                    Sfd.Filter = "Excel files|*.xlsx";
                    Sfd.AddExtension = true;
                    Sfd.CheckPathExists = true;
                    Sfd.Title = "Chọn vị trí lưu file";
                    if (Sfd.ShowDialog() == DialogResult.OK)
                    {
                        saveFile = Sfd.FileName;
                        Workbook CurWB = Globals.ThisAddIn.Application.ActiveWorkbook;
                        Worksheet InputSheet = Globals.ThisAddIn.Application.ActiveSheet;
                        InputSheet.Name = "Input Data";

                        CurWB.SaveAs(saveFile);
                        // Định dạng dòng đầu tiên
                        #region
                        InputSheet.Range["A1", "J1"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Orange);
                        InputSheet.Range["A1", "J1"].Borders.LineStyle = Excel.Constants.xlSolid;
                        InputSheet.Range["A1", "J1"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        InputSheet.Range["A1", "J1"].Font.Bold = true;
                        #endregion
                        // Copy BeamForce
                        #region
                        object[,] rawData = new object[dtb.Rows.Count + 1, dtb.Columns.Count];
                        for (int col = 0; col < dtb.Columns.Count; col++)
                        {
                            rawData[0, col] = dtb.Columns[col].ColumnName;
                        }
                        for (int col = 0; col < dtb.Columns.Count; col++)
                        {
                            for (int row = 0; row < dtb.Rows.Count; row++)
                            {
                                rawData[row + 1, col] = dtb.Rows[row].ItemArray[col];
                            }
                        }
                        string finalColLetter = string.Empty;
                        string colCharset = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
                        int colCharsetLen = colCharset.Length;

                        if (dtb.Columns.Count > colCharsetLen)
                        {
                            finalColLetter = colCharset.Substring(
                                (dtb.Columns.Count - 1) / colCharsetLen - 1, 1);
                        }

                        finalColLetter += colCharset.Substring(
                                (dtb.Columns.Count - 1) % colCharsetLen, 1);
                        
                        string excelRange = string.Format("A1:{0}{1}", finalColLetter, dtb.Rows.Count + 1);
                        InputSheet.get_Range(excelRange, Type.Missing).Value2 = rawData;
                        InputSheet.get_Range(excelRange, Type.Missing).Font.Name = "Times New Roman";
                        InputSheet.get_Range(excelRange, Type.Missing).Font.Size = 11;
                        InputSheet.get_Range(excelRange, Type.Missing).Cells.HorizontalAlignment =Excel.XlHAlign.xlHAlignCenter;
                        #endregion

                        //Copy Section
                        #region
                        object[,] rawData2 = new object[dtb2.Rows.Count + 1, dtb2.Columns.Count];
                        for (int col = 0; col < dtb2.Columns.Count; col++)
                        {
                            rawData2[0, col] = dtb2.Columns[col].ColumnName;
                        }
                        for (int col = 0; col < dtb2.Columns.Count; col++)
                        {
                            for (int row = 0; row < dtb2.Rows.Count; row++)
                            {
                                rawData2[row + 1, col] = dtb2.Rows[row].ItemArray[col];
                            }
                        }
                        string finalColLetter2 = string.Empty;
                        string colCharset2 = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
                        int colCharsetLen2 = colCharset2.Length;

                        if (dtb2.Columns.Count > colCharsetLen2)
                        {
                            finalColLetter2 = colCharset2.Substring(
                                (dtb2.Columns.Count - 1) / colCharsetLen2 - 1, 1);
                        }

                        finalColLetter2 += colCharset2.Substring(
                                (dtb2.Columns.Count - 1) % colCharsetLen2, 1);
                        string excelRange2 = string.Format("L1:{0}{1}", "N", dtb2.Rows.Count + 1);
                        InputSheet.get_Range(excelRange2, Type.Missing).Value2 = rawData2;
                        InputSheet.get_Range(excelRange2, Type.Missing).Font.Name = "Times New Roman";
                        InputSheet.get_Range(excelRange2, Type.Missing).Font.Size = 11;
                        InputSheet.get_Range(excelRange2, Type.Missing).Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        #endregion

                        //Copy Section Properties
                        #region
                        object[,] rawData3 = new object[dtb3.Rows.Count + 1, dtb3.Columns.Count];
                        for (int col = 0; col < dtb3.Columns.Count; col++)
                        {
                            rawData3[0, col] = dtb3.Columns[col].ColumnName;
                        }
                        for (int col = 0; col < dtb3.Columns.Count; col++)
                        {
                            for (int row = 0; row < dtb3.Rows.Count; row++)
                            {
                                rawData3[row + 1, col] = dtb3.Rows[row].ItemArray[col];
                            }
                        }
                        string finalColLetter3 = string.Empty;
                        string colCharset3 = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
                        int colCharsetLen3 = colCharset3.Length;

                        if (dtb3.Columns.Count > colCharsetLen3)
                        {
                            finalColLetter3 = colCharset3.Substring(
                                (dtb3.Columns.Count - 1) / colCharsetLen3 - 1, 1);
                        }

                        finalColLetter3 += colCharset3.Substring(
                                (dtb3.Columns.Count - 1) % colCharsetLen3, 1);
                        string excelRange3 = string.Format("P1:{0}{1}", "R", dtb3.Rows.Count + 1);
                        InputSheet.get_Range(excelRange3, Type.Missing).Value2 = rawData3;
                        InputSheet.get_Range(excelRange3, Type.Missing).Font.Name = "Times New Roman";
                        InputSheet.get_Range(excelRange3, Type.Missing).Font.Size = 11;
                        InputSheet.get_Range(excelRange3, Type.Missing).Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        #endregion

                        //Copy Story data
                        #region
                        object[,] rawData4 = new object[dtb4.Rows.Count + 1, dtb4.Columns.Count];
                        for (int col = 0; col < dtb4.Columns.Count; col++)
                        {
                            rawData4[0, col] = dtb4.Columns[col].ColumnName;
                        }
                        for (int col = 0; col < dtb4.Columns.Count; col++)
                        {
                            for (int row = 0; row < dtb4.Rows.Count; row++)
                            {
                                rawData4[row + 1, col] = dtb4.Rows[row].ItemArray[col];
                            }
                        }
                        string finalColLetter4 = string.Empty;
                        string colCharset4 = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
                        int colCharsetLen4 = colCharset4.Length;

                        if (dtb4.Columns.Count > colCharsetLen4)
                        {
                            finalColLetter4 = colCharset4.Substring(
                                (dtb4.Columns.Count - 1) / colCharsetLen4 - 1, 1);
                        }

                        finalColLetter4 += colCharset4.Substring(
                                (dtb4.Columns.Count - 1) % colCharsetLen4, 1);
                        string excelRange4 = string.Format("T1:{0}{1}", "W", dtb4.Rows.Count + 1);
                        InputSheet.get_Range(excelRange4, Type.Missing).Value2 = rawData4;
                        InputSheet.get_Range(excelRange4, Type.Missing).Font.Name = "Times New Roman";
                        InputSheet.get_Range(excelRange4, Type.Missing).Font.Size = 11;
                        InputSheet.get_Range(excelRange4, Type.Missing).Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        #endregion

                        //Copy Beam Connectivity
                        #region
                        object[,] rawData5 = new object[dtb5.Rows.Count + 1, dtb5.Columns.Count];
                        for (int col = 0; col < dtb5.Columns.Count; col++)
                        {
                            rawData5[0, col] = dtb5.Columns[col].ColumnName;
                        }
                        for (int col = 0; col < dtb5.Columns.Count; col++)
                        {
                            for (int row = 0; row < dtb5.Rows.Count; row++)
                            {
                                rawData5[row + 1, col] = dtb5.Rows[row].ItemArray[col];
                            }
                        }
                        string finalColLetter5 = string.Empty;
                        string colCharset5 = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
                        int colCharsetLen5 = colCharset5.Length;

                        if (dtb5.Columns.Count > colCharsetLen5)
                        {
                            finalColLetter5 = colCharset5.Substring(
                                (dtb5.Columns.Count - 1) / colCharsetLen5 - 1, 1);
                        }

                        finalColLetter5 += colCharset5.Substring(
                                (dtb5.Columns.Count - 1) % colCharsetLen5, 1);
                        string excelRange5 = string.Format("Y1:{0}{1}", "AB", dtb5.Rows.Count + 1);
                        InputSheet.get_Range(excelRange5, Type.Missing).Value2 = rawData5;
                        InputSheet.get_Range(excelRange5, Type.Missing).Font.Name = "Times New Roman";
                        InputSheet.get_Range(excelRange5, Type.Missing).Font.Size = 11;
                        InputSheet.get_Range(excelRange5, Type.Missing).Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        #endregion


                        //Copy Joint Coordinates
                        #region
                        object[,] rawData6 = new object[dtb6.Rows.Count + 1, dtb6.Columns.Count];
                        for (int col = 0; col < dtb6.Columns.Count; col++)
                        {
                            rawData6[0, col] = dtb6.Columns[col].ColumnName;
                        }
                        for (int col = 0; col < dtb6.Columns.Count; col++)
                        {
                            for (int row = 0; row < dtb6.Rows.Count; row++)
                            {
                                rawData6[row + 1, col] = dtb6.Rows[row].ItemArray[col];
                            }
                        }
                        string finalColLetter6 = string.Empty;
                        string colCharset6 = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
                        int colCharsetLen6 = colCharset6.Length;

                        if (dtb6.Columns.Count > colCharsetLen6)
                        {
                            finalColLetter6 = colCharset6.Substring(
                                (dtb6.Columns.Count - 1) / colCharsetLen6 - 1, 1);
                        }

                        finalColLetter6 += colCharset6.Substring(
                                (dtb6.Columns.Count - 1) % colCharsetLen6, 1);
                        string excelRange6 = string.Format("AI1:{0}{1}", "AN", dtb6.Rows.Count + 1);
                        InputSheet.get_Range(excelRange6, Type.Missing).Value2 = rawData6;
                        InputSheet.get_Range(excelRange6, Type.Missing).Font.Name = "Times New Roman";
                        InputSheet.get_Range(excelRange6, Type.Missing).Font.Size = 11;
                        InputSheet.get_Range(excelRange6, Type.Missing).Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        #endregion

                        // Định dạng bảng tính
                        #region
                        InputSheet.Columns.NumberFormat = "0.00";
                        InputSheet.Columns[4].NumberFormat = "0.000";

                        InputSheet.Columns.ColumnWidth = 8;
                        InputSheet.Columns[1].AutoFit();
                        InputSheet.Columns[2].AutoFit();
                        InputSheet.Columns[3].AutoFit();
                        InputSheet.Columns[12].AutoFit();
                        InputSheet.Columns[13].AutoFit();
                        InputSheet.Columns[14].AutoFit();
                        InputSheet.Columns[16].AutoFit();
                        InputSheet.Columns[20].AutoFit();

                        InputSheet.Columns[11].ColumnWidth = 5;
                        InputSheet.Columns[15].ColumnWidth = 5;
                        InputSheet.Columns[19].ColumnWidth = 5;
                        #endregion
                        CurWB.Save();
                        btnThongsodam.Enabled = true;

                        //Chuyển sang sheet2
                        Worksheet CalSheet = CurWB.Worksheets.Add();
                        CalSheet.Name = "Calculation";
                        //CalSheet.Activate();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: " + ex);
                }


            }
        }

        private void btnOpenDam_Click(object sender, RibbonControlEventArgs e)
        {
            OpenFileDialog OFD = new OpenFileDialog();
            OFD.Filter = "Excel files|*.xlsx; *.xlsm";
            OFD.AddExtension = true;
            OFD.CheckPathExists = true;
            OFD.Title = "Chọn một File đã có sẵn dữ liệu";
            if (OFD.ShowDialog() == DialogResult.OK)
            {
                string WBname = OFD.FileName;
                Globals.ThisAddIn.Application.ActiveWorkbook.Close();
                Globals.ThisAddIn.Application.Workbooks.Open(WBname);
                btnThongsodam.Enabled = true;
            }
        }

        private void btnBeamData_Click(object sender, RibbonControlEventArgs e)
        {
            Workbook WB = Globals.ThisAddIn.Application.ActiveWorkbook;
            Worksheet Input = WB.Worksheets["Input Data"];
            for (double i = 2; Input.Range["AB" + i].Value!=null; i++)
            {
                string s = Convert.ToString(Input.Range["AB" + i].Value);
                string [] sub =s.Split(';');
                Input.Range["AB" + i].Value = sub[0];
                Input.Range["AC" + i].Value = sub[1];
            }
            Input.Range["AB1"].Value = "Start Point";
            Input.Range["AC1"].Value = "End Point";
            Input.Columns[27].NumberFormat = "0";
            Input.Columns["AB"].NumberFormat = "0";
            Input.Columns["AC"].NumberFormat = "0";
            Input.Columns["AJ"].NumberFormat = "0";
            Input.Columns["AK"].NumberFormat = "0";
        }
    }



}
