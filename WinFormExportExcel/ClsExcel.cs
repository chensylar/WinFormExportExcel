using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace WinFormExportExcel
{
    class ClsExcel
    {

        private Microsoft.Office.Interop.Excel.Application xApp;
        private Microsoft.Office.Interop.Excel.Workbook xWorkBook;
        public Microsoft.Office.Interop.Excel.Worksheet xSheet;
        private object m_objMiss = System.Reflection.Missing.Value;

        public Microsoft.Office.Interop.Excel.Application X_App
        {
            get
            {
                return xApp;
            }
            set
            {
                xApp = value;
            }
        }

        public Excel.Workbook X_WorkBook
        {
            get
            {
                return xWorkBook;
            }
            set
            {
                xWorkBook = value;
            }
        }

        public Excel.Worksheet X_Sheet
        {
            get
            {
                return xSheet;
            }
            set
            {
                xSheet = value;
            }
        }

        public object ObjMiss
        {
            get
            {
                return m_objMiss;
            }
        }

        /// <summary>
        /// 构造函数
        /// </summary>
        public ClsExcel()
        {
            try
            {
                xApp = new Excel.Application();
                xWorkBook = xApp.Workbooks.Add(true);

                //xWorkBook.Sheets.Add(null, null, 0, null);
                xSheet = (Excel.Worksheet)xWorkBook.ActiveSheet;
                // 
            }
            catch (Exception e)
            {
                throw (new Exception("[ClsExcel]:" + e.Message));
            }
        }
        /// <summary>
        /// 打开所要保存的文件
        /// </summary>
        /// <returns></returns>
        public static string GetSaveFileName(string fileName)
        {
            string filename = string.Empty;
            SaveFileDialog dlgSave = new SaveFileDialog();

            dlgSave.Title = "导出文件为";
            dlgSave.Filter = "Microsoft Office Excel 工作簿(*.xls)|*.xls";


            dlgSave.FileName = fileName;

            if (dlgSave.ShowDialog() == DialogResult.Cancel)
            {
                return string.Empty;
            }

            filename = dlgSave.FileName;
            dlgSave.OverwritePrompt = true;
            return filename;
        }
        public void ExportToExcel(ref DataGridView dgv, string fileName, string sheetName)
        {
            ExportToExcel(ref dgv, 0, 0, fileName, sheetName);
        }
        public void ExportToExcel(ref DataGridView dgv, int rowBeginNum,
                                           int colBeginNum, string fileName, string sheetName)
        {
            ExportToExcel(ref dgv, rowBeginNum, colBeginNum, dgv.RowCount, dgv.ColumnCount, fileName, sheetName);
        }
        public void ExportToExcel(ref DataGridView dgv, int rowBeginNum,
                                           int colBeginNum, int rowEndNum, int colEndNum, string fileName, string sheetName)
        {
            string theFineName = string.Empty;
            theFineName = GetSaveFileName(fileName);
            int rowNum = rowEndNum;
            int columnNum = colEndNum;
            int rowIndex = 1;
            int columnIndex = 1;
            //FrmProgressBar fmPB = new FrmProgressBar();
            try
            {
                //fmPB = ClsPubFuctions.GetProgressBarForm("耐心等待。。。", rowNum + 10);
                //fmPB.Show();
                Application.DoEvents();
                //写入头
                for (int j = colBeginNum; j < columnNum; j++, columnIndex++)
                {
                    SetValue(rowIndex, columnIndex, dgv.Columns[j].HeaderText);
                }
                //写入每一行
                rowIndex = 1;
                for (int i = rowBeginNum; i < rowNum; i++)
                {
                    rowIndex++;
                    columnIndex = 0;
                    Application.DoEvents();
                    for (int j = colBeginNum; j < columnNum; j++) // 列循环
                    {
                        columnIndex++;
                        if (null != dgv.Rows[i].Cells[j].Value)
                        {
                            string temp = dgv.Rows[i].Cells[j].Value.ToString();
                            SetValue(rowIndex, columnIndex, temp);
                        }
                    }
                    //fmPB.ProgressBar.Value += 1;
                }
                //fmPB.Close();

                //FormatHeadText(1, columnIndex);


                //SDX-修改于-2012-06-22
                //SetTimesTextFormat( 1, 1, rowIndex, rowIndex);
                //SetTimesTextFormat(1, 1, rowIndex, columnIndex);
                SetSheetName(sheetName);
                SaveCopyAs(theFineName);
                MessageBox.Show("导出完成");
            }
            catch (Exception ex)
            {
                MessageBox.Show("ExportDataGridView2Excel:" + ex.ToString());
            }
            finally
            {
                //if (fmPB != null)
                //{
                //    fmPB.Close();
                //}
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="cExcel"></param>
        /// <param name="beginRowIndex"></param>
        /// <param name="beginColIndex"></param>
        /// <param name="endRowIndex"></param>
        /// <param name="endColIndex"></param>
        public void SetTimesTextFormat(int beginRowIndex, int beginColIndex,
                              int endRowIndex, int endColIndex)
        {
            Excel.Worksheet xSheet = X_Sheet;
            Excel.Range xRange;
          
            try
            {
                xRange = (Excel.Range)xSheet.get_Range(xSheet.Cells[beginRowIndex, beginColIndex],
                                 xSheet.Cells[endRowIndex, endColIndex]);
                xRange.Borders.LineStyle = 1;//边框大小
                xSheet.Columns.AutoFit();//自动调整宽度
            }
            catch (Exception e)
            {
                throw (new Exception("SetTimesTextFormat:" + e.Message));
            }
            finally
            {
                xRange = null;
            }
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="cExcel"></param>
        /// <param name="beginRowIndex"></param>
        /// <param name="beginColIndex"></param>
        /// <param name="endRowIndex"></param>
        /// <param name="endColIndex"></param>
        public void SetTimesTextFormat1(int beginRowIndex, int beginColIndex,
                              int endRowIndex, int endColIndex)
        {
            Excel.Worksheet xSheet = X_Sheet;
            Excel.Range xRange;

            try
            {
                xRange = (Excel.Range)xSheet.get_Range(xSheet.Cells[beginRowIndex, beginColIndex],
                                 xSheet.Cells[endRowIndex, endColIndex]);
                xRange.Borders.LineStyle = 1;//边框大小
                ((Excel.Range)xSheet.Columns["B:B", System.Type.Missing]).ColumnWidth = 10;
                ((Excel.Range)xSheet.Columns["D:D", System.Type.Missing]).ColumnWidth = 11;
            }
            catch (Exception e)
            {
                throw (new Exception("SetTimesTextFormat:" + e.Message));
            }
            finally
            {
                xRange = null;
            }
        }
        /// <summary>
        /// 格式化列标题
        /// </summary>
        public void FormatHeadText(int rowIndex, int colIndex)
        {
            Excel.Worksheet xSheet = X_Sheet;
            Excel.Range xRange;

            try
            {
                xRange = (Excel.Range)xSheet.get_Range(xSheet.Cells[rowIndex, 1],
                                 xSheet.Cells[rowIndex, colIndex]);
                xRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                xRange.Font.Size = 12;
                //
                xRange.Borders.LineStyle = 1; // 显示边框
            }
            catch (Exception e)
            {
                throw (new Exception("FormatHeadText:" + e.Message));
            }
            finally
            {
                xRange = null;
            }
        }
        public void AddSheet()
        {
            xSheet = (Excel.Worksheet)xWorkBook.Sheets.Add(Type.Missing, xSheet, Type.Missing, Type.Missing);
            xSheet.Activate();
        }
        public void SetSheetName(string name)
        {
            xSheet.Name = name;
        }
        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="sExcelFile">文件路径</param>
        /// <param name="iIndex">Worksheet的索引，从1开始</param>
        public ClsExcel(string filePath, int index)
        {
            try
            {
                xApp = new Excel.Application();
                xWorkBook = xApp.Workbooks.Open(filePath, 3, false, 5, "", "", true,
                                                Excel.XlPlatform.xlWindows, "", true, false, null,
                                                false, false, Excel.XlCorruptLoad.xlNormalLoad);
                //xWorkBook = xApp.Workbooks.OpenfilePath);
                xSheet = ((Excel.Worksheet)xWorkBook.Worksheets[index]);
                xApp.DisplayAlerts = false;
            }
            catch (Exception e)
            {
                throw (new Exception("[ClsExcel]:" + e.Message));
            }
        }

        /// <summary>
        /// 返回本excel文件中有几个工作表
        /// </summary>
        /// <returns></returns>
        public int WorkSheetCount()
        {
            return xWorkBook.Worksheets.Count;
        }

        /// <summary>
        /// 给某个单元格符值，索引从1开始
        /// </summary>
        /// <param name="iRow"></param>
        /// <param name="iCol"></param>
        /// <param name="sValue"></param>
        public void SetValue(int row, int col, string value)
        {
            try
            {
                ((Excel.Range)xSheet.Cells[row, col]).Value2 = value;

            }
            catch (Exception e)
            {
                throw (new Exception("[SetValue]:" + e.Message));
            }
        }

        public void SetValue(int row, int col, string value, string time)
        {
            try
            {
                ((Excel.Range)xSheet.Cells[row, col]).Value2 = value;
                ((Excel.Range)xSheet.Cells[row + 2, col]).Value2 = time;
                //((Excel.Range)xSheet.Cells[row, col]).Font.Size = 12;
                Excel.Range range = xSheet.get_Range(xSheet.Cells[1, 1], xSheet.Cells[3, 3]);
                range.Font.Size = 14;

            }
            catch (Exception e)
            {
                throw (new Exception("[SetValue]:" + e.Message));
            }
        }
        /// <summary>
        /// 获得某个单元格去掉空白字符后的现实文本，索引从1开始
        /// </summary>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <returns></returns>
        public string GetTextTrim(int row, int col)
        {
            try
            {
                return Convert.ToString(((Excel.Range)xSheet.Cells[row, col]).Text).Trim();
            }
            catch (Exception e)
            {
                throw (new Exception("[GetTextTrim]:" + e.Message));
            }
        }

        /// <summary>
        /// 设置当前的工作表
        /// </summary>
        /// <param name="iIndex"></param>
        public void SetActiveSheet(int iIndex)
        {
            try
            {
                xSheet = null;
                xSheet = ((Excel.Worksheet)xWorkBook.Worksheets[iIndex]);
            }
            catch (Exception e)
            {

                throw (new Exception("[SetActiveSheet]:" + e.Message));
            }
        }

        /// <summary>
        /// 关闭WorkBook
        /// </summary>
        /// <param name="save"></param>
        public void Close(bool save)
        {
            try
            {
                if (xWorkBook != null)
                {
                    xWorkBook.Close(save, Missing.Value, Missing.Value);
                }

                xSheet = null;
                xWorkBook = null;

                if (xApp != null)
                {
                    xApp.Quit();
                }

                xApp = null;
                GC.Collect();
            }
            catch (Exception e)
            {
                MessageBox.Show("[Close]:" + e.Message);
            }
        }

        /// <summary>
        /// 复制当前Sheet到指定的sheet之后
        /// </summary>
        /// <param name="iIndex">定位的索引</param>
        /// <param name="bAfter">在此之后</param>
        public void CopySheet(int iIndex, bool bAfter)
        {
            if (bAfter)
            {
                xSheet.Copy(Missing.Value, xWorkBook.Worksheets[iIndex]);
            }
            else
            {
                xSheet.Copy(xWorkBook.Worksheets[iIndex], Missing.Value);
            }
        }

        /// <summary>
        /// 删除指定索引的Sheet
        /// </summary>
        /// <param name="sheetIndex"></param>
        public void DeleteSheet(int sheetIndex)
        {
            try
            {
                ((Excel.Worksheet)xWorkBook.Sheets[sheetIndex]).Delete();
            }
            catch (Exception ex)
            {
                throw new Exception("[DeleteSheet]: " + ex.Message);
            }
        }

        /// <summary>
        /// 打印当前工作表
        /// </summary>
        public void PrintOut()
        {
            xSheet.PrintOut(System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                            System.Reflection.Missing.Value, false, System.Reflection.Missing.Value,
                            false, false, System.Reflection.Missing.Value);
            //Excel.Application.DoEvents();
        }

        public void PrintPreview()
        {
            xSheet.Application.Visible = true;
            xSheet.PrintPreview(false);
            //Excel.Application.DoEvents();
        }

        /// <summary>
        /// 打印Excel文档
        /// </summary>
        /// <param name="excel"></param>
        public void PrintOut(ClsExcel excel)
        {
            int iCounter = excel.WorkSheetCount();

            for (int i = 1; i <= iCounter; i++)
            {
                excel.SetActiveSheet(i);
                excel.PrintOut();
            }

            excel.Close(false);
        }

        /// <summary>
        /// 打印Excel文档
        /// </summary>
        /// <param name="excel"></param>
        /// <param name="isLandcape">是否横向打印 true 为是</param>
        public void PrintOut(ClsExcel excel, bool isLandcape)
        {
            int iCounter = excel.WorkSheetCount();

            for (int i = 1; i <= iCounter; i++)
            {
                excel.SetActiveSheet(i);

                if (true == isLandcape)
                {
                    xSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;
                }

                excel.PrintOut();
            }

            excel.Close(false);
        }

        /// <summary>
        /// 判断列是否在列表中
        /// </summary>
        /// <param name="colName"></param>
        /// <param name="colNameList"></param>
        /// <returns></returns>
        public static bool CheckColIsInColList(string colName, List<string> colNameList)
        {
            foreach (string col in colNameList)
            {
                if (col == colName)
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// 将 DataGridView 中的数据导出到Excel
        /// </summary>
        /// <param name="dgv"></param>
        /// <param name="psFileName"></param>
        public static void ExportDataGridView2Excel(ref DataGridView dgv, ref DataGridView dgv2, string fileName,
                                                    string frmBarName, string sheet1Name, string sheet2Name)
        {
            ExportDataGridView2Excel(ref dgv, ref dgv2, 0, 0, fileName, frmBarName, sheet1Name, sheet2Name);
        }

        /// <summary>
        /// 将 DataGridView 中的数据导出到Excel
        /// </summary>
        /// <param name="dgv"></param>
        /// <param name="rowBeginNum"></param>
        /// <param name="colBeginNum"></param>
        /// <param name="fileName"></param>
        public static void ExportDataGridView2Excel(ref DataGridView dgv, ref DataGridView dgv2, int rowBeginNum,
                                            int colBeginNum, string fileName, string frmBarName, string sheet1Name, string sheet2Name)
        {
            if (null == fileName || 0 == fileName.Length)
            {
                return;
            }

            string openPeriodOne = sheet1Name + "以后未返校学生列表";
            string openPeriodTwo = sheet2Name + "以后未返校学生列表";

            int rowNum = dgv.Rows.Count;
            int columnNum = dgv.Columns.Count;
            int rowIndex = 1 + rowBeginNum;
            int columnIndex = colBeginNum;

            int rowNum2 = dgv2.Rows.Count;
            int columnNum2 = dgv2.Columns.Count;
            int rowIndex2 = 1 + rowBeginNum;
            int columnIndex2 = colBeginNum;

            ClsExcel cExcel = null;
            //FrmProgressBar fmPB = new FrmProgressBar();

            try
            {
                //fmPB = ClsPubFuctions.GetProgressBarForm(frmBarName, (rowNum + columnNum + rowNum2 + columnNum2));
                //fmPB.Show();

                //Excel.Application.DoEvents();
                //ClsSystem.mainForm.Cursor = Cursors.WaitCursor;
                cExcel = new ClsExcel();

                // 复制Sheet
                cExcel.CopySheet(1, true);
                cExcel.SetActiveSheet(1);


                #region --openPeriodOne--

                for (int j = 0; j < columnNum; j++)
                {
                    if (!dgv.Columns[j].Visible || dgv.Columns[j].Name == "Check")
                    {
                        continue;
                    }

                    columnIndex++;
                    cExcel.SetValue(rowIndex, columnIndex, dgv.Columns[j].HeaderText);

                    if (0 != dgv.Rows.Count && null != dgv.Rows[0].Cells[j].Value)
                    {
                        cExcel.FormatHeadText(1, columnNum);
                    }


                }

                // 行循环
                for (int i = 0; i < rowNum; i++)
                {
                    rowIndex++;
                    columnIndex = 0;

                    for (int j = 0; j < columnNum; j++) // 列循环
                    {
                        if (!dgv.Columns[j].Visible || dgv.Columns[j].Name == "Check")
                        {
                            continue;
                        }

                        columnIndex++;

                        if (null != dgv.Rows[i].Cells[j].Value)
                        {
                            cExcel.SetValue(rowIndex, columnIndex, dgv.Rows[i].Cells[j].Value.ToString().Replace("\r\n", ""));
                        }
                    }

                    //fmPB.ProgressBar.Value += 1;
                }

                cExcel.X_Sheet.Name = openPeriodOne;
                #endregion

                // -------------------------------------以上为openPeriodOne-----------------------------

                #region --openPeriodTwo--
                cExcel.SetActiveSheet(2);

                for (int j = 0; j < columnNum2; j++)
                {
                    if (!dgv2.Columns[j].Visible || dgv2.Columns[j].Name == "Check")
                    {
                        continue;
                    }

                    columnIndex2++;
                    cExcel.SetValue(rowIndex2, columnIndex2, dgv2.Columns[j].HeaderText);

                    if (0 != dgv2.Rows.Count && null != dgv2.Rows[0].Cells[j].Value)
                    {
                        cExcel.FormatHeadText(1, columnNum2);
                    }

                    //fmPB.ProgressBar.Value += 1;
                }

                // 行循环
                for (int i = 0; i < rowNum2; i++)
                {
                    rowIndex2++;
                    columnIndex2 = 0;

                    for (int j = 0; j < columnNum2; j++) // 列循环
                    {
                        if (!dgv2.Columns[j].Visible || dgv2.Columns[j].Name == "Check")
                        {
                            continue;
                        }

                        columnIndex2++;

                        if (null != dgv2.Rows[i].Cells[j].Value)
                        {
                            cExcel.SetValue(rowIndex2, columnIndex2, dgv2.Rows[i].Cells[j].Value.ToString().Replace("\r\n", ""));
                        }
                    }

                    //fmPB.ProgressBar.Value += 1;
                }

                cExcel.X_Sheet.Name = openPeriodTwo;

                #endregion

                cExcel.xWorkBook.SaveCopyAs(fileName);
                //ClsSystem.mainForm.Cursor = Cursors.Default;
            }
            catch (Exception ex)
            {
                MessageBox.Show("ExportDataGridView2Excel:" + ex.ToString());
            }
            finally
            {
                //if (fmPB != null)
                //{
                //    fmPB.Close();
                //}

                //ClsSystem.mainForm.Cursor = Cursors.Default;

                if (cExcel != null)
                {
                    cExcel.Close(false);
                }
            }
        }


        /// <summary>
        /// 格式化列标题
        /// </summary>
        /// <param name="rowIndex"></param>
        /// <param name="dt"></param>
        //public void FormatHeadText(int rowIndex, int colIndex)
        //{
        //    Excel.Range xRange;

        //    try
        //    {
        //        xRange = (Excel.Range)xSheet.get_Range(xSheet.Cells[rowIndex, 1],
        //                         xSheet.Cells[rowIndex, colIndex]);
        //        xSheet.Columns.AutoFit();
        //        xRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter; // 居中对齐
        //        xRange.Font.Bold = true;
        //        xRange.WrapText = true;

        //    }
        //    catch (Exception e)
        //    {
        //        throw (new Exception("SetTextFormat:" + e.Message));
        //    }
        //    finally
        //    {
        //        xRange = null;
        //    }
        //}

        /// <summary>
        /// 合并单元格
        /// </summary>
        /// <param name="cExcel"></param>
        /// <param name="beginRowIndex">起始行索引</param>
        /// <param name="beginColIndex">起始列索引</param>
        /// <param name="endRowIndex">结束行索引</param>
        /// <param name="endColIndex">结束列索引</param>
        public void FormatMergeCell(ClsExcel cExcel, int beginRowIndex, int beginColIndex,
                                            int endRowIndex, int endColIndex)
        {
            Excel.Worksheet xSheet = cExcel.X_Sheet;
            Excel.Range xRange;

            try
            {
                xRange = (Excel.Range)xSheet.get_Range(xSheet.Cells[beginRowIndex, beginColIndex],
                                            xSheet.Cells[endRowIndex, endColIndex]);
                xRange.MergeCells = true;
            }
            catch (Exception ex)
            {
                throw new Exception("[FormatMergeCell]:" + ex.Message);
            }
            finally
            {
                xRange = null;
            }
        }

        /// <summary>
        /// 格式化空白行
        /// </summary>
        /// <param name="cExcel"></param>
        /// <param name="rowIndex">行索引</param>
        /// <param name="rowHeight">行高</param>
        /// <param name="border">边框样式</param>
        public void FormatBlank(ClsExcel cExcel, int rowIndex, double rowHeight, int border)
        {
            Excel.Worksheet xSheet = cExcel.X_Sheet;
            Excel.Range xRange;

            try
            {
                xRange = (Excel.Range)xSheet.get_Range(xSheet.Cells[rowIndex, 1],
                                 xSheet.Cells[rowIndex, 1]);
                xRange.RowHeight = rowHeight;
                xRange.Borders.LineStyle = border;
            }
            catch (Exception ex)
            {
                throw (new Exception("[FormatBlank]:" + ex.Message));
            }
            finally
            {
                xRange = null;
            }
        }

        public void SaveCopyAs(string fileName)
        {
            xWorkBook.SaveCopyAs(fileName);
        }
        /// <summary>
        /// 快速将DataTable导出到EXCEL，可能并不稳定，容易产生异常，特别是当数据量特大时，
        /// </summary>
        /// <param name="dt">所要导出的DataTable</param>
        /// <param name="beginColumnIndex">开始导出的列</param>
        /// <param name="replaceTabColumnIndex">要去除空格的列，可以为空，表示没有列需要替换</param>
        public static string ConvertToString(DataTable dt, int beginColumnIndex, string replaceTabColumnName)
        {
            try
            {
                StringBuilder sb = new StringBuilder();
                for (int k = beginColumnIndex; k < dt.Columns.Count; k++)
                {
                    sb.Append(dt.Columns[k].ColumnName.ToString().Trim() + "\t");
                }
                sb.Append(Environment.NewLine);

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    System.Windows.Forms.Application.DoEvents();

                    for (int j = beginColumnIndex; j < dt.Columns.Count; j++)
                    {
                        string str = dt.Rows[i][j].ToString().Trim();
                        if (dt.Columns[j].ColumnName.ToString().Trim() == replaceTabColumnName)
                        {
                            str = Regex.Replace(dt.Rows[i][j].ToString().Trim(), @"\s", "");
                        }
                        sb.Append(str + "\t");
                    }
                    sb.Append(Environment.NewLine);
                }

                return sb.ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            return string.Empty;
        }

        public static string ConvertToString(byte[] bt, int toBase)
        {
            try
            {
                StringBuilder sb = new StringBuilder();
                for (int k = 0; k < bt.Length; k++)
                {
                    if (k % 10 == 0)
                    {
                        sb.Append(Convert.ToString(bt[k], toBase) + "\r");
                    }
                    else
                    {
                        sb.Append(Convert.ToString(bt[k], toBase) + "\t");
                    }
                }

                return sb.ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            return string.Empty;
        }

        public void PasteToExcel(string data)
        {
            System.Windows.Forms.Clipboard.Clear();
            System.Windows.Forms.Clipboard.SetDataObject(data);
            xSheet.Activate();
            Excel.Range xRange = (Excel.Range)xSheet.get_Range(xSheet.Cells[1, 1], xSheet.Cells[1, 1]);
            int wait = 50 + data.Length / 100;
            //MessageBox.Show(wait.ToString());
            Application.DoEvents();
            Thread.Sleep(wait);
            xSheet.Paste(xRange, false);
            System.Windows.Forms.Clipboard.Clear();
        }
        public void PageSetup(int H)
        {
            xSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;
            xSheet.PageSetup.Zoom = false;
            xSheet.PageSetup.FitToPagesWide = 1;
            xSheet.PageSetup.FitToPagesTall = H;

        }
        public static void WriteTXTFile(string data, string fileName)
        {
            try
            {
                StreamWriter sw = new StreamWriter(fileName, false, Encoding.GetEncoding("gb2312"));
                sw.Write(data);
                sw.Flush();
                sw.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        public static void ExportToExcel(System.Data.DataTable table, string saveFileName)
        {
            try
            {
                if (table == null)
                {
                    return;
                }
                //bool fileSaved = false;

                //ExcelApp xlApp = new ExcelApp();

                Excel.Application xlApp = new Excel.Application();

                if (xlApp == null)
                {
                    return;
                }

                saveFileName = GetSaveFileName(saveFileName);

                Excel.Workbooks workbooks = (Excel.Workbooks)xlApp.Workbooks;
                Excel.Workbook workbook = workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
                Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Worksheets[1];//取得sheet1

                long rows = table.Rows.Count;

                /*下边注释的两行代码当数据行数超过行时，出现异常：异常来自HRESULT:0x800A03EC。因为：Excel 2003每个sheet只支持最大行数据

                //Range fchR = worksheet.get_Range(worksheet.Cells[1, 1], worksheet.Cells[table.Rows.Count+2, gridview.Columns.View.VisibleColumns.Count+1]);

                //fchR.Value2 = datas;*/

                if (rows > 65535)
                {

                    long pageRows = 60000;//定义每页显示的行数,行数必须小于

                    int scount = (int)(rows / pageRows);

                    if (scount * pageRows < table.Rows.Count)//当总行数不被pageRows整除时，经过四舍五入可能页数不准
                    {
                        scount = scount + 1;
                    }

                    for (int sc = 1; sc <= scount; sc++)
                    {
                        if (sc > 1)
                        {

                            object missing = System.Reflection.Missing.Value;

                            worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets.Add(

                           missing, missing, missing, missing);//添加一个sheet

                        }

                        else
                        {
                            worksheet = (Excel.Worksheet)workbook.Worksheets[sc];//取得sheet1
                        }

                        string[,] datas = new string[pageRows + 1, table.Columns.Count + 1];


                        for (int i = 0; i < table.Columns.Count; i++) //写入字段
                        {
                            datas[0, i] = table.Columns[i].Caption;
                        }

                        Excel.Range range = worksheet.get_Range(worksheet.Cells[1, 1], worksheet.Cells[1, table.Columns.Count]);
                        range.Interior.ColorIndex = 15;//15代表灰色
                        range.Font.Bold = true;
                        range.Font.Size = 9;

                        int init = int.Parse(((sc - 1) * pageRows).ToString());
                        int r = 0;
                        int index = 0;
                        int result;

                        if (pageRows * sc >= table.Rows.Count)
                        {
                            result = table.Rows.Count;
                        }
                        else
                        {
                            result = int.Parse((pageRows * sc).ToString());
                        }
                        for (r = init; r < result; r++)
                        {
                            index = index + 1;
                            for (int i = 0; i < table.Columns.Count; i++)
                            {
                                if (table.Columns[i].DataType == typeof(DateTime))
                                {
                                    object obj = table.Rows[r][table.Columns[i].ColumnName];
                                    datas[index, i] = obj == null ? "" : "'" + obj.ToString().Trim();//在obj.ToString()前加单引号是为了防止自动转化格式

                                }
                                else
                                {
                                    object obj = table.Rows[r][table.Columns[i].ColumnName];
                                    datas[index, i] = obj == null ? "" : obj.ToString().Trim();//在obj.ToString()前加单引号是为了防止自动转化格式

                                }

                            }
                        }

                        Excel.Range fchR = worksheet.get_Range(worksheet.Cells[1, 1], worksheet.Cells[index + 2, table.Columns.Count + 1]);

                        fchR.Value2 = datas;
                        worksheet.Columns.EntireColumn.AutoFit();//列宽自适应。

                        range = worksheet.get_Range(worksheet.Cells[1, 1], worksheet.Cells[index + 1, table.Columns.Count]);

                        //15代表灰色

                        range.Font.Size = 9;
                        range.RowHeight = 14.25;
                        range.Borders.LineStyle = 1;
                        range.HorizontalAlignment = 1;

                    }

                }

                else
                {

                    string[,] datas = new string[table.Rows.Count + 2, table.Columns.Count + 1];
                    for (int i = 0; i < table.Columns.Count; i++) //写入字段         
                    {
                        datas[0, i] = table.Columns[i].Caption;
                    }

                    Excel.Range range = worksheet.get_Range(worksheet.Cells[1, 1], worksheet.Cells[1, table.Columns.Count]);
                    range.Interior.ColorIndex = 15;//15代表灰色
                    range.Font.Bold = true;
                    range.Font.Size = 9;

                    int r = 0;
                    for (r = 0; r < table.Rows.Count; r++)
                    {
                        for (int i = 0; i < table.Columns.Count; i++)
                        {

                            if (table.Columns[i].DataType == typeof(DateTime))
                            {
                                object obj = table.Rows[r][table.Columns[i].ColumnName];
                                datas[r + 1, i] = obj == null ? "" : "'" + obj.ToString().Trim();//在obj.ToString()前加单引号是为了防止自动转化格式

                            }
                            else
                            {
                                object obj = table.Rows[r][table.Columns[i].ColumnName];
                                datas[r + 1, i] = obj == null ? "" : obj.ToString().Trim();//在obj.ToString()前加单引号是为了防止自动转化格式

                            }
                        }

                        //System.Windows.Forms.Application.DoEvents();


                    }

                    Excel.Range fchR = worksheet.get_Range(worksheet.Cells[1, 1], worksheet.Cells[table.Rows.Count + 2, table.Columns.Count + 1]);

                    fchR.Value2 = datas;

                    worksheet.Columns.EntireColumn.AutoFit();//列宽自适应。


                    range = worksheet.get_Range(worksheet.Cells[1, 1], worksheet.Cells[table.Rows.Count + 1, table.Columns.Count]);

                    //15代表灰色

                    range.Font.Size = 9;
                    range.RowHeight = 14.25;
                    range.Borders.LineStyle = 1;
                    range.HorizontalAlignment = 1;
                }

                if (saveFileName != "")
                {
                    try
                    {
                        workbook.Saved = true;
                        workbook.SaveCopyAs(saveFileName);
                        // fileSaved = true;


                    }

                    catch
                    {
                        //  fileSaved = false;
                    }

                }

                else
                {

                    //fileSaved = false;

                }

                xlApp.Quit();
                MessageBox.Show("导出完成");
                GC.Collect();//强行销毁   
                //web后台谈框框，winform可以使用messagebox.. 大家都懂的
                // System.Web.HttpContext.Current.Response.Write("<Script Language=JavaScript>...alert('Export Success! File path in D disk root directory.');</Script>");
            }
            catch (Exception)
            {
                // System.Web.HttpContext.Current.Response.Write("<Script Language=JavaScript>...alert('Export Error! please contact administrator!');</Script>");
            }
        }


    } //ClsExcel
}
