using System;
using System.IO;
using NPOI.HSSF.UserModel; 
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System.Text;
using System.Collections.Generic;
using System.Data;

namespace NetCommonTools
{

    public class Demo
    {

        /// <summary>
        /// 导出示例，需要预先准备好EXCEL 模版
        /// 程序会使用该模版进行修改，并另存为新的Excel
        /// </summary>
        /// <param name="templatePath"></param>
        /// <param name="dt"></param>
        /// <param name="exportFilePath"></param>
        public static void ExportByTemplate(string templatePath, DataTable dt, string exportFilePath)
        {

            using (NPOIExcelHelper excel = new NPOIExcelHelper(templatePath))
            {
                excel.CurrentSheetName = "Sheet1";
                int i = 1;
                foreach (DataRow item in dt.Rows)
                {
                    int j = 0;
                    excel.SetValue(i, j++, item[0]);
                    excel.SetValue(i, j++, item[1]);
                    excel.SetValue(i, j++, item[2]);
                    excel.SetValue(i, j++, item[5]);

                }
                excel.SaveAs(exportFilePath);
            }
        }

        public static void TestCopyArea(string TemplatePath, string newFilePath)
        {

            using (NPOIExcelHelper source = new NPOIExcelHelper(TemplatePath))
            {

                source.CurrentSheetName = "请款单";
                if (source.CurrentSheet != null)
                {
                    for (int j = 1; j <= 50; j++)
                    {
                        int startRow = 15 * j;

                        var sheet = source.CurrentSheet;
                        for (int i = 0; i < 13; i++)
                        {
                            IRow row = sheet.GetRow(i);
                            IRow newRow = row.CopyRowTo(i + startRow);
                            newRow.Height = row.Height;
                        }

                        CellRangeAddress cra1 = new CellRangeAddress(0 + startRow, 0 + startRow, 0, 13);
                        source.CurrentSheet.AddMergedRegion(cra1);
                        CellRangeAddress cra2 = new CellRangeAddress(1 + startRow, 1 + startRow, 0, 13);
                        source.CurrentSheet.AddMergedRegion(cra2);
                        CellRangeAddress cra3 = new CellRangeAddress(2 + startRow, 2 + startRow, 0, 13);
                        source.CurrentSheet.AddMergedRegion(cra3);
                        CellRangeAddress cra4 = new CellRangeAddress(3 + startRow, 4 + startRow, 0, 1);
                        source.CurrentSheet.AddMergedRegion(cra4);
                        //金               额（人民币）										
                        CellRangeAddress cra5 = new CellRangeAddress(3 + startRow, 3 + startRow, 3, 13);
                        source.CurrentSheet.AddMergedRegion(cra5);
                        CellRangeAddress cra6 = new CellRangeAddress(5 + startRow, 8 + startRow, 2, 2);
                        source.CurrentSheet.AddMergedRegion(cra6);
                        CellRangeAddress cra7 = new CellRangeAddress(9 + startRow, 9 + startRow, 0, 1);
                        source.CurrentSheet.AddMergedRegion(cra7);

                        //收款单位名称：北京京东世纪信息技术有限公司	
                        CellRangeAddress cra8 = new CellRangeAddress(10 + startRow, 10 + startRow, 0, 1);
                        source.CurrentSheet.AddMergedRegion(cra8);
                        //"开户行：招商银行北京青年路支行   账号：xxxxxxxx"											
                        CellRangeAddress cra9 = new CellRangeAddress(10 + startRow, 10 + startRow, 2, 13);
                        source.CurrentSheet.AddMergedRegion(cra9);
                        //"请于      年      月      日前完成付款。 备注说明：（若无法在申请付款时间内付款，请提前告知原因及可付款日期。）"		 
                        CellRangeAddress cra10 = new CellRangeAddress(11 + startRow, 11 + startRow, 0, 13);
                        source.CurrentSheet.AddMergedRegion(cra10);
                        //公司领导：                         财务经理：                    会计：                    部门负责人：                   经办人：													
                        CellRangeAddress cra11 = new CellRangeAddress(12 + startRow, 12 + startRow, 0, 13);
                        source.CurrentSheet.AddMergedRegion(cra11);
                        source.SaveAs(newFilePath);
                    } 
                }
            }
        }

    }
    public class NPOIExcelHelper : IDisposable
    {


        /// <summary>
        /// 区域拷贝定义
        /// </summary>
        public class CopyRegionSettings
        {
            public ISheet FromSheet { get; set; }
            public int StartCell { get; set; }
            public int StartRow { get; set; }
            public int EndRow { get; set; }

            public int EndCell { get; set; }
            public ISheet targetSheet { get; set; }
            /// <summary>
            /// 目标Sheet 起始行
            /// </summary>
            public int ToRowIndex { get; set; }

            public int ToCellIndex { get; set; }
            public bool CopyAll { get; set; }
            public bool CopyData { get; set; }
            public bool CopyStyle { get; set; }
            public bool CopyFormula { get; set; }
            /// <summary>
            /// 需要避开写入的行
            /// </summary>
            public List<int> SkipRows { get; set; }
            public List<int> SkipCells { get; set; }

        }
        /// <summary>
        /// 文件ID SalesportalUploadlogModel.FileID
        /// </summary>
        public Guid FileID { get; set; }
        private readonly HSSFWorkbook workbook;
        private ISheet sheet;
        private readonly Stream stream;
        public NPOIExcelHelper(Stream excelStream)
        {
            stream = excelStream;

            try
            {
                workbook = new HSSFWorkbook(stream);
                workbook.SetActiveSheet(0);
                CurrentSheet = workbook.GetSheetAt(0);

            }
            catch (Exception ex)
            {
                if (stream != null)
                {
                    stream.Close();
                }
                throw new Exception("NPOI实例化错误:" + ex);
            }

        }
        public NPOIExcelHelper(string templatePath)
        {
            try
            {
                stream = new FileStream(templatePath, FileMode.Open, FileAccess.Read);
                workbook = new HSSFWorkbook(stream);
                workbook.SetActiveSheet(0);

                CurrentSheet = workbook.GetSheetAt(0);
            }
            catch (Exception ex)
            {
                throw new Exception("NPOI实例化错误:" + ex.ToString());
            }
        }

        /// <summary>
        /// 设置/获取当前Sheet
        /// </summary>
        public string CurrentSheetName
        {
            get { return sheet.SheetName; }
            set
            {
                sheet = workbook.GetSheet(value);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="startColumnIndex">开始列</param>
        /// <param name="IndexStartWith1">默认为0  </param>
        /// <param name="startRowIndex">开始行</param>
        /// <param name="endRowIndex"></param>
        /// <param name="endColumnIndex">结束列</param>
        /// <returns></returns>
        public Array ConverToArray(int startRowIndex = 0, int endRowIndex = 0, int startColumnIndex = 0, int endColumnIndex = 0)
        {

            endColumnIndex = (endColumnIndex == 0 ? ColumnCount : endColumnIndex);
            endRowIndex = (endRowIndex == 0 ? RowCount : endRowIndex);

            Array array = new object[(endRowIndex - startRowIndex + 1), (endColumnIndex - startColumnIndex + 1)];

            for (int i = startRowIndex; i < endRowIndex + 1; i++)
            {
                for (int j = startColumnIndex; j < endColumnIndex + 1; j++)
                {
                    array.SetValue(GetValue(i, j).Trim(), i - startRowIndex, j - startColumnIndex);
                }
            }
            return array;

        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="pkIndex">指定主键列，如果为空则不往下读取-1为忽略</param>
        /// <param name="startColumnIndex">开始列</param>
        /// <param name="IndexStartWith1">默认为0  </param>
        /// <param name="startRowIndex">开始行</param>
        /// <param name="endRowIndex"></param>
        /// <param name="endColumnIndex">结束列</param>
        /// <returns></returns>
        public DataTable ConverToDataTable(DataTable dt, int pkIndex, int startRowIndex = 0, int endRowIndex = 0, int startColumnIndex = 0, int endColumnIndex = 0)
        {

            endColumnIndex = (endColumnIndex == 0 ? ColumnCount : endColumnIndex);
            endRowIndex = (endRowIndex == 0 ? RowCount : endRowIndex);

            for (int i = startRowIndex; i < endRowIndex + 1; i++)
            {
                if (pkIndex != -1 && string.IsNullOrEmpty(GetValue(i, pkIndex).Trim()))
                {
                    break;
                }
                DataRow row = dt.NewRow();
                for (int j = startColumnIndex; j < endColumnIndex + 1; j++)
                {
                    row[j - startColumnIndex] = GetValue(i, j).Trim();
                }
                dt.Rows.Add(row);
            }
            return dt;
        }

        /// <summary>
        /// 设置/获取当前Sheet
        /// </summary>
        public ISheet CurrentSheet
        {
            get { return sheet; }
            set { sheet = value; }
        }
        public IRow GetRow(int rowIndex)
        {

            return sheet.GetRow(rowIndex);
        }

        public int RowCount
        {
            get
            {    //重新定义最后行下标 
                return CleanEmptyRow();
            }
        }

        /// <summary>
        /// 从下往上,从左到右, 找到一个单元格不为空,则取此行下标为当前sheet最大下标
        /// </summary>
        private int CleanEmptyRow()
        {
            for (int i = sheet.LastRowNum; i > 0; i--)
            {
                for (int j = 0; j < sheet.GetRow(i).LastCellNum; j++)
                {
                    if (!string.IsNullOrEmpty(this.GetValue(i, j)))
                    {
                        return i;
                    }
                }
            }
            return 0;
        }
        /// <summary> 
        /// 取第一行的最后一列的下标,不是最大列数
        /// </summary>
        public int ColumnCount
        {
            get { return sheet.GetRow(0).LastCellNum; }
        }

        public ICell GetCell(IRow row, int cellIndex)
        {
            return row.GetCell(cellIndex) ?? row.CreateCell(cellIndex);

        }
        public ICell GetCell(int rowIndex, int cellIndex)
        {
            IRow row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            ICell cell = row.GetCell(cellIndex) ?? row.CreateCell(cellIndex);
            return cell;
        }
        public decimal? GetNullableDecimal(int rowIndex, int cellIndex)
        {
            var cell = GetCell(rowIndex, cellIndex);
            if (!string.IsNullOrEmpty(cell.ToString()))
            {
                return Convert.ToDecimal(cell.NumericCellValue);
            }
            else
            {
                return null;
            }
        }
        public string GetValue(int rowIndex, int cellIndex)
        {

            var cell = GetCell(rowIndex, cellIndex);
            return cell.ToString();
        }
        public DateTime GetDate(int rowIndex, int cellIndex)
        {

            var cell = GetCell(rowIndex, cellIndex);
            return cell.DateCellValue;
        }
        public string ColumnIndexToName(int colNum)
        {
            colNum += 1;
            StringBuilder sb = new StringBuilder();
            int cycleNum = (colNum - 1) / 26;
            int withinNum = colNum - (cycleNum * 26);
            if (cycleNum > 0)
                sb.Append((char)(cycleNum - 1 + 'A'));
            sb.Append((char)(withinNum - 1 + 'A'));
            return sb.ToString();
        }

        public int ColumnNameToIndex(string colName)
        {
            int result = 0;
            string lcColName = colName.ToUpper();
            for (int ctr = 0; ctr < lcColName.Length; ++ctr)
                result = (result * 26) + (lcColName[ctr] - 'A');
            return result;
        }

        public void SetValue(int rowIndex, int cellIndex, object value)
        {
            IRow row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            ICell cell = row.GetCell(cellIndex) ?? row.CreateCell(cellIndex);
            string val = null == value ? string.Empty : value.ToString();
            if (value != null)
                switch (value.GetType().FullName)
                {
                    case "System.String": //字符串类型
                        cell.SetCellValue(val);
                        break;
                    case "System.DateTime": //日期类型
                        DateTime dateV;
                        DateTime.TryParse(val, out dateV);
                        cell.SetCellValue(dateV);
                        break;
                    case "System.Boolean": //布尔型
                        bool boolV;
                        bool.TryParse(val, out boolV);
                        cell.SetCellValue(boolV);
                        break;
                    case "System.Int16": //整型
                    case "System.Int32":
                    case "System.Int64":
                    case "System.Byte":
                        int intV;
                        int.TryParse(val, out intV);
                        cell.SetCellValue(intV);
                        break;
                    case "System.Decimal": //浮点型
                    case "System.Double":
                        double doubV;
                        double.TryParse(val, out doubV);
                        cell.SetCellValue(doubV);
                        break;
                    case "System.DBNull": //空值处理
                        cell.SetCellValue("");
                        break;
                    default:
                        cell.SetCellValue("");
                        break;
                }
        }

        /// <summary>
        /// 设置有效性
        /// </summary>
        /// <param name="colName">列名</param>
        /// <param name="startRowIndex">开始行索引 0起</param>
        /// <param name="validSheetName">有效性 目标sheet</param>
        /// <param name="validateColName">有效性列名</param>
        /// <param name="validateRowStart">有效性开始行 0起</param>
        /// <param name="validateRowEnd">有效性结束行 0起</param>
        public void SetValidation(string colName, int startRowIndex,
            string validSheetName, string validateColName, int validateRowStart, int validateRowEnd)
        {
            var columnIndex = ColumnNameToIndex(colName);
            //设置数据有效性作用域
            var regions = new CellRangeAddressList(startRowIndex, 65535, columnIndex, columnIndex);
            //设置名称管理器管理数据源范围
            var range = workbook.CreateName();
            //                          验证页,              验证列名                  验证开始行
            range.RefersToFormula = validSheetName + "!$" + validateColName + "$" + (validateRowStart + 1) +
                //验证结束列             //验证结束行
                ":$" + validateColName + "$" + (validateRowStart + validateRowEnd);
            range.NameName = "dicRange" + columnIndex;
            //根据名称生成下拉框内容
            DVConstraint constraint = DVConstraint.CreateFormulaListConstraint("dicRange" + columnIndex);
            //绑定下拉框和作用区域
            var dataValidate = new HSSFDataValidation(regions, constraint);
            sheet.AddValidationData(dataValidate);
        }


        /// <summary>
        /// 区域复制函数
        /// </summary>
        /// <param name="region"></param> 
        public static void CopyArea(CopyRegionSettings region)
        {

            int toRowIndex = region.ToRowIndex;

            for (int fromRowIndex = region.StartRow; fromRowIndex <= region.EndRow; fromRowIndex++, toRowIndex++)
            {
                if (region.SkipRows != null && region.SkipRows.Contains(fromRowIndex))
                {
                    continue;
                }
                int toColIndex = region.ToCellIndex;
                for (int fromColIndex = region.StartCell; fromColIndex <= region.EndCell; fromColIndex++, toColIndex++)
                {
                    if (region.SkipCells != null && region.SkipCells.Contains(fromColIndex))
                    {
                        continue;
                    }
                    IRow sourceRow = region.FromSheet.GetRow(fromRowIndex);
                    ICell source = sourceRow.GetCell(fromColIndex);
                    if (sourceRow != null && source != null)
                    {
                        IRow changingRow = null;
                        ICell target = null;
                        changingRow = region.targetSheet.GetRow(toRowIndex);
                        if (changingRow == null)
                            changingRow = region.targetSheet.CreateRow(toRowIndex);
                        target = changingRow.GetCell(toColIndex);
                        if (target == null)
                            target = changingRow.CreateCell(toColIndex);
                        if (region.CopyData)//仅数据
                        {
                            //对单元格的值赋值
                            switch (source.CellType)
                            {
                                case CellType.Unknown:
                                    break;
                                case CellType.Numeric:
                                    target.SetCellValue(source.NumericCellValue);
                                    break;
                                case CellType.String:
                                    target.SetCellValue(source.StringCellValue);
                                    break;
                                case CellType.Formula:
                                    target.SetCellFormula(source.CellFormula);
                                    //target = e.EvaluateInCell(target);

                                    break;
                                case CellType.Blank:
                                    break;
                                case CellType.Boolean:
                                    target.SetCellValue(source.BooleanCellValue);
                                    break;
                                case CellType.Error:
                                    break;
                                default:
                                    target.SetCellValue(source.ToString());
                                    break;
                            }
                        }

                        if (region.CopyStyle)
                        {
                            //单元格的格式
                            target.CellStyle = source.CellStyle;
                        }
                    }
                }
            }
        }
        public void SaveAs(Stream stream)
        {
            //将内容写入数据流,以供下载 
            workbook.ForceFormulaRecalculation = true;
            workbook.Write(stream);
        }

        public void SaveAs(string filePath)
        {
            //将内容写入临时文件,以供下载
            using (var fsTemp = new FileStream(filePath, FileMode.OpenOrCreate))
            {
                workbook.ForceFormulaRecalculation = true;
                workbook.Write(fsTemp);
                fsTemp.Close();
                fsTemp.Dispose();
            }
        }
        public string Save()
        {
            workbook.ForceFormulaRecalculation = true;
            string filepath = Path.GetTempFileName();
            SaveAs(filepath);
            return filepath;
        }

        public void Dispose()
        {
            stream.Close();
            GC.Collect();
        }
    }
}
