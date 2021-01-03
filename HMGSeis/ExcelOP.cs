using System;

using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using Microsoft.Office.Interop.Excel;

namespace HMGSeis
{
    /// <summary>
    /// 
    /// 
    /// </summary>
    public class ExcelBE
    {
        private int _row = 0;
        private int _col = 0;
        private string _text = string.Empty;
        private string _startCell = string.Empty;
        private string _endCell = string.Empty;
        private string _interiorColor = string.Empty;
        private bool _isMerge = false;
        private int _size = 0;
        private string _fontColor = string.Empty;
        private string _format = string.Empty;
        /// <summary>
        /// 
        /// </summary>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <param name="text"></param>
        /// <param name="startCell"></param>
        /// <param name="endCell"></param>
        /// <param name="interiorColor"></param>
        /// <param name="isMerge"></param>
        /// <param name="size"></param>
        /// <param name="fontColor"></param>
        /// <param name="format"></param>
        public ExcelBE(int row, int col, string text, string startCell, string endCell, string interiorColor, bool isMerge, int size, string fontColor, string format)
        {
            _row = row;
            _col = col;
            _text = text;
            _startCell = startCell;
            _endCell = endCell;
            _interiorColor = interiorColor;
            _isMerge = isMerge;
            _size = size;
            _fontColor = fontColor;
            _format = format;
        }

        public ExcelBE()
        { }

        public int Row
        {
            get { return _row; }
            set { _row = value; }
        }

        public int Col
        {
            get { return _col; }
            set { _col = value; }
        }

        public string Text
        {
            get { return _text; }
            set { _text = value; }
        }

        public string StartCell
        {
            get { return _startCell; }
            set { _startCell = value; }
        }

        public string EndCell
        {
            get { return _endCell; }
            set { _endCell = value; }
        }

        public string InteriorColor
        {
            get { return _interiorColor; }
            set { _interiorColor = value; }
        }

        public bool IsMerge
        {
            get { return _isMerge; }
            set { _isMerge = value; }
        }

        public int Size
        {
            get { return _size; }
            set { _size = value; }
        }

        public string FontColor
        {
            get { return _fontColor; }
            set { _fontColor = value; }
        }

        public string Formart
        {
            get { return _format; }
            set { _format = value; }
        }

    }
    /// <summary>
    /// BASE FUNCTION OF EXCEL
    /// </summary>
    public class ExcelBase:IDisposable
    {

        private Microsoft.Office.Interop.Excel.Application app = null;
        private Microsoft.Office.Interop.Excel.Workbook workbook = null;
        private Microsoft.Office.Interop.Excel.Worksheet worksheet = null;
        private Microsoft.Office.Interop.Excel.Range workSheet_range = null;
        private string excelFilePath = null;
        private string excelFileName = null;
        private string excelFullFileName = null;

        public string ExcelFilePath { get => excelFilePath; set => excelFilePath = value; }
        public string ExcelFileName { get => excelFileName; set => excelFileName = value; }
        public string ExcelFullFileName { get => excelFullFileName; set => excelFullFileName = value; }
        public Range WorkSheet_range { get => workSheet_range; set => workSheet_range = value; }
        public Worksheet Worksheet { get => worksheet; set => worksheet = value; }
        public Workbook Workbook { get => workbook; set => workbook = value; }
        public Application App { get => app; set => app = value; }

        public ExcelBase()
        {
            createDoc();
        }
        public ExcelBase(string _ModelDirectory, string _ExcelFileName)
        {
            OpenDOC(_ModelDirectory, _ExcelFileName);
        }

        public void createDoc()
        {
            try
            {
                App = new Microsoft.Office.Interop.Excel.Application();
                App.Visible = true;
                Workbook = App.Workbooks.Add(1);
                Worksheet = (Microsoft.Office.Interop.Excel.Worksheet)Workbook.Sheets[1];
            }
            catch (Exception e)
            {
                Console.Write("Error");
            }
            finally
            {
            }
        }
        public void OpenDOC(string _ModelDirectory, string _ExcelFileName)
        {

            try
            {
                App = new Microsoft.Office.Interop.Excel.Application();
                App.Visible = true;
                excelFilePath = _ModelDirectory;
                excelFileName = _ExcelFileName;
                excelFullFileName = _ModelDirectory + System.IO.Path.DirectorySeparatorChar + _ExcelFileName;
                Workbook = App.Workbooks.Open(excelFullFileName, Type.Missing);
                Worksheet = (Microsoft.Office.Interop.Excel.Worksheet)Workbook.Sheets[1];

            }
            catch (Exception e)
            {
                Console.Write("Error");
            }
            finally
            {
            }



        }

        public void InsertData(ExcelBE be)
        {
            Worksheet.Cells[be.Row, be.Col] = be.Text;
            WorkSheet_range = Worksheet.get_Range(be.StartCell, be.EndCell);
            WorkSheet_range.MergeCells = be.IsMerge;
            WorkSheet_range.Interior.Color = GetColorValue(be.InteriorColor);
            WorkSheet_range.Borders.Color = System.Drawing.Color.Black.ToArgb();
            WorkSheet_range.ColumnWidth = be.Size;
            WorkSheet_range.Font.Color = string.IsNullOrEmpty(be.FontColor) ? System.Drawing.Color.White.ToArgb() : System.Drawing.Color.Black.ToArgb();
            WorkSheet_range.NumberFormat = be.Formart;
        }
        public string  GetSheet_Cell_Text(ExcelBE be,int _row,int _column)
        {
            be.Row = _row; be.Col = _column;
            
            return Worksheet.Cells[be.Row, be.Col].Text;
        }
        public double GetSheet_Cell_Double(ExcelBE be, int _row, int _column)
        {
            be.Row = _row; be.Col = _column;
            
            return Worksheet.Cells[be.Row, be.Col].Value2;
        }

        private int GetColorValue(string interiorColor)
        {
            switch (interiorColor)
            {
                case "YELLOW":
                    return System.Drawing.Color.Yellow.ToArgb();
                case "GRAY":
                    return System.Drawing.Color.Gray.ToArgb();
                case "GAINSBORO":
                    return System.Drawing.Color.Gainsboro.ToArgb();
                case "Turquoise":
                    return System.Drawing.Color.Turquoise.ToArgb();
                case "PeachPuff":
                    return System.Drawing.Color.PeachPuff.ToArgb();

                default:
                    return System.Drawing.Color.White.ToArgb();
            }
        }


        #region IDisposable Support
        private bool disposedValue = false; // 要检测冗余调用

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    workbook.Close(Type.Missing,Type.Missing,Type.Missing);
                    app.Quit(); // TODO: 释放托管状态(托管对象)。
                    app = null;
                }

                // TODO: 释放未托管的资源(未托管的对象)并在以下内容中替代终结器。
                // TODO: 将大型字段设置为 null。
                
                
                workbook = null;
                worksheet = null;
                workSheet_range = null;
                excelFilePath = null;
                excelFileName = null;
                excelFullFileName = null;
                disposedValue = true;               
            }
        }

        // TODO: 仅当以上 Dispose(bool disposing) 拥有用于释放未托管资源的代码时才替代终结器。
        // ~ExcelBase() {
        //   // 请勿更改此代码。将清理代码放入以上 Dispose(bool disposing) 中。
        //   Dispose(false);
        // }

        // 添加此代码以正确实现可处置模式。
        void IDisposable.Dispose()
        {
            // 请勿更改此代码。将清理代码放入以上 Dispose(bool disposing) 中。
            Dispose(true);
            // TODO: 如果在以上内容中替代了终结器，则取消注释以下行。
            // GC.SuppressFinalize(this);
        }
        #endregion
    }

    /// <summary>
    /// 字符串转换类
    /// </summary>
    public static class ParseChange
    {
        /// <summary>
        /// 字符串转int
        /// </summary>
        /// <param name="inStr">要进行转换的字符串</param>
        /// <param name="defaultValue">默认值，默认：0</param>
        /// <returns></returns>
        public static int ParseInt(string inStr, int defaultValue = 0)
        {
            int parseInt;
            if (int.TryParse(inStr, out parseInt))
                return parseInt;
            return defaultValue;
        }
        public static double ParseDouble(string inStr, double defaultValue = 0.0)
        {
            double parsedouble;
            if (double.TryParse(inStr, out parsedouble))
                return parsedouble;
            return defaultValue;
        }

    }
}
