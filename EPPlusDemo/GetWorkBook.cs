using OfficeOpenXml;
using System.Collections.Generic;
using System.IO;

namespace EPPlusDemo
{
    public class GetWorkBook
    {
        protected int WorkSheetIndex = -1;

        private string ExcelPath = "";
        private List<List<ExcelModel>> workBook = new List<List<ExcelModel>>();
        private bool FirstRowIsHeader = true;

        public GetWorkBook(string excelPath, int workSheetIndex = -1)
        {
            ExcelPath = excelPath;
            WorkSheetIndex = workSheetIndex;
        }

        public List<List<ExcelModel>> SelectWorkBook()
        {
            workBook = new List<List<ExcelModel>>();
            using (var package = new ExcelPackage(new FileInfo(ExcelPath)))
            {
                if (WorkSheetIndex > 0)
                {
                    var workSheet = package.Workbook.Worksheets[WorkSheetIndex];
                    ProcessWorkSheet(workSheet);
                }
                else
                {
                    foreach (var workSheet in package.Workbook.Worksheets)
                    {
                        ProcessWorkSheet(workSheet);
                    }
                }
            }
            return workBook;
        }

        private void ProcessWorkSheet(ExcelWorksheet workSheet)
        {
            var list = new List<ExcelModel>();

            var startRow = workSheet.Dimension.Start.Row;
            if (FirstRowIsHeader)
                startRow++;

            for (int i = startRow; i <= workSheet.Dimension.End.Row; i++)
            {
                var obj = new ExcelModel();
                obj.Id = int.Parse(workSheet.Cells[i, 1].Text);
                obj.Name = workSheet.Cells[i, 2].Text;
                obj.Surname = workSheet.Cells[i, 3].Text;
                obj.CellPhone = long.Parse(workSheet.Cells[i, 4].Text);
                obj.EmailAddress = workSheet.Cells[i, 5].Text;
                list.Add(obj);
            }
            workBook.Add(list);
        }
    }
}
