using System.Collections.Generic;

namespace EPPlusDemo
{
    public class GetWorkSheet : GetWorkBook
    {
        public GetWorkSheet(string excelPath, int workSheetIndex = 1) 
            : base(excelPath, workSheetIndex)
        {
            WorkSheetIndex = workSheetIndex;
        }

        public List<ExcelModel> SelectWorkSheet()
        {
            var workBook = base.SelectWorkBook();
            if (workBook.Count > 0)
                return workBook[0];
            else
                return new List<ExcelModel>();
        }
    }
}
