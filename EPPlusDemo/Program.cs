namespace EPPlusDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            var path = @"C:\development-code-school\EPPlusDemo\Data\EPP Demo.xlsx";

            var workBook = new GetWorkBook(path).SelectWorkBook();

            var workSheet1 = new GetWorkSheet(path).SelectWorkSheet();
            var workSheet2 = new GetWorkSheet(path, 2).SelectWorkSheet();
        }
    }
}
