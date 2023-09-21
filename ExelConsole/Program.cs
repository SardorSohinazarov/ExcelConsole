using OfficeOpenXml;

internal class Program
{
    private static void Main(string[] args)
    {
        var listOfStudents = new List<string>()
        {
            "Sardor Sohinazarov",
            "Sarvar Sohinazarov",
            "Sanjar Sohinazarov",
            "Komil Sohinazarov",
            "Karim Sohinazarov",
            "Alisher Sohinazarov",
            "Bobur Sohinazarov",
        };

        // If you are a commercial business and have
        // purchased commercial licenses use the static property
        // LicenseContext of the ExcelPackage class:
        //ExcelPackage.LicenseContext = LicenseContext.Commercial;

        // If you use EPPlus in a noncommercial context
        // according to the Polyform Noncommercial license:
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        var path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop),"ListOfUsers.xlsx");
        

        using (var package = new ExcelPackage(path))
        {
            #region Bu yerda faqatgina file bor yo'qligi
            FileInfo file = new FileInfo(path);
            if (!file.Exists)
            {
                file.Create();
            }
            #endregion

            var sheet = package.Workbook.Worksheets.Add("My Sheet");

            sheet.Cells[$"A1"].Value = "Nomer";
            sheet.Cells[$"B1"].Value = "Ism-familiya";

            for (int i = 0;i < listOfStudents.Count; i++)
            {
                sheet.Cells[$"A{i+2}"].Value = i+1;
                sheet.Cells[$"B{i+2}"].Value = listOfStudents[i].ToString();
            }

            package.Save();
        }
    }
}