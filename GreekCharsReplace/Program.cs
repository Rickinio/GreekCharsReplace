using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace GreekCharsReplace
{
    internal class Program
    {
        static void Main(string[] args)
        {
            ReplaceGreekLettersInExcel(@"C:\Temp\Stamatopoulos_Product_19042023.xlsx");
        }

        public static void ReplaceGreekLettersInExcel(string filePath)
        {
            IWorkbook workbook;
            using (var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite))
            {
                string fileExtension = Path.GetExtension(filePath);
                if (fileExtension == ".xls")
                {
                    workbook = new HSSFWorkbook(fileStream);
                }
                else if (fileExtension == ".xlsx")
                {
                    workbook = new XSSFWorkbook(fileStream);
                }
                else
                {
                    throw new ArgumentException("Invalid file format. Only .xls and .xlsx files are supported.");
                }

                ISheet worksheet = workbook.GetSheetAt(0); // Assuming the first worksheet is the target sheet

                for (int row = 0; row <= worksheet.LastRowNum; row++)
                {
                    IRow currentRow = worksheet.GetRow(row);
                    if (currentRow != null)
                    {
                        ICell greekLetterCell = currentRow.GetCell(1);
                        if (greekLetterCell != null)
                        {
                            string greekLetter = greekLetterCell.StringCellValue;
                            string englishEquivalent = ReplaceGreekWithEnglish(greekLetter);
                            ICell englishEquivalentCell = currentRow.GetCell(3) ?? currentRow.CreateCell(3);
                            englishEquivalentCell.SetCellValue(englishEquivalent);
                        }
                    }
                }

                using (var outputFileStream = new FileStream(filePath, FileMode.Create))
                {
                    workbook.Write(outputFileStream, false);
                }
            }
        }

        public static string ReplaceGreekWithEnglish(string input)
        {
            // Define Greek to English mapping
            var greekToEnglish = new Dictionary<string, string>()
            {
                { "Α", "A" },
                { "Β", "B" },
                { "Ε", "E" },
                { "Ζ", "Z" },
                { "Η", "H" },
                { "Ι", "I" },
                { "Κ", "K" },
                { "Μ", "M" },
                { "Ν", "N" },
                { "Ο", "O" },
                { "Ρ", "P" },
                { "Τ", "T" },
                { "Υ", "Y" },
                { "Χ", "X" },
                // Add more mappings as needed
            };

            // Replace Greek letters with English equivalents
            foreach (var kvp in greekToEnglish)
            {
                input = input.Replace(kvp.Key, kvp.Value);
            }

            return input;
        }
    }
}