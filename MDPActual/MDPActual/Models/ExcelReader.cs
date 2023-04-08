using ExcelDataReader;
using System.Data;

namespace MDPActual.Models
{
    //Класс для читания эксель файла
    public class ExcelReader
    {
        private int pageNumber;
        private string fileName;
        //конструктор для чтения нужной таблицы
        public ExcelReader(int pageNumber, string fileName)
        {
            this.pageNumber = pageNumber;
            this.fileName = fileName;
        }

        public List<string> SearchNameOfItems(List<string> array = null)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            using (var stream = System.IO.File.Open(fileName, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateOpenXmlReader(stream))
                {
                    DataTable specificWorkSheet = reader.AsDataSet().Tables[pageNumber];
                    int[] position = new int[2];
                    position = SearchPosition(specificWorkSheet, "наименование");
                    int startRow = position[0];
                    int startColumn = position[1];
                    int currentRow = 0;

                    foreach (var row in specificWorkSheet.Rows)
                    {
                        if (currentRow > startRow && ((DataRow)row)[startColumn].ToString() != "")
                            array.Add(((DataRow)row)[startColumn].ToString());
                        currentRow++;
                    }
                }
            }
            return array;
        }

        public int[] SearchPosition(DataTable dt, string name)
        {
            int currentRow = 0;
            int[] position = new int[2];

            foreach (var row in dt.Rows)
            {
                for (int currentColumn = 0; currentColumn <= 5; currentColumn++)
                    if (((DataRow)row)[currentColumn].ToString().ToLower() == name.ToLower())
                    {
                        position[0] = currentRow;
                        position[1] = currentColumn;
                        break;
                    }
                currentRow++;
            }
            return position;
        }
    }
}
