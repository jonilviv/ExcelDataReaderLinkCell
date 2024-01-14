using ExcelDataReader;
using System.Data;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace ExcelDataReaderLinkCell
{
    public class UnitTest1
    {
        public UnitTest1()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        }

        [Fact]
        public async Task Reader_link()
        {
            string fileName = "Book1.xlsx";
            await using FileStream stream = File.OpenRead(fileName);
            using IExcelDataReader reader = ExcelReaderFactory.CreateOpenXmlReader(stream);

            do
            {
                while (reader.Read())
                {
                    object val = reader.GetValue(0);

                    Assert.Equal("https://google.com", val);
                }
            } while (reader.NextResult());
        }

        [Fact]
        public async Task DataSet_link()
        {
            string fileName = "Book1.xlsx";
            await using FileStream stream = File.OpenRead(fileName);
            using IExcelDataReader reader = ExcelReaderFactory.CreateOpenXmlReader(stream);

            DataSet dataSet = reader.AsDataSet();

            foreach (DataTable table in dataSet.Tables)
            {
                foreach (DataRow row in table.Rows)
                {
                    object val = row[0];

                    Assert.Equal("https://google.com", val);
                }
            }
        }
    }
}