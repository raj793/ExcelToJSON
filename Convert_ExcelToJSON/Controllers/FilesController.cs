using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Web;
using System.Web.Http;

namespace Convert_ExcelToJSON.Controllers
{
    [RoutePrefix("files")]
    public class FilesController : ApiController
    {
        [HttpPost]
        [Route("Input/{tab}")]
        public List<List<List<string>>> ExcelToString([FromUri]int tab=1)
        {
            var httpRequest = HttpContext.Current.Request;
            var postFile = httpRequest.Files[0];
            if (postFile == null)
                throw new Exception("File is null");
            Workbook wb = new Workbook(postFile.InputStream);
            var originalFilename = postFile.FileName;
            var indexOfFileSeparator = originalFilename.IndexOf('.');
            var fileName = originalFilename.Substring(0, indexOfFileSeparator);
            var tabCount = wb.Worksheets.Count;
            if (tab > tabCount)
                throw new Exception("This worksheet has only " + tabCount + " tabs.");

            var tabArray = new List<List<List<string>>>();
            try
            {

                for (int t = 0; t < tab; t++)
                {
                    Worksheet worksheet = wb.Worksheets[t];
                    var ColumnRange = worksheet.Cells.Columns.Count;
                    var RowRange = worksheet.Cells.Rows.Count;
                    if (ColumnRange == 0 || RowRange == 0)
                        throw new Exception("No Rows or Columns detected in tab no. " + t);
                    DataTable dataTable = worksheet.Cells.ExportDataTable(0, 0, RowRange, ColumnRange);
                    var itemArray = new List<List<string>>();
                    foreach (DataRow r in dataTable.Rows)
                    {
                        var myString = r.ItemArray;
                        var list = new List<string>();
                        for (int i = 0; i < myString.LongLength; i++)
                        {
                            if (myString[i] == null)
                                list.Add("Null");
                            else
                                list.Add(myString.GetValue(i).ToString());
                        }
                        itemArray.Add(list);
                    }
                    tabArray.Add(itemArray);
                }
            }
            catch(Exception e)
            {
                throw new Exception("Exception:" + e.Message);
            }
            return tabArray;
        }
    }
}
        

