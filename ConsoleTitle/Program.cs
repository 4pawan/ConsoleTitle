// See https://aka.ms/new-console-template for more information

using System.Diagnostics;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using HtmlAgilityPack;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

Console.WriteLine("Hello, World!");
int k = 0;

ISheet sheet;
string path = @"C:\Project\Title\ConsoleTitle\ConsoleTitle\input\Book2.xlsx";
string filePath = @"C:\Project\Title\ConsoleTitle\ConsoleTitle\input\Book234.xlsx";
string txtPath = @"C:\Project\Title\ConsoleTitle\ConsoleTitle\input\saved.txt";

using (var stream = new FileStream(path, FileMode.Open, FileAccess.ReadWrite))
{
    stream.Position = 0;
    XSSFWorkbook xssWorkbook = new XSSFWorkbook(stream);
    sheet = xssWorkbook.GetSheetAt(0);

    for (int i = (sheet.FirstRowNum + 1); i <= sheet.LastRowNum; i++)
    {
        IRow row = sheet.GetRow(i);
        if (row == null) continue;
        if (row.Cells.All(d => d.CellType == CellType.Blank)) continue;
        for (int j = row.FirstCellNum; j < 1; j++)
        {
            if (row.GetCell(j) != null)
            {
                if (!string.IsNullOrEmpty(row.GetCell(j).ToString()) && !string.IsNullOrWhiteSpace(row.GetCell(j).ToString()))
                {
                    var cellValue = row.GetCell(j).ToString();
                    var response = await GetResponseString(cellValue);
                    string data = GetBreadCrumb(response);
                    k++;
                    Debug.WriteLine("----- >" + k);
                    var cell1 = row.CreateCell(1);
                    cell1.SetCellValue(data);
                    Console.WriteLine("----- >" + k);
                    AppendDataToTextFile(data);

                }
            }
        }
    }


    FileStream fs = new FileStream(filePath, FileMode.Create);
    xssWorkbook.Write(fs);

}
async Task<string> GetResponseString(string Url)
{
    var httpClient = new HttpClient();
    Uri myUri = new Uri(Url, UriKind.Absolute);
    int effort = 1;
    string contents = null;
    while (effort < 4)
    {
        try
        {
            var response = await httpClient.GetAsync(myUri);
            contents = await response.Content.ReadAsStringAsync();

            if (response.StatusCode == HttpStatusCode.Moved)
            {
                var redirectedUrl = response.Headers.Location.AbsoluteUri;
                var responseredirected = await httpClient.GetAsync(redirectedUrl);
                contents = await responseredirected.Content.ReadAsStringAsync();
            }


            break;
        }
        catch (Exception ex)
        {
            effort++;
            Console.WriteLine("Error..retrying " + effort + " url :" + Url);
        }
    }

    return contents;
}

static string GetTitle(string file)
{
    Match m = Regex.Match(file, @"<title>\s*(.+?)\s*</title>");
    if (m.Success)
    {
        return m.Groups[1].Value;
    }
    else
    {
        return "";
    }
}
static string GetBreadCrumb(string file)
{

    //string pattern2 = @"<div role=""navigation"" aria-label=""Breadcrumb"" class=""breadcrumb"">\s*(.+?)\s*</div>";
    //string pattern3 = @"<nav class=""task-breadcrumbs"" aria-label=""Breadcrumb"">\s*(.+?)\s*</nav>";
    StringBuilder result = new StringBuilder();

    HtmlDocument doc = new HtmlDocument();
    doc.LoadHtml(file);

    var pattern1 = doc.DocumentNode.SelectNodes("//div[@class='breadcrumb']");
    if (pattern1 != null)
    {
        var list = pattern1.Descendants("li");
        foreach (var item in list)
        {
            result.Append(item.InnerText + "/");
        }

        return result.ToString().TrimEnd('/');
    }

    var pattern2 = doc.DocumentNode.SelectNodes("//div/nav[@aria-label='Breadcrumb']");
    if (pattern2 != null)
    {
        var list = pattern2.Descendants("li");
        foreach (var item in list)
        {
            result.Append(item.InnerText + "/");
        }

        return result.ToString().TrimEnd('/'); ;

    }
    return null;
}

void AppendDataToTextFile(string title)
{

    using (var sw = new StreamWriter(txtPath, true))
    {
        sw.WriteLine(title);
    }

}



Console.WriteLine("Press any key to exit...");
Console.ReadKey();