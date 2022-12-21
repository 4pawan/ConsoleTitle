// See https://aka.ms/new-console-template for more information

using System.Diagnostics;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using HtmlAgilityPack;
using NPOI.HPSF;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

Console.WriteLine("Hello, World!");
int k = 0;

ISheet sheet;
string rootpath = @"C:\Project\ConsoleTitle\ConsoleTitle\input";
string path = @$"{rootpath}\Book2.xlsx";
string filePath = @$"{rootpath}\Book234.xlsx";
string txtTitlePath = @$"{rootpath}\title.txt";
string txtBreadcrumbPath = @$"{rootpath}\breadcrumb.txt";
string txtAssetCountPath = @$"{rootpath}\count.txt";
string txtDetailedListForGivenLinkPath = @"{0}\links\{1}.txt"; // {rootpath}\links\guid.txt


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
                    string title = GetTitle(response);
                    string breadcrumb = GetBreadCrumb(response);
                    string guid = Guid.NewGuid().ToString("N");
                    string dataToBeSavedinExcel = GetAssetDetails(response, guid);

                    k++;
                    Debug.WriteLine("----- >" + k);
                    var cell1 = row.CreateCell(1);
                    cell1.SetCellValue(dataToBeSavedinExcel);
                    Console.WriteLine("----- >" + k);
                    AppendDataToTextFile(dataToBeSavedinExcel, txtAssetCountPath);
                    AppendDataToTextFile(title, txtTitlePath);
                    AppendDataToTextFile(breadcrumb, txtBreadcrumbPath);
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

    var pattern3 = doc.DocumentNode.SelectNodes("//ol[@class='breadcrumb']");
    if (pattern3 != null)
    {
        var list = pattern3.Descendants("li");
        foreach (var item in list)
        {
            var span = item.ChildNodes["span"].InnerText;
            result.Append(span + "/");
        }

        return result.ToString().TrimEnd('/'); ;

    }



    return null;
}

string GetAssetDetails(string file, string guid)
{
    StringBuilder result = new StringBuilder();
    string AssetCount = "{0}@Pdf-{1}|doc-{2}";
    HtmlDocument doc = new HtmlDocument();
    doc.LoadHtml(file);

    var links = doc.DocumentNode.SelectNodes("//a[@href]");

    if (links != null)
    {
        var pdflist = links.Where(link => link.GetAttributeValue("href", null).EndsWith(".pdf"));
        var doclist = links.Where(link => link.GetAttributeValue("href", null).EndsWith(".doc"));
        StringBuilder builder = new StringBuilder();

        int pdfcount = 0;
        int doccount = 0;
        string detailedPath = string.Format(txtDetailedListForGivenLinkPath, rootpath, guid);

        foreach (var item in pdflist)
        {
            var medialink = item.GetAttributeValue("href", null);
            builder.AppendLine(medialink);
            pdfcount++;
        }

        foreach (var item in doclist)
        {
            var medialink = item.GetAttributeValue("href", null);
            builder.AppendLine(medialink);
            doccount++;
        }
        string details = string.Format(AssetCount, guid, pdfcount, doccount);
        AppendDataToTextFile(builder.ToString(), detailedPath);
        return details;
    }


    return null;
}


void AppendDataToTextFile(string data, string txtPath)
{

    if (!File.Exists(txtPath))
    {
        using (StreamWriter sw = File.CreateText(txtPath))
        {
            sw.WriteLine(data);
        }

    }
    else
    {
        using (var sw = new StreamWriter(txtPath, true))
        {
            sw.WriteLine(data);
        }
    }
}



Console.WriteLine("Press any key to exit...");
Console.ReadKey();
