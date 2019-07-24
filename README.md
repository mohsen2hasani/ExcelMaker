install "EPPlus.Core" nuget and use this class to get excel file from any list of objects in .net core projects:

example:

```
public class TestController : Controller
{
    public IActionResult Index()
    {
        //var model = list of any classes
        var excel = model.GetExcel("ExcelFileName");
        return File(excel.FileContents, excel.ContentType, excel.FileDownloadName);
    }
}
```
