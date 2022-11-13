string? jsonFileName = args.FirstOrDefault();
Console.WriteLine("read json file: " + jsonFileName);

if (string.IsNullOrEmpty(jsonFileName))
{
    Console.WriteLine("please start with file");
    return;
}

if (!File.Exists(jsonFileName))
{
    Console.WriteLine("no such file: {0}", jsonFileName);
    return;
}

XlsxData? jsonData = XlsxHelper.LoadJsonFile(jsonFileName);
if (jsonData == null)
{
    Console.WriteLine("no JSON-File");
    return;
}

var xlsx = new Xlsx(jsonData);
xlsx.Convert();


// Console.WriteLine("press any key to end the program");
// Console.ReadLine();
Console.WriteLine("END");
