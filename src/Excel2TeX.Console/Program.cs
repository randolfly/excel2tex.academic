using Cocona;
using Excel2TeX;

var app = CoconaApp.Create();

app.AddCommand(
    (
        [Option('s', Description = "src excel file name")] string? src,
        [Option('o', Description = "output tex file name")] string? output,
        [Option('d', Description = "src excel directory")] string? dir,
        [Option(Description = "convert excel files in directory and sub-directories recursively")] bool recursive = false
    ) =>
    {
        // single file case
        if (src is not null)
        {
            var fullPath = Path.GetFullPath(src, Environment.CurrentDirectory);
            if (!File.Exists(fullPath))
            {
                Console.WriteLine($"file: {fullPath} not exists!");
                return;
            }
            if (!fullPath.EndsWith(AppConfig.SourceFileSuffix))
            {
                Console.WriteLine($"file: {fullPath} is not excel file!");
                return;
            }
            //TODO single excel file convertion
            output ??= fullPath.Replace(AppConfig.SourceFileSuffix, AppConfig.TargetFileSuffix);
            Console.WriteLine($"src file: {src}, output file: {output}");
            return;
        }
        // directory case
        if (dir is not null)
        {
            var fullPath = Path.GetFullPath(dir, Environment.CurrentDirectory);
            if (!Directory.Exists(fullPath))
            {
                Console.WriteLine($"file: {fullPath} not exists!");
                return;
            }
            var fileList = Directory.GetFiles(fullPath)
                                    .Where(f => f.EndsWith(AppConfig.SourceFileSuffix))
                                    .ToList();
            //TODO multiple excel file convertion
            return;
        }
    }
);

app.Run();