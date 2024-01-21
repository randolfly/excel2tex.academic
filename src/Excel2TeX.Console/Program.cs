using Cocona;

var app = CoconaApp.Create();

app.AddCommand(
    (
        [Option('s', Description = "src excel file name")] string? src,
        [Option('o', Description = "output tex file name")] string? output,
        [Option('d', Description = "src excel directory")] string? directory,
        [Option(Description = "convert excel files in directory and sub-directories recursively")] bool recursive = false
    ) =>
    {
        output ??= src;
        Console.WriteLine($"src file: {src}, output file: {output}");
    }
);

app.Run();