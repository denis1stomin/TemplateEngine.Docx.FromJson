using System;
using System.IO;
using CommandLine;
using CommandLine.Text;

namespace TemplateEngine.Docx.Runner
{
    partial class Program
    {
        static void Main(string[] args)
        {
            var parser = new Parser(x => {
                x.CaseSensitive = false;
                x.IgnoreUnknownArguments = true;
                x.AutoHelp = true;
            });

            var parseRes = parser.ParseArguments<CmdParam>(args)
                .WithParsed(MainInner);

            if (parseRes.Tag == ParserResultType.NotParsed)
                NotParsed(parseRes);
        }

        static void MainInner(CmdParam param)
        {
            var outputAbsolutePath = Path.GetFullPath($"{param.OutputPath}");
            var outputAbsolutePathInfo = new FileInfo(outputAbsolutePath);

            if (!Directory.Exists(outputAbsolutePathInfo.Directory.FullName))
                Directory.CreateDirectory(outputAbsolutePathInfo.Directory.FullName);

            if (param.Force && File.Exists(outputAbsolutePath))
                File.Delete(outputAbsolutePath);

            File.Copy(param.TemplatePath, outputAbsolutePath);

            JustLogic.ResolveTemplate(outputAbsolutePath, Path.GetFullPath(param.SourcePath), param.FinalizeTemplate);

            if (!string.IsNullOrWhiteSpace(param.ConvertToFormat))
            {
                var outputFolder = Path.GetFullPath(Path.GetDirectoryName(param.OutputPath));
                DocumentFormatConvertor.Convert(outputAbsolutePath, outputFolder, param.ConvertToFormat);
            }
        }

        static void NotParsed(ParserResult<CmdParam> res)
        {
            var appName = AppDomain.CurrentDomain.FriendlyName;
            Console.WriteLine($"{appName} - Word document generator which uses TemplateEngine.Docx package under the hood.");
            Console.WriteLine(HelpText.AutoBuild(res));

            Environment.Exit(1);
        }
    }
}
