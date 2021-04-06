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
            var finalPath = Path.GetFullPath($"{param.OutputPath}");
            var finalPathInfo = new FileInfo(finalPath);

            if (!Directory.Exists(finalPathInfo.Directory.FullName))
                Directory.CreateDirectory(finalPathInfo.Directory.FullName);

            if (param.Force && File.Exists(finalPath))
                File.Delete(finalPath);

            File.Copy(param.TemplatePath, finalPath);

            JustLogic.ResolveTemplate(param.OutputPath, param.SourcePath, param.FinalizeTemplate);
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
