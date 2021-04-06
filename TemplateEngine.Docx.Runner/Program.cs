using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using CommandLine;
using CommandLine.Text;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

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

            if (param.Force && File.Exists(finalPath))
                File.Delete(finalPath);
            File.Copy(param.SourcePath, finalPath);

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
