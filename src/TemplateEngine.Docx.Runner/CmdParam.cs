using System;
using CommandLine;

namespace TemplateEngine.Docx.Runner
{
    class CmdParam
    {
        [Option('s', "source", Required = true, HelpText = "Path to a srouce data file.")]
        public string SourcePath { get; set; }

        [Option('t', "template", Required = true, HelpText = "Path to a template Word document.")]
        public string TemplatePath { get; set; }

        [Option('o', "output", Default = "generated_document.docx", Required = false, HelpText = "Output path to a generated document.")]
        public string OutputPath { get; set; }

        [Option('f', "force", Default = false, Required = false, HelpText = "Rewrite output document if exists.")]
        public bool Force { get; set; }

        [Option("finalize", Required = false, HelpText = "If set to True removes content controls from the output document.")]
        public bool FinalizeTemplate { get; set; }

        [Option("convert-to", Default = null, Required = false, HelpText = "Converts final document to a target format. See https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word.wdsaveformat?view=word-pia")]
        public string ConvertToFormat { get; set; }
    }
}
