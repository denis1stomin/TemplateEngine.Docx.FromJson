using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Diagnostics;

namespace TemplateEngine.Docx.Runner
{
    public static class JustLogic
    {
        public const string TypeAttributeName = "__type";
        public const string FieldSelectorAttributeName = "controlTagSelector";
        public const string FieldValueAttributeName = "controlValue";
        public const string TableRowsAttributeName = "tablesRows";
        public const string PicturePathAttributeName = "picturePath";

        public static void ResolveTemplate(string templatePath, string dataSourcePath, bool removeContentControls)
        {
            var timer = Stopwatch.StartNew();

            var contentToFill = GetContent(dataSourcePath, new List<IContentItem>()
            {
                // TODO : move format to app.settings
                new FieldContent("CreationDateTime", DateTimeOffset.Now.Date.ToShortDateString())
            });

            using (var outputDocument = new TemplateProcessor(templatePath)
                .SetRemoveContentControls(removeContentControls)
                .SetNoticeAboutErrors(false))
            {
                outputDocument.FillContent(contentToFill);
                outputDocument.SaveChanges();
            }

            timer.Stop();
            Console.WriteLine($"Output path: '{templatePath}'");
            Console.WriteLine($"Time spent: '{timer.Elapsed}'");
        }

        private static Content GetContent(string sourceJsonPath, IReadOnlyList<IContentItem> additionalContent)
        {
            var fields = new List<IContentItem>(additionalContent);

            var jArr = ParseJsonSource(sourceJsonPath);
            foreach (var jItem in jArr)
            {
                var type = jItem[TypeAttributeName].ToString();
                if (type.Equals(nameof(FieldContent), StringComparison.OrdinalIgnoreCase))
                {
                    var selector = jItem[FieldSelectorAttributeName].ToString();

                    if (fields.Any(x => x.Name.Equals(selector, StringComparison.OrdinalIgnoreCase)))
                        continue;

                    var value = jItem[FieldValueAttributeName].ToString();

                    fields.Add(new FieldContent(selector, value));
                }
                else if (type.Equals(nameof(TableContent), StringComparison.OrdinalIgnoreCase))
                {
                    var selector = jItem[FieldSelectorAttributeName].ToString();
                    var table = new TableContent(selector);

                    var rows = jItem[TableRowsAttributeName].ToObject<JArray>();
                    foreach (var row in rows)
                    {
                        var rowColumns = GetTableRow(row);
                        table.AddRow(rowColumns.ToArray());
                    }

                    fields.Add(table);
                }
                else if (type.Equals(nameof(ImageContent), StringComparison.OrdinalIgnoreCase))
                {
                    var selector = jItem[FieldSelectorAttributeName].ToString();
                    var picturePath = jItem[PicturePathAttributeName].ToString();

                    var picture = new ImageContent(selector, File.ReadAllBytes(picturePath));
                    fields.Add(picture);
                }
                else
                {
                    throw new NotSupportedException($"Content type '{type}' is not supported yet.");
                }
            }

            return new Content(fields.ToArray());
        }

        private static IReadOnlyList<IContentItem> GetTableRow(JToken rowObject)
        {
            var result = new List<IContentItem>();

            foreach (JProperty prop in rowObject)
                result.Add(new FieldContent(prop.Name, prop.Value.ToString()));

            return result;
        }

        private static JArray ParseJsonSource(string sourceJsonPath)
        {
            using (var reader = new StreamReader(
                new FileStream(sourceJsonPath, FileMode.Open, FileAccess.Read, FileShare.Read)))
            {
                var strContent = reader.ReadToEnd();
                var objContent = JsonConvert.DeserializeObject<JArray>(strContent);

                return objContent;
            }
        }
    }
}
