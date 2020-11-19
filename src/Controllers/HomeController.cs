using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using replace_text_docx.Models;
using Microsoft.AspNetCore.Http;
using System.IO;
using OpenXmlPowerTools;
using DocumentFormat.OpenXml.Packaging;
using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;
using YamlDotNet.RepresentationModel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace replace_text_docx.Controllers
{
    public class HomeController : Controller
    {

        private void ProcessSimplePlaceholders(StreamReader yamlStream, WordprocessingDocument document)
        {
            // prepare YAML document
            TextReader yamlReader = yamlStream;
            var deserializer = new YamlDotNet.Serialization.Deserializer();
            var studentObject = deserializer.Deserialize<dynamic>(yamlReader)["student"];

            var cachedTexts = new Dictionary<string, string>();

            // students
            foreach (var key in studentObject.Keys)
            {
                var value = studentObject[key];
                var isString = value.GetType() == typeof(String);
                if (isString)
                {
                    string placeholder = "{student." + key + "}";
                    string replace_text = value;

                    cachedTexts.Add(placeholder, replace_text);
                }
            }

            // school
            var schoolObject = studentObject["school"];
            foreach (var key in schoolObject.Keys)
            {
                var value = schoolObject[key];
                var isString = value.GetType() == typeof(String);
                if (isString)
                {
                    string placeholder = "{student.school." + key + "}";
                    string replace_text = value;

                    cachedTexts.Add(placeholder, replace_text);
                }
            }

            // replace simple placeholders
            foreach (var item in cachedTexts)
            {
                SearchAndReplacer.SearchAndReplace(document, item.Key, item.Value, true);
            }
        }

        private void ProcessTables(StreamReader yamlStream, WordprocessingDocument document)
        {
            // prepare YAML document
            TextReader yamlReader = yamlStream;
            var deserializer = new YamlDotNet.Serialization.Deserializer();
            var studentObject = deserializer.Deserialize<dynamic>(yamlReader)["student"];
            var subjectsObject = studentObject["subjects"];

            // process docx file
            Body body = document.MainDocumentPart.Document.Body;
            foreach (Table table in body.Descendants<Table>())
            {
                if (table.InnerText.Contains("{student.subjects"))
                {
                    var columnTemplates = new Dictionary<string, int>();

                    TableRow row = table.Elements<TableRow>().ElementAt(1);
                    var columns = row.Elements<TableCell>().ToList();
                    for (var i = 0; i < columns.Count; i++)
                    {
                        var template = columns[i].InnerText;
                        if (template.Contains("{student.subjects."))
                        {
                            template = template.Replace("{student.subjects.", "");
                            template = template.Replace("}", "");

                            columnTemplates.Add(template.ToLower(), i);
                        }
                        else
                        {
                            columnTemplates.Add(string.Empty, i);
                        }
                    }

                    if (columnTemplates.Count > 0)
                    {
                        // remove template row
                        row.Remove();

                        // subjects
                        var subjects = studentObject["subjects"];
                        foreach (var subjectRow in subjects)
                        {
                            var cachedColumns = new Dictionary<string, string>();
                            var subjectColumns = subjectRow.Keys;
                            foreach (var column in subjectColumns)
                            {
                                var value = subjectRow[column];
                                var isString = value.GetType() == typeof(String);
                                if (isString)
                                {
                                    string placeholder = column;
                                    string replace_text = value;

                                    cachedColumns.Add(placeholder.ToLower(), replace_text);
                                }
                            }

                            TableRow newTableRow = new TableRow();
                            foreach (var columnTemplate in columnTemplates)
                            {
                                if (columnTemplate.Key == string.Empty)
                                {
                                    newTableRow.AppendChild(new TableCell(new Paragraph(new Run(new Text(string.Empty)))));
                                }
                                else
                                {
                                    var text = cachedColumns[columnTemplate.Key];
                                    newTableRow.AppendChild(new TableCell(new Paragraph(new Run(new Text(text)))));
                                }
                            }

                            table.AppendChild(newTableRow);
                        }
                    }
                }
            }
        }

        public IActionResult Index()
        {
            // for local testing

            // FileStream fileStream = new FileStream("report_template.docx", FileMode.Open);
            // using (var ms = new MemoryStream())
            // {
            //     fileStream.CopyTo(ms);

            //     using (WordprocessingDocument document = WordprocessingDocument.Open(ms, true))
            //     {
            //         ProcessSimplePlaceholders(new StreamReader("input_file.yaml"), document);
            //         ProcessTables(new StreamReader("input_file.yaml"), document);
            //     }

            //     using (var fileStreamW = new FileStream("CHECK_RESULT.docx", FileMode.Create))
            //     {
            //         fileStreamW.Write(ms.ToArray());
            //     }
            // }

            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }

        [HttpPost]
        public IActionResult AddFile(IFormFileCollection uploads)
        {
            IFormFile docxFile = uploads.Where(x => x.FileName.ToLower().Contains(".docx")).FirstOrDefault();
            IFormFile yamlFile = uploads.Where(x => x.FileName.ToLower().Contains(".yaml")).FirstOrDefault();

            if (docxFile == null || yamlFile == null)
            {
                throw new Exception("No files");
            }

            if (docxFile.Length > 0 && yamlFile.Length > 0)
            {
                //TextReader yamlReader = new StreamReader(yamlFile.OpenReadStream());
                using (var memoryStream = new MemoryStream())
                {
                    docxFile.CopyTo(memoryStream);

                    using (WordprocessingDocument document = WordprocessingDocument.Open(memoryStream, true))
                    {
                        ProcessSimplePlaceholders(new StreamReader(yamlFile.OpenReadStream()), document);
                        ProcessTables(new StreamReader(yamlFile.OpenReadStream()), document);
                    }

                    return new FileContentResult(memoryStream.ToArray(), "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                    {
                        FileDownloadName = $"updated_document.docx"
                    };
                }
            }

            return RedirectToAction("Index");
        }
    }
}
