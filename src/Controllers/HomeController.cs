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
            var rootLevel = deserializer.Deserialize<dynamic>(yamlReader);
            var cachedTexts = new Dictionary<string, string>();

            foreach (var key in rootLevel.Keys)
            {
                var value = rootLevel[key];
                var isString = value.GetType() == typeof(String);
                if (isString)
                {
                    string placeholder = "{" + key + "}";
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
            var rootLevel = deserializer.Deserialize<dynamic>(yamlReader);

            // GridName => Array of column name to value 
            var parsedTables = new Dictionary<string, List<Dictionary<string, string>>>();

            foreach (string gridName in rootLevel.Keys)
            {
                var grid = rootLevel[gridName];
                var isString = grid.GetType() == typeof(String);
                if (isString)
                {
                    continue;
                }
                else
                {
                    var parsedRows = new List<Dictionary<string, string>>();
                    foreach (var row in grid)
                    {
                        var columns = row.Keys;
                        var rowWithValues = new Dictionary<string, string>();

                        foreach (var column in columns)
                        {
                            rowWithValues.Add(column, row[column]);
                        }

                        parsedRows.Add(rowWithValues);
                    }

                    parsedTables.Add(gridName.ToLower(), parsedRows);
                }
            }

            // process docx file
            Body body = document.MainDocumentPart.Document.Body;
            foreach (Table table in body.Descendants<Table>())
            {
                var rowInnerText = table.InnerText.ToLower();

                var hasFoundGrid = false;

                // find grid
                foreach (string gridName in parsedTables.Keys)
                {
                    var startedFrom = "{" + gridName + ".";

                    if (rowInnerText.Contains(startedFrom))
                    {
                        hasFoundGrid = true;

                        // we have found YAML Grid for DOCX table
                        var docxColumnTemplates = new List<string>();

                        TableRow row = table.Elements<TableRow>().ElementAt(1);
                        var columns = row.Elements<TableCell>().ToList();
                        for (var i = 0; i < columns.Count; i++)
                        {
                            var template = columns[i].InnerText.ToLower();
                            

                            if (template.Contains(startedFrom))
                            {
                                template = template.Replace(startedFrom, "");
                                template = template.Replace("}", "");

                                docxColumnTemplates.Add(template);
                            }
                            else
                            {
                                docxColumnTemplates.Add(string.Empty);
                            }
                        }

                        if (docxColumnTemplates.Count > 0)
                        {
                            // remove template row
                            row.Remove();

                            var parsedRows = parsedTables[gridName];
                            foreach (var parsedRow in parsedRows)
                            {
                                var newTableRow = new TableRow();
                                foreach (var docxColumn in docxColumnTemplates)
                                {
                                    var newValue = "";
                                    if (!parsedRow.ContainsKey(docxColumn))
                                    {
                                        newValue = string.Empty;
                                    }else
                                    {
                                        newValue = parsedRow[docxColumn];
                                    }

                                    newTableRow.AppendChild(new TableCell(new Paragraph(new Run(new Text(newValue)))));
                                }

                                table.AppendChild(newTableRow);
                            }
                        }

                    }

                    if (hasFoundGrid) break;
                }
            }
        }

        public IActionResult Index()
        {
            // only for local testing
            // const string yaml_file_name = "input_file_upd.yaml";
            // const string docx_file_name = "input_report.docx";
            // const string output_report_file = "CHECK_RESULT.docx";


            // //FileStream fileStream = new FileStream("report_template.docx", FileMode.Open);
            // FileStream fileStream = new FileStream(docx_file_name, FileMode.Open);
            // using (var ms = new MemoryStream())
            // {
            //     fileStream.CopyTo(ms);

            //     using (WordprocessingDocument document = WordprocessingDocument.Open(ms, true))
            //     {
            //         ProcessSimplePlaceholders(new StreamReader(yaml_file_name), document);
            //         ProcessTables(new StreamReader(yaml_file_name), document);
            //     }

            //     using (var fileStreamW = new FileStream(output_report_file, FileMode.Create))
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
