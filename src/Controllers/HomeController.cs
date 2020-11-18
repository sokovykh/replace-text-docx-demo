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

namespace replace_text_docx.Controllers
{
    public class HomeController : Controller
    {
        public IActionResult Index()
        {
            // TextReader tr = new StreamReader("input_file.yaml");

            // var cachedTexts = new Dictionary<string, string>();

            // var deserializer = new YamlDotNet.Serialization.Deserializer();

            // // student 
            // var studentObject = deserializer.Deserialize<dynamic>(tr)["student"];
            // foreach (var key in studentObject.Keys)
            // {
            //     var value = studentObject[key];
            //     var isString = value.GetType() == typeof(String);
            //     if (isString)
            //     {
            //         string placeholder = "{student." + key + "}";
            //         string replace_text = value;

            //         cachedTexts.Add(placeholder, replace_text);
            //     }
            // }

            // // school
            // var schoolObject = studentObject["school"];
            // foreach (var key in schoolObject.Keys)
            // {
            //     var value = schoolObject[key];
            //     var isString = value.GetType() == typeof(String);
            //     if (isString)
            //     {
            //         string placeholder = "{student.school." + key + "}";
            //         string replace_text = value;

            //         cachedTexts.Add(placeholder, replace_text);
            //     }
            // }

            // FileStream fileStream = new FileStream("report_template.docx", FileMode.Open);
            // using (var ms = new MemoryStream())
            // {
            //     fileStream.CopyTo(ms);

            //     using (WordprocessingDocument doc = WordprocessingDocument.Open(ms, true))
            //     {
            //         foreach (var item in cachedTexts)
            //         {
            //             SearchAndReplacer.SearchAndReplace(doc, item.Key, item.Value, true);
            //         }
            //     }

            //     using(var fileStreamW = new FileStream("RESULT.docx", FileMode.Truncate))
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

            // if (docxFile == null || yamlFile == null)
            // {
            //     throw new Exception("No files");
            // }

            if (docxFile.Length > 0 && yamlFile.Length > 0)
            {
                TextReader tr = new StreamReader(yamlFile.OpenReadStream());

                var cachedTexts = new Dictionary<string, string>();

                var deserializer = new YamlDotNet.Serialization.Deserializer();

                // student 
                var studentObject = deserializer.Deserialize<dynamic>(tr)["student"];
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

                using (var ms = new MemoryStream())
                {
                    docxFile.CopyTo(ms);

                    using (WordprocessingDocument doc = WordprocessingDocument.Open(ms, true))
                    {
                        foreach (var item in cachedTexts)
                        {
                            SearchAndReplacer.SearchAndReplace(doc, item.Key, item.Value, true);
                        }
                    }

                    // using (var fileStreamW = new FileStream("RESULT.docx", FileMode.Truncate))
                    // {
                    //     fileStreamW.Write(ms.ToArray());
                    // }

                    return new FileContentResult(ms.ToArray(), "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                    {
                        FileDownloadName = $"updated_document.docx"
                    };


                }
                


            }








            foreach (var uploadedFile in uploads)
            {
                // путь к папке Files
                string path = "~/Files/" + uploadedFile.FileName;
                // сохраняем файл в папку Files в каталоге wwwroot
                // using (var fileStream = new FileStream(_appEnvironment.WebRootPath + path, FileAccess.ReadWrite))
                // {
                //     await uploadedFile.CopyToAsync(fileStream);
                // }
                //FileModel file = new FileModel { Name = uploadedFile.FileName, Path = path };
                //_context.Files.Add(file);
            }
            //_context.SaveChanges();

            return RedirectToAction("Index");
        }
    }
}
