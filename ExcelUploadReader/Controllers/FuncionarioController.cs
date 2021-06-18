using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading.Tasks;
using ExcelDataReader;
using ExcelUploadReader.Models;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Hosting;

namespace ExcelUploadReader.Controllers
{
    public class FuncionarioController : Controller
    {
        private IWebHostEnvironment webHostEnvironment;
        public FuncionarioController(IWebHostEnvironment _webHostEnvironment)
        {
            webHostEnvironment = _webHostEnvironment;
        }


        [HttpGet]
        public IActionResult Index(List<Funcionario> funcionarios = null)
        {

            funcionarios = funcionarios == null ? new List<Funcionario>() : funcionarios;

            return View(funcionarios);
        }

        [HttpPost]
        public IActionResult Index(IFormFile file, [FromServices] IWebHostEnvironment hostEnvironment)
        {
            string filename = $"{hostEnvironment.WebRootPath}\\files\\{file.FileName}";
            using (FileStream fileStream = System.IO.File.Create(filename))
            {
                file.CopyTo(fileStream);
                fileStream.Flush();
            }
            var funcionarios = this.Getfuncionario(file.FileName);

            var arrtest = new string[] { "aaaa", "bbb", "cccc" };

            SingleUpload(file);

            return Index(funcionarios);
        }

        private List<Funcionario> Getfuncionario(string fName)
        {
            List<Funcionario> funcionarios = new List<Funcionario>();
            var fileName = $"{Directory.GetCurrentDirectory()}{@"\\wwwroot\files"}" + "\\" + fName;
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            using (var stream = System.IO.File.Open(fileName, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    while (reader.Read())
                    {
                        funcionarios.Add(new Funcionario()
                        {
                            nome = reader.GetValue(0).ToString(),
                            funcao = reader.GetValue(1).ToString(),
                            pais = reader.GetValue(2).ToString(),
                            filial = reader.GetValue(3).ToString(),
                            telefone = reader.GetValue(4).ToString(),
                            celular = reader.GetValue(5).ToString(),
                            email = reader.GetValue(6).ToString()
                        });
                    }
                }
            }
            return funcionarios;
        }

        [Route("SingleUpload")]
        [HttpPost]
        public IActionResult SingleUpload(IFormFile photo)
        {
            var path = Path.Combine(this.webHostEnvironment.WebRootPath,"images" ,photo.FileName);

            var stream = new FileStream(path, FileMode.Create);
            photo.CopyToAsync(stream);
            return View("Index");

        }



        //protected FileStreamResult DownloadFolder(string path, string[] names, int count)
        //{

        //    var contentRootPath = $"{Directory.GetCurrentDirectory()}{@"\\wwwroot\files"}" ;

        //    FileStreamResult fileStreamResult;
        //    var tempPath = Path.Combine(Path.GetTempPath(), "temp.zip");
        //    if (names.Length == 1)
        //    {
        //        path = path.Remove(path.Length - 1);
        //        ZipFile.CreateFromDirectory(path, tempPath, CompressionLevel.Fastest, true);
        //        FileStream fileStreamInput = new FileStream(tempPath, FileMode.Open, FileAccess.Read, FileShare.Delete);
        //        fileStreamResult = new FileStreamResult(fileStreamInput, "APPLICATION/octet-stream");
        //        fileStreamResult.FileDownloadName = names[0] + ".zip";
        //    }
        //    else
        //    {
        //        string extension;
        //        string currentDirectory;
        //        ZipArchiveEntry zipEntry;
        //        ZipArchive archive;
        //        if (count == 0)
        //        {
        //            string directory = Path.GetDirectoryName(path);
        //            string rootFolder = Path.GetDirectoryName(directory);
        //            using (archive = ZipFile.Open(tempPath, ZipArchiveMode.Update))
        //            {
        //                for (var i = 0; i < names.Length; i++)
        //                {
        //                    currentDirectory = Path.Combine(rootFolder, names[i]);
        //                    foreach (var filePath in Directory.GetFiles(currentDirectory, "*.*", SearchOption.AllDirectories))
        //                    {
        //                        zipEntry = archive.CreateEntryFromFile(contentRootPath + "\\" + filePath, names[i] + filePath.Substring(currentDirectory.Length), CompressionLevel.Fastest);
        //                    }
        //                }
        //            }
        //        }
        //        else
        //        {
        //            string lastSelected = names[names.Length - 1];
        //            string selectedExtension = Path.GetExtension(lastSelected);
        //            if (selectedExtension == "")
        //            {
        //                path = Path.GetDirectoryName(Path.GetDirectoryName(path));
        //                path = path.Replace("\\", "/") + "/";

        //            }
        //            using (archive = ZipFile.Open(tempPath, ZipArchiveMode.Update))
        //            {
        //                for (var i = 0; i < names.Length; i++)
        //                {
        //                    extension = Path.GetExtension(names[i]);
        //                    currentDirectory = Path.Combine(path, names[i]);
        //                    if (extension == "")
        //                    {
        //                        foreach (var filePath in Directory.GetFiles(currentDirectory, "*.*", SearchOption.AllDirectories))
        //                        {
        //                            zipEntry = archive.CreateEntryFromFile(contentRootPath + "\\" + filePath, filePath.Substring(path.Length), CompressionLevel.Fastest);
        //                        }
        //                    }
        //                    else
        //                    {
        //                        zipEntry = archive.CreateEntryFromFile(contentRootPath + "\\" + currentDirectory, names[i], CompressionLevel.Fastest);
        //                    }
        //                }
        //            }

        //        }
        //        FileStream fileStreamInput = new FileStream(tempPath, FileMode.Open, FileAccess.Read, FileShare.Delete);
        //        fileStreamResult = new FileStreamResult(fileStreamInput, "APPLICATION/octet-stream");
        //        fileStreamResult.FileDownloadName = "folders.zip";
        //    }
        //    if (System.IO.File.Exists(tempPath))
        //    {
        //        System.IO.File.Delete(tempPath);
        //    }

        //    return fileStreamResult;

        //}


    }
}
