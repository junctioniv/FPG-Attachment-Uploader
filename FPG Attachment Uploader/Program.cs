using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using FPG_Attachment_Uploader.Properties;
using iTextSharp.text;
using iTextSharp.text.pdf;
using Microsoft.WindowsAPICodePack.Dialogs;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace FPG_Attachment_Uploader
{
	class Program
	{
		[STAThread]
		static void Main(string[] args)
		{
			var userName = Settings.Default.UserName;
			var password = Settings.Default.Password;
			var clientId = Settings.Default.ClientId;
			var excelPath = string.Empty;
			var pdfPath = string.Empty;
			var outputPath = string.Empty;
			var count = 0;
			var numSuccess = 0;
			var numFail = 0;
			IWorkbook workbook;
			List<Report> reports = new List<Report>();

			try
			{
				//var list = ConcurClient.GetReportIds();

				while (string.IsNullOrWhiteSpace(excelPath))
				{
					Console.WriteLine("Press Enter to Select the Excel Import File:");
					Console.ReadLine();

					using (var dialog = new CommonOpenFileDialog
					{
						Title = "Select Excel Document to Process",
						InitialDirectory = "C:\\",
						EnsurePathExists = true,
						Multiselect = false
					})
					{
						var result = dialog.ShowDialog();

						if (result == CommonFileDialogResult.Ok && !string.IsNullOrWhiteSpace(dialog.FileName))
						{
							excelPath = dialog.FileName;
							Console.WriteLine($"Path Selected: {excelPath}");
						}
						else
						{
							Console.WriteLine("The Path you selected was not valid. Please select a valid path.");
						}
					}
				}

				while (string.IsNullOrWhiteSpace(pdfPath))
				{
					Console.WriteLine("Press Enter to Select Egencia Pdfs Folder:");
					Console.ReadLine();

					using (var dialog = new CommonOpenFileDialog
					{
						Title = "Select a Folder to Process",
						InitialDirectory = "C:\\",
						EnsurePathExists = true,
						IsFolderPicker = true,
						Multiselect = false
					})
					{
						var result = dialog.ShowDialog();

						if (result == CommonFileDialogResult.Ok && !string.IsNullOrWhiteSpace(dialog.FileName))
						{
							pdfPath = dialog.FileName;
							Console.WriteLine($"Path Selected: {pdfPath}");

							Directory.CreateDirectory(pdfPath + "\\Attached");
						}
						else
						{
							Console.WriteLine("The Path you selected was not valid. Please select a valid path.");
						}
					}
				}

				while (string.IsNullOrWhiteSpace(outputPath))
				{
					Console.WriteLine("Press Enter to Select the PDF output direcotry:");
					Console.ReadLine();

					using (var dialog = new CommonOpenFileDialog
					{
						Title = "Select a Folder to Process",
						InitialDirectory = "C:\\",
						EnsurePathExists = true,
						IsFolderPicker = true,
						Multiselect = false
					})
					{
						var result = dialog.ShowDialog();

						if (result == CommonFileDialogResult.Ok && !string.IsNullOrWhiteSpace(dialog.FileName))
						{
							outputPath = dialog.FileName;
							Console.WriteLine($"Path Selected: {outputPath}");

							Directory.CreateDirectory(outputPath + "\\Attached");
						}
						else
						{
							Console.WriteLine("The Path you selected was not valid. Please select a valid path.");
						}
					}
				}

				using (var file = new FileStream(excelPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
				{
					var extension = Path.GetExtension(excelPath);
					if (extension.Contains("xlsx"))
					{
						workbook = new XSSFWorkbook(file);
					}
					else
					{
						workbook = new HSSFWorkbook(file);
					}
				}

				var sheet = workbook.GetSheet(workbook.GetSheetName(workbook.ActiveSheetIndex));
				var oathToken = ConcurClient.Login(userName, password, clientId);
				for (var i = 1; i <= sheet.LastRowNum; i++)
				{
					var row = sheet.GetRow(i);
					var invoiceNum = row.GetCell(0)?.StringCellValue ?? string.Empty;
					var reportId = row.GetCell(2)?.StringCellValue ?? string.Empty;
					var idString = row.GetCell(1)?.StringCellValue?? string.Empty;
					var idStringTokens = idString.Split('|');

					if(idStringTokens.Length <= 1) continue;

					var report = reports.SingleOrDefault(q => q.ReportId == reportId) ?? new Report{ReportId = reportId, Id = invoiceNum};
					var entryId = idStringTokens[0].Trim();
					var entry = new ReportEntry{Id = entryId};
					if (entry.Id.Length >= 36)
					{
						entry.Type = ReportEntryType.Concur;
						entry.Image = ConcurClient.GetReciptImageByEntryId(entryId);
					}
					else
					{
						entry.Type = ReportEntryType.Egencia;
						entry.Path = $"{pdfPath}/{entry.Id}.pdf";
					}

					report.Entries.Add(entry);
					if (reports.All(q => q.ReportId != reportId))
					{
						reports.Add(report);
					}
				}

				foreach (var report in reports)
				{
					GenerateReportPdf(report, outputPath);
				}
				
				var files = Directory.GetFiles(outputPath, "* *.pdf", SearchOption.TopDirectoryOnly);
				Console.WriteLine($"{files.Length} Files to Process.");

				foreach (var file in files)
				{
					count++;

					var filename = Path.GetFileName(file);
					if (string.IsNullOrEmpty(filename))
					{
						ErrorLogger.LogError(filename, $"Processing PDF {count}/{files.Length} - File was empty");
						continue;
					}

					var invoiceNum = filename.Split(' ')[0];

					Console.WriteLine($"Processing PDF {count}/{files.Length} - Invoice# {invoiceNum}");

					if (!IntacctClient.UploadAttachment(file)) continue;

					numSuccess++;
				}

				numFail = count - numSuccess;
				if (numFail > 0)
				{
					ErrorLogger.SendErrorEmail(numSuccess, numFail, reports);
				}

				Console.WriteLine($"Processing Finished. {numSuccess} Attachments Uploaded. {numFail} Attachments Failed. Press any Enter to Close.");
				Console.ReadLine();
			}
			catch (Exception e)
			{
				Console.WriteLine(e);
			}
		}

		private static void GenerateReportPdf(Report report, string outputDirectory)
		{
			var path = $"{outputDirectory}/{report.Id}.pdf";
			using (var outputStream = new FileStream(path, FileMode.Create))
			{
				using (var doc = new Document())
				{
					using (var merge = new PdfSmartCopy(doc, outputStream))
					{
						doc.Open();
						foreach (var entry in report.Entries)
						{
							try
							{
								if (entry.Type == ReportEntryType.Concur)
								{
									var receipt = entry.Image;
									ConcurClient.DownloadImage(receipt);
									switch (receipt.ContentType)
									{
										case "pdf":
											merge.AddDocument(new PdfReader(receipt.Data));
											break;
										case "image":
										default:
											using (var output = new MemoryStream())
											{
												using (var document = new Document())
												{
													PdfWriter.GetInstance(document, output);
													document.Open();

													var image = Image.GetInstance(receipt.Data);
													image.ScaleToFit(document.PageSize);
													document.Add(image);
													document.NewPage();
												}
												var bytes = output.ToArray();
												merge.AddDocument(new PdfReader(bytes));
											}
											break;
									}
								}
								else
								{
									var bytes = File.ReadAllBytes(entry.Path);
									if (bytes.Length == 0)
									{
										ErrorLogger.LogError(entry.Path, "File Was Empty");
										report.HasError = true;
										break;
									}
									merge.AddDocument(new PdfReader(bytes));
								}
								
							}
							catch (Exception e)
							{
								report.HasError = true;
								Console.WriteLine(e);
								break;
							}
						}
					}
				}
			}

			if (report.HasError)
			{
				//Lets get rid of our problem child
				File.Delete(path);
			}
		}
	}
}
