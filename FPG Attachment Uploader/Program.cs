using System;
using System.Collections.Generic;
using System.Globalization;
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
using NLog;
using NLog.Config;
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
			var intacctExcelPath = string.Empty;
			var concurExcelPath = string.Empty;
			var pdfPath = string.Empty;
			var outputPath = string.Empty;
			var count = 0;
			var numSuccess = 0;
			var numFail = 0;
			IWorkbook concurWorkbook;
			var reports = new List<Report>();
			var reportEntries = new List<ReportEntry>();

			try
			{
				//var list = ConcurClient.GetReportIds();

				while (string.IsNullOrWhiteSpace(intacctExcelPath))
				{
					Console.WriteLine("Press Enter to Select the Intacct Excel Import File:");
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
							intacctExcelPath = dialog.FileName;
							Console.WriteLine($"Path Selected: {intacctExcelPath}");
						}
						else
						{
							Console.WriteLine("The Path you selected was not valid. Please select a valid path.");
						}
					}
				}

				while (string.IsNullOrWhiteSpace(concurExcelPath))
				{
					Console.WriteLine("Press Enter to Select the Concur Excel Import File:");
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
							concurExcelPath = dialog.FileName;
							Console.WriteLine($"Path Selected: {concurExcelPath}");
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

				IWorkbook intacctWorkbook;
				using (var file = new FileStream(intacctExcelPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
				{
					var extension = Path.GetExtension(intacctExcelPath);
					if (extension.Contains("xlsx"))
					{
						intacctWorkbook = new XSSFWorkbook(file);
					}
					else
					{
						intacctWorkbook = new HSSFWorkbook(file);
					}
				}

				using (var file = new FileStream(concurExcelPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
				{
					var extension = Path.GetExtension(intacctExcelPath);
					if (extension.Contains("xlsx"))
					{
						concurWorkbook = new XSSFWorkbook(file);
					}
					else
					{
						concurWorkbook = new HSSFWorkbook(file);
					}
				}

				var sheet = concurWorkbook.GetSheet(concurWorkbook.GetSheetName(intacctWorkbook.ActiveSheetIndex));
				Console.WriteLine("Processing Concur Excel File.");
				for (var i = 1; i <= sheet.LastRowNum; i++)
				{
					var row = sheet.GetRow(i);
					var key = row.GetCell(24)?.NumericCellValue.ToString(CultureInfo.InvariantCulture) ?? string.Empty;
					var reportId = row.GetCell(25)?.StringCellValue ?? string.Empty;
					var vendor = row.GetCell(6)?.StringCellValue ?? string.Empty;
					var memo = row.GetCell(7)?.StringCellValue ?? string.Empty;
					var date = row.GetCell(5)?.DateCellValue ?? new DateTime();
					var amount = row.GetCell(9)?.NumericCellValue ?? 0;

					var entry = new ReportEntry
					{
						Key = key,
						ReportId = reportId,
						VendorName = vendor,
						Memo = memo,
						TransactionDate = date,
						Amount = amount
					};
					 reportEntries.Add(entry);
				}

				sheet = intacctWorkbook.GetSheet(intacctWorkbook.GetSheetName(intacctWorkbook.ActiveSheetIndex));
				var oathToken = ConcurClient.Login(userName, password, clientId);

				Console.WriteLine("Processing Intacct Excel File.");
				for (var i = 4; i <= sheet.LastRowNum; i++)
				{
					var row = sheet.GetRow(i);
					var invoiceNum = "";
					var key = "";
					var report = new Report();
					try
					{
						invoiceNum = row.GetCell(0)?.StringCellValue ?? string.Empty;
						report = reports.SingleOrDefault(q => q.Id == invoiceNum) ?? new Report { Id = invoiceNum, ReportId = ""};

						var keyString = row.GetCell(1)?.StringCellValue ?? string.Empty;
						var keyTokens = keyString.Split('|');

						if (keyTokens.Length <= 1) continue;

						key = keyTokens[0].Trim();

						var entry = new ReportEntry() { Key = key };
						if (key.Length < 11)
						{
							//Find the existing ReportEntry
							entry = reportEntries.SingleOrDefault(q => q.Key == key);
							if (entry == null)
							{
								throw new Exception($"Expense Entry Key#{key} could not be found.");
							}

							report.ReportId = entry.ReportId;
							entry.Type = ReportEntryType.Concur;

							//If a file exists for this key in the pdf folder, we want to add pull from it instead of concur
							if (File.Exists($"{pdfPath}\\{entry.Key}.pdf)"))
							{
								entry.Path = $"{pdfPath}\\{entry.Key}.pdf";
							}
							else
							{
								entry.Image = ConcurClient.GetEntryReceiptImageFromMetaData(entry);
							}	
						}
						else
						{
							entry.Type = ReportEntryType.Egencia;
							entry.Path = $"{pdfPath}\\{entry.Key}.pdf";
						}
						report.Entries.Add(entry);
					}
					catch (Exception e)
					{
						report.HasError = true;
						ErrorLogger.LogError(!string.IsNullOrEmpty(key) ? key : $"Row# {i}", e.Message + e.StackTrace);
					}

					//If We have not stored this Report, store it.
					if (reports.All(q => q.Id != report.Id))
					{
						reports.Add(report);
					}
				}

				Console.WriteLine("Finished Processing Excel files. Generating PDF Reports.");

				var cont = true;
				foreach (var report in reports)
				{
					if (report.HasError)
					{
						cont = false;
						continue;
					}
					Console.WriteLine($"Generating PDF for Invoice#{report.Id}");
					GenerateReportPdf(report, outputPath);
				}

				if (!cont || reports.Any(q => q.HasError))
				{
					ErrorLogger.SendErrorEmail(numSuccess, numFail, reports);
					Console.WriteLine("There were errors with generating the PDFs. Please See Error Email. Press enter to continue and upload and attach all generated PDFs or Close the program to abort.");
					Console.ReadLine();
				}
				else
				{
					Console.WriteLine("All PDFs Generated without errors. Please press enter to upload and attach all generated PDFs.");
					Console.ReadLine();
				}

				var files = Directory.GetFiles(outputPath, "*-*.pdf", SearchOption.TopDirectoryOnly);
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
				if (numFail > 0 || ErrorLogger.HasErrors())
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
			var path = $"{outputDirectory}\\{report.Id} Receipts.pdf";
			try
			{
				using (var outputStream = new FileStream(path, FileMode.Create, FileAccess.ReadWrite))
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
									if (!string.IsNullOrEmpty(entry.Path)) //If we have already set a path for the entry, just pull that file
									{
										var bytes = File.ReadAllBytes(entry.Path);
										if (bytes.Length == 0)
										{
											throw new Exception($"File Was Empty: {entry.Path}");
										}
										merge.AddDocument(new PdfReader(bytes));
									}
									else if (entry.Type == ReportEntryType.Concur)//If we don't have a path and we are a concur receipt, process that image
									{
										if(entry.Image == null) { throw new Exception("No Image Found for this Entry.");}

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
									else //If we don't have a path, and we aren't a concur document, wtf...
									{
										throw new Exception($"No File for: {entry.Path}");
									}

								}
								catch (Exception e)
								{
									report.HasError = true;
									ErrorLogger.LogError(entry.Key, e.Message);
								}
							}

							if (merge.CurrentPageNumber <= 1)
							{
								merge.PageEmpty = true;
								merge.NewPage();
							}
						}
					}
				}
			}
			catch (Exception e)
			{
				report.HasError = true;
				ErrorLogger.LogError(report.Id, e.Message);
			}
			
			if (report.HasError)
			{
				//Lets get rid of our problem child
				File.Delete(path);
			}
		}
	}
}
