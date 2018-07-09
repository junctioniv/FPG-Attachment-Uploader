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

namespace FPG_Attachment_Uploader
{
	class Program
	{
		static void Main(string[] args)
		{
			var userName = Settings.Default.UserName;
			var password = Settings.Default.Password;
			var clientId = Settings.Default.ClientId;
			var path = "C:/Users/James/Desktop/pdfs/";


			try
			{
				var oathToken = ConcurClient.Login(userName, password, clientId);
				var list = ConcurClient.GetReportIds();
				
				foreach (var report in list)
				{
					GenerateReportPdf(report, path);
				}



				var count = 0;
				var numSuccess = 0;
				var numFail = 0;
				
				var files = Directory.GetFiles(path, "* *.pdf", SearchOption.TopDirectoryOnly);
				Console.WriteLine($"{files.Length} Files to Process.");

				foreach (var file in files)
				{
					count++;
					try
					{
						var filename = Path.GetFileName(file);
						if (string.IsNullOrEmpty(filename))
						{
							ErrorLogger.LogError(filename, $"Processing PDF {count}/{files.Length} - File was empty");
							continue;
						}

						var invoiceNum = filename.Split(' ')[0];
						var attachmentsId = $"Att {invoiceNum}";
						var attachmentsName = filename.Split('.')[0];

						Console.WriteLine($"Processing PDF {count}/{files.Length} - Invoice# {invoiceNum}");

						var bytes = File.ReadAllBytes(file);
						if (bytes.Length == 0)
						{
							ErrorLogger.LogError(filename, "File Was Empty");
							continue;
						}

						//Check to see if the Invoice Already Exists, and grab the invoiceRecordNo if it does.
						if (!IntacctClient.InvoiceExists(invoiceNum, out var invoice)) continue;
						var invoiceRecordNo = invoice.Element("RECORDNO")?.Value;

						if (invoiceRecordNo == null)
						{
							continue;
						}

						if (!IntacctClient.GetInvoice(Convert.ToInt32(invoiceRecordNo), invoiceNum, out invoice)) continue;
						var invoiceLineItem = invoice.Element("SODOCUMENTENTRIES")?.Element("sodocumententry");
						var salesInvoiceDocumentId = invoice.Element("DOCID")?.Value;

						//Upload the Attachments
						if (!IntacctClient.UploadAttachment(attachmentsId, attachmentsName, invoiceNum, file)) continue;

						//Update the Invoice to have the new Attachment
						if (!IntacctClient.UpdateInvoice(salesInvoiceDocumentId, attachmentsId, invoiceNum, invoiceLineItem)) continue;


						//Move the succesfully Attached Invoice to the "Attached" Directory
						File.WriteAllBytes($"{path}\\Attached\\{filename}", bytes);
						File.Delete(file);
					}
					catch (Exception e)
					{
						Console.WriteLine($"There was an error processing File:{file} - {e}");
						continue;
					}

					numSuccess++;
				}

				numFail = count - numSuccess;
				if (numFail > 0)
				{
					ErrorLogger.SendErrorEmail(numSuccess, numFail);
				}

				Console.WriteLine($"Processing Finished. {numSuccess} Attachments Uploaded. {numFail} Attachments Failed. Press any Enter to Close.");
				Console.ReadLine();


			}
			catch (Exception e)
			{
				Console.WriteLine(e);
			}
		}

		private static void GenerateReportPdf(KeyValuePair<string, List<ReceiptImage>> report,string outputDirectory)
		{
			using (var outputStream = new FileStream($"{outputDirectory}/{report.Key}.pdf", FileMode.Create))
			{
				using (var doc = new Document())
				{
					using (var merge = new PdfSmartCopy(doc, outputStream))
					{
						doc.Open();
						foreach (var receipt in report.Value)
						{
							try
							{
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
							catch (Exception e)
							{
								Console.WriteLine(e);
							}
						}
					}
				}

			}
		}
	}
}
