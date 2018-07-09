﻿using System;
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

			try
			{
				var oathToken = ConcurClient.Login(userName, password, clientId);
				var list = ConcurClient.GetReportIds();
				
				foreach (var report in list)
				{
					GenerateReportPdf(report, "C:/Users/James/Desktop/pdfs/");
				}
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
