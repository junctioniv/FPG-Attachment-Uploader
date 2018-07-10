using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using FPG_Attachment_Uploader.Properties;
using Intacct.SDK;
using Intacct.SDK.Functions.AccountsReceivable;
using Intacct.SDK.Functions.Common;
using Intacct.SDK.Functions.Common.Query.Comparison.EqualTo;
using Intacct.SDK.Functions.Common.Query.Logical;
using Intacct.SDK.Functions.Company;
using Intacct.SDK.Functions.OrderEntry;
using NLog;
using NLog.Config;
using NLog.Targets;


namespace FPG_Attachment_Uploader
{
	class IntacctClient
	{

		public static FileTarget Target = new FileTarget
		{
			FileName = "$C:/Intacct/logs/intacct.log"
		};

		public static OnlineClient Client = new OnlineClient(ClientConfig);

		private static ClientConfig _clientConfig = null;
		private static ClientConfig ClientConfig  // Create the config and set a session ID
		{
			get
			{
				if (_clientConfig == null)
				{
					SimpleConfigurator.ConfigureForTargetLogging(Target, LogLevel.Debug);
					var logger = LogManager.GetLogger("intacct-sdk-net");
					_clientConfig = new ClientConfig
					{
						SenderId = Properties.Intacct.Default.SenderId,
						SenderPassword = Properties.Intacct.Default.SenderPassword,
						CompanyId = Properties.Intacct.Default.CompanyId,
						UserId = Properties.Intacct.Default.UserId,
						UserPassword = Properties.Intacct.Default.UserPassword,
						LocationId = Properties.Intacct.Default.LocationId,
						Logger = logger,
						LogLevel = LogLevel.FromString(Properties.Intacct.Default.LogLevel)
					};
				}

				return _clientConfig;
			}	
		}

		public static bool InvoiceExists(string invoiceNum, out System.Xml.Linq.XElement invoice)
		{
			invoice = null; //Set this justincase the function explodes
			try
			{
				var query = new ReadByQuery
				{
					ObjectName = "ARINVOICE",
					Fields = { "*" },
					Query = new AndCondition { Conditions = { new EqualToString { Field = "RECORDID", Value = invoiceNum }, new EqualToString { Field = "MODULEKEY", Value = "8.SO" } } }
					//Query = new EqualToString { Field = "RECORDID", Value = invoiceNum }
				};

				var task = Client.Execute(query);
				task.Wait();
				var response = task.Result;

				if (response.Results.Count == 0 || response.Results[0].Data.Count == 0)
				{
					ErrorLogger.LogError(invoiceNum, $"No Invoice Exists with Invoice#: {invoiceNum}");
					return false;
				}

				invoice = response.Results[0]?.Data[0];
			}
			catch (Exception e)
			{
				ErrorLogger.LogError(invoiceNum, $"Getting Invoice# {invoiceNum} Errored: {e.Message}");
				return false;
			}

			return true;
		}

		public static bool GetInvoice(int invoiceRecordNo, string invoiceNum, out System.Xml.Linq.XElement invoice)
		{
			invoice = null; //Set this justincase the function explodes
			try
			{
				var query = new ReadByQuery
				{
					ObjectName = "SODOCUMENT",
					Fields = { "*" },
					Query = new EqualToString { Field = "PRRECORDKEY", Value = invoiceRecordNo.ToString() }
				};

				var task = Client.Execute(query);
				task.Wait();
				var response = task.Result;

				if (response.Results.Count == 0)
				{
					ErrorLogger.LogError(invoiceNum, $"No Invoice Exists with RECORDNO: {invoiceRecordNo}");
					return false;
				}

				//invoice = response.Results[0]?.Data[0];
				var recordNo = response.Results[0]?.Data[0]?.Element("RECORDNO")?.Value;

				if (recordNo == null)
				{
					ErrorLogger.LogError(invoiceNum, "Cannot Find Sales Invoice.");
					return false;
				}

				var invoiceQuery = new Read
				{
					ObjectName = "SODOCUMENT",
					Fields = { "*" },
					Keys = { Convert.ToInt32(recordNo) },
					DocParId = "Sales Invoice"
				};

				var invoiceTask = Client.Execute(invoiceQuery);
				invoiceTask.Wait();
				var invoiceResponse = invoiceTask.Result;

				if (invoiceResponse.Results.Count == 0)
				{
					ErrorLogger.LogError(invoiceNum, $"No Sales Invoice Exists with RECORDNO: {invoiceRecordNo}");
					return false;
				}

				invoice = invoiceResponse.Results[0]?.Data[0];
			}
			catch (Exception e)
			{
				ErrorLogger.LogError(invoiceNum, $"Getting Invoice with RECORDNO: {invoiceRecordNo} Errored: {e.Message}");
				return false;
			}

			return true;
		}

		public static bool UploadAttachment(string attachmentsId, string attachmentsName, string invoiceNum, string path)
		{
			try
			{
				var create = new AttachmentsCreate
				{
					AttachmentsId = attachmentsId,
					AttachmentFolderName = "Default",
					AttachmentsName = attachmentsName,
					Files = { new AttachmentFile { Extension = "pdf", FileName = attachmentsName, FilePath = path } }
				};

				var task = Client.Execute(create);
				task.Wait();

				var response = task.Result;
				var result = response.Results[0];
				if (result is null || result.Status != "success")
				{
					ErrorLogger.LogError(invoiceNum, "Document Upload Failed. Result was null.");
					return false;
				}

				var attachmentId = result.Key;
				if (attachmentId is null)
				{
					ErrorLogger.LogError(invoiceNum, "AttachmentId was null. Upload failed.");
					return false;
				}
			}
			catch (Exception e)
			{
				ErrorLogger.LogError(invoiceNum, $"Uploading {attachmentsName} Errored: {e.Message}");
				return false;
			}

			return true;
		}

		public static bool UpdateInvoice(string doucmentId, string attachmentsId, string invoiceNum, System.Xml.Linq.XElement invoiceLineItem)
		{
			var lineNo = Convert.ToInt32(invoiceLineItem?.Element("LINE_NUM")?.Value ?? "1");
			var memo = invoiceLineItem?.Element("MEMO")?.Value ?? string.Empty;
			var itemid = invoiceLineItem?.Element("ITEMID")?.Value ?? string.Empty;
			var quantity = Convert.ToInt32(invoiceLineItem?.Element("QUANTITY")?.Value ?? "1");
			try
			{
				var updateInvoice = new OrderEntryTransactionUpdate
				{
					TransactionId = doucmentId,
					AttachmentsId = attachmentsId,
					Lines =
						new List<AbstractOrderEntryTransactionLine>
						{
							new OrderEntryTransactionLineUpdate {LineNo = lineNo, Memo = memo, ItemId = itemid, Quantity = quantity}
						} //A line Item update is required to call this function. Dumb...
				};
				var task = Client.Execute(updateInvoice);
				task.Wait();

				var response = task.Result;
				var result = response.Results[0];
				if (result is null || result.Status != "success")
				{
					ErrorLogger.LogError(invoiceNum, "Dailed to update Invoice.");
					return false;
				}
			}
			catch (Exception e)
			{
				ErrorLogger.LogError(invoiceNum, $"Updating Invoice #{invoiceNum} Errored: {e.Message}");
				return false;
			}

			return true;
		}

		public static bool UploadAttachment(string path)
		{
			try
			{
				var filename = Path.GetFileName(path);
				if (string.IsNullOrEmpty(filename))
				{
					return false;
				}

				var invoiceNum = filename.Split(' ')[0];
				var attachmentsId = $"Att {invoiceNum}";
				var attachmentsName = filename.Split('.')[0];

				var bytes = File.ReadAllBytes(path);
				if (bytes.Length == 0)
				{
					ErrorLogger.LogError(filename, "File Was Empty");
					return false;
				}

				//Check to see if the Invoice Already Exists, and grab the invoiceRecordNo if it does.
				if (!IntacctClient.InvoiceExists(invoiceNum, out var invoice)) return false;
				var invoiceRecordNo = invoice.Element("RECORDNO")?.Value;

				if (invoiceRecordNo == null)
				{
					return false;
				}

				if (!IntacctClient.GetInvoice(Convert.ToInt32(invoiceRecordNo), invoiceNum, out invoice)) return false;
				var invoiceLineItem = invoice.Element("SODOCUMENTENTRIES")?.Element("sodocumententry");
				var salesInvoiceDocumentId = invoice.Element("DOCID")?.Value;

				//Upload the Attachments
				if (!IntacctClient.UploadAttachment(attachmentsId, attachmentsName, invoiceNum, path)) return false;

				//Update the Invoice to have the new Attachment
				if (!IntacctClient.UpdateInvoice(salesInvoiceDocumentId, attachmentsId, invoiceNum, invoiceLineItem))
					return false;

				//Move the succesfully Attached Invoice to the "Attached" Directory
				File.WriteAllBytes($"{path}\\Attached\\{filename}", bytes);
				File.Delete(path);
			}
			catch (Exception e)
			{
				Console.WriteLine($"There was an error processing File:{path} - {e}");
				return false;
			}
			return true;
		}
	}
}
