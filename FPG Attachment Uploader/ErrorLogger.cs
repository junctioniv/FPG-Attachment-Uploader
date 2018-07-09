using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using FPG_Attachment_Uploader.Properties;

namespace FPG_Attachment_Uploader
{
	public class ErrorLogger
	{
		private static Dictionary<string, List<string>> _errors = new Dictionary<string, List<string>>();

		public static void LogError(string key, string message)
		{
			if (_errors.ContainsKey(key))
			{
				_errors[key].Add(message);
			}
			else
			{
				_errors.Add(key, new List<string> { message });
			}

			Console.WriteLine(message);
		}

		public static void SendErrorEmail(int numInserted, int numFailed)
		{
			using (var client = new SmtpClient
			{
				Port = 587,
				Host = "smtp.gmail.com",
				EnableSsl = true,
				Timeout = 10000,
				DeliveryMethod = SmtpDeliveryMethod.Network,
				UseDefaultCredentials = false,
				Credentials = new NetworkCredential("fpg.error.reporter@gmail.com", "Fr0ntl1n3!"),
			})
			{
				var body =
					$"There were Errors on the last FPG Intacct Attachment Uploader Run. {numInserted} Attachments uploaded succssefully. {numFailed} Attachmetns failed to Upload. Please see the following files that errored:\n\n";

				foreach (var key in _errors.Keys)
				{
					var list = _errors[key];

					foreach (var error in list)
					{
						body += $"{key} - {error} \n";
					}
				}

				using (var message = new MailMessage
				{
					From = new MailAddress("fpg.error.reporter@gmail.com"),
					To = { new MailAddress(Settings.Default.ErrorEmailAddress) },
					Subject = $"FPG Intacct Attachment Uploader Errors - {DateTime.Now:g}",
					Body = body,
					BodyEncoding = Encoding.UTF8,
					DeliveryNotificationOptions = DeliveryNotificationOptions.OnFailure
				})
				{
					client.Send(message);
				}
			}
		}
	}
}
