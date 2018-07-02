using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using FPG_Attachment_Uploader.Properties;

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
			}
			catch (Exception e)
			{
				Console.WriteLine(e);
				throw;
			}
		}
	}
}
