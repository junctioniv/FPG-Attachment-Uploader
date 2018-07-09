﻿using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Runtime.Remoting.Messaging;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;
using FPG_Attachment_Uploader.Properties;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Drawing;
using iTextSharp.text;

namespace FPG_Attachment_Uploader
{
	public class ConcurClient
	{
		private static readonly string ClientSecret = Settings.Default.ClientSecret;
		private static OAuthTokenDetail AuthToken = null;

		/// <summary>
		/// This facade method shows how to use the SDK to obtain an oauth token (with some extra details) out of given LoginID/Password/ClientID parameters.
		/// It also shows how you can instantiate service objects for versions 1.0 asnd 3.0 of Concur Platform API.
		/// </summary>
		/// <param name="loginId">This is the Email/UserName entered by the user when he/she signin at https://developer.concur.com/ </param>
		/// <param name="password">This is the Password used by the user when he/she signin at https://developer.concur.com/ </param>
		/// <param name="clientId">This is the oauth client id (required to identify the application) and needed to generate an oauth token </param>
		/// <returns>OAuthTokenDetail object, which encapsulates details of the oauth token obtained for the current user</returns>
		public static OAuthTokenDetail Login(string loginId, string password, string clientId)
		{
			//TODO: Add Code to reuse Oauth token once we have it. Also we want to run all the calls through this class.
			if (AuthToken != null)
			{
				return AuthToken;
			}

			using (var client = new HttpClient())
			{
				var pairs = new List<KeyValuePair<string, string>>
				{
					new KeyValuePair<string, string>("client_id", clientId),
					new KeyValuePair<string, string>("client_secret", ClientSecret),
					new KeyValuePair<string, string>("grant_type", "password"),
					new KeyValuePair<string, string>("username", loginId),
					new KeyValuePair<string, string>("password", password)
				};
				var content = new FormUrlEncodedContent(pairs);

				var request = new HttpRequestMessage
				{
					RequestUri = new Uri("https://us.api.concursolutions.com/oauth2/v0/token"),
					Method = HttpMethod.Post,
					Content = content
				};

				var response = client.SendAsync(request).Result;

				if (!response.IsSuccessStatusCode)
				{
					//TODO: Fix this case
					throw new Exception("Could Not Login");
				}
				var json = response.Content.ReadAsStringAsync().Result;
				var oathToken = JsonConvert.DeserializeObject<OAuthTokenDetail>(json);

				return AuthToken = oathToken;
			}
		}

		public static string Call(string url, List<KeyValuePair<string, string>> content = null)
		{
			using (var client = new HttpClient())
			{
				var request = new HttpRequestMessage
				{
					RequestUri = new Uri(AuthToken.Geolocation + url),
					Method = HttpMethod.Get,
				};

				if (content != null)
				{
					request.Content = new FormUrlEncodedContent(content);
				}

				//Add Athentication header in correspondence with https://developer.concur.com/api-reference/authentication/getting-started.html
				request.Headers.Add("Authorization", $"{AuthToken.TokenType} {AuthToken.AccessToken}");
				request.Headers.Add("Accept", "application/json");

				var response = client.SendAsync(request).Result;

				if (!response.IsSuccessStatusCode)
				{
					//TODO: Fix this case
					throw new Exception("Could Not Login");
				}

				var json = response.Content.ReadAsStringAsync().Result;
				return json;
			} 
		}

		public static ReceiptImage GetReciptImageByEntryId(string id)
		{
			var json = Call($"/api/image/v1.0/expenseentry/{id}");
			var image = JsonConvert.DeserializeObject<ReceiptImage>(json);

			return image;
		}

		public static Dictionary<string, List<ReceiptImage>> GetReportIds()
		{
			var list = new Dictionary<string, List<ReceiptImage>>();
			var json = Call($"/api/v3.0/expense/reports?user=all&userDefinedDateAfter={DateTime.Today.AddDays(-60):yyyy-MM-dd}&limit=1");
			var items = JsonConvert.DeserializeObject<JObject>(json)["Items"];

			var total = items.Children().Count();
			var count = 1;
			foreach (var obj in items)
			{
				Console.WriteLine($"{DateTime.Now:G}	Processing {count++}/{total} - {obj["Name"]}");
				var id = (string) obj["ID"];
				json = Call($"/api/expense/expensereport/v2.0/report/{id}");

				var expenseEntriesList = JsonConvert.DeserializeObject<JObject>(json)["ExpenseEntriesList"];
				foreach (var entry in expenseEntriesList)
				{
					var entryImageId = entry["EntryImageID"]?.ToString() ?? "";
					var entryId = entry["ReportEntryID"]?.ToString() ?? "";

					if (string.IsNullOrEmpty(entryImageId)) continue;

					var image = GetReciptImageByEntryId(entryId);

					if (list.ContainsKey(id))
					{
						list[id].Add(image);
					}
					else
					{
						list.Add(id, new List<ReceiptImage>{image});
					}
				}
			}
			return list;
		}

		public static ReceiptImage DownloadImage(ReceiptImage image)
		{
			using(var client = new HttpClient())
			{
				var response = client.GetAsync(image.ImageUrl).Result;
				var filetype = response.Content.Headers.ContentType.MediaType;
				var data = response.Content.ReadAsByteArrayAsync().Result;

				image.ContentType = filetype.StartsWith("image") ? "image" : "pdf";
				image.Data = data;
			}

			return image;
		}
	}
	
	public class ReceiptImage
	{
		[JsonProperty("Id")]
		public string Id;

		[JsonProperty("Url")]
		public string ImageUrl;

		public string ContentType = "";
		public byte[] Data = {};
	}

	public class OAuthTokenDetail
	{
		[JsonProperty("access_token")]
		public string AccessToken;

		[JsonProperty("geolocation")]
		public string Geolocation;

		[JsonProperty("expires_in")]
		public string ExpirationDate;

		[JsonProperty("refresh_token")]
		public string RefreshToken;

		[JsonProperty("scope")]
		public string Scope;

		[JsonProperty("token_type")]
		public string TokenType;

		[JsonProperty("id_token")]
		public string IdToken;
	}
}
