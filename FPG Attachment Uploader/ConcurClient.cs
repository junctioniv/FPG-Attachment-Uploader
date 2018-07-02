using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;
using FPG_Attachment_Uploader.Properties;
using Newtonsoft.Json;

namespace FPG_Attachment_Uploader
{
	public class ConcurClient
	{
		private static readonly string ClientSecret = Settings.Default.ClientSecret;

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

				return oathToken;
			}
		}
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
