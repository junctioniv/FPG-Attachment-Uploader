using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Concur.Authentication;

namespace FPG_Attachment_Uploader
{
	class Program
	{
		static void Main(string[] args)
		{
			var authService = new AuthenticationService();
			var oauthDetail = authService.GetOAuthTokenAsync(loginId, password, clientId).Result;

			serviceV3 = new Concur.Connect.V3.ConnectService(oauthDetail.AccessToken, oauthDetail.InstanceUrl);
			serviceV1 = new Concur.Connect.V1.ConnectService(oauthDetail.AccessToken, oauthDetail.InstanceUrl);
			return oauthDetail;
		}
	}
}
