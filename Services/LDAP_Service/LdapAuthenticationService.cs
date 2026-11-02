using OYA_Web_Application.Models.Services;
using OYA_Web_Application.Services.KeyVault_Service;
using System.DirectoryServices.Protocols;
using System.Text;

namespace OYA_Web_Application.Services.LDAP_Service
{
	public class LdapAuthenticationService
	{
		public LdapResponse ValidateCredentials(string ldapServer, string ldapUsername, string ldapPassword)
		{
			// LDAP server details
			//string ldapServer = "DC19.heliasllc.local";
			// string ldapServer = "DC-X220.heliasllc.local";
			//string ldapServer = "azrwusprddc01.heliasllc.local";

			// LDAP admin connection credentials
			// string ldapUsername = "heliasadmin";
			// string ldapPassword = "Hel@@dmin!16";

			// LDAP User connection credentials
			//string ldapUsername = "testuser@helias.com";
			//string ldapPassword = "welcome@123";

			int ldapPort = 389; // Default LDAP port

			try
			{
				// Create LDAP connection
				using (LdapConnection connection = new LdapConnection(ldapServer + ":" + ldapPort))
				{
					connection.SessionOptions.ProtocolVersion = 3;

					// Bind with user credentials
					connection.AuthType = AuthType.Basic;
					connection.Credential = new System.Net.NetworkCredential(ldapUsername, ldapPassword);
					connection.Bind();
					return GetUsernameFromAD(connection, ldapUsername);
				}
			}
			catch (LdapException ex)
			{
				Console.WriteLine($"Authentication failed: {ex.Message}");
				return new LdapResponse();
			}
			catch (Exception ex)
			{
				return new LdapResponse();
			}
		}

		private LdapResponse GetUsernameFromAD(LdapConnection ldapConnection, string emailAddress)
		{
			try
			{
				var containerDn = "DC=heliasllc,DC=local";
				var user = "test user";
				var searchUser = new SearchRequest(
				containerDn,
				$"(&(objectClass=user)(userPrincipalName={emailAddress}))", // Filter for the specific email address
				SearchScope.Subtree, // Search only within the specified container, not recursively
				null // Retrieve all attributes
				);

				var searchResponseUser = (SearchResponse)ldapConnection.SendRequest(searchUser);

				if (searchResponseUser.Entries.Count == 1)
				{
					var userEntry = searchResponseUser.Entries[0];
					var cn = userEntry.Attributes["cn"]?[0]?.ToString();
					var userPrincipalName = userEntry.Attributes["userPrincipalName"]?[0]?.ToString();
					var sAMAccountName = userEntry.Attributes["sAMAccountName"]?[0]?.ToString();

					var memberOf = userEntry.Attributes["memberOf"];
					AzureKeyVaultService service = new AzureKeyVaultService();
					string groupName = service.RetrieveSecret("ADGroup-name");
					bool isAuthorizeduser = false;
					if (memberOf != null)
					{
						foreach (byte[] group in memberOf)
						{
							var groupDn = Encoding.UTF8.GetString(group);
							// var groupDn = group.ToString();
							Console.WriteLine(groupDn);
							Console.WriteLine(group.ToString());
							if (groupDn.ToString().Contains(groupName))
							{
								isAuthorizeduser = true;
								//return true;
							}
						}
					}

					// Add more attributes as needed
					LdapResponse ldapResponse = new LdapResponse
					{
						isAuthenticated = true,
						isAuthorizeduser = isAuthorizeduser,
						cn = cn,
						UserPrincipalName = userPrincipalName,
						SAMAccountName = sAMAccountName,
					};
					return ldapResponse;
				}
				return new LdapResponse();
			}
			catch (Exception ex)
			{
				Console.WriteLine($"Error: {ex.Message}");
				return new LdapResponse();
			}
		}
	}
}
