using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.IO;

/// <summary>
/// Defines TeamDynamix REST API helper functions. These include 
/// getting an API client, an authentication helper and a
/// submit file helper. This class relies on the Microsoft.AspNet.WebApi.Client
/// (HttpClient) Nuget package.
/// </summary>
class TDApiHelpers
{

	/// <summary>
	/// Gets the an HttpClient with the proper settings for interacting with the
	/// TeamDynamix Web API.
	/// </summary>
	/// <param name="apiBaseUrl">The base URL for the TeamDynamix REST API.
	/// This should be in the format [your TeamDynamix domain]/TDWebApi/.</param>
	/// <returns></returns>
	public static HttpClient GetTDApiClient(string apiBaseUrl)
	{

	  // Initialize a RESTful API client.
	  HttpClient apiClient = new HttpClient();
	  apiClient.Timeout = new TimeSpan(0, 2, 0); // 2 minute timeout.
	  apiClient.BaseAddress = new Uri(apiBaseUrl.Trim().TrimEnd('/', '\\') + "/");

	  // Add JSON return types.
	  apiClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
	  apiClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("text/json"));

	  // Return the API client
	  return apiClient;

	}

	/// <summary>
	/// Calls the /TDWebApi/api/auth/loginadmin endpoint with the specified 
	/// TeamDynamix Web API client and web service settings from TDAdmin.
	/// </summary>
	/// <param name="apiClient">The API client.</param>
	/// <param name="apiBeid">The Web Services BEID from TDAdmin.</param>
	/// <param name="apiWebServicesKey">The Web Services Key from TDAdmin.</param>
	/// <returns></returns>
	public static bool AdminAuthenticateAgainstTDApi(
	  HttpClient apiClient,
	  string apiBeid,
	  string apiWebServicesKey)
	{

	  // Initialize an admin login object.
	  var loginParams = new {
		BEID = apiBeid,
		WebServicesKey = apiWebServicesKey
	  };

	  // Post to the api/auth/login endpoint
	  HttpResponseMessage result = apiClient.PostAsJsonAsync("api/auth/loginadmin", loginParams).Result;

	  // If our call failed, log why and exit out.
	  if (!result.IsSuccessStatusCode)
	  {

		// Log an error, the status code, the error phrase and the error message.
		// Then exit out and try again next time.
		Program.Log.ErrorFormat($"Error authenticating against the TeamDynamix Web API. See error details below.");
		Program.Log.ErrorFormat($"Authenticate: Status Code: {result.StatusCode.ToString()}");
		Program.Log.ErrorFormat($"Authenticate: Error Phrase: {result.ReasonPhrase}");
		Program.Log.ErrorFormat($"Authenticate: Error message: {result.Content.ReadAsStringAsync().Result}");
		return false;

	  }

	  // We think our call succeeded. Let's try to get the auth token out of it now.
	  Program.Log.InfoFormat($"Successfully authenticated against the TeamDynamix Web API.");
	  string authToken = result.Content.ReadAsStringAsync().Result;

	  // If we don't actually have an auth token, something screwed up. Log it and abort.
	  if (string.IsNullOrWhiteSpace(authToken))
	  {

		Program.Log.ErrorFormat($"Authentication to the TeamDynamix Web API succeeded but a token was unable to be obtained." +
		  " The application cannot proceed without an authorization token. Will try again in the next processing loop.");

		return false;

	  }

	  // Since we got a non-empty token, set the API client Authorization header
	  // to "Bearer [token]".
	  apiClient.DefaultRequestHeaders.Authorization =
		  new AuthenticationHeaderValue("Bearer", authToken);

	  // If we got to here, return true.
	  return true;

	}

	/// <summary>
	/// Calls the /TDWebApi/api/auth/login endpoint with the specified 
	/// TeamDynamix Web API client and credentials for a named user account.
	/// </summary>
	/// <param name="apiClient">The API client.</param>
	/// <param name="username">The username.</param>
	/// <param name="password">The password.</param>
	/// <returns></returns>
	public static bool UserAuthenticateAgainstTDApi(
	  HttpClient apiClient,
	  string username,
	  string password)
	{

	  // Initialize an admin login object.
	  var loginParams = new {
		UserName = username,
		Password = password
	  };

	  // Post to the api/auth/login endpoint
	  HttpResponseMessage result = apiClient.PostAsJsonAsync("api/auth/login", loginParams).Result;

	  // If our call failed, log why and exit out.
	  if (!result.IsSuccessStatusCode)
	  {

		// Log an error, the status code, the error phrase and the error message.
		// Then exit out and try again next time.
		Program.Log.ErrorFormat($"Error authenticating against the TeamDynamix Web API. See error details below.");
		Program.Log.ErrorFormat($"Authenticate: Status Code: {result.StatusCode.ToString()}");
		Program.Log.ErrorFormat($"Authenticate: Error Phrase: {result.ReasonPhrase}");
		Program.Log.ErrorFormat($"Authenticate: Error message: {result.Content.ReadAsStringAsync().Result}");
		return false;

	  }

	  // We think our call succeeded. Let's try to get the auth token out of it now.
	  Program.Log.InfoFormat($"Successfully authenticated against the TeamDynamix Web API.");
	  string authToken = result.Content.ReadAsStringAsync().Result;

	  // If we don't actually have an auth token, something screwed up. Log it and abort.
	  if (string.IsNullOrWhiteSpace(authToken))
	  {

		Program.Log.ErrorFormat($"Authentication to the TeamDynamix Web API succeeded but a token was unable to be obtained." +
		  " The application cannot proceed without an authorization token. Will try again in the next processing loop.");

		return false;

	  }

	  // Since we got a non-empty token, set the API client Authorization header
	  // to "Bearer [token]".
	  apiClient.DefaultRequestHeaders.Authorization =
		  new AuthenticationHeaderValue("Bearer", authToken);

	  // If we got to here, return true.
	  return true;

	}

	/// <summary>
	/// Calls the /TDWebApi/api/people/import endpoint with the specified 
	/// TeamDynamix Web API client and file to upload. This code could be 
	/// to point at any file upload endpoint in the TeamDynamix REST API though.
	/// All TeamDynamix file uploads will be multi-part form post. This particular
	/// example also includes how to add a second string form part named notifyEmailAddress
	/// because the people import endpoint accepts it, but that would not be required in
	/// more generic file-only upload endpoints.
	/// </summary>
	/// <param name="apiClient">The API client.</param>
	/// <param name="fileToUpload">The file to upload.</param>
	/// <param name="notifyEmailAddress">The notify email address.</param>
	/// <returns></returns>
	public static HttpResponseMessage UploadFile(HttpClient apiClient, FileInfo fileToUpload, string notifyEmailAddress)
	{

	  // Initialize a new multi-part form content object.
	  MultipartFormDataContent formContent = new MultipartFormDataContent();

	  // Add the file content to the form data for this call.
	  var fileContent = new ByteArrayContent(File.ReadAllBytes(pendingFile.FullName));
	  formContent.Add(fileContent, "attachment", Path.GetFileName(pendingFile.Name));

	  // If we have a notify email address, add this as form data named notifyEmail.
	  if (!string.IsNullOrWhiteSpace(notifyEmailAddress))
	  {

		Program.Log.InfoFormat($"Import job will be set to notify {notifyEmailAddress} upon completion.");
		var data = new StringContent(notifyEmailAddress);
		formContent.Add(data, "notifyEmail");

	  }
	  
	  // Post the file to the API!
	  Program.Log.Info($"Submitting import file {pendingFile.Name} to the TeamDynamix Web API.");
	  return apiClient.PostAsync($"api/people/import", formContent).Result;

	}

}
