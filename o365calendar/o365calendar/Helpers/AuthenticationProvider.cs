using System;
using System.Threading.Tasks;
using Windows.Security.Authentication.Web;
using Windows.Storage;
using Microsoft.Azure.ActiveDirectory.GraphClient;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.Office365.Discovery;
using Microsoft.Office365.OutlookServices;

namespace o365calendar.Helpers
{

    /// <summary>
    /// :exclamation:This is the implementation of the O365 API Authentication and does not use the unified API
    /// Please note if possible use the new API Endpoints
    /// </summary>
    internal class AuthenticationProvider
    {
        private static readonly string ClientId =
            Windows.UI.Xaml.Application.Current.Resources["ida:ClientID"].ToString();

        private static Uri _returnUri = WebAuthenticationBroker.GetCurrentApplicationCallbackUri();

        private static readonly string CommonAuthority =
            Windows.UI.Xaml.Application.Current.Resources["ida:AADInstance"] + @"Common";

        private static readonly Uri DiscoveryServiceUri = new Uri("https://api.office.com/discovery/v1.0/me/");
        private const string DiscoveryResourceId = "https://api.office.com/discovery/";
        private const string AadServiceResourceId = @"https://graph.windows.net/";

        private static ActiveDirectoryClient _graphClient;
        private static OutlookServicesClient _outlookClient;

        //:blue_book:Settings Container and Properties to access it.

        #region Settings Stuff to keep

        /// <summary>
        /// :smile: holds the value of the properties below
        /// </summary>
        private static ApplicationDataContainer _settings = ApplicationData.Current.LocalSettings;

        /// <summary>
        /// Property Containing the Last Authority when set otherwise returns string.empty
        /// </summary>
        public static string LastAuthority
        {
            get
            {
                return _settings.Values.ContainsKey("LastAuthority") && _settings.Values["LastAuthority"] != null
                    ? _settings.Values["LastAuthority"].ToString()
                    : string.Empty;
            }
            set { _settings.Values["LastAuthority"] = value; }
        }

        /// <summary>
        /// Property for storing the tenant id so that we can pass it to the ActiveDirectoryClient
        /// </summary>
        internal static string TenantId
        {
            get
            {
                return _settings.Values.ContainsKey("TenantId") && _settings.Values["TenantId"] != null
                    ? _settings.Values["TenantId"].ToString()
                    : string.Empty;
            }
            set { _settings.Values["TenantId"] = value; }
        }

        public static string LoggedInUser
        {
            get
            {
                return _settings.Values.ContainsKey("LoggedInUser") && _settings.Values["LoggedInUser"] != null
                    ? _settings.Values["LoggedInUser"].ToString()
                    : string.Empty;
            }
            set { _settings.Values["LoggedInUser"] = value; }
        }

        #endregion

        //:key: The Authentication Context used to sign in the user
        public static AuthenticationContext AuthenticationContext { get; set; }

        public static async Task<ActiveDirectoryClient> GetGraphClientAsync()
        {
            //:one: Check if client is available otherwise create new one
            if (_graphClient != null)
            {
                return _graphClient;
            }
            else
            {
                //ActiveDirectory Service Endpoint for UWP Apps
                Uri AadServiceEndpointUri = new Uri(AadServiceResourceId);

                try
                {
                    //:two:First look for the last authority used during authentication, if there is no value use common
                    string authority = string.IsNullOrEmpty(LastAuthority) ? CommonAuthority : LastAuthority;
                    //:three:Create AuthenticationContext using authority
                    AuthenticationContext = new AuthenticationContext(authority);

                    //:office: There is a Enterprise SSO feature in Windows for CorporateNetwork Users. If you want to use this Feature you need to add the Enterprise Authentication, Pricate Networks, and Shared User Certificates capabilites in the App Manifest Package_appxmanifest
                    //AuthenticationContext.UseCorporateNetwork = true;

                    //:four:Get the Accesstoken
                    var token = await GetTokenHelperAsync(AuthenticationContext, AadServiceResourceId);

                    //:five:Validate the Token
                    if (!string.IsNullOrEmpty(token))
                    {
                        _graphClient = new ActiveDirectoryClient(new Uri(AadServiceEndpointUri, TenantId),
                            async () => await GetTokenHelperAsync(AuthenticationContext, AadServiceResourceId));
                        return _graphClient;
                    }
                    else
                    {
                        //User cancelled sign-in or smth went wrong
                        return null;
                    }
                }
                catch (Exception exception)
                {
                    MessageDialogHelper.DisplayException(exception as Exception);
                    return null;
                }
                finally
                {
                    AuthenticationContext.TokenCache.Clear();
                }
            }
        } //GetGraphClientAsync

        /// <summary>
        /// Checks and makes sure that an OutlookClient object is available.
        /// </summary>
        /// <param name="capability">Checks for the Capability</param>
        /// <returns>Outlook Services Client Object</returns>
        public static async Task<OutlookServicesClient> GetOutlookClientAsync(string capability)
        {
            //:one:Check if this client already exists, otherwise ceate a new one
            if (_outlookClient != null)
            {
                return _outlookClient;
            }
            else
            {
                try
                {
                    //:two: Check if capability is available
                    CapabilityDiscoveryResult result = await GetDiscoveryCapabilityResultAsync(capability);

                    //:three: Create OutlookClient
                    _outlookClient = new OutlookServicesClient(result.ServiceEndpointUri,
                        async () => await GetTokenHelperAsync(AuthenticationContext, result.ServiceResourceId));
                    //:four: Return the OutlookClient
                    return _outlookClient;
                }
                catch (DiscoveryFailedException dfe)
                {
                    MessageDialogHelper.DisplayException(dfe as Exception);
                    return null;
                }
                catch (ArgumentException e)
                {
                    MessageDialogHelper.DisplayException(e as Exception);
                    return null;
                }
                finally
                {
                    AuthenticationContext.TokenCache.Clear();
                }
            }
        } //GetOutlookClientAsync 

        /// <summary>
        /// Signs the user out of the service
        /// </summary>
        public static async Task SignOutAsync()
        {
            if (string.IsNullOrEmpty(LoggedInUser))
            {
                return;
            }
            //Clean the mess
            AuthenticationContext.TokenCache.Clear();
            _graphClient = null;
            _outlookClient = null;
            //Clear stored values from last authentication.
            _settings.Values["TenantId"] = null;
            _settings.Values["LastAuthority"] = null;
        }

        /// <summary>
        /// This is the Main Authentication functionality to get the access Token for the desired services
        /// :bangbang: Get an access Token for the given context and resourceId. Attempt to aquire token silently, if that fails try to acquire the token prompting for user credentials.
        /// </summary>
        /// <param name="context">AuthenticationContext</param>
        /// <param name="resourceId">Endpoint</param>
        /// <returns></returns>
        private static async Task<string> GetTokenHelperAsync(AuthenticationContext context, string resourceId)
        {
            AuthenticationResult result = null;

            result = await context.AcquireTokenAsync(resourceId, ClientId, _returnUri);

            if (result.Status == AuthenticationStatus.Success)
            {
                string accessToken = result.AccessToken;
                //:floppy_disk: Store values for logged in user
                //they can be reused if user re-opens the app without disconnecting
                _settings.Values["LoggedInUser"] = result.UserInfo.UniqueId;
                _settings.Values["TenantId"] = result.TenantId;
                _settings.Values["LastAuthority"] = context.Authority;

                return accessToken;
            }
            return null;
        }

        public static async Task<DiscoveryServiceCache> CreateAndSaveDiscoveryServiceCacheAsync()
        {
            DiscoveryServiceCache discoveryCache = null;

            var discoveryClient = new DiscoveryClient(DiscoveryServiceUri,
                async () => await GetTokenHelperAsync(AuthenticationContext, DiscoveryResourceId));

            var discoveryCapabilityResult = await discoveryClient.DiscoverCapabilitiesAsync();

            discoveryCache = await DiscoveryServiceCache.CreateAndSaveAsync(LoggedInUser, discoveryCapabilityResult);

            return discoveryCache;
        }

        public static async Task<CapabilityDiscoveryResult> GetDiscoveryCapabilityResultAsync(string capability)
        {
            var cacheResult = await DiscoveryServiceCache.LoadAsync();

            CapabilityDiscoveryResult discoveryCapabilityResult = null;

            if (cacheResult != null && cacheResult.DiscoveryInfoForServices.ContainsKey(capability))
            {
                discoveryCapabilityResult = cacheResult.DiscoveryInfoForServices[capability];

                if (LoggedInUser != cacheResult.UserId)
                {
                    // cache is for another user
                    cacheResult = null;
                }
            }

            //Cache might exist from previous calls, but it might not include newly added capabilities.
            else if (cacheResult != null && !cacheResult.DiscoveryInfoForServices.ContainsKey(capability))
            {
                cacheResult = null;
                cacheResult = await CreateAndSaveDiscoveryServiceCacheAsync();
                discoveryCapabilityResult = cacheResult.DiscoveryInfoForServices[capability];
            }

            if (cacheResult == null)
            {
                cacheResult = await CreateAndSaveDiscoveryServiceCacheAsync();
                discoveryCapabilityResult = cacheResult.DiscoveryInfoForServices[capability];
            }

            return discoveryCapabilityResult;
        }
    }
}