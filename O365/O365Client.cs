using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using Hyperfish.ImportExport.Helpers;
using Hyperfish.ImportExport.O365.Exceptions;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Search.Query;
using Microsoft.SharePoint.Client.UserProfiles;
using static Hyperfish.ImportExport.Log;

namespace Hyperfish.ImportExport.O365
{

    public partial class O365Client
    {
        private const int SpoRetries = 5;
        private const int SpoBackoff = 20000;

        // max image size for EXO 3.8 megabytes
        private const long MaxImageSizeInBytes = 3800000;

        private O365Settings _settings;

        public O365Client(O365Settings settings)
        {
            _settings = settings;
        }

        public MemoryStream DownloadProfilePhoto(O365Profile person, O365Service service)
        {

            switch (service)
            {
                case O365Service.Spo:

                    var clientContext = GetSpoClientContextForSite(SpoSite.MySite);

                    var uri = person.SpoProfilePictureUrlLarge;
                    var serverrelativePath = uri.AbsolutePath;

                    var memoryStream = new MemoryStream();

                    using (FileInformation f = Microsoft.SharePoint.Client.File.OpenBinaryDirect(clientContext, serverrelativePath))
                    {
                        f.Stream.CopyTo(memoryStream);
                    }

                    return memoryStream;

                case O365Service.Exo:

                    var exo = GetExoClientContext();
                    var results = exo.GetUserPhoto(person.Upn.Upn, "HR648x648", string.Empty);

                    if (results.Status == GetUserPhotoStatus.PhotoReturned)
                    {
                        var picStream = new MemoryStream(results.Photo);
                        return picStream;
                    }
                    else
                    {
                        throw new FileNotFoundException($"Failed to get photo from Exo for {person.Upn.Upn}");
                    }

                default:
                    throw new ArgumentOutOfRangeException(nameof(service), service, null);
            }

        }

        public void UpdateUserProfilePhoto(O365Profile person, O365Service service, Stream photoStream)
        {

            switch (service)
            {
                case O365Service.Spo:

                    var clientContext = GetSpoClientContextForSite(SpoSite.MySite);

                    UpdateSpoPhoto(clientContext, person.Upn, photoStream);

                    return;

                case O365Service.Exo:

                    UpdateExoPhoto(person.Upn, photoStream);
                    break;

                default:
                    throw new ArgumentOutOfRangeException(nameof(service), service, null);
            }
        }

        public void PageThroughAllUsers(IEnumerable<O365Attribute> properties, int pageSize, Func<IList<O365Profile>, bool> ProcessPage)
        {
            Stopwatch sw = new Stopwatch();
            sw.Start();

            Logger?.Debug($"O365 paging through all users.");

            // get the list of users to audit
            var loginNames = GetUserLoginNamesFromSearch();

            // batch profile requests into 100s
            var batches = SplitList(loginNames, 100);

            Logger?.Debug($"[O365] Batching profile fetch into {batches.Count}. Elapsed time: {sw.Elapsed.TotalSeconds} seconds");

            foreach (var batch in batches)
            {
                // get the user profiles from SPO
                var userProfiles = GetProfilesForUsers(batch);

                // make O365Profiles for each user
                var o365Profiles = userProfiles
                        .Where(p => p.ServerObjectIsNull != null && p.ServerObjectIsNull.Value != true)
                        .Select(u =>
                        {
                            var upn = u.AccountName.Contains('|') ? u.AccountName.Split('|').Last() : string.Empty;

                            return new O365Profile(new UpnIdentifier(upn))
                            {
                                Properties = u.UserProfileProperties.Where(p => !string.IsNullOrEmpty(p.Value)).ToDictionary(p => p.Key, p => (object)p.Value)
                            };
                        })
                        .Where(u => !string.IsNullOrEmpty(u.Upn.Upn)).ToList();

                Logger?.Debug($"[O365] Found {o365Profiles.Count}. Elapsed time: {sw.Elapsed.TotalSeconds} seconds");

                // process this page
                var shouldContinue = ProcessPage(o365Profiles);

                if (!shouldContinue)
                {
                    break;
                }

                Logger?.Debug($"[O365] Completed batch. Elapsed time: {sw.Elapsed.TotalSeconds} seconds");
            }

            sw.Stop();
            Logger?.Debug($"[O365] Completed batching {batches.Count}. Elapsed time: {sw.Elapsed.TotalSeconds} seconds");

        }

        private List<string> GetUserLoginNamesFromSearch()
        {
            Logger?.Debug($"[O365] Getting user list from search");

            try
            {
                List<string> loginNames = new List<string>();

                var ctx = GetSpoClientContextForSite(SpoSite.RootSite);
                SearchExecutor searchExecutor = new SearchExecutor(ctx);

                int currentPage = 0;
                int totalRows = -1;
                int startRow = 1;
                int rowLimit = 10;
                do
                {
                    startRow = (rowLimit * currentPage) + 1;

                    // http://www.techmikael.com/2015/01/how-to-query-using-result-source-name.html
                    KeywordQuery qry = new KeywordQuery(ctx);
                    qry.Properties["SourceName"] = "Local People Results";
                    qry.Properties["SourceLevel"] = "SSA";

                    qry.QueryText = "*";
                    qry.RowLimit = rowLimit;
                    qry.StartRow = startRow;

                    ClientResult<ResultTableCollection> results = searchExecutor.ExecuteQuery(qry);
                    ctx.ExecuteQueryWithIncrementalRetry(SpoRetries, SpoBackoff, Logger);

                    var resultTable = results.Value[0];

                    if (currentPage == 0)
                    {
                        totalRows = resultTable.TotalRows;
                    }

                    foreach (var resultRow in resultTable.ResultRows)
                    {
                        loginNames.Add(resultRow["AccountName"].ToString());
                    }

                    currentPage++;

                } while (startRow + rowLimit < totalRows);

                return loginNames;
            }
            catch (MaximumRetryAttemptedException ex)
            {
                // Exception handling for the Maximum Retry Attempted
                Logger?.Error($"[O365] Max retries / throttle for SPO reached getting site user list. Message {ex.Message}");
                throw new O365Exception($"Max retries / throttle for SPO reached. Message {ex.Message}", ex);
            }
            catch (Exception ex)
            {
                Logger?.Error($"[O365] Problem getting root site user list. Message {ex.Message}");
                throw new O365Exception($"Problem getting root site user list. Message {ex.Message}", ex);
            }

        }


        private List<User> GetRootSiteUserList()
        {
            Logger?.Debug($"[O365] Getting root site user list");

            var ctx = GetSpoClientContextForSite(SpoSite.RootSite);

            try
            {
                var web = ctx.Web;

                // load the site users
                ctx.Load(web, w => w.SiteUsers);

                // get users who are users, not groups etc...
                var usersQuery = from user in web.SiteUsers
                                                 where user.PrincipalType == Microsoft.SharePoint.Client.Utilities.PrincipalType.User
                                                         && user.IsShareByEmailGuestUser == false
                                                         && user.IsEmailAuthenticationGuestUser == false
                                                 select user;

                var siteUsers = ctx.LoadQuery(usersQuery);

                // execute the query to the service
                ctx.ExecuteQueryWithIncrementalRetry(SpoRetries, SpoBackoff, Logger);

                return siteUsers.ToList();
            }
            catch (MaximumRetryAttemptedException ex)
            {
                // Exception handling for the Maximum Retry Attempted
                Logger?.Error($"[O365] Max retries / throttle for SPO reached getting site user list. Message {ex.Message}");
                throw new O365Exception($"Max retries / throttle for SPO reached. Message {ex.Message}", ex);
            }
            catch (Exception ex)
            {
                Logger?.Error($"[O365] Problem getting root site user list. Message {ex.Message}");
                throw new O365Exception($"Problem getting root site user list. Message {ex.Message}", ex);
            }
        }

        private List<PersonProperties> GetProfilesForUsers(IList<string> loginNames)
        {
            var ctx = GetSpoClientContextForSite(SpoSite.Admin);
            var peopleManager = new PeopleManager(ctx);
            var userProfiles = new List<PersonProperties>();

            Logger?.Debug($"[O365] Getting profiles for {loginNames.Count} users");

            try
            {
                foreach (var loginName in loginNames)
                {
                    var userProfile = peopleManager.GetPropertiesFor(loginName);
                    ctx.Load(userProfile);
                    userProfiles.Add(userProfile);
                }

                ctx.ExecuteQueryWithIncrementalRetry(SpoRetries, SpoBackoff, Logger);
            }
            catch (MaximumRetryAttemptedException ex)
            {
                // Exception handling for the Maximum Retry Attempted
                Logger?.Error($"[O365] Max retries / throttle for SPO reached getting profiles. Message {ex.Message}");
                throw new O365Exception($"Max retries / throttle for SPO reached. Message {ex.Message}", ex);
            }
            catch (Exception ex)
            {
                Logger?.Error($"[O365] Propblem getting profiles for users. Message: {ex.Message}");
            }

            return userProfiles;
        }

        private ClientContext GetSpoClientContextForSite(SpoSite type)
        {
            string siteUri;
            var adminUser = _settings.Username;
            var adminPassword = _settings.Password;

            switch (type)
            {
                case SpoSite.MySite:
                    siteUri = _settings.MySiteHost;
                    break;

                case SpoSite.Admin:
                    siteUri = _settings.AdminUri;
                    break;

                case SpoSite.RootSite:
                    siteUri = _settings.RootSiteHost;
                    break;

                default:
                    throw new ArgumentOutOfRangeException(nameof(type), type, null);
            }

            try
            {
                Logger?.Debug($"Initializing service object for SPO Client API: " + siteUri);

                var clientContext = new ClientContext(siteUri);

                clientContext.Credentials = new SharePointOnlineCredentials(adminUser, adminPassword);

                // Add our User Agent information
                clientContext.ExecutingWebRequest += delegate (object sender, WebRequestEventArgs e)
                {
                    e.WebRequestExecutor.WebRequest.UserAgent = "ISV|Hyperfish|HyperfishImportExport/1.0";
                };

                return clientContext;
            }
            catch (Exception ex)
            {
                Logger?.Error($"Error creating client context for SPO: {siteUri}, {ex.Message}");
                throw;
            }
        }

        private ExchangeService GetExoClientContext()
        {
            var adminUser = _settings.Username;
            var adminPassword = _settings.Password;

            ExchangeService exo = new ExchangeService(ExchangeVersion.Exchange2016);
            exo.Credentials = new WebCredentials(adminUser, adminPassword.UnSecure());

            exo.Url = new Uri(_settings.ExoServiceUri);

            return exo;
        }

        private void UpdateBasicSpoAttribute(ClientContext clientContext, UpnIdentifier upn, string attributeName, string value)
        {
            var targetUser = O365Settings.SpoProfilePrefix + upn.Upn;
            PeopleManager peopleManager = new PeopleManager(clientContext);

            peopleManager.SetSingleValueProfileProperty(targetUser, attributeName, value.ToString());
        }

        private void UpdateSpoPhoto(ClientContext clientContext, UpnIdentifier upn, Stream photoStream)
        {
            const string library = @"/User Photos/Profile Pictures";

            var mySitePhotos = new List<ProfilePhoto>()
                        {
                                new ProfilePhoto(PhotoSize.S), new ProfilePhoto(PhotoSize.M), new ProfilePhoto(PhotoSize.L)
                        };

            var mySiteHostClientContext = GetSpoClientContextForSite(SpoSite.MySite);

            try
            {
                foreach (var mySitePhoto in mySitePhotos)
                {
                    Logger?.Debug($"Resizing image: {mySitePhoto.Size.ToString()} to {mySitePhoto.MaxWidth} max width");

                    using (Stream thumb = ImageHelper.ResizeImage(photoStream, mySitePhoto.MaxWidth))
                    {
                        if (thumb != null)
                        {
                            var url = mySitePhoto.Url(upn.Upn, library);

                            Logger?.Debug($"Uploading {mySitePhoto.Size.ToString()} photo to {url}");

                            Microsoft.SharePoint.Client.File.SaveBinaryDirect(mySiteHostClientContext, url, thumb, true);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger?.Error($"Problem resizing/uploading photo. Message: {ex.Message}");
                throw new O365Exception($"Problem resizing/uploading photo. Message: {ex.Message}", ex);
            }

            try
            {
                Logger?.Debug($"Setting the PictureUrl and SPS-PicturePlaceholderState");

                var relativeMedPhotoPath = mySitePhotos.First(p => p.Size == PhotoSize.M).Url(upn.Upn, library);
                var medPhotoPath = new Uri(new Uri(mySiteHostClientContext.Url), relativeMedPhotoPath).AbsoluteUri;

                // update the profile with the URL for the photo and the placeholderstate
                this.UpdateBasicSpoAttribute(clientContext, upn, "PictureURL", medPhotoPath);
                this.UpdateBasicSpoAttribute(clientContext, upn, "SPS-PicturePlaceholderState", "0");

                // make the changes
                clientContext.ExecuteQueryWithIncrementalRetry(SpoRetries, SpoBackoff, Logger);

            }
            catch (MaximumRetryAttemptedException ex)
            {
                // Exception handling for the Maximum Retry Attempted
                Logger?.Error($"[O365] Max retries / throttle for SPO reached updating SPO photo. Message {ex.Message}");
                throw new O365Exception($"Max retries / throttle for SPO reached. Message {ex.Message}", ex);
            }
            catch (Exception ex)
            {
                Logger?.Error($"Problem setting the PictureUrl and SPS-PicturePlaceholderState. Message: {ex.Message}");
                throw new O365Exception($"Problem setting the PictureUrl and SPS-PicturePlaceholderState. Message: {ex.Message}", ex);
            }
        }

        private void UpdateExoPhoto(UpnIdentifier upn, Stream photoStream)
        {
            try
            {
                Logger?.Debug($"Resizing image");

                var newImage = FormatAndSizeImageForExo(photoStream);

                // upload             
                var service = GetExoClientContext();

                Logger?.Debug($"Uploading photo to Exo");

                service.SetUserPhoto(upn.Upn, newImage);

                Console.WriteLine("Uploaded photo to Exo");

            }
            catch (Exception ex)
            {
                Logger?.Error($"Problem resizing/uploading EXO photo. Message: {ex.Message}");
                throw new O365Exception($"Problem resizing/uploading EXO photo. Message: {ex.Message}", ex);
            }

        }

        private static byte[] FormatAndSizeImageForExo(Stream photoStream)
        {
            try
            {
                using (var scaledImage = ImageHelper.ResizeImage(photoStream, MaxImageSizeInBytes, 93))
                {
                    return scaledImage.ToArray();
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Failed to process image", ex);
            }
        }

        private IList<List<T>> SplitList<T>(IList<T> source, int groupSize)
        {
            return source
                    .Select((x, i) => new { Index = i, Value = x })
                    .GroupBy(x => x.Index / groupSize)
                    .Select(x => x.Select(v => v.Value).ToList<T>())
                    .ToList();
        }
    }

    public enum O365Service
    {
        Spo,
        Exo
    }

    public enum SpoSite
    {
        MySite,
        Admin,
        RootSite
    }


}
