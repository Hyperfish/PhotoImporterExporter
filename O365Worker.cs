using log4net;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Hyperfish.ImportExport.O365;
using static Hyperfish.ImportExport.Log;

namespace Hyperfish.ImportExport
{
    class O365Worker
    {
        private const string PhotoDirectory = "photos";

        private readonly O365Client _o365;
        private readonly Dictionary<string, ImportExportRecord> _allPeople = new Dictionary<string, ImportExportRecord>();
        
        public O365Worker(O365Settings o365Settings)
        {
            _o365 = new O365Client(o365Settings);
        }

        public Dictionary<string, ImportExportRecord> Export(O365Service service)
        {
            EnsurePhotosDirectory();

            // Export all the images from SPO
            _o365.PageThroughAllUsers(null, 10, (page) => ProcessO365PageOfUsers(page, service));
            
            return _allPeople;
        }

        public void Import(Dictionary<string, ImportExportRecord> people, O365Service service)
        {
            foreach (var person in people)
            {
                if (string.IsNullOrEmpty(person.Value.PhotoLocation))
                {
                    // no photo .. skipping
                    Logger.Debug($"No photo found for {person.Value.Upn} ... skipping");
                    continue;
                }

                using (var fs = File.Open(person.Value.PhotoLocation, FileMode.Open))
                {
                    var profile = new O365Profile(new UpnIdentifier(person.Value.Upn));
                    _o365.UpdateUserProfilePhoto(profile, service, fs);

                    Logger.Info($"Uploaded photo for {person.Value.Upn}");
                }
            }
        }

        private bool ProcessO365PageOfUsers(IList<O365Profile> page, O365Service service)
        {
            foreach (var person in page)
            {
                if (person.Upn == null)
                {
                    Logger.Info($"Upn for user not found. Moving on");
                    continue;
                }

                if (service == O365Service.Spo && !person.HasSpoProfilePicture)
                {
                    Logger.Info($"Found user {person.Upn}, but skipping as they dont have a photo set");
                    continue;
                }
                
                Logger.Info($"Found user {person.Upn} downloading photo");

                try
                {
                    var uri = new Uri(person.SpoProfilePictureUrl.ToString());
                    var extension = Path.GetExtension(uri.AbsolutePath);
                    
                    var filename = SanitizeFileName(person.Upn.Upn) + extension; //uri.Segments.Last();
                    
                    var fullPath = Path.Combine(PhotoDirectory, filename);

                    // download the file
                    var pictureStream = _o365.DownloadProfilePhoto(person, service);

                    // save the picture
                    using (FileStream fs = new FileStream(fullPath, FileMode.Create))
                    {
                        if(pictureStream.CanSeek) pictureStream.Seek(0, SeekOrigin.Begin);
                        pictureStream.CopyTo(fs);
                    }
                    
                    // save them for later
                    if (!_allPeople.ContainsKey(person.Upn.Upn))
                    {
                        _allPeople[person.Upn.Upn] = new ImportExportRecord() {PhotoLocation = fullPath, Upn = person.Upn.Upn};
                    }
                    
                }
                catch (Exception e)
                {
                    Logger.Error($"Error downloading photo for user {person.Upn}, Message: {e.Message}");
                    // move on
                }
            }

            return true;
        }

        private string SanitizeFileName(string filename)
        {
            var badChars = new List<char>();
            badChars.AddRange(Path.GetInvalidFileNameChars());
            badChars.Add('@');
            badChars.Add(':');
            badChars.Add('.');

            foreach (var badChar in badChars)
            {
                filename = filename.Replace(badChar, '_');
            }

            return filename;
        }

        private void EnsurePhotosDirectory()
        {
            Directory.CreateDirectory("photos");
        }
        

    }
}
