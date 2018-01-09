using log4net;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using Hyperfish.ImportExport.O365;
using static Hyperfish.ImportExport.Log;

namespace Hyperfish.ImportExport
{
    class O365Worker
    {
        private const string PhotoDirectory = "photos";
        private static byte[] _silouetteProfilePicLarge;

        private readonly O365Client _o365;
        private readonly Dictionary<string, ImportExportRecord> _allPeople = new Dictionary<string, ImportExportRecord>();
        
        public O365Worker(O365Settings o365Settings)
        {
            _o365 = new O365Client(o365Settings);
        }

        static O365Worker()
        {
            //_silouetteProfilePicLarge = File.ReadAllBytes("silhouette_LThumb.jpg");

            var assembly = Assembly.GetExecutingAssembly();
            var resourceName = "Hyperfish.ImportExport.O365.assets.silhouette_LThumb.jpg";
            
            var memoryStream = new MemoryStream();

            using (var stream = assembly.GetManifestResourceStream(resourceName))
            {
                stream.Seek(0, SeekOrigin.Begin);
                stream.CopyTo(memoryStream);

                _silouetteProfilePicLarge = memoryStream.ToArray();
            }
            
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
                    var uri = new Uri(person.SpoProfilePictureUrlLarge.ToString());
                    var extension = Path.GetExtension(uri.AbsolutePath);
                    
                    var filename = SanitizeFileName(person.Upn.Upn) + extension;
                    var fullPath = Path.Combine(PhotoDirectory, filename);

                    var record = new ImportExportRecord() { PhotoLocation = String.Empty, Upn = person.Upn.Upn };

                    // download the file
                    var pictureStream = _o365.DownloadProfilePhoto(person, service);
                    
                    if (IsPhotoLargeSilouette(pictureStream))
                    {
                        // this is a silouette photo ... move on
                        continue;
                    }
                    else
                    {
                        // save the picture
                        using (FileStream fs = new FileStream(fullPath, FileMode.Create))
                        {
                            if (pictureStream.CanSeek) pictureStream.Seek(0, SeekOrigin.Begin);
                            pictureStream.CopyTo(fs);
                        }

                        record.PhotoLocation = fullPath;
                    }

                    // keep their record
                    _allPeople[person.Upn.Upn] = record;

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
        
        private static bool IsPhotoLargeSilouette(MemoryStream pictureToCompare)
        {
            if (_silouetteProfilePicLarge.Length != pictureToCompare.Length)
                return false;

            pictureToCompare.Position = 0;
            var pictureBytes = pictureToCompare.ToArray();

            return _silouetteProfilePicLarge.SequenceEqual(pictureBytes);
        }

        private void EnsurePhotosDirectory()
        {
            Directory.CreateDirectory("photos");
        }
        

    }
}
