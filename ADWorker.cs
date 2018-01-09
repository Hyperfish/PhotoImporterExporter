using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using Hyperfish.ImportExport.AD;
using static Hyperfish.ImportExport.Log;

namespace Hyperfish.ImportExport
{
    class AdWorker
    {
        private readonly ActiveDirectoryClient  _ad;

        public AdWorker(AdSettings adSettings)
        {
            _ad = new ActiveDirectoryClient(adSettings);
        }

        public void ImportPicsToAd(Dictionary<string, ImportExportRecord> peopleToImport)
        {
            foreach (var upn in peopleToImport.Keys)
            {
                var upnIdentifier = new UpnIdentifier(upn);
                var picture = peopleToImport[upn].PhotoLocation;

                Stream photoStream;

                Logger.Info($"Importing photo for {upn}, {picture}");

                try
                {
                    photoStream = File.OpenRead(picture);
                }
                catch (Exception e)
                {
                    Logger.Error($"Couldn't read photo from disk: {picture}, Message: {e.Message}. Skipping");
                    continue;
                }

                try
                {
                    _ad.UpdateAttribute(upnIdentifier, ActiveDirectoryClient.AdPhotoAttribute, photoStream);
                }
                catch (Exception e)
                {
                    Logger.Error($"Problem updating AD for {upn}: {e.Message}. Skipping");
                }
            }
        }

    }
}
