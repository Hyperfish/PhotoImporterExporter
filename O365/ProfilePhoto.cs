using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Hyperfish.ImportExport.O365
{
    class ProfilePhoto
    {
        //sizes for profile pictures
        private const int SmallThumbWidth = 48;
        private const int MediumThumbWidth = 72;
        private const int LargeThumbWidth = 200;

        public ProfilePhoto(PhotoSize size)
        {
            this.Size = size;
        }

        public PhotoSize Size { get; set; }

        public string Url(string username, string library)
        {
            string photoPrefix = username.Replace("@", "_").Replace(".", "_");
            string photoPath = string.Concat(library.TrimEnd('/'), "/{0}_{1}Thumb.jpg");

            return string.Format(photoPath, photoPrefix, this.Size.ToString());
        }

        public int MaxWidth
        {
            get
            {
                switch (this.Size)
                {
                    case PhotoSize.S:
                        return SmallThumbWidth;
                    case PhotoSize.M:
                        return MediumThumbWidth;
                    case PhotoSize.L:
                        return LargeThumbWidth;
                    default:
                        return LargeThumbWidth;
                }
            }
        }
    }
    
    public enum PhotoSize
    {
        S,
        M,
        L
    }
    
}
