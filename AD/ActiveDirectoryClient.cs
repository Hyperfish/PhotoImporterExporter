using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.DirectoryServices;
using System.DirectoryServices.ActiveDirectory;
using System.DirectoryServices.Protocols;
using System.IO;
using System.Linq;
using System.Net;
using System.Security;
using System.Text;
using Hyperfish.ImportExport.O365;
using log4net;
using DirectorySynchronizationOptions = System.DirectoryServices.DirectorySynchronizationOptions;
using SearchOption = System.DirectoryServices.Protocols.SearchOption;
using SearchScope = System.DirectoryServices.Protocols.SearchScope;
using static Hyperfish.ImportExport.Log;

namespace Hyperfish.ImportExport.AD
{
    public class ActiveDirectoryClient
    {
        private const int MaxImageSizeInBytes = 100000;
        public static AdAttribute AdPhotoAttribute = new AdAttribute("thumbnailphoto");

        private LdapConnection _connection;
        private readonly AdSettings _settings;

        public ActiveDirectoryClient(AdSettings settings)
        {
            _settings = settings;
        }

        public LdapConnection Connection
        {
            get
            {
                // if we already have a connection return it
                if (_connection != null)
                    return _connection;

                // build up the connection

                LdapDirectoryIdentifier directoryIdentifier = new LdapDirectoryIdentifier(_settings.AdServer);

                // if we are to connect using a username/pwd 
                if (!string.IsNullOrEmpty(_settings.AdCredentials?.Username))
                {
                    Logger?.Debug($"Get LdapConnection with creds: {_settings.AdCredentials.Username}, {_settings.AdCredentials.Domain}, {_settings.AdServer}");
                    NetworkCredential credentials = new NetworkCredential(_settings.AdCredentials.Username, _settings.AdCredentials.Password, _settings.AdCredentials.Domain);

                    _connection = new LdapConnection(directoryIdentifier, credentials);
                    _connection.SessionOptions.ProtocolVersion = 3;
                    _connection.AuthType = AuthType.Negotiate;
                    _connection.Bind();
                }
                else
                {
                    Logger?.Debug($"Get LdapConnection with current user credentials. {_settings.AdServer}");
                    _connection = new LdapConnection(directoryIdentifier);
                }

                return _connection;
            }
        }

        public void UpdateAttribute(UpnIdentifier upn, AdAttribute propertyName, object propertyValue)
        {
            this.SetAdInfo(upn, propertyName, propertyValue);
        }

        public DistinguishedName BaseDn
        {
            get
            {
                Logger?.Debug($"Finding BaseDN.");

                try
                {
                    // grab it
                    var first = GetBaseAttributes();

                    var baseDn = (string)first.Attributes["defaultNamingContext"][0];

                    Logger?.Debug($"Found default naming context for directory {_settings.AdServer}: {baseDn}");

                    return new DistinguishedName(baseDn);
                }
                catch (LdapException ldapException)
                {
                    throw new AdConnectionException(ldapException);
                }
                catch (Exception exception)
                {
                    throw new AdException("Error finding base DN", exception);
                }
            }
        }

        private SearchResultEntry GetBaseAttributes()
        {
            var request = new SearchRequest(
                string.Empty,
                "(objectclass=*)",
                SearchScope.Base, "defaultNamingContext", "highestCommittedUSN");

            if (!(Connection.SendRequest(request) is SearchResponse response) || response.Entries.Count <= 0)
            {
                throw new AdException("Failed to find the Active Directory root.");
            }

            // grab it
            var first = response.Entries[0];
            return first;
        }
        
        private SearchResultEntry FindUser(IAdIdentifier identifier, IEnumerable<AdAttribute> properties)
        {
            string searchProperty;
            string searchTerm = string.Empty;

            switch (identifier)
            {
                case UpnIdentifier upn:
                    searchProperty = "userPrincipalName";
                    searchTerm = upn.Upn;
                    break;

                case DistinguishedName dn:
                    searchProperty = "distinguishedName";
                    searchTerm = dn.Dn;
                    break;

                default:
                    throw new ArgumentOutOfRangeException(nameof(identifier), identifier, "Not a supported identifier type");
            }

            Logger?.Debug($"Finding user by {searchProperty}, {searchTerm}, under {BaseDn.Dn}");

            var propertyList = properties.Select(p => p.Name).ToArray();

            var request = new SearchRequest(BaseDn.Dn, "(" + searchProperty + "=" + searchTerm + ")", SearchScope.Subtree, propertyList);

            // ensure these fields are always requested
            if (!request.Attributes.Contains("distinguishedName")) request.Attributes.Add("distinguishedName");
            if (!request.Attributes.Contains("objectGUID")) request.Attributes.Add("objectGUID");
            if (!request.Attributes.Contains("userPrincipalName")) request.Attributes.Add("userPrincipalName");

            // do the search
            if (Connection.SendRequest(request) is SearchResponse response && response.Entries.Count > 0)
            {
                // get the first
                var entry = response.Entries[0];

                Logger?.Debug("Found User DN: " + entry.DistinguishedName);

                return entry;
            }
            else
            {
                Logger?.Debug("Didn't find: " + request.Filter.ToString());
                return null;
            }
        }

        private void SetAdInfo(IAdIdentifier identifier, AdAttribute property, object propertyValue)
        {
            try
            {
                var entryToUpdate = FindUser(identifier, new List<AdAttribute>() { property });

                if (entryToUpdate == null)
                {
                    throw new AdException($"User not found {identifier.ToString()}");
                }

                var dn = new DistinguishedName(entryToUpdate.DistinguishedName);

                // if its trying to be removed...
                if (propertyValue == null || (propertyValue is string && string.IsNullOrEmpty((string)propertyValue)))
                {
                    // do nothing
                }
                else
                {
                    // add or update
                    switch (property.Name.ToLower())
                    {
                        case "thumbnailphoto":
                            if (!(propertyValue is Stream photoStream)) throw new ArgumentException("Stream must be passed when updating thrumnailphoto", nameof(propertyValue));
                            var resizedPhoto = FormatAndSizeImageForAd(photoStream);
                            propertyValue = resizedPhoto.ToArray();
                            break;

                        default:
                            break;
                    }

                    if (entryToUpdate.Attributes.Contains(property.Name) || entryToUpdate.Attributes.Contains(property.Name.ToLower()))
                    {
                        // property already exists
                        this.UpdateAttribute(dn, property, propertyValue);
                    }
                    else
                    {
                        // property doesnt already exist
                        this.AddAttribute(dn, property, propertyValue);
                    }
                }

                Logger?.Debug($"Updated {identifier.ToString()}, {property.Name}");
            }
            catch (AdException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new AdException($"Error updating AD: {ex.Message}", ex);
            }
        }

        private void AddAttribute(DistinguishedName dn, AdAttribute attribute, object attributeValue)
        {
            try
            {
                ModifyRequest modRequest;

                if (attribute.IsMultiValued && attributeValue is IList values)
                {
                    object[] array = new object[values.Count];
                    values.CopyTo(array, 0);

                    modRequest = new ModifyRequest(dn.Dn, DirectoryAttributeOperation.Add, attribute.Name, array);
                }
                else
                {
                    modRequest = new ModifyRequest(dn.Dn, DirectoryAttributeOperation.Add, attribute.Name, attributeValue);
                }

                // example of modifyrequest not using the response object...
                this.Connection.SendRequest(modRequest);

                Logger?.Debug($"{attribute.Name} added successfully.");
            }
            catch (DirectoryOperationException e)
            {
                Logger?.Debug($"Failed to add {attribute.Name} in: {dn}. Message: {e.Message}");
                throw;
            }
        }

        private void UpdateAttribute(DistinguishedName dn, AdAttribute attribute, object attributeValue)
        {
            try
            {
                ModifyRequest modRequest;

                if (attribute.IsMultiValued && attributeValue is IList values)
                {
                    object[] array = new object[values.Count];
                    values.CopyTo(array, 0);

                    modRequest = new ModifyRequest(dn.Dn, DirectoryAttributeOperation.Replace, attribute.Name, array);
                }
                else
                {
                    modRequest = new ModifyRequest(dn.Dn, DirectoryAttributeOperation.Replace, attribute.Name, attributeValue);
                }

                this.Connection.SendRequest(modRequest);

                Logger?.Debug($"The {attribute.Name} attribute in: {dn} replaced successfully with a value");
            }
            catch (DirectoryOperationException e)
            {
                Logger?.Debug($"Failed to update {attribute.Name} in: {dn}. Message: {e.Message}");
                throw;
            }
        }
        
        /// <summary>
        /// Scales and formats images to fit AD well
        /// </summary>
        /// <param name="imageBytes">The image to process and resize</param>
        /// <returns>Resized image in JPG format</returns>
        private static MemoryStream FormatAndSizeImageForAd(Stream photoStream)
        {
            try
            {
                var scaledImage = ImageHelper.ResizeImage(photoStream, MaxImageSizeInBytes, 93);
                return scaledImage;
            }
            catch (Exception ex)
            {
                throw new AdException("Failed to process image", ex);
            }
        }


    }
}
