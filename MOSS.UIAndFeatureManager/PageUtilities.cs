using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Publishing;

namespace MOSS.UIAndFeatureManager.UIAndFeatureManager
{
    public class PageUtilities
    {
        public List<string> GetAllPagesInWeb(SPWeb web, bool includeSubsites)
        {
            if (web == null) return new List<string>();
            if (!includeSubsites) return GetAllPagesInWeb(web);
            if (web.Webs.Count == 0) return GetAllPagesInWeb(web);

            List<string> allPages = GetAllPagesInWeb(web);
            foreach (SPWeb subWeb in web.Webs)
            {
                // Recursively search for pages within subsites
                allPages.AddRange(GetAllPagesInWeb(subWeb, true));
                subWeb.Dispose();
            }
            return allPages;
        }

        public List<string> GetAllPagesInWeb(SPWeb web)
        {
            if (PublishingWeb.IsPublishingWeb(web))
            {
                return GetAllPublishingPages(PublishingWeb.GetPublishingWeb(web));
            }
            else
            {
                return GetAllPagesFromRoot(web);
            }
        }

        private List<string> GetAllPagesFromRoot(SPWeb web)
        {
            if (web == null) return new List<string>();
            List<string> allPages = new List<string>();

            // Get all pages from the root directory
            try
            {
                foreach (SPFile file in web.RootFolder.Files)
                {
                    if (Path.GetExtension(file.Url) == ".aspx") allPages.Add(file.Url);
                }
            }
            catch (Exception) { throw; }

            return allPages;
        }

        private List<string> GetAllPublishingPages(PublishingWeb web)
        {
            if (web == null) return new List<string>();
            List<string> allPages = new List<string>();

            // Get all pages from the pages library
            try
            {
                foreach (PublishingPage pubPage in web.GetPublishingPages())
                {
                    if (Path.GetExtension(pubPage.Url) == ".aspx") allPages.Add(pubPage.Url);
                }
            }
            catch (Exception) { throw; }

            return allPages;
        }
    }
}