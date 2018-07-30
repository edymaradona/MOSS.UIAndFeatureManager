using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Web;
using System.Web.UI.WebControls.WebParts;
using System.Windows.Forms;
using Microsoft.SharePoint;
using Microsoft.SharePoint.WebPartPages;

namespace MOSS.UIAndFeatureManager.UIAndFeatureManager
{
    public class WebPartUtilities
    {
        private PageUtilities pagesUtilities = new PageUtilities();
        private List<string> sharePoint2007WebParts;

        public WebPartUtilities()
        {
            BuildWebPartsLists();
        }

        public List<WebPartToDisplay> GetAllWebPartsInWeb(SPWeb web, bool includeSubsites, bool includeSharePointWebParts, bool includeCustomWebParts)
        {
            if (web == null) return new List<WebPartToDisplay>();
            if (!includeSubsites) return GetAllWebPartsInWeb(web, includeSharePointWebParts, includeCustomWebParts);
            if (web.Webs.Count == 0) return GetAllWebPartsInWeb(web, includeSharePointWebParts, includeCustomWebParts);

            List<WebPartToDisplay> allWebParts = GetAllWebPartsInWeb(web, includeSharePointWebParts, includeCustomWebParts);
            foreach (SPWeb subWeb in web.Webs)
            {
                // Recursively search for web parts within subsites
                allWebParts.AddRange(GetAllWebPartsInWeb(subWeb, true, includeSharePointWebParts, includeCustomWebParts));
                subWeb.Dispose();
            }
            return allWebParts;
        }

        public List<WebPartToDisplay> GetAllWebPartsInWeb(SPWeb web, bool includeSharePointWebParts, bool includeCustomWebParts)
        {
            if (web == null) return new List<WebPartToDisplay>();

            List<string> allPages = pagesUtilities.GetAllPagesInWeb(web);
            if (allPages == null || allPages.Count == 0) return new List<WebPartToDisplay>();

            List<WebPartToDisplay> allWebParts = new List<WebPartToDisplay>();
            foreach (string pageUrl in allPages)
            {
                try
                {
                    allWebParts.AddRange(GetAllWebPartsOnPage(web, pageUrl, includeSharePointWebParts, includeCustomWebParts));
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error in GetAllWebPartsInWeb:\n" + ex.ToString() + "\n\n" +
                                    "Web Url = " + web.Url + "\n" +
                                    "Page Url = " + pageUrl, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    continue;
                }
            }
            return allWebParts;
        }

        public List<WebPartToDisplay> GetAllWebPartsOnPage(SPWeb web, string pageUrl, bool includeSharePointWebParts, bool includeCustomWebParts)
        {
            if (web == null) return new List<WebPartToDisplay>();
            if (string.IsNullOrEmpty(pageUrl)) return new List<WebPartToDisplay>();

            bool isContextNull = false;
            if (HttpContext.Current == null)
            {
                isContextNull = true;
                HttpRequest request = new HttpRequest(string.Empty, web.Url, string.Empty);
                HttpContext.Current = new HttpContext(request, new HttpResponse(new StringWriter()));
                HttpContext.Current.Items["HttpHandlerSPWeb"] = web;
            }

            List<WebPartToDisplay> allWebParts = new List<WebPartToDisplay>();
            using (SPLimitedWebPartManager webpartManager = web.GetFile(pageUrl).GetLimitedWebPartManager(PersonalizationScope.Shared))
            {
                foreach (object webpartObject in webpartManager.WebParts)
                {
                    try
                    {
                        WebPartToDisplay webpartToDisplay = new WebPartToDisplay();
                        if (webpartObject is Microsoft.SharePoint.WebPartPages.WebPart)
                        {
                            // This is a SharePoint web part
                            Microsoft.SharePoint.WebPartPages.WebPart sharepointWebPart = webpartObject as Microsoft.SharePoint.WebPartPages.WebPart;
                            webpartToDisplay.Title = sharepointWebPart.Title;
                            webpartToDisplay.Description = sharepointWebPart.Description;
                            webpartToDisplay.Type = sharepointWebPart.GetType().ToString();
                            webpartToDisplay.Zone = sharepointWebPart.ZoneID;
                            webpartToDisplay.PageUrl = web.Url + "/" + pageUrl;
                            webpartToDisplay.Visible = sharepointWebPart.Visible;
                            webpartToDisplay.Category = GetWebPartCategory(webpartToDisplay.Type);
                        }
                        else if (webpartObject is System.Web.UI.WebControls.WebParts.WebPart)
                        {
                            // This is a ASP.NET web part
                            System.Web.UI.WebControls.WebParts.WebPart aspnetWebPart = webpartObject as System.Web.UI.WebControls.WebParts.WebPart;
                            webpartToDisplay.Title = aspnetWebPart.Title;
                            webpartToDisplay.Description = aspnetWebPart.Description;
                            webpartToDisplay.Type = aspnetWebPart.GetType().ToString();
                            webpartToDisplay.Zone = webpartManager.GetZoneID(aspnetWebPart);
                            webpartToDisplay.PageUrl = web.Url + "/" + pageUrl;
                            webpartToDisplay.Visible = aspnetWebPart.Visible;
                            webpartToDisplay.Category = GetWebPartCategory(webpartToDisplay.Type);
                        }

                        if (webpartToDisplay.Category == WebPartCategory.SharePoint && !includeSharePointWebParts) continue;
                        if (webpartToDisplay.Category == WebPartCategory.Custom && !includeCustomWebParts) continue;

                        allWebParts.Add(webpartToDisplay);
                    }
                    catch (Exception)
                    {
                        throw;
                    }
                    finally
                    {
                        if (isContextNull) HttpContext.Current = null;
                    }
                }
                webpartManager.Dispose();
            }

            if (isContextNull) HttpContext.Current = null;

            return allWebParts;
        }

        public WebPartCategory GetWebPartCategory(string webpartType)
        {
            // Get the namespace of the web part
            string webpartNamespace = webpartType.Substring(0, webpartType.LastIndexOf("."));
            return sharePoint2007WebParts.Contains(webpartNamespace) ? WebPartCategory.SharePoint : WebPartCategory.Custom;
        }

        /// <summary>
        /// Compiles Lists of established Web Parts resource files. 
        /// Create new resource list text files under Project 'Resources ->  Add..'
        /// then go to 'Properties -> Build Action = Embedded'
        /// </summary>
        private void BuildWebPartsLists()
        {
            Assembly assembly = Assembly.GetExecutingAssembly();
            StreamReader stream = new StreamReader(assembly.GetManifestResourceStream("MOSS.UIAndFeatureManager.Resources.sharePoint2007WebParts.txt"));

            // Get SharePoint 2007 Web Parts
            sharePoint2007WebParts = new List<string>();
            while (!stream.EndOfStream)
            {
                sharePoint2007WebParts.Add(stream.ReadLine());
            }
        }
    }

    public class WebPartToDisplay
    {
        public string Title { get; set; }
        public string Description { get; set; }
        public string Type { get; set; }
        public string Zone { get; set; }
        public string PageUrl { get; set; }
        public bool Visible { get; set; }
        public WebPartCategory Category { get; set; }
    }

    public enum WebPartCategory
    {
        SharePoint,
        Custom
    }
}
