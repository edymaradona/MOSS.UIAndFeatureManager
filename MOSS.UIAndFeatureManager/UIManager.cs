using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Windows.Forms;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using MOSS.UIAndFeatureManager.UIAndFeatureManager;

namespace MOSS.UIAndFeatureManager
{
/// <summary>
/// *******************************************************************************************
/// Site Information tool.  Displays the content of all sites under a selected site collection.
/// Author: Mark Shipman (mark.e.shipman@live.com)
/// Last Edited: 1/25/2012, 5:27 PM
/// *******************************************************************************************
/// </summary>

#region UIManager Class
//====================================================================================
public partial class UIManager : Form
{
    #region Properties

        int featuresDisplayed = 0;
        private List<FeatureToDisplay> featuresToDisplay = new List<FeatureToDisplay>();
        private List<string> sharePoint2007Features = new List<string>();
        private List<string> nintexFeatures = new List<string>();
        private List<string> nextDocsFeatures = new List<string>();
        DataGridView data = null;
        DataGridView unfiltered_data = null;
        string filename = null;
        DataGridView filetype = null;

    #endregion

    #region Constructor
        /// <summary>
        /// Constructor. Initializes connection to SharePoint farm. 
        ///  * Builds the list of sites in drop down menu.
        ///  * Generates list of known features.
        /// </summary>
        public UIManager()
        {
            InitializeComponent();
            BuildFeaturesLists();
            data = UIGrid;
            try
            {
                foreach (SPService service in SPFarm.Local.Services)
                {
                    if (service is SPWebService)
                    {
                        foreach (SPWebApplication app in ((SPWebService)service).WebApplications)
                        {
                            if (app.IsAdministrationWebApplication) continue;
                            cbWebApplications.Items.Add(app.GetResponseUri(SPUrlZone.Default).AbsoluteUri);
                        }
                    }
                }
            }
            catch (Exception constr) { MessageBox.Show("Could not access SharePoint Farm!\n" +constr.ToString(), "Error"); }
        }
    #endregion

    #region Menu Items
    //====================================================================================
        private void saveAsToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (saveFileDialog1.ShowDialog() == DialogResult.OK && saveFileDialog1.FileName != "")
            {
                filename = @saveFileDialog1.FileName.ToString();
                filetype = data;
                Save(filename, filetype);
            }
        }
        private void saveToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if ((filename == null && filetype == null) || filetype != data)
            {
                if (saveFileDialog1.ShowDialog() == DialogResult.OK && saveFileDialog1.FileName != "")
                {
                    filename = @saveFileDialog1.FileName.ToString();
                    filetype = data;
                    Save(filename, data);
                }
            }
            else
            {
                Save(filename, filetype);
            }
        }
        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("UIManager v1.0\nCreated by: Mark Shipman (mark.e.shipman@live.com)", "About..",
            MessageBoxButtons.OKCancel, MessageBoxIcon.Asterisk);
        }
        private void closeToolStripMenuItem1_Click(object sender, EventArgs e) 
        {
            if (MessageBox.Show("Are you sure you want to close the program?", "Exit", MessageBoxButtons.YesNo) == DialogResult.Yes)
                this.Close();
        }
        private void copyToolStripMenuItem1_Click(object sender, EventArgs e) { SendKeys.SendWait("^(C)"); }
        private void pasteToolStripMenuItem1_Click(object sender, EventArgs e) { SendKeys.SendWait("^(V)"); }
        private void cutToolStripMenuItem1_Click(object sender, EventArgs e) { SendKeys.SendWait("^(X)"); }
    //====================================================================================    
    #endregion

    #region Buttons
    //====================================================================================
        /// <summary>
        /// Export UI Data Grid to Excel (.xls) file. 
        /// Writes dataGridView table in HTML table format saved to .xls format in order to support saving file on machines without Excel installed.
        /// </summary>
        private void ExportExcelButton_Click(object sender, EventArgs e)
        {
            saveFileDialog1.DefaultExt = "xls";
            saveFileDialog1.ShowDialog();

            if (saveFileDialog1.FileName != "")
            {
                using (System.IO.StreamWriter file = new System.IO.StreamWriter(@saveFileDialog1.FileName.ToString(), false))
                {
                    file.WriteLine("<html><body><table>");

                    // Create header columns
                    file.WriteLine("<tr>");
                    for (int i = 0; i < UIGrid.Columns.Count; i++) 
                        file.WriteLine("<td style=\"background-color:#c0c0c0\">" + UIGrid.Columns[i].Name.ToString() +"</td>");
                    file.WriteLine("</tr>");

                    for (int x = 0; x < UIGrid.Rows.Count-1; x++)
                    {
                        file.WriteLine("<tr>");
                        for (int y = 0; y < UIGrid.Columns.Count; y++)
                        {
                            file.WriteLine("<td style=\"background-color:#ffffff\">"+ UIGrid.Rows[x].Cells[y].Value.ToString() +"</td>");
                        }
                        file.WriteLine("</tr>");
                    }
                    file.WriteLine("</table></body></html>");
                }
            }
            saveFileDialog1.DefaultExt = "html";
        }

        /// <summary>
        /// Next Button to Display more results (if enabled)
        /// </summary>
        private void nextButton_Click(object sender, EventArgs e) { this.DisplayFeatures(); }
    //====================================================================================
    #endregion

    #region UI Manager
    //====================================================================================
        /// <summary>
        /// Display the UI Settings for the currently selected site
        /// </summary>
        private void btnDisplayUISettings_Click(object sender, EventArgs e)
        {
            UIGrid.Rows.Clear();

            SPWebApplication webApp = GetSelectedWebApplication();
            if (webApp == null) return;

            try
            {
                progressBar1.Value = 0;
                collectionLabel.Text = "Total Sites: Loading...Please Wait";
                this.Invalidate();
                this.Update();

                foreach (SPSite site in webApp.Sites)
                {
                    foreach (SPWeb web in site.AllWebs)
                    {
                        int rowIndex = UIGrid.Rows.Add();
                        DataGridViewRow row = UIGrid.Rows[rowIndex];
                        row.Cells["URL"].Value = web.Url.ToString();
                        row.Cells["Theme"].Value = web.Theme.ToString() == "" ? "Default Theme" : web.Theme.ToString();
                        row.Cells["Master_URL"].Value = System.IO.Path.GetFileName(web.MasterUrl.ToString());//.Substring(web.MasterUrl.LastIndexOf('/'));
                        row.Cells["Custom_Master_URL"].Value = System.IO.Path.GetFileName(web.CustomMasterUrl.ToString());//.Substring(web.MasterUrl.LastIndexOf('/'));
                        row.Cells["Alternate_CSS_URL"].Value = string.IsNullOrEmpty(web.AlternateCssUrl) ? "N/A" : web.AlternateCssUrl;
                        web.Dispose();
                    }
                    site.Dispose();
                }
                UIGrid.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
                collectionLabel.Text = "Total Sites: " + (UIGrid.Rows.Count - 1);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
    //====================================================================================
    #endregion

    #region Feature Manager
    //====================================================================================
        /// <summary>
        /// Add Features to the DataGridView for display
        /// </summary>
        private void DisplayFeatures()
        {
            if (featuresToDisplay == null) return;
            if (featuresToDisplay.Count <= 0) return;

            // User specified number of features to display per page
            int numFeaturesToDisplay = featuresToDisplay.Count;
            if (fm_index.Text != string.Empty)
            {
                try
                {
                    numFeaturesToDisplay = int.Parse(fm_index.Text);
                }
                catch
                {
                    fm_indexError.Visible = true;
                    return;
                }
            }

            // Loop end point
            int ceiling = (featuresDisplayed + numFeaturesToDisplay) > featuresToDisplay.Count ? featuresToDisplay.Count : (featuresDisplayed + numFeaturesToDisplay);

            // If not displaying all features, show 'Display More..' button
            if (numFeaturesToDisplay < (featuresToDisplay.Count - featuresDisplayed))
            {
                this.displayMoreFeatures.Visible = true;
            }

            // Add features and Url data to grid
            for (int i = featuresDisplayed; featuresDisplayed < ceiling; featuresDisplayed++)
            {
                int progress = (featuresDisplayed + 1) * 100 / ceiling;
                progressBar1.Value = progress;
                progressLabel.Text = (featuresDisplayed + 1).ToString() + " / " + featuresToDisplay.Count.ToString();

                int rowIndex = fm_grid.Rows.Add();
                DataGridViewRow row = fm_grid.Rows[rowIndex];
                row.Cells["Type"].Value = featuresToDisplay[featuresDisplayed].Type;
                row.Cells["Scope"].Value = featuresToDisplay[featuresDisplayed].Scope;
                row.Cells["FeatureSite"].Value = featuresToDisplay[featuresDisplayed].ParentSiteUrl;
                row.Cells["Feature"].Value = featuresToDisplay[featuresDisplayed].FolderName;
                row.Cells["GUID"].Value = featuresToDisplay[featuresDisplayed].GUID;
            }
            fm_grid.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
        }

        /// <summary>
        /// Get Feature Type 
        /// </summary>
        public FeatureType GetFeatureType(SPFeatureDefinition featureDefinition)
        {
            if (sharePoint2007Features.Contains(featureDefinition.DisplayName))
                return FeatureType.SharePoint;
            if (nintexFeatures.Contains(featureDefinition.DisplayName))
                return FeatureType.Nintex;
            if (nextDocsFeatures.Contains(featureDefinition.DisplayName))
                return FeatureType.NextDocs;
            else return FeatureType.Custom;
        }

        /// <summary>
        /// Feature Search
        /// </summary>
        private void fm_search_Click(object sender, EventArgs e)
        {
            featuresToDisplay.Clear();
            fm_grid.Rows.Clear();
            featuresDisplayed = 0;

            SPWebApplication webApp = GetSelectedWebApplication();
            if (webApp == null) return;

            try
            {
                if (fm_search_textbox.Text == "<GUID / Feature Name>") fm_search_textbox.Clear();
                progressBar1.Value = 0;
                collectionLabel.Text = "Feature Occurences: Loading...Please Wait";
                this.Invalidate();
                this.Update();

                // Farm Features
                if (fm_scope_Farm.Checked)
                {
                    SPWebServiceCollection webServices = new SPWebServiceCollection(SPFarm.Local);
                    foreach (SPWebService webService in webServices)
                    {
                        foreach (SPFeature feature in webService.Features)
                        {
                            AddFeature(fm_search_textbox.Text, feature, FeatureScope.Farm, "N/A", fm_type_Custom.Checked);
                        }
                    }
                }

                // Web Application Features
                if (fm_scope_WebApplication.Checked)
                {
                    foreach (SPFeature feature in webApp.Features)
                    {
                        AddFeature(fm_search_textbox.Text, feature, FeatureScope.WebApplication, webApp.GetResponseUri(SPUrlZone.Default).AbsoluteUri, fm_type_Custom.Checked);
                    }
                }

                foreach (SPSite site in webApp.Sites)
                {
                    // Site Collection features..
                    if (fm_scope_SiteCollection.Checked)
                    {
                        foreach (SPFeature feature in site.Features) AddFeature(fm_search_textbox.Text, feature, FeatureScope.SiteCollection, site.Url, fm_type_Custom.Checked);
                    }

                    // Site features..
                    if (fm_scope_Site.Checked)
                    {
                        foreach (SPWeb web in site.AllWebs)
                        {
                            foreach (SPFeature feature in web.Features) AddFeature(fm_search_textbox.Text, feature, FeatureScope.Site, web.Url, fm_type_Custom.Checked);
                            web.Dispose();
                        }
                    }
                    site.Dispose();
                }
                DisplayFeatures();
                collectionLabel.Text = "Feature Occurences: " + (fm_grid.Rows.Count - 1);
                progressBar1.Value = 0;
            }
            catch (Exception ex) { if (DEBUG_MODE.Checked) MessageBox.Show(ex.ToString()); }
        }

        /// <summary>
        /// Event Handler for initially clicking inside the Feature search box.
        /// </summary>
        private void fm_search_textbox_Enter(object sender, EventArgs e)
        {
            if (fm_search_textbox.Text == "<GUID / Feature Name>")
                fm_search_textbox.Clear();
            else
                fm_search_textbox.SelectAll();
        }

        /// <summary>
        /// Compiles Lists of established resource files. 
        /// Create new resource list text files under Project 'Resources ->  Add..'
        /// then go to 'Properties -> Build Action = Embedded'
        /// </summary>
        private void BuildFeaturesLists()
        {
            System.Reflection.Assembly assembly = System.Reflection.Assembly.GetExecutingAssembly();
            StreamReader stream;

            // Get MOSS 2007 Features
            try
            {
                stream = new StreamReader(assembly.GetManifestResourceStream("MOSS.UIAndFeatureManager.Resources.sharePoint2007Features.txt"));
                while (!stream.EndOfStream)
                {
                    string line = stream.ReadLine();
                    line = line.LastIndexOf(',') > 0 ? line.Substring(0, line.LastIndexOf(',')) : line;
                    sharePoint2007Features.Add(line);
                }
            }
            catch
            {
                if (DEBUG_MODE.Checked)
                    MessageBox.Show("Error accessing SharePoint 2007 Features resource!");

            }

            // Get NextDocs Features
            try
            {
                stream = new StreamReader(assembly.GetManifestResourceStream("MOSS.UIAndFeatureManager.Resources.nextDocsFeatures.txt"));
                while (!stream.EndOfStream)
                {
                    nextDocsFeatures.Add(stream.ReadLine());
                }
            }
            catch
            {
                if (DEBUG_MODE.Checked)
                    MessageBox.Show("Error accessing NextDocs Features resource!");
            }

            // Get Nintex Features
            try
            {
                stream = new StreamReader(assembly.GetManifestResourceStream("MOSS.UIAndFeatureManager.Resources.nintexFeatures.txt"));
                while (!stream.EndOfStream)
                {
                    nintexFeatures.Add(stream.ReadLine());
                }
            }
            catch
            {
                if (DEBUG_MODE.Checked)
                    MessageBox.Show("Error accessing Nintex Features resource!");
            }

            // List of all resource assembly file names
            //var auxList= System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceNames();
        }

        /// <summary>
        /// Add a feature to the Features List
        /// </summary>
        public void AddFeature(string featureToCheck, SPFeature activatedFeature, FeatureScope activatedFeatureScope, string parentUrl, bool customFeaturesOnly)
        {
            if (activatedFeature == null) return;
            if (activatedFeature.Definition == null) return;
            try
            {
                if (featureToCheck.Equals(activatedFeature.Definition.DisplayName, StringComparison.InvariantCultureIgnoreCase) ||
                    featureToCheck.Equals(activatedFeature.Definition.Name, StringComparison.InvariantCultureIgnoreCase) ||
                    featureToCheck.Equals(activatedFeature.Definition.Id.ToString(), StringComparison.InvariantCultureIgnoreCase) ||
                    featureToCheck == "<GUID / Feature Name>" ||
                    featureToCheck == string.Empty)
                {
                    FeatureType featureType = GetFeatureType(activatedFeature.Definition);
                    if (!fm_type_Custom.Checked && featureType == FeatureType.Custom) return;
                    if (!fm_type_ThirdParty.Checked && (featureType == FeatureType.NextDocs || featureType == FeatureType.Nintex)) return;
                    if (!fm_type_SharePoint2007.Checked && featureType == FeatureType.SharePoint) return;

                    FeatureToDisplay featureToDisplay = new FeatureToDisplay();
                    featureToDisplay.Type = featureType;
                    featureToDisplay.Scope = activatedFeatureScope;
                    featureToDisplay.ParentSiteUrl = parentUrl;
                    featureToDisplay.FolderName = activatedFeature.Definition.DisplayName;
                    featureToDisplay.GUID = activatedFeature.Definition.Id;
                    featuresToDisplay.Add(featureToDisplay);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
    //====================================================================================
    #endregion Feature Manager

    #region Web Part Manager
    //====================================================================================
        /// <summary>
        /// Event Handler for entering the WP Search Textbox
        /// </summary>
        private void wp_search_textbox_Enter(object sender, EventArgs e)
        {
            if (wp_webpart_search.Text == "<Web Part Search>")
                wp_webpart_search.Clear();
            else
                wp_webpart_search.SelectAll();
        }

        /// <summary>
        /// Web Part Search button
        /// </summary>
        private void wp_search_click(object sender, EventArgs e)
        {
            webPartGrid.Rows.Clear();

            SPWebApplication webApp = GetSelectedWebApplication();
            if (webApp == null) return;

            if (wp_activeFeatures.Checked)
                DisplayActiveWebParts(webApp);
            if (wp_featureTemplates.Checked)
                DisplayWebPartTemplates(webApp);
        }

        /// <summary>
        /// Display all Web Part Templates to the Web Part grid for a specified Web Application
        /// </summary>
        private void DisplayWebPartTemplates(SPWebApplication webApp)
        {
            try
            {
                if (webPartGrid.Columns.Contains("Zone")) webPartGrid.Columns.Remove("Zone");
                if (webPartGrid.Columns.Contains("WebPartTitle")) webPartGrid.Columns.Remove("WebPartTitle");
                if (webPartGrid.Columns.Contains("Description")) webPartGrid.Columns.Remove("Description");

                progressBar1.Value = 0;
                collectionLabel.Text = "Total Web Parts: Loading...Please Wait";
                this.Invalidate();
                this.Update();

                // Site Collection Web Part Templates
                foreach (SPSite site in webApp.Sites)
                {
                    SPList siteParts = site.GetCatalog(SPListTemplateType.WebPartCatalog);
                    {
                        foreach (SPListItem item in siteParts.Items)
                        {
                            if ((siteParts != null && wp_webpart_search.Text.Contains(item.Name) ||
                                wp_webpart_search.Text == "" ||
                                wp_webpart_search.Text == "<Web Part Search>" ) &&
                                (item.Web.Url.StartsWith(wp_url_search.Text) ||
                                wp_url_search.Text == "<Site Collection / Site>" ||
                                wp_url_search.Text == "" ))
                            {
                                int rowIndex = webPartGrid.Rows.Add();
                                DataGridViewRow row = webPartGrid.Rows[rowIndex];
                                row.Cells["webURL"].Value = item.Web.Url;
                                row.Cells["WebPart"].Value = item.Name;
                                row.Cells["webType"].Value = item.Name.EndsWith(".dwp") ? "SharePoint" : "ASP.NET";
                                row.Cells["WebScope"].Value = "Site Collection";
                            }
                        }
                    }

                    // Site Web Part Templates
                    foreach (SPWeb web in site.AllWebs)
                    {
                        SPList webParts = web.GetCatalog(SPListTemplateType.WebPartCatalog);
                        if (webParts != null)
                        {
                            foreach (SPListItem item in webParts.Items)
                            {
                                if (siteParts != null && wp_webpart_search.Text == item.Name ||
                                    wp_webpart_search.Text == item.Web.Url ||
                                    wp_webpart_search.Text == "" ||
                                    wp_webpart_search.Text == "<Web Part Search>")
                                {
                                    int rowIndex = webPartGrid.Rows.Add();
                                    DataGridViewRow row = webPartGrid.Rows[rowIndex];
                                    row.Cells["webURL"].Value = item.Web.Url;
                                    row.Cells["WebPart"].Value = item.Name;
                                    row.Cells["webType"].Value = item.Name.EndsWith(".dwp") ? "SharePoint" : ".NET";
                                    row.Cells["WebScope"].Value = "Site";
                                }
                            }
                        }
                        web.Dispose();
                    }
                    site.Dispose();
                }
                webPartGrid.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
                collectionLabel.Text = "Total Web Parts: " + (webPartGrid.Rows.Count - 1);
            }
            catch (Exception ex)
            {
                collectionLabel.Text = "Total Web Parts: 0";
                if (DEBUG_MODE.Checked)
                {
                    MessageBox.Show(ex.ToString());
                }
            }
        }

        /// <summary>
        /// Display all active Web Parts to the Web Part grid for the specified Web Application
        /// </summary>
        private void DisplayActiveWebParts(SPWebApplication webApp)
        {
            try
            {
                // Process search criteria
                bool includeSubSites = wp_type_subsites.Checked;
                bool includeSharePointWebParts = wp_type_SharePoint2007.Checked;
                bool includeCustomWebParts = wp_type_custom.Checked;
                string urlFilter = wp_url_search.Text == "<Site Collection / Site>" ? string.Empty : wp_url_search.Text.Trim();
                string keywordFilter = wp_webpart_search.Text == "<Web Part Search>" ? string.Empty : wp_webpart_search.Text.Trim().ToLower();

                try
                {
                    // Validate URL filter format
                    if (urlFilter != string.Empty) new Uri(urlFilter);
                }
                catch (UriFormatException)
                {
                    MessageBox.Show("Please enter a valid URL filter", "URL Format Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                if (urlFilter == string.Empty || includeSubSites)
                {
                    DialogResult dialogResult = MessageBox.Show("Your search criteria may generate a large number of results, do you want to continue?", "Search Criteria Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                    if (dialogResult != DialogResult.Yes) return;
                }

                collectionLabel.Text = "Total Web Parts: Loading...Please Wait";
                this.Invalidate();
                this.Update();

                WebPartUtilities webpartUtilities = new WebPartUtilities();
                List<WebPartToDisplay> allWebParts = new List<WebPartToDisplay>();
                if (urlFilter != string.Empty)
                {
                    // Apply a URL filter
                    using (SPSite site = new SPSite(urlFilter))
                    {
                        using (SPWeb web = site.OpenWeb())
                        {
                            allWebParts = webpartUtilities.GetAllWebPartsInWeb(web, includeSubSites, includeSharePointWebParts, includeCustomWebParts);
                        }
                    }
                }
                else
                {
                    // No URL filter, go through each site collection
                    foreach (SPSite site in webApp.Sites)
                    {
                        using (SPWeb web = site.RootWeb)
                        {
                            allWebParts.AddRange(webpartUtilities.GetAllWebPartsInWeb(web, includeSubSites, includeSharePointWebParts, includeCustomWebParts));
                        }
                        site.Dispose();
                    }
                }

                if (!webPartGrid.Columns.Contains("Zone")) webPartGrid.Columns.Add("Zone", "Zone");
                if (!webPartGrid.Columns.Contains("WebPartTitle")) webPartGrid.Columns.Add("WebPartTitle", "Title");
                if (!webPartGrid.Columns.Contains("Description")) webPartGrid.Columns.Add("Description", "Description");

                foreach (WebPartToDisplay webpartToDisplay in allWebParts)
                {
                    if (keywordFilter != string.Empty)
                    {
                        if (!webpartToDisplay.Title.ToLower().Contains(keywordFilter) &&
                            !webpartToDisplay.Type.ToLower().Contains(keywordFilter))
                            continue;
                    }

                    int rowIndex = webPartGrid.Rows.Add();
                    DataGridViewRow row = webPartGrid.Rows[rowIndex];
                    row.Cells["webURL"].Value = webpartToDisplay.PageUrl;
                    row.Cells["WebPart"].Value = webpartToDisplay.Type;
                    row.Cells["webType"].Value = webpartToDisplay.Category;
                    row.Cells["WebScope"].Value = webpartToDisplay.Visible ? "Visible" : "Hidden";
                    row.Cells["Description"].Value = webpartToDisplay.Description;
                    row.Cells["Zone"].Value = webpartToDisplay.Zone;
                    row.Cells["WebPartTitle"].Value = webpartToDisplay.Title;
                }

                webPartGrid.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
                collectionLabel.Text = "Total Web Parts: " + (webPartGrid.Rows.Count - 1);
            }
            catch (Exception ex)
            {
                collectionLabel.Text = "Total Web Parts: 0";
                if (DEBUG_MODE.Checked)
                {
                    MessageBox.Show("Error in DisplayActiveWebParts:\n" + ex.ToString());
                }
            }
        }
    //====================================================================================
    #endregion

    #region Event Handlers
    //====================================================================================
        /// <summary>
        /// Get the selected SPWebApplication from drop down menu
        /// </summary>
        public SPWebApplication GetSelectedWebApplication()
        {
            siteSelectErr.Visible = false;
            if (cbWebApplications.Text == "<Web Application>" ||
                cbWebApplications.SelectedItem == null)
            {
                siteSelectErr.Visible = true;
                return null;
            }

            SPWebApplication webApp = null;
            foreach (SPService service in SPFarm.Local.Services)
            {
                if (service is SPWebService)
                {
                    foreach (SPWebApplication app in ((SPWebService)service).WebApplications)
                    {
                        if (app.IsAdministrationWebApplication) continue;
                        if (app.GetResponseUri(SPUrlZone.Default).AbsoluteUri == cbWebApplications.SelectedItem.ToString())
                        {
                            webApp = app;
                            break;
                        }
                    }
                }
            }
            return webApp;
        }

        /// <summary>
        /// Event Handler for button presses inside the Feature search box.  
        /// </summary>
        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                this.fm_search_Click(sender, e);
        }

        /// <summary>
        /// Event Handler for switching Tabs inside the form.
        /// </summary>
        private void tabControl1_Selecting(object sender, TabControlCancelEventArgs e)
        {
            if (e.TabPage.Name.Equals("Features"))
            {
                this.collectionLabel.Text = "Feature Occurences: " + (fm_grid.Rows.Count - 1);
                data = fm_grid;
            }
            else if (e.TabPage.Name.Equals("SiteCollections"))
            {
                this.collectionLabel.Text = "Total Sites: " + (UIGrid.Rows.Count - 1);
                data = UIGrid;
            }
            else if (e.TabPage.Name.Equals("webPartManager"))
            {
                this.collectionLabel.Text = "Web Parts: " + (webPartGrid.Rows.Count - 1);
                data = webPartGrid;
            }
        }
        /// <summary>
        /// EventHandler for mouse clicks on the Features Grid
        /// </summary>
        private void grid_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            DataGridView dgv = (DataGridView)sender;
            if (e.Button == MouseButtons.Right)
            {
                data.ContextMenuStrip = contextMenu;
                data.ContextMenuStrip.Show(this, this.PointToClient(Cursor.Position));
            }
        }

        /// <summary>
        /// EventHandler for mouse clicks on the UI Grid
        /// </summary>
        private void UIGrid_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            DataGridView dgv = (DataGridView)sender;
            if (e.Button == MouseButtons.Right)
            {
                UIGrid.ContextMenuStrip = contextMenu;
                UIGrid.ContextMenuStrip.Show(this, this.PointToClient(Cursor.Position));
            }
        }

        /// <summary>
        /// Event Handler for context menu opening
        /// </summary>
        private void contextMenu_Opening(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (unfiltered_data != null && unfiltered_data.Parent == data.Parent)
                removeFiltersToolStripMenuItem.Enabled = true;
            else
                removeFiltersToolStripMenuItem.Enabled = false;
        }


        /// <summary>
        /// Select current DataGridViewCell when perfoming a Context Menu right-click.  
        /// Required as a work around for bug in VS2008 3.5
        /// </summary>
        private void grid_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
                data.CurrentCell = data.Rows[e.RowIndex].Cells[e.ColumnIndex];
        }
    //====================================================================================
    #endregion

    #region Grid Context Menu
    //====================================================================================
        /// <summary>
        /// Filter the features grid by the selected cell value
        /// </summary>
        private void filterByValueToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                if ((unfiltered_data == null && data.Rows.Count > 1) || (unfiltered_data != null && data.Rows.Count > 1 && unfiltered_data.Parent != data.Parent))
                    unfiltered_data = Copy(data);
                DataGridViewCell current = data.CurrentCell;
                int colIndex = data.CurrentCell.ColumnIndex;
                int max = data.Rows.Count - 2;

                for (int i = max; i >= 0; i--)
                {
                    if (data.Rows[i].Cells[colIndex].Value.ToString() != current.Value.ToString())
                    {
                        data.Rows.RemoveAt(i);
                    }
                }
            }
            catch (Exception filter) { if (DEBUG_MODE.Checked) MessageBox.Show(filter.ToString()); }
        }

        /// <summary>
        /// Remove the applied context filter(s)
        /// </summary>
        private void removeFiltersToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (unfiltered_data != null && data.Parent == unfiltered_data.Parent)
                {
                    data.Rows.Clear();
                    data.Rows.Add(unfiltered_data.Rows.Count - 1);
                    foreach (DataGridViewRow row in unfiltered_data.Rows)
                    {
                        foreach (DataGridViewCell cell in unfiltered_data.Rows[row.Index].Cells)
                        {
                            data.Rows[row.Index].Cells[cell.ColumnIndex].Value = cell.Value;
                        }
                    }
                    unfiltered_data = null;
                }
            }
            catch (Exception removeFilter) { if (DEBUG_MODE.Checked) MessageBox.Show(removeFilter.ToString()); }
        }

        private void copyToolStripMenuItem2_Click(object sender, EventArgs e) { Clipboard.SetText(data.CurrentCell.Value.ToString()); }
        private void selectAllToolStripMenuItem_Click(object sender, EventArgs e) { data.SelectAll(); }
    //====================================================================================
    #endregion

    #region Helper Methods
    //====================================================================================
        /// <summary>
        /// Save DataGridView to file (HTML Table)
        /// </summary>
        private void Save(string filename, DataGridView data)
        {
            using (System.IO.StreamWriter file = new System.IO.StreamWriter(filename, false))
            {
                file.WriteLine("<html><body><table>");

                // Create header columns
                file.WriteLine("<tr>");
                for (int i = 0; i < data.Columns.Count; i++)
                    file.WriteLine("<td style=\"background-color:#c0c0c0\">" + data.Columns[i].Name.ToString() + "</td>");
                file.WriteLine("</tr>");

                for (int x = 0; x < data.Rows.Count - 1; x++)
                {
                    file.WriteLine("<tr>");
                    for (int y = 0; y < data.Columns.Count; y++)
                    {
                        file.WriteLine("<td style=\"background-color:#ffffff\">" + data.Rows[x].Cells[y].Value.ToString() + "</td>");
                    }
                    file.WriteLine("</tr>");
                }
                file.WriteLine("</table></body></html>");
            }
        }

        /// <summary>
        /// Perform a deep copy of a specified DataGridView table and return the copy.
        /// </summary>
        private DataGridView Copy(DataGridView dg)
        {
            DataGridView copy = new DataGridView();
            foreach (DataGridViewColumn column in dg.Columns)
            {
                copy.Columns.Add(column.Name, column.HeaderText);
            }

            copy.Rows.Add(dg.Rows.Count - 1);
            foreach (DataGridViewRow row in dg.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    copy.Rows[row.Index].Cells[cell.ColumnIndex].Value = dg.Rows[row.Index].Cells[cell.ColumnIndex].Value;
                }
            }
            copy.Parent = dg.Parent;
            return copy;
        }
    //====================================================================================
    #endregion
}
//====================================================================================
#endregion

#region Feature Class
//====================================================================================
    /// <summary>
    /// Feature Types
    /// </summary>
    public enum FeatureType
    {
        SharePoint,
        NextDocs,
        Nintex,
        Custom
    }
    
    /// <summary>
    /// Feature Scopes
    /// </summary>
    public enum FeatureScope
    {
        Site,
        SiteCollection,
        WebApplication,
        Farm
    }

    /// <summary>
    /// Class representing a Feature
    /// </summary>
    public class FeatureToDisplay
    {
        public FeatureType Type { get; set; }
        public FeatureScope Scope { get; set; }
        public string ParentSiteUrl { get; set; }
        public string FolderName { get; set; }
        public Guid GUID { get; set; }
    }
//====================================================================================
#endregion

}
