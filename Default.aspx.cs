using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Web.UI;
using Microsoft.SharePoint.Client;
using System.Net;

namespace WebApplication6
{
    public partial class _Default : Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            {
                string assetid = Request.QueryString["AssetID"];
                if (!string.IsNullOrEmpty(assetid))
                {
                    LoadData(assetid);
                    // ErrorMessage.Visible = true;
                    // ErrorMessage.Text = "Details for Asset ID:" + assetid;
                    ContentPanel.Visible = true;
                }
                else
                {
                    ErrorMessage.Visible = true;
                    ErrorMessage.Text = "No Asset ID Supplied";
                    ContentPanel.Visible = false;
                    //MainPanel.Update();
                }
                // MainPanel.Update();
            }
        }
        private void LoadData(string assetid)
        {
            var spUrl = Config.GetConfigvalue("SharePointURL");
            var spAssetList = Config.GetConfigvalue("SharePointAssetListName");
            var imageTempFolder = Config.GetConfigvalue("ImageTempFolder");

            ClientContext clientContext = new ClientContext(spUrl);
            clientContext.Credentials = new NetworkCredential("Avinash","Avinash$2468");
            Web oWebsite = clientContext.Web;
            ListCollection collList = oWebsite.Lists;
            List oList = collList.GetByTitle(spAssetList); //"Asset Register");

            CamlQuery camlQuery = new CamlQuery();
            camlQuery.ViewXml = string.Format("<View><Query><Where><Eq>" +
                                "<FieldRef Name='ID'/><Value Type='Number'>" + assetid + "</Value>" +
                                "</Eq></Where></Query><RowLimit>50</RowLimit></View>");
            Microsoft.SharePoint.Client.ListItemCollection collListItem = oList.GetItems(camlQuery);
            clientContext.Load(collListItem,
                items => items.IncludeWithDefaultProperties(
                    item => item.DisplayName));
            clientContext.ExecuteQuery();
            if (collListItem.Count > 0)
            {
                Microsoft.SharePoint.Client.ListItem oListItem = collListItem[0];
                //get first record, should only be one as we finding by ID
                // foreach (Microsoft.SharePoint.Client.ListItem oListItem in collListItem)
                // {
                //Deprecated
                // this.Panel1.Controls.Add(new LiteralControl("<br/><br/><br/><strong>TAB1 (Details)</strong><br/><br/><br/>"));
               #region "Image Code"
               try
               {
                    //Fix to overcome issue on krith server where IP address is not known on server, need to use machine name
                    if (oListItem.FieldValues.ContainsKey("Picture_x0020_Link"))
                    {
                        if (oListItem["Picture_x0020_Link"] != null)
                        {
                            //string picUrl =
                            //    ((Microsoft.SharePoint.Client.FieldUrlValue)(oListItem["Picture_x0020_Link"])).ToString(); /*oListItem["Picture_x0020_Link"].ToString();*/
                            string picUrl = ((Microsoft.SharePoint.Client.FieldUrlValue)(oListItem["Picture_x0020_Link"])).Url;
                            if (picUrl.Contains(""))
                           {
                               picUrl.Replace("http://illovo.gwsa.co.za/", "ILLOVO");
                                //picUrl.Replace("http://illovosp2010/", "sp2013"); 
                                //This is a terrible hack, dont do this ever again :-)
                            }
                            //  Uri uri2 = new Uri(@picUrl); // Code original broke here. Added UrlBuilder for path of pic on sharepoint 
                            UriBuilder builder = new UriBuilder(@picUrl);
                            Uri uri = builder.Uri;
                         
                            string server = uri.AbsoluteUri.Replace(uri.AbsolutePath, "");
                            string serverabsolute = uri.AbsolutePath;
                            //filename = uri.Segments[2].ToString();
                            FileInformation f = Microsoft.SharePoint.Client.File.OpenBinaryDirect(clientContext, serverabsolute);
                        
                            string[] fileparts = uri.LocalPath.Split('/');
                            var filename = fileparts[fileparts.Length - 1];

                            using (var fileStream = new FileStream((imageTempFolder + filename), FileMode.Create))
                            {
                                f.Stream.CopyTo(fileStream);
                            }
                           //Create Thumb
                           Bitmap bmp = CreateThumbnail(imageTempFolder + filename, 150, 150);

                           if (bmp == null)
                           {
                               //ErrorResult();
                               return;
                           }
                           string outputFilename = null;
                           outputFilename = imageTempFolder + "thumb_" + filename;
                           bmp.Save(outputFilename);

                           var image = new System.Web.UI.WebControls.Image();
                           image.ImageUrl = @"\Images\thumb_" + filename;
                           image.Width = System.Web.UI.WebControls.Unit.Pixel(150);
                           image.AlternateText = "Image of Asset";
                           this.ImagePanel.Controls.Add(image);
                       
                           //this.spaimage.ImageUrl = imagepath;
                           //AssetThumbnail.ImageUrl = image.ImageUrl.ToString();
                           //this.Panel1.Controls.Add(image);
                       }
                   }
               }
               catch (Exception ex)
               {
                   ErrorMessage.Visible = true;
                   ErrorMessage.Text = "No image found for Asset ID: " + assetid;
                   ErrorDetails.Text = ex.ToString();
                   ContentPanel.Visible = true;
                   // MainPanel.Update();
               }
                #endregion
      
        #region "TAB 1 (Details) - Fields"
        string description = "";
                if (oListItem.FieldValues.ContainsKey("Title"))
                {
                    if (oListItem["Title"] != null)
                    {
                        description = oListItem["Title"].ToString() + " ";

                    }
                }
                string location = "";
                if (oListItem.FieldValues.ContainsKey("Location"))
                {
                    if (oListItem["Location"] != null)
                    {
                        location = oListItem["Location"].ToString() + " ";
                    }
                }
                string level = "";
                if (oListItem.FieldValues.ContainsKey("Level"))
                {
                    if (oListItem["Level"] != null)
                    {
                        level = oListItem["Level"].ToString();
                    }
                }
                string model = "";
                if (oListItem.FieldValues.ContainsKey("Model_x0020_Number"))
                {
                    if (oListItem["Model_x0020_Number"] != null)
                    {
                        model = model + oListItem["Model_x0020_Number"].ToString();
                    }
                }
                if (oListItem.FieldValues.ContainsKey("Serial_x0020_Number"))
                {
                    if (oListItem["Serial_x0020_Number"] != null)
                    {
                        model = model + " / " + oListItem["Serial_x0020_Number"].ToString();
                    }
                }
                string assetType = "";
                if (oListItem.FieldValues.ContainsKey("Asset_x0020_Type"))
                {
                    if (oListItem["Asset_x0020_Type"] != null)
                    {
                        assetType = assetType + oListItem["Asset_x0020_Type"].ToString();
                    }
                }
                string itemType = "";
                if (oListItem.FieldValues.ContainsKey("Item_x0020_Type"))
                {
                    if (oListItem["Item_x0020_Type"] != null)
                    {
                        itemType = itemType + oListItem["Item_x0020_Type"].ToString();
                    }
                }

                DescriptionField.Text = description;
                LocationField.Text = location + " " + level + " ";
                ModelNumberField.Text = model;
                AssetTypeField.Text = assetType;
                //ItemTypeField.Text = _item_type;
                #endregion
        #region "TAB 2 (More) - Fields"
                
                string supplier = "";
                if (oListItem.FieldValues.ContainsKey("Supplier"))
                {
                    if (oListItem["Supplier"] != null)
                    {
                        supplier = oListItem["Supplier"].ToString() + " ";
                    }
                }
                string _ordernum = "";
                if (oListItem.FieldValues.ContainsKey("Order_x0020_Number"))
                {
                    if (oListItem["Order_x0020_Number"] != null)
                    {
                        _ordernum = oListItem["Order_x0020_Number"].ToString();
                    }
                }
                string _assetid = "";
                if (oListItem.FieldValues.ContainsKey("Asset_x0020_ID"))
                {
                    if (oListItem["Asset_x0020_ID"] != null)
                    {
                        _assetid = _assetid+ oListItem["Asset_x0020_ID"].ToString();
                    }
                }
               
                string _quantity = "";
                if (oListItem.FieldValues.ContainsKey("Quantity"))
                {
                    if (oListItem["Quantity"] != null)
                    {
                        _quantity = oListItem["Quantity"].ToString();
                    }
                }

                string _purchaseprice = "";
                if (oListItem.FieldValues.ContainsKey("Purchase_x0020_Price"))
                {
                    if (oListItem["Purchase_x0020_Price"] != null)
                    {
                        _purchaseprice = oListItem["Purchase_x0020_Price"].ToString();
                    }
                }
                SupplierField.Text = supplier;
                AssetIDField.Text = _assetid;
                OrderNumberField.Text = _ordernum;
                QuantityField.Text = _quantity;
                //OtherInformationField.Text = _otherinfo;

                #endregion
            }
            else
            {
                ErrorMessage.Text = "Asset not found matching provided ID";
                // MainPanel.Update();
            }
        }
        public static Bitmap CreateThumbnail(string lcFilename, int lnWidth, int lnHeight)
        {
            Bitmap bmpOut = null;
            try
            {
                Bitmap loBMP = new Bitmap(lcFilename);
                ImageFormat loFormat = loBMP.RawFormat;
                decimal lnRatio;
                int lnNewWidth = 0;
                int lnNewHeight = 0;

                //*** If the image is smaller than a thumbnail just return it
                if (loBMP.Width < lnWidth && loBMP.Height < lnHeight)
                    return loBMP;

                if (loBMP.Width > loBMP.Height)
                {
                    lnRatio = (decimal) lnWidth/loBMP.Width;
                    lnNewWidth = lnWidth;
                    decimal lnTemp = loBMP.Height*lnRatio;
                    lnNewHeight = (int) lnTemp;
                }
                else
                {
                    lnRatio = (decimal) lnHeight/loBMP.Height;
                    lnNewHeight = lnHeight;
                    decimal lnTemp = loBMP.Width*lnRatio;
                    lnNewWidth = (int) lnTemp;
                }
                // System.Drawing.Image imgOut = 
                //      loBMP.GetThumbnailImage(lnNewWidth,lnNewHeight,
                //                              null,IntPtr.Zero);
                // *** This code creates cleaner (though bigger) thumbnails and properly
                // *** and handles GIF files better by generating a white background for
                // *** transparent images (as opposed to black)
                bmpOut = new Bitmap(lnNewWidth, lnNewHeight);
                Graphics g = Graphics.FromImage(bmpOut);
                g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                g.FillRectangle(Brushes.White, 0, 0, lnNewWidth, lnNewHeight);
                g.DrawImage(loBMP, 0, 0, lnNewWidth, lnNewHeight);

                loBMP.Dispose();
            }
            catch
            {
                return null;
            }
            return bmpOut;
        }
    }
}