using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.WebParts;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using SPX = Microsoft.SharePoint.Client;

namespace HB.SP.Provisioning.Tool
{
    class Program
    {
        private static string url, title, description, fieldXmlFormat, internalName, jsLinkPath;
        private static bool exit = false;
        static void Main(string[] args)
        {
            try
            {
                while (exit != true)
                {
                    Console.WriteLine("Choose one item from menu:");
                    Console.WriteLine("1-Create list.");
                    Console.WriteLine("2-Delete list.");
                    Console.WriteLine("3-Add field to list.");
                    Console.WriteLine("4-Retrieve fields from list.");
                    Console.WriteLine("5-Hide field from list.");
                    Console.WriteLine("6-Add JSLink to list.");
                    Console.WriteLine("7-Add fields from xml schema.");
                    Console.WriteLine("8-Add page to list.");
                    Console.WriteLine("Press 0 to exit");
                    string choice = Console.ReadLine();

                    switch (choice)
                    {
                        case "0":
                            exit = true;
                            break;
                        case "1":
                            Console.WriteLine("Insert the SharePoint site url to create :");
                            url = Console.ReadLine();
                            Console.WriteLine("Insert the list's title :");
                            title = Console.ReadLine();
                            Console.WriteLine("Insert the list's description : (Optional parameter)");
                            description = Console.ReadLine();
                            CreateList(url, title, description);
                            break;
                        case "2":
                            Console.WriteLine("Insert the SharePoint site url to delete :");
                            url = Console.ReadLine();
                            Console.WriteLine("Insert the list's title :");
                            title = Console.ReadLine();
                            DeleteList(url, title);
                            break;
                        case "3":
                            Console.WriteLine("Insert the SharePoint site url :");
                            url = Console.ReadLine();
                            Console.WriteLine("Insert the list's title :");
                            title = Console.ReadLine();
                            Console.WriteLine("Insert the xml format of the field :");
                            fieldXmlFormat = Console.ReadLine();
                            AddFieldToList(url, title, fieldXmlFormat);
                            break;
                        case "4":
                            //DrawProgressBar(0, 100, 100, '§');
                            Console.WriteLine("Insert the SharePoint site url :");
                            url = Console.ReadLine();
                            Console.WriteLine("Insert the list's title :");
                            title = Console.ReadLine();
                            RetrieveFieldsFromList(url, title);
                            break;
                        case "5":
                            Console.WriteLine("Insert the SharePoint site url :");
                            url = Console.ReadLine();
                            Console.WriteLine("Insert the list's title :");
                            title = Console.ReadLine();
                            Console.WriteLine("Insert the internal name:");
                            internalName = Console.ReadLine();
                            HideField(url, title, internalName);
                            break;
                        case "6":
                            Console.WriteLine("Insert the SharePoint site url :");
                            url = Console.ReadLine();
                            Console.WriteLine("Insert the list's title :");
                            title = Console.ReadLine();
                            Console.WriteLine("Insert the JsLink Path:");
                            jsLinkPath = Console.ReadLine();
                            AddJsLinkToList(url, title, jsLinkPath);
                            break;
                        case "7":
                            Console.WriteLine("Insert the SharePoint site url :");
                            url = Console.ReadLine();
                            Console.WriteLine("Insert the list's title :");
                            title = Console.ReadLine();
                            AddFieldsToList(url, title);
                            break;
                        case "8":
                            Console.WriteLine("Insert the SharePoint site url :");
                            url = Console.ReadLine();
                            PushPageToList(url);
                            break;
                        default:
                            Console.WriteLine("Character not recognized !");
                            break;
                    }

                    DrawProgressBar(100, 100, 100, '§');

                    //Console.WriteLine("Press enter to close...");
                    //Console.ReadLine();
                }

            }
            catch (Exception ex)
            {
                Console.Write(ex.Message);
                Console.WriteLine("Press enter to close...");
                Console.ReadLine();
                throw;
            }
        }

        #region Page
        public static bool PushPageToList(string url)
        {
            bool result = false;

            ClientContext context = new ClientContext(url);

            var sitePageLib = context.Web.Lists.GetByTitle("Site Pages");
            FileCreationInformation supportFileInfo = new FileCreationInformation();
            supportFileInfo.Url = "ProjectRoom365.aspx";
            supportFileInfo.Overwrite = true;

            #region Support-Content

            string supportContent = @"
                      <%@ Page language=""C#"" MasterPageFile=""~masterurl/default.master"" Inherits=""Microsoft.SharePoint.WebPartPages.WebPartPage,Microsoft.SharePoint,Version=15.0.0.0,Culture=neutral,PublicKeyToken=71e9bce111e9429c"" meta:webpartpageexpansion=""full"" meta:progid=""SharePoint.WebPartPage.Document"" %>
                      <%@ Register Tagprefix=""SharePoint"" Namespace=""Microsoft.SharePoint.WebControls"" Assembly=""Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c"" %> 
                      <%@ Register Tagprefix=""Utilities"" Namespace=""Microsoft.SharePoint.Utilities"" Assembly=""Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c"" %> 
                      <%@ Import Namespace=""Microsoft.SharePoint"" %> 
                      <%@ Assembly Name=""Microsoft.Web.CommandUI, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c"" %> 
                      <%@ Register Tagprefix=""WebPartPages"" Namespace=""Microsoft.SharePoint.WebPartPages"" Assembly=""Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c"" %>

                      <asp:Content ContentPlaceHolderId=""PlaceHolderPageTitle"" runat=""server"">
                          Project Room 365 for External Partners - Dashboard
                      </asp:Content>
                      <asp:Content ContentPlaceHolderId=""PlaceHolderAdditionalPageHead"" runat=""server"">
                          <meta name=""GENERATOR"" content=""Microsoft SharePoint"" />
                          <meta name=""ProgId"" content=""SharePoint.WebPartPage.Document"" />
                          <meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" />
                          <meta name=""CollaborationServer"" content=""SharePoint Team Web Site"" />

                          <SharePoint:CssRegistration runat=""server"" Name=""<% $SPUrl:~sitecollection/TemplateFiles/hb.support.css%>"" After=""corev15.css"" />

                          <SharePoint:ScriptLink runat=""server"" ID=""jQuery"" Language=""javascript"" Name=""~sitecollection/Style Library/jquery-1.11.3.js""/>
                          <SharePoint:ScriptLink runat=""server"" ID=""ScriptLinkAdal"" Language=""javascript"" Name=""~sitecollection/Style Library/adal.js""/>

                          <SharePoint:ScriptBlock runat=""server"">
                              var navBarHelpOverrideKey = ""WSSEndUser""; 
                          </SharePoint:ScriptBlock>

                              <SharePoint:StyleBlock runat=""server"">
                              body #s4-leftpanel {
                                  display:none;
                              }

                              .s4-ca {
                                  margin-left:0px;
                              }
                          </SharePoint:StyleBlock>
                          </asp:Content>
                          <asp:Content ContentPlaceHolderId=""PlaceHolderSearchArea"" runat=""server"">
                          <SharePoint:DelegateControl runat=""server"" ControlId=""SmallSearchInputBox""/>
                          </asp:Content>
                          <asp:Content ContentPlaceHolderId=""PlaceHolderPageTitleInTitleArea"" runat=""server"">
                            Project Room 365 for External Partners - Dashboard
                          </asp:Content>
                          <asp:Content ContentPlaceHolderId=""PlaceHolderPageDescription"" runat=""server"">
                          <SharePoint:ProjectProperty Property=""Description"" runat=""server""/>
                          </asp:Content>
                          <asp:Content ContentPlaceHolderId=""PlaceHolderMain"" runat=""server"">
                            <div class=""row"" id=""_main"">
                                <div>
                                    <a href=""javascript:;"" id=""signInLink"">Sign In</a>
                                    <a href=""javascript:;"" id=""signOutLink"">Sign Out</a>
                                    <p>
                                        <a href=""javascript:;"" id=""_getGroups"">Get groups</a>
                                    </p>
                                </div>
                                <div>
                                    <p id=""loginMessage""></p>
                                </div>
                            <div>

                          <SharePoint:ScriptLink runat=""server"" ID=""ScriptLinkAdal"" Language=""javascript"" Name=""~sitecollection/Style Library/hb.projectroom365dashboard.js""/>
                          <footer></footer>        
                      </asp:Content>";

            #endregion

            supportFileInfo.Content = Encoding.ASCII.GetBytes(supportContent);
            Microsoft.SharePoint.Client.File supportFile = sitePageLib.RootFolder.Files.Add(supportFileInfo);

            context.ExecuteQuery();

            result = true;
            return result;
        }
        #endregion
        #region List
        private static bool CreateList(string url, string title, string description = "")
        {
            bool result = false;

            // Starting with ClientContext, the constructor requires a URL to the 
            // server running SharePoint. 
            ClientContext context = new ClientContext(url);
            context.Credentials = System.Net.CredentialCache.DefaultNetworkCredentials;

            // The SharePoint web at the URL.
            Web web = context.Web;

            ListCreationInformation creationInfo = new ListCreationInformation();
            creationInfo.Title = title;
            creationInfo.TemplateType = (int)ListTemplateType.GenericList;
            List list = web.Lists.Add(creationInfo);
            list.Description = description;

            list.Update();
            context.ExecuteQuery();

            result = true;

            return result;
        }
        private static bool DeleteList(string url, string title)
        {
            bool result = false;

            // Starting with ClientContext, the constructor requires a URL to the 
            // server running SharePoint. 
            ClientContext context = new ClientContext(url);

            // The SharePoint web at the URL.
            Web web = context.Web;

            List list = web.Lists.GetByTitle(title);
            list.DeleteObject();

            context.ExecuteQuery();
            result = true;

            return result;
        }
        private static bool AddJsLinkToList(string url, string title, string JsLinkPath)
        {
            bool result = false;
            ClientContext context = new ClientContext(url);

            //Get the list

            List list = context.Web.Lists.GetByTitle(title);
            context.Load(list);
            context.Load(list.Forms);
            context.ExecuteQuery();

            //Get all the forms
            foreach (var spForm in list.Forms)
            {
                //Get the edit form
                if (spForm.ServerRelativeUrl.Contains("EditForm.aspx") 
                    || spForm.ServerRelativeUrl.Contains("DispForm.aspx")
                    || spForm.ServerRelativeUrl.Contains("NewForm.aspx"))
                {

                    SPX.File file = context.Web.GetFileByServerRelativeUrl(spForm.ServerRelativeUrl);
                    LimitedWebPartManager wpm = file.GetLimitedWebPartManager(PersonalizationScope.Shared);
                    context.Load(wpm.WebParts,
                    wps => wps.Include(
                        wp => wp.WebPart.Title));
                    context.ExecuteQuery();

                    //Set the properties for all web parts
                    foreach (WebPartDefinition wpd in wpm.WebParts)
                    {
                        WebPart wp = wpd.WebPart;
                        wp.Properties["JSLink"] = JsLinkPath;
                        wpd.SaveWebPartChanges();
                        context.ExecuteQuery();

                        result = true;
                    }

                }
            }

            return result;
        }
        #endregion
        #region Field
        private static bool AddFieldToList(string url, string listTitle, string fieldXmlFormat)
        {
            bool result = false;
            
            // Starting with ClientContext, the constructor requires a URL to the 
            // server running SharePoint. 
            ClientContext context = new ClientContext(url);

            SPX.List list = context.Web.Lists.GetByTitle(listTitle);

            SPX.Field field = list.Fields.AddFieldAsXml(fieldXmlFormat,
                                                       true,
                                                       AddFieldOptions.AddFieldInternalNameHint);

            SPX.FieldText fld = context.CastTo<FieldText>(field);

            fld.Update();

            context.ExecuteQuery();
            result = true;

            return result;
        }

        private static bool AddFieldsToList(string url, string listTitle)
        {
            bool result = false;
            string elements = Path.Combine(Environment.CurrentDirectory, "elements.xml");

            XmlReader xmlReader = XmlReader.Create(elements);
            while (xmlReader.ReadToFollowing("List"))
            {
                if ((xmlReader.NodeType == XmlNodeType.Element) && (xmlReader.Name == "Field"))
                {
                    if (xmlReader.HasAttributes)
                        Console.WriteLine(xmlReader.GetAttribute("currency") + ": " + xmlReader.GetAttribute("rate"));
                }
            }
            Console.ReadKey();

            result = true;

            return result;
        }

        private static bool HideField(string url, string listTitle, string internalName)
        {
            bool result = false;

            ClientContext clientContext = new ClientContext(url);
            List listSupportFiles = clientContext.Web.Lists.GetByTitle(listTitle);
            Field field = listSupportFiles.Fields.GetByInternalNameOrTitle(internalName);
            field.SetShowInDisplayForm(false);
            field.SetShowInNewForm(false);
            field.SetShowInNewForm(false);
            field.Hidden = true;
            field.Update();
            clientContext.Load(field);
            clientContext.ExecuteQuery();
            result = true;

            return result;
        }

        private static bool RetrieveFieldsFromList(string url, string listTitle)
        {
            bool result = false;

            // Starting with ClientContext, the constructor requires a URL to the 
            // server running SharePoint. 
            ClientContext context = new ClientContext(url);

            SPX.List list = context.Web.Lists.GetByTitle(listTitle);
            context.Load(list.Fields);

            // We must call ExecuteQuery before enumerate list.Fields. 
            context.ExecuteQuery();

            //
            IQueryable<Field> fields = list.Fields.Where(x => x.StaticName == "hbLinkImage");
            //
            foreach (SPX.Field field in fields)
            {
                if (field.Group.ToLower() != "custom columns")
                {
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine(field.InternalName);
                    Console.ForegroundColor = ConsoleColor.DarkGreen;
                    Console.Write(field.SchemaXml);
                    Console.WriteLine("");
                }
            }
            result = true;

            return result;
        }
        #region Support Solutions

        //public static readonly SPMeta2.Definitions.FieldDefinition SupportTitleField = new SPMeta2.Definitions.FieldDefinition
        //{
        //    Title = "Title",
        //    Id = new Guid("{BC1F6A88-F307-4961-92AB-7534CC1F369B}"),
        //    InternalName = "hbSupportTitle",
        //    AddFieldOptions = AddFieldOptions.AddFieldInternalNameHint,
        //    FieldType = FieldType.Text,
        //    Required = true,
        //    AddToDefaultView = true,
        //};

        //public static readonly SPMeta2.Definitions.FieldDefinition SupportURLField = new SPMeta2.Definitions.FieldDefinition
        //{
        //    Title = "Go to",
        //    Id = new Guid("{EB7C7B6B-68AA-444B-A4C7-6593FC6DB1BF}"),
        //    InternalName = "hbSupportURL",
        //    AddFieldOptions = AddFieldOptions.AddFieldInternalNameHint,
        //    FieldType = FieldType.URL,
        //    Required = false,
        //    AddToDefaultView = true,
        //};

        //public static readonly SPMeta2.Definitions.Fields.NoteFieldDefinition SupportDescriptionField = new SPMeta2.Definitions.Fields.NoteFieldDefinition
        //{
        //    Title = "Description",
        //    Id = new Guid("{E77FCC5A-7F7E-44AF-9D5B-D3D12AD3C954}"),
        //    InternalName = "hbSupportDescription",
        //    AddFieldOptions = AddFieldOptions.AddFieldInternalNameHint,
        //    FieldType = FieldType.HTML,
        //    NumberOfLines = 100,
        //    UnlimitedLengthInDocumentLibrary = true,
        //    RichText = true,
        //    RichTextMode = BuiltInRichTextMode.FullHtml,
        //    Required = true,
        //    AddToDefaultView = true,
        //};

        #endregion
        #endregion
        #region Utility
        private static void DrawProgressBar(int complete, int maxVal, int barSize, char progressCharacter)
        {
            Console.CursorVisible = false;
            int left = Console.CursorLeft;
            decimal perc = (decimal)complete / (decimal)maxVal;
            int chars = (int)Math.Floor(perc / ((decimal)1 / (decimal)barSize));
            string p1 = String.Empty, p2 = String.Empty;

            for (int i = 0; i < chars; i++) p1 += progressCharacter;
            for (int i = 0; i < barSize - chars; i++) p2 += progressCharacter;

            Console.ForegroundColor = ConsoleColor.Green;
            Console.Write(p1);
            Console.ForegroundColor = ConsoleColor.DarkGreen;
            Console.Write(p2);

            Console.ResetColor();
            Console.Write(" {0}%", (perc * 100).ToString("N2"));
            Console.CursorLeft = left;
        }
        #endregion
    }
}
