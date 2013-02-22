using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;


using System.Diagnostics;
using Spring.IO;
using Spring.Social.OAuth1;

using Spring.Social.Dropbox.Api;
using Spring.Social.Dropbox.Connect;


using System.IO;
using System.Threading;

using Outlook = Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools.Ribbon;

namespace outlook_drop
{
    public partial class BrowserControl : UserControl
    {
        private DropboxServiceProvider dropboxServiceProvider;
        private OAuthToken oauthToken;
        public event EventHandler Login_EventHandler;
        private IDropbox dropbox;

        private string consumerKey = "",
                       consumerSecret = "";

        public BrowserControl()
        {
            InitializeComponent();
        }

        public void Init()
        {
            dropboxServiceProvider = new DropboxServiceProvider(consumerKey, consumerSecret, AccessLevel.AppFolder);
            oauthToken = dropboxServiceProvider.OAuthOperations.FetchRequestTokenAsync(null , null).Result;
            OAuth1Parameters parameters = new OAuth1Parameters();
            string authenticateUrl = dropboxServiceProvider.OAuthOperations.BuildAuthorizeUrl(oauthToken.Value, parameters);

            webBrowser.DocumentCompleted += new WebBrowserDocumentCompletedEventHandler(webBrowser_DocumentCompleted);
            webBrowser.Navigate(authenticateUrl);
        }

        private void webBrowser_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            Debug.WriteLine(e.Url);
            if (e.Url.AbsolutePath.Equals("/1/oauth/authorize"))
            {
                AuthorizedRequestToken requestToken = new AuthorizedRequestToken(oauthToken, null);
                OAuthToken oauthAccessToken = dropboxServiceProvider.OAuthOperations.ExchangeForAccessTokenAsync(requestToken, null).Result;
                dropbox = dropboxServiceProvider.GetApi(oauthAccessToken.Value, oauthAccessToken.Secret);

                EventHandler handler = Login_EventHandler;
                if (handler != null)
                {
                    handler(this, EventArgs.Empty);
                }
                webBrowser.DocumentCompleted -= new WebBrowserDocumentCompletedEventHandler(webBrowser_DocumentCompleted);

            }
        }

        public string GetShareLink(string path = "")
        {
            DropboxLink shareableLink = dropbox.GetShareableLinkAsync(path).Result;
            return shareableLink.Url;
        }

        public void Upload(string name, string path)
        {
            dropbox.UploadFileAsync(new FileResource(path), name)
                .ContinueWith(t =>
                    Debug.WriteLine(name + " uploaded")
                );
        }
     }
}
