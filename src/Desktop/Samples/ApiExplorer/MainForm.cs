// ------------------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation
// 
// Permission is hereby granted, free of charge, to any person obtaining a copy
//  of this software and associated documentation files (the "Software"), to deal
//  in the Software without restriction, including without limitation the rights
//  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
//  copies of the Software, and to permit persons to whom the Software is
//  furnished to do so, subject to the following conditions:
// 
// The above copyright notice and this permission notice shall be included in
//  all copies or substantial portions of the Software.
// 
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
//  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
//  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
//  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
//  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
//  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
//  THE SOFTWARE.
// ------------------------------------------------------------------------------

using System;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Live;

using System.Runtime.InteropServices;


namespace Microsoft.Live.Desktop.Samples.ApiExplorer
{
    public partial class MainForm : Form//, IRefreshTokenHandler
    {
        // Update the ClientID with your app client Id that you created from https://account.live.com/developers/applications.
        
        private const string ClientID = "000000004C1338B6";
        private const string ClientSecret = "77X2kYOiGjfcBk2eNUvtUd5Dxnl4YcPw";

//        private const string ClientID = "000000004C133E94";
//        private const string ClientSecret = null;

        private LiveAuthForm authForm;
        private LiveAuthClient liveAuthClient;
        private LiveConnectClient _liveConnectClient;
        private SkyDriveClient _skyDriveClientClient;
//        private RefreshTokenInfo refreshTokenInfo;

        public MainForm()
        {
            if (ClientID.Contains('%'))
            {
                throw new ArgumentException("Update the ClientID with your app client Id that you created from https://account.live.com/developers/applications.");
            }

            InitializeComponent();
        }

        private LiveAuthClient AuthClient
        {
            get
            {
                if (this.liveAuthClient == null)
                {
                    this.AuthClient = new LiveAuthClient(ClientID, ClientSecret);
                }

                return this.liveAuthClient;
            }

            set
            {
                if (this.liveAuthClient != null)
                {
                    this.liveAuthClient.PropertyChanged -= this.liveAuthClient_PropertyChanged;
                }

                this.liveAuthClient = value;
                if (this.liveAuthClient != null)
                {
                    this.liveAuthClient.PropertyChanged += this.liveAuthClient_PropertyChanged;
                }

                this.LiveConnectClient = null;
            }
        }

        private async void AssignLiveConnectClient()
        {
            try
            {
                if (LiveConnectClient != null)
                {
                    LiveOperationResult meRs = await LiveConnectClient.GetAsync("me");
                    dynamic meData = meRs.Result;
                    this.meNameLabel.Text = meData.name;

                    LiveDownloadOperationResult meImgResult = await LiveConnectClient.DownloadAsync("me/picture");
                    this.mePictureBox.Image = Image.FromStream(meImgResult.Stream);
                }
                UpdateUIElements();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, @"Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private LiveConnectClient LiveConnectClient
        {
            get
            {
                return _liveConnectClient;
            }
            set
            {
                if (_liveConnectClient == value) return;
                _liveConnectClient = value;
                _skyDriveClientClient = null;
                AssignLiveConnectClient();
            }
        }
        private SkyDriveClient SkyDriveClient
        {
            get
            {
                return LiveConnectClient == null ? null : (_skyDriveClientClient ?? (_skyDriveClientClient = new SkyDriveClient(LiveConnectClient)));
            }
        }

        private LiveConnectSession AuthSession
        {
            get
            {
                return this.AuthClient.Session;
            }
        }

        private void liveAuthClient_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == "Session")
            {
                this.UpdateUIElements();
            }
        }

        private void UpdateUIElements()
        {
            LiveConnectSession session = this.AuthSession ?? (LiveConnectClient == null ? null : LiveConnectClient.Session);
            bool isSignedIn = session != null;
            
            this.signOutButton.Enabled = isSignedIn;
            this.connectGroupBox.Enabled = isSignedIn;
            this.currentScopeTextBox.Text = isSignedIn ? string.Join(" ", session.Scopes ?? new string[0]) : string.Empty;
            if (!isSignedIn)
            {
                this.meNameLabel.Text = string.Empty;
                this.mePictureBox.Image = null;
            }
        }

        private async void SigninButton_Click(object sender, EventArgs e)
        {
            if (this.authForm == null)
            {
                string startUrl = this.AuthClient.GetLoginUrl(this.GetAuthScopes());
                var endUrl = this.AuthClient.RedirectUrl;
                this.authForm = new LiveAuthForm(
                    startUrl,
                    endUrl,
                    this.OnAuthCompleted);
                this.authForm.FormClosed += AuthForm_FormClosed;
                this.authForm.ShowDialog(this);
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            AuthClient = new LiveAuthClient(textToken.Text, ClientID, ClientSecret);
            LiveConnectClient = new LiveConnectClient(AuthClient);
        }

        private string[] GetAuthScopes()
        {
            string[] scopes = new string[this.scopeListBox.SelectedItems.Count];
            this.scopeListBox.SelectedItems.CopyTo(scopes, 0);
            return scopes;
        }

        void AuthForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            this.CleanupAuthForm();
        }

        private void CleanupAuthForm()
        {
            if (this.authForm != null)
            {
                this.authForm.Dispose();
                this.authForm = null;
            }
        }

        private void LogOutput(string text)
        {
            this.outputTextBox.Text += text + "\r\n";
        }

        private async void OnAuthCompleted(AuthResult result)
        {
            this.CleanupAuthForm();
            if (result.AuthorizeCode != null)
            {
                try
                {
                    LiveConnectSession session = await this.AuthClient.ExchangeAuthCodeAsync(result.AuthorizeCode);
                    this.LiveConnectClient = new LiveConnectClient(session);
/*
                    LiveOperationResult meRs = await this.liveConnectClient.GetAsync("me");
                    dynamic meData = meRs.Result;
                    this.meNameLabel.Text = meData.name;

                    LiveDownloadOperationResult meImgResult = await this.liveConnectClient.DownloadAsync("me/picture");
                    this.mePictureBox.Image = Image.FromStream(meImgResult.Stream);
*/
                }
                catch (LiveAuthException aex)
                {
                    this.LogOutput("Failed to retrieve access token. Error: " + aex.Message);
                }
                catch (LiveConnectException cex)
                {
                    this.LogOutput("Failed to retrieve the user's data. Error: " + cex.Message);
                }
            }
            else
            {
                this.LogOutput(string.Format("Error received. Error: {0} Detail: {1}", result.ErrorCode, result.ErrorDescription));
            }
        }

        private void SignOutButton_Click(object sender, EventArgs e)
        {
            this.signOutWebBrowser.Navigate(this.AuthClient.GetLogoutUrl());
            this.AuthClient = null;
            this.UpdateUIElements();
        }

        private void ClearOutputButton_Click(object sender, EventArgs e)
        {
            this.outputTextBox.Text = string.Empty;
        }

        private void ScopeListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.signinButton.Enabled = this.scopeListBox.SelectedItems.Count > 0;
        }

        private async void ExecuteButton_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(this.pathTextBox.Text) && this.methodComboBox.Text != @"UPLOAD")
            {
                this.LogOutput("Path cannot be empty.");
                return;
            }

            try
            {
                LiveOperationResult result = null;
                switch (this.methodComboBox.Text)
                {
                    case "GET":
                        result = await this.LiveConnectClient.GetAsync(this.pathTextBox.Text);
                        break;
                    case "PUT":
                        result = await this.LiveConnectClient.PutAsync(this.pathTextBox.Text, this.requestBodyTextBox.Text);
                        break;
                    case "POST":
                        result = await this.LiveConnectClient.PostAsync(this.pathTextBox.Text, this.requestBodyTextBox.Text);
                        break;
                    case "DELETE":
                        result = await this.LiveConnectClient.DeleteAsync(this.pathTextBox.Text);
                        break;
                    case "MOVE":
                        result = await this.LiveConnectClient.MoveAsync(this.pathTextBox.Text, this.destPathTextBox.Text);
                        break;
                    case "COPY":
                        result = await this.LiveConnectClient.CopyAsync(this.pathTextBox.Text, this.destPathTextBox.Text);
                        break;
                    case "UPLOAD":
                        if (string.IsNullOrWhiteSpace(this.pathTextBox.Text))
                        {
                            var root = SkyDriveClient.RootFolder;
                            this.pathTextBox.Text = root.Id;
                        }
                        result = await this.UploadFile(this.pathTextBox.Text);
                        break;
                    case "DOWNLOAD":
                        await this.DownloadFile(this.pathTextBox.Text);
                        this.LogOutput("The download operation is completed.");
                        break;
                }

                if (result != null)
                {
                    this.LogOutput(result.RawResult);
                }
            }
            catch (Exception ex)
            {
                this.LogOutput("Received an error. " + ex.Message);
            }
        }


        private async Task<LiveOperationResult> UploadFile(string path)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            Stream stream = null;
            dialog.RestoreDirectory = true;

            if (dialog.ShowDialog() != DialogResult.OK)
            {
                throw new InvalidOperationException("No file is picked to upload.");
            }
            try
            {
                if ((stream = dialog.OpenFile()) == null)
                {
                    throw new Exception("Unable to open the file selected to upload.");
                }
                using (stream)
                {
                    return await this.LiveConnectClient.UploadAsync(path, dialog.SafeFileName, stream, OverwriteOption.DoNotOverwrite);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        private async Task<SkyDriveClient.SkyDriveFileInfo> UploadFile2(string path)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            Stream stream = null;
            dialog.RestoreDirectory = true;

            if (dialog.ShowDialog() != DialogResult.OK)
            {
                throw new InvalidOperationException("No file is picked to upload.");
            }
            try
            {
                if ((stream = dialog.OpenFile()) == null)
                {
                    throw new Exception("Unable to open the file selected to upload.");
                }

                return await SkyDriveClient.UploadFileAsync(dialog.FileName, path);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private async Task DownloadFile(string path)
        {
            SaveFileDialog dialog = new SaveFileDialog();
            Stream stream = null;
            dialog.RestoreDirectory = true;

            if (dialog.ShowDialog() != DialogResult.OK)
            {
                throw new InvalidOperationException("No file is picked to upload.");
            }
            try
            {
                if ((stream = dialog.OpenFile()) == null)
                {
                    throw new Exception("Unable to open the file selected to upload.");
                }

                using (stream)
                {
                    LiveDownloadOperationResult result = await this.LiveConnectClient.DownloadAsync(path);
                    if (result.Stream != null)
                    {
                        using (result.Stream)
                        {
                            await result.Stream.CopyToAsync(stream);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private async void MainForm_Load(object sender, EventArgs e)
        {
            this.methodComboBox.SelectedIndex = 0;

            try
            {
                LiveLoginResult loginResult = await this.AuthClient.InitializeAsync();
                if (loginResult.Session != null)
                {
                    this.LiveConnectClient = new LiveConnectClient(loginResult.Session);
                }
            }
            catch (Exception ex)
            {
                this.LogOutput("Received an error during initializing. " + ex.Message);
            }
        }

        private async void button2_Click(object sender, EventArgs e)
        {
            try
            {
//                var l = (await SkyDriveClient.GetSkyDriveItemsAsync()).ToList();
//                var l = (await SkyDriveClient.GetSkyDriveFilesAsync()).ToList();
//                var el = await SkyDriveClient.GetSharedEditLinkAsync(l[2].Id);
//                var vl = await SkyDriveClient.GetSharedReadLinkAsync(l[2].Id);
//                var embl = await SkyDriveClient.GetEmbedLinkAsync(l[2].Id);
                if (string.IsNullOrWhiteSpace(this.pathTextBox.Text))
                {
                    var root = SkyDriveClient.RootFolder;
                    this.pathTextBox.Text = root.Id;
                }
                var result = await this.UploadFile2(this.pathTextBox.Text);
                if (result != null)
                {
                    this.LogOutput(string.Format("Id: {0}\r\nName: {1}\r\nLink: {2}\r\nSource: {3}", result.Id, result.Name, result.Link, result.Source));

                    await SkyDriveClient.DownloadFileAsync(@"C:\temp\xxxx.xlsx", result.Id);
                }
            }
            catch (AggregateException ex)
            {
                this.LogOutput("Received an error. " + ex.InnerExceptions.Aggregate("", (res, err) => res += "\r\n" + err.InnerException.Message));
            }
            catch (Exception ex)
            {
                this.LogOutput("Received an error. " + ex.Message);
            }
        }

    }
}