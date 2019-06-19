/*************************************************************************
 * Description: Winforms Sample to demonstrate how to use Microsoft Graph 
 * Author: Titus Gicheru
 * Email: d-gicheru@outlook.com
 * 
 * Microsoft.Client.Identity v2.7.1
 * Microsoft.Graph v1.15.0
 * Microsoft.Graph.Auth v0.1.0-preview.2
 *************************************************************************/

using GraphLib;
using MetroFramework.Controls;
using Microsoft.Graph;
using System;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Linq;
using System.Drawing;
using System.IO;
using System.Drawing.Drawing2D;

namespace GraphApp
{
    public partial class frmMain : MetroFramework.Forms.MetroForm
    {
        private static GraphServiceClient graphClient = null;
        private static MetroProgressSpinner _loader;
        private static MetroLabel _userDisplayLabel;
        private static PictureBox _pictureBoxProfilePhoto;

        private static MetroLabel _displayNameValue;
        private static MetroLabel _mobileNumberValue;
        private static MetroLabel _userPrincipalNameValue;
        private static MetroLabel _jobTitleValue;
        private static MetroLabel _officeLocationValue;

        private static FlowLayoutPanel _filesLayout;
        private static DataGridView _dataGridViewMessages;

        public static IUserMessagesCollectionPage _messages;

        public frmMain()
        {
            InitializeComponent();

            _loader = loader;
            _loader.Visible = false;

            _userDisplayLabel = metroLabelDisplayName;
            _pictureBoxProfilePhoto = pictureBoxProfilePhoto;

            _displayNameValue = lblDisplayNameValue;
            _mobileNumberValue = lblMobileNumberValue;
            _userPrincipalNameValue = lblUserPrincipalNameValue;
            _jobTitleValue = lblJobTitleValue;
            _officeLocationValue = lblOfficeLocationValue;

            _filesLayout = flowLayoutPanelFiles;
            _dataGridViewMessages = dataGridViewMessages;

            ClearControls();
        }

        private async void MetroButtonSignIn_Click(object sender, EventArgs e)
        {
            await DisplayUser();
            await GetMessages();
            await ListRecentFiles();
            await BindProfilePhoto();

            _loader.Spinning = false;
            _loader.Visible = false;

            if (graphClient != null)
            {
                metroButtonSignIn.Text = "Refresh";
            }
            else
            {
                metroButtonSignIn.Text = "Sign In";
            }
        }

        private void ClearControls()
        {
            _userDisplayLabel.Text = "";
            _displayNameValue.Text = "";
            _mobileNumberValue.Text = "";
            _userPrincipalNameValue.Text = "";
            _jobTitleValue.Text = "";
            _officeLocationValue.Text = "";
        }
        
        #region Microsoft Graph
        private static async Task DisplayUser()
        {
            _loader.Visible = true;
            _loader.Spinning = true;
            var me = await GetMeAsync();

            if (me != null)
            {
                _userDisplayLabel.Text = $"Hey! {me.GivenName}";
                _displayNameValue.Text = me.DisplayName;

                if (me.BusinessPhones != null)
                {
                    _mobileNumberValue.Text = me.BusinessPhones.FirstOrDefault();
                }

                _userPrincipalNameValue.Text = me.UserPrincipalName;
                _jobTitleValue.Text = me.JobTitle;
                _officeLocationValue.Text = me.OfficeLocation;             
            }
            else
            {
                MessageBox.Show("Did not find user");
            }
        }

        private static async Task BindProfilePhoto()
        {
            var photoStream = await GetMePhotoAsync();
            _pictureBoxProfilePhoto.BackgroundImage = new Bitmap(photoStream);
        }

        private static async Task<User> GetMeAsync()
        {
            User currentUser = null;
            try
            {
                graphClient = GraphServiceClientProvider.GetAuthenticatedClient();
                // Request to get the current logged in user object from Microsoft Graph
                currentUser = await graphClient.Me.Request().GetAsync();
                return currentUser;
            }

            catch (ServiceException e)
            {
                MessageBox.Show("We could not get the current user: " + e.Error.Message, "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }

        private static async Task<Stream> GetMePhotoAsync()
        {
            Stream profilePhotoStream = null;
            try
            {
                graphClient = GraphServiceClientProvider.GetAuthenticatedClient();
                profilePhotoStream = await graphClient.Me.Photo.Content.Request().GetAsync();
                return profilePhotoStream;
            }

            catch (ServiceException e)
            {
                MessageBox.Show("We could not get the current user photo: " + e.Error.Message, "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }

        private static async Task GetMessages()
        {
            var messages = await GetMessagesAsync();

            if (messages != null)
            {
                var filteredMessages = (from r in messages select new { r.From, r.Subject, r.IsRead, r.IsDraft, r.HasAttachments, r.CreatedDateTime });
                _dataGridViewMessages.DataSource = filteredMessages.ToList();
            }
        }

        private static async Task<IUserMessagesCollectionPage> GetMessagesAsync()
        {
            IUserMessagesCollectionPage messages = null;
            try
            {
                graphClient = GraphServiceClientProvider.GetAuthenticatedClient();
                messages = await graphClient.Me
                    .Messages
                    .Request()
                    .Select("From,Subject,IsRead,IsDraft,HasAttachments,CreatedDateTime")
                    .GetAsync();
                return messages;
            }
            catch (ServiceException e)
            {
                MessageBox.Show("We could not get the any messages: " + e.Error.Message, "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }

        private static async Task ListRecentFiles()
        {
            var recentFiles = await GetFilesAsync();

            if (recentFiles != null)
            {
                var currentPage = recentFiles.CurrentPage;

                _filesLayout.Controls.Clear();
                foreach (var driveItem in currentPage)
                {
                    var fileType = GetIcon(driveItem.Name);
                    var driveControl = CreateFileControl(driveItem.Name, fileType, driveItem.Id);
                    _filesLayout.Controls.Add(driveControl);
                }
            }
        }

        private static async Task<IDriveRecentCollectionPage> GetFilesAsync()
        {
            IDriveRecentCollectionPage driveRecentCollectionPage = null;

            try
            {
                graphClient = GraphServiceClientProvider.GetAuthenticatedClient();
                driveRecentCollectionPage = await graphClient.Me.Drive.Recent().Request().GetAsync();
                return driveRecentCollectionPage;
            }
            catch (ServiceException e)
            {
                MessageBox.Show("We could not get the any Files: " + e.Error.Message, "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }
        #endregion

        /***********************************************************
        * Added ability to create nice looking controls at runtime
        * for recent files from Graph
        ************************************************************/
        #region RuntimeControls
        private static Panel CreateFileControl(string fileName, Bitmap fileType, string id)
        {
            // 
            // pictureBoxDriveControl
            // 
            PictureBox pictureBoxDriveControl = new PictureBox();
            pictureBoxDriveControl.BackgroundImage = fileType;
            pictureBoxDriveControl.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            pictureBoxDriveControl.Location = new System.Drawing.Point(50, 19);
            pictureBoxDriveControl.Name = "pictureBoxDriveControl";
            pictureBoxDriveControl.Size = new System.Drawing.Size(156, 168);
            pictureBoxDriveControl.TabIndex = 1;
            pictureBoxDriveControl.TabStop = false;

            // 
            // metroLabelDriveControl
            // 
            Label metroLabelDriveControl = new Label();
            metroLabelDriveControl.AutoSize = false;
            metroLabelDriveControl.ForeColor = System.Drawing.Color.White;
            metroLabelDriveControl.BackColor = System.Drawing.Color.Transparent;
            metroLabelDriveControl.Location = new System.Drawing.Point(8, 195);
            metroLabelDriveControl.Name = "mDC" + id;
            metroLabelDriveControl.Size = new System.Drawing.Size(240, 61);
            metroLabelDriveControl.Text = fileName;
            metroLabelDriveControl.TextAlign = System.Drawing.ContentAlignment.TopCenter;

            // 
            // panelDriveControl
            // 
            Panel panelDriveControl = new Panel();
            panelDriveControl.Controls.Add(pictureBoxDriveControl);
            panelDriveControl.Controls.Add(metroLabelDriveControl);
            panelDriveControl.BackColor = System.Drawing.Color.Transparent;
            panelDriveControl.Name = "pDC" + id;
            panelDriveControl.Size = new System.Drawing.Size(256, 266);
            panelDriveControl.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;



            return panelDriveControl;
        }

        private static string GetFileExtension(string fileName)
        {
            var result = fileName.Split('.');
            return result[1];
        }

        private static Bitmap GetIcon(string fileType)
        {
            var exetension = GetFileExtension(fileType);

            switch (exetension)
            {
                case "xlx":
                    return Properties.Resources.excel;
                case "xlsx":
                    return Properties.Resources.excel;
                case "doc":
                    return Properties.Resources.word;
                case "docx":
                    return Properties.Resources.word;
                case "ppt":
                    return Properties.Resources.powerpoint;
                case "pptx":
                    return Properties.Resources.powerpoint;
                case "one":
                    return Properties.Resources.onenote;
                default:
                    return Properties.Resources.unknown;
            }
        }
        #endregion               

        /*******************************************************************************************
         * Unecessary Code - Just for aesthetic purpose 
         * the picturebox just looks cool when its round
         * https://stackoverflow.com/questions/7731855/rounded-edges-in-picturebox-c-sharp
        ********************************************************************************************/
        private void PictureBoxProfilePhoto_Paint(object sender, PaintEventArgs e)
        {
            GraphicsPath gp = new GraphicsPath();
            gp.AddEllipse(0, 0, pictureBoxProfilePhoto.Width - 3, pictureBoxProfilePhoto.Height - 3);
            Region rg = new Region(gp);
            pictureBoxProfilePhoto.Region = rg;
        }
    }
}