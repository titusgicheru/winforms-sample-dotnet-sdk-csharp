/*************************************************************************
 * Description: Winforms Sample to demonstrate how to use Microsoft Graph 
 * Author: Titus Gicheru
 * Email: d-gicheru@outlook.com
 * 
 * Microsoft.Client.Identity v2.7.1
 * Microsoft.Graph v1.15.0
 * Microsoft.Graph.Auth v0.1.0-preview.2
 *************************************************************************/

using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using GraphApp.Properties;
using GraphLib;
using MetroFramework.Controls;
using MetroFramework.Forms;
using Microsoft.Graph;

namespace GraphApp
{
    public partial class frmMain : MetroForm
    {
        private static GraphServiceClient graphClient;
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

            metroButtonSignIn.Text = graphClient != null ? "Refresh" : "Sign In";
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
            if (photoStream != null) _pictureBoxProfilePhoto.BackgroundImage = new Bitmap(photoStream);
        }

        private static async Task<User> GetMeAsync()
        {
            try
            {
                graphClient = GraphServiceClientProvider.GetAuthenticatedClient();
                // Request to get the current logged in user object from Microsoft Graph
                var currentUser = await graphClient.Me.Request().GetAsync();
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
            try
            {
                graphClient = GraphServiceClientProvider.GetAuthenticatedClient();
                var profilePhotoStream = await graphClient.Me.Photo.Content.Request().GetAsync();
                return profilePhotoStream;
            }

            catch (ServiceException e)
            {
                //MessageBox.Show("We could not get the current user photo: " + e.Error.Message, "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
            try
            {
                graphClient = GraphServiceClientProvider.GetAuthenticatedClient();
                var messages = await graphClient.Me
                    .Messages
                    .Request()
                    .Select("From,Subject,IsRead,IsDraft, HasAttachments,CreatedDateTime")
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
            try
            {
                graphClient = GraphServiceClientProvider.GetAuthenticatedClient();
                var driveRecentCollectionPage = await graphClient.Me.Drive.Recent().Request().GetAsync();
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
            PictureBox pictureBoxDriveControl = new PictureBox
            {
                BackgroundImage = fileType,
                BackgroundImageLayout = ImageLayout.Zoom,
                Location = new Point(50, 19),
                Name = "pictureBoxDriveControl",
                Size = new Size(156, 168),
                TabIndex = 1,
                TabStop = false
            };

            // 
            // metroLabelDriveControl
            // 
            Label metroLabelDriveControl = new Label
            {
                AutoSize = false,
                ForeColor = Color.White,
                BackColor = Color.Transparent,
                Location = new Point(8, 195),
                Name = "mDC" + id,
                Size = new Size(240, 61),
                Text = fileName,
                TextAlign = ContentAlignment.TopCenter
            };

            // 
            // panelDriveControl
            // 
            Panel panelDriveControl = new Panel();
            panelDriveControl.Controls.Add(pictureBoxDriveControl);
            panelDriveControl.Controls.Add(metroLabelDriveControl);
            panelDriveControl.BackColor = Color.Transparent;
            panelDriveControl.Name = "pDC" + id;
            panelDriveControl.Size = new Size(256, 266);
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
                    return Resources.excel;
                case "xlsx":
                    return Resources.excel;
                case "doc":
                    return Resources.word;
                case "docx":
                    return Resources.word;
                case "ppt":
                    return Resources.powerpoint;
                case "pptx":
                    return Resources.powerpoint;
                case "one":
                    return Resources.onenote;
                default:
                    return Resources.unknown;
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