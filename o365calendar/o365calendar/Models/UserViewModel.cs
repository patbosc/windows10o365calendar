using System;
using System.Threading.Tasks;
using Windows.UI.Xaml.Media.Imaging;
using Microsoft.Azure.ActiveDirectory.GraphClient;
using o365calendar.Common;
using o365calendar.Helpers;

namespace o365calendar.Models
{
    public class UserViewModel : ViewModelBase
    {

        private IUser _currentUser;
        private string _mailAddress;
        private string _id;
        private string _displayName = "(not connected)";
        private string _jobTitle;
        private bool _signedIn;
        private string _logOnCaption = "Connect to Office 365";
        private static readonly BitmapImage _signedOutImage = new BitmapImage(new Uri("ms-appx:///assets/UserDefault.png", UriKind.RelativeOrAbsolute));
        private BitmapImage _avatar = _signedOutImage;
        private RelayCommand _toggleSignInCommand;
        private UserOperations _userOperations = new UserOperations();

        /// <summary>
        /// Gets the Id of the user.
        /// </summary>
        public string Id
        {
            get
            {
                return _id;
            }

            private set
            {
                SetProperty(ref _id, value);
            }
        }

        /// <summary>
        /// True if the user is signed in; Otherwise, false.
        /// </summary>
        public bool SignedIn
        {
            get
            {
                return _signedIn;
            }

            private set
            {
                SetProperty(ref _signedIn, value);
            }
        }

        /// <summary>
        /// The display name of the user.
        /// </summary>
        public string DisplayName
        {
            get
            {
                return _displayName;
            }

            private set
            {
                SetProperty(ref _displayName, value);
            }
        }

        /// <summary>
        /// The job title of the user.
        /// </summary>
        public string JobTitle
        {
            get
            {
                return _jobTitle;
            }

            private set
            {
                SetProperty(ref _jobTitle, value);
            }
        }

        /// <summary>
        /// Caption to show depending on the whether the user is signed in or not. 
        /// </summary>
        public string LogOnCaption
        {
            get
            {
                return _logOnCaption;
            }

            set
            {
                SetProperty(ref _logOnCaption, value);
            }
        }

        /// <summary>
        /// The user's avatar.
        /// </summary>
        public BitmapImage Avatar
        {
            get
            {
                return _avatar;
            }

            set
            {
                SetProperty(ref _avatar, value);
            }
        }

        public string MailAddress
        {
            get
            {
                return _mailAddress;
            }

        }

        private bool _isBusy;

        /// <summary>
        /// True when we are in the process of logging in; Otherwise, false.
        /// </summary>
        public bool IsBusy
        {
            get
            {
                return _isBusy;
            }
            set
            {
                SetProperty(ref _isBusy, value);
            }

        }

        /// <summary>
        /// Command to sign the user in if he is not already signed in or to sign the user out.
        /// </summary>
        public RelayCommand ToggleSignInCommand
        {
            get
            {
                if (_toggleSignInCommand == null)
                {
                    _toggleSignInCommand = new RelayCommand
                    (
                        async () => {
                            if (!SignedIn)
                            {
                                IsBusy = true;
                                await SignInCurrentUserAsync();
                                IsBusy = false;
                            }
                            else
                            {
                                IsBusy = true;
                                await SignOutAsync();
                                IsBusy = false;
                            }
                        },
                        null
                    );
                }

                return _toggleSignInCommand;
            }
        }

        private async Task SignOutAsync()
        {
            if (!SignedIn)
                return;

            await _userOperations.SignOutAsync();

            Avatar = _signedOutImage;

            DisplayName = "(not connected)";
            JobTitle = String.Empty;

            SignedIn = false;
            LogOnCaption = "Connect to Office 365";
        }

        /// <summary>
        /// Signs in the current user.
        /// </summary>
        /// <returns></returns>
        public async Task SignInCurrentUserAsync()
        {
            _currentUser = await _userOperations.GetCurrentUserAsync();

            if (_currentUser != null)
            {
                DisplayName = _currentUser.DisplayName;
                JobTitle = _currentUser.JobTitle;
                Avatar = await _userOperations.GetUserThumbnailPhotoAsync(_currentUser);
                LogOnCaption = "Disconnect from Office 365";
                Id = _currentUser.ObjectId;
                _mailAddress = _currentUser.Mail;
                SignedIn = true;
            }
        }
    }

}
