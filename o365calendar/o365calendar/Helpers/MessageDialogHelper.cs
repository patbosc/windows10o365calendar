﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Windows.UI.Popups;
using Microsoft.Office365.OAuth;

namespace o365calendar.Helpers
{
    internal static class MessageDialogHelper
    {

        internal static async Task<bool> ShowYesNoDialogAsync(string content, string title)
        {
            bool result = false;
            MessageDialog messageDialog = new MessageDialog(content, title);

            messageDialog.Commands.Add(new UICommand(
                "Yes",
                new UICommandInvokedHandler((cmd) => result = true)
                ));
            messageDialog.Commands.Add(new UICommand(
               "No",
               new UICommandInvokedHandler((cmd) => result = false)
               ));

            // Set the command that will be invoked by default 
            messageDialog.DefaultCommandIndex = 0;

            // Set the command to be invoked when escape is pressed 
            messageDialog.CancelCommandIndex = 1;

            await messageDialog.ShowAsync();

            return result;
        }

        internal static async void ShowDialogAsync(string content, string title)
        {
            MessageDialog messageDialog = new MessageDialog(content, title);
            messageDialog.Commands.Add(new UICommand(
               "OK",
               null
               ));

            await messageDialog.ShowAsync();
        }

        #region Exception display helpers
        // Display details of the exception in a message dialog.
        // We are doing this here to help you, as a developer, understand exactly
        // what exception was received. In a real app, you would
        // handle exceptions within your code and give a more user-friendly behavior.
        internal static void DisplayException(AuthenticationFailedException exception)
        {
            var title = "Authentication failed";
            StringBuilder content = new StringBuilder();
            content.AppendLine("We were unable to connect to Office 365. Here's the exception we received:");
            content.AppendFormat("Exception: {0}\n", exception.ErrorCode);
            content.AppendFormat("Description: {0}\n\n", exception.ErrorDescription);
            content.AppendLine("Suggestion: Make sure you have added the Connected Services to this project as outlined in the Readme file");
            MessageDialogHelper.ShowDialogAsync(content.ToString(), title);
            Debug.WriteLine(content.ToString());
        }

        // Display details of the exception. 
        // We are doing this here to help you, as a developer, understand exactly
        // what exception was received. In a real app, you would
        // handle exceptions within your code and give a more user-friendly behavior.
        internal static void DisplayException(MissingConfigurationValueException exception)
        {
            var title = "Connected Services configuration failure";
            StringBuilder content = new StringBuilder();
            content.AppendLine("We were unable to connect to Office 365. Here's the exception we received:");
            content.AppendFormat("Exception: {0}\n\n", exception.Message);
            content.AppendLine("Suggestion: Make sure you have added the Connected Services to this project as outlined in the Readme file.");
            MessageDialogHelper.ShowDialogAsync(content.ToString(), title);
        }

        // Display details of the exception. 
        // We are doing this here to help you, as a developer, understand exactly
        // what exception was received. In a real app, you would
        // handle exceptions within your code and give a more user-friendly behavior.
        internal static void DisplayException(Exception exception)
        {
            var title = "Connected Services configuration failure";
            StringBuilder content = new StringBuilder();
            content.AppendLine("We were unable to connect to Office 365. Here's the exception we received:");
            content.AppendFormat("Exception: {0}\n\n", exception.Message);
            content.AppendLine("Suggestion: Make sure you have added the Connected Services to this project as outlined in the Readme file.");
            MessageDialogHelper.ShowDialogAsync(content.ToString(), title);
        }
        #endregion
    }

}
