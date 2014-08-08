using System.Threading.Tasks;

namespace DynamicOutlook
{
    /// <summary>
    /// This class does all the work of sending emails using an installed version of Outlook
    /// </summary>
    public abstract class OutlookHelper
    {
        /// <summary>
        /// Sends an email using Outlook. Starts Outlook if not running already (in which case an Outlook exit is attempted, and Outlook is left showing and running on failure).
        /// </summary>
        /// <param name="to">recipient(s)</param>
        /// <param name="subject">subject of the email</param>
        /// <param name="body">body of the email, in plain text</param>
        /// <param name="attachments">files to attach to the email</param>
        /// <returns>true if the email was sent, false otherwise</returns>
        /// <exception cref="COMException">this method surfaces Outlook's exceptions transparently</exception>
        public static async Task<bool> SendAsync(string to, string subject, string body, params string[] attachments)
        {
            bool sent = false;

            if (IsOutlookInstalled)
            {
                OutlookApplication outlook = null;
                OutlookMailItem email = null;

                try
                {
                    await Task.Run(
                        () =>
                        {
                            outlook = new OutlookApplication();
                            email = outlook.CreateMailItem();
                        });

                    email.Recipients.Add(to);
                    email.Subject = subject;
                    email.Body = body;
                    foreach (string attachment in attachments)
                    {
                        email.Attachments.Add(attachment);
                    }

                    sent = await email.SendAsync(5000);
                }
                finally
                {
                    if (email != null)
                    {
                        email.Dispose();
                        email = null;
                    }
                    if (outlook != null)
                    {
                        outlook.Dispose();
                        outlook = null;
                    }
                }
            }

            return sent;
        }

        /// <summary>
        /// Gets whether Outlook is installed
        /// </summary>
        public static bool IsOutlookInstalled
        {
            get
            {
                return (OutlookApplication.OutlookApplicationType != null);
            }
        }

    }
}
