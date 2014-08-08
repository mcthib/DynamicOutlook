using System;
using System.Threading.Tasks;
using System.Runtime.InteropServices;

namespace DynamicOutlook
{
    /// <summary>
    /// This represents an email(as it's being authored)
    /// </summary>
    class OutlookMailItem : IDisposable
    {
        /// <summary>
        /// Ctor
        /// </summary>
        /// <param name="underlyingObject">a COM object for the mail item</param>
        public OutlookMailItem(dynamic underlyingObject)
        {
            _underlyingObject = underlyingObject;
        }

        /// <summary>
        /// The underlying COM object
        /// </summary>
        private dynamic _underlyingObject = null;

        /// <summary>
        /// Gets the recipients of the email message
        /// </summary>
        public OutlookRecipients Recipients
        {
            get
            {
                if (_recipients == null)
                {
                    _recipients = new OutlookRecipients(_underlyingObject.Recipients);
                }

                return _recipients;
            }
        }
        private OutlookRecipients _recipients = null;

        /// <summary>
        /// Gets the attachments of the email message
        /// </summary>
        public OutlookAttachments Attachments
        {
            get
            {
                if (_attachments == null)
                {
                    _attachments = new OutlookAttachments(_underlyingObject.Attachments);
                }

                return _attachments;
            }
        }
        private OutlookAttachments _attachments = null;

        /// <summary>
        /// Gets or sets the subject of the email
        /// </summary>
        public string Subject
        {
            get
            {
                return _underlyingObject.Subject as string;
            }
            set
            {
                _underlyingObject.Subject = value;
            }
        }

        /// <summary>
        /// Gets or sets the body of the email
        /// </summary>
        public string Body
        {
            get
            {
                return _underlyingObject.Body as string;
            }
            set
            {
                _underlyingObject.Body = value;
            }
        }

        /// <summary>
        /// Saves the mail item - must do prior to sending
        /// </summary>
        public void Save()
        {
            _underlyingObject.Save();
        }

        /// <summary>
        /// Sends the mail item
        /// </summary>
        /// <param name="milliseconds">timeout in ms to wait for email to be reported as sent</param>
        /// <returns>true if email was sent beofre timeout, false otherwise</returns>
        public async Task<bool> SendAsync(int milliseconds)
        {
            bool sent = false;
            DateTime startTime = DateTime.Now;
            TimeSpan timeout = TimeSpan.FromMilliseconds((double)milliseconds);

            // Save first (as draft)
            // Disabled - maybe we don't want the email to show up unless it'sbeen sent. The idea is that we're sending mail programmatically
            // and therefore a degeneration on the part of the caller could cause a lot of mail to pile up in the Drafts folder...
            // Save();

            _underlyingObject.Send();

            // Asynchronously wait for the mail status to go to Sent - check back every 250 ms.
            while (DateTime.Now.Subtract(startTime) < timeout)
            {
                try
                {
                    if (_underlyingObject.Sent)
                    {
                        sent = true;
                        break;
                    }
                }
                catch (COMException ex)
                {
                    if (ex.ErrorCode == unchecked((int)0x8004010A))
                    {
                        // "Item was moved or deleted" - i.e. sent
                        sent = true;
                        break;
                    }
                    else
                    {
                        throw;
                    }
                }
                
                await Task.Delay(250);
            }

            return sent;
        }

        /// <summary>
        /// Releases native / COM resources
        /// </summary>
        public void Dispose()
        {
            if (_recipients != null)
            {
                _recipients.Dispose();
                _recipients = null;
            }

            if (_attachments != null)
            {
                _attachments.Dispose();
                _attachments = null;
            }

            if (_underlyingObject != null)
            {
                Marshal.ReleaseComObject(_underlyingObject);
                _underlyingObject = null;
            }
        }
    }
}
