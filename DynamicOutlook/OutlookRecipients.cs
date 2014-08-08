using System;
using System.Runtime.InteropServices;

namespace DynamicOutlook
{
    /// <summary>
    /// Abstracts a collection of recipients
    /// </summary>
    class OutlookRecipients : IDisposable
    {
        /// <summary>
        /// Ctor
        /// </summary>
        /// <param name="underlyingObject">a COM object for the recipients collection</param>
        public OutlookRecipients(dynamic underlyingObject)
        {
            _underlyingObject = underlyingObject;
        }

        /// <summary>
        /// The underlying COM object
        /// </summary>
        private dynamic _underlyingObject = null;

        /// <summary>
        /// Addsa recipient to the list
        /// </summary>
        /// <param name="recipient">a recipient, in a form outlook can understand</param>
        public void Add(string recipient)
        {
            _underlyingObject.Add(recipient);
        }

        /// <summary>
        /// Releases COM / native resources
        /// </summary>
        public void Dispose()
        {
            if (_underlyingObject != null)
            {
                Marshal.ReleaseComObject(_underlyingObject);
                _underlyingObject = null;
            }
        }
    }
}
