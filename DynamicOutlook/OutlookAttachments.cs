using System;
using System.Runtime.InteropServices;

namespace DynamicOutlook
{
    /// <summary>
    /// Abstracts an attachment collection
    /// </summary>
    class OutlookAttachments : IDisposable
    {
        /// <summary>
        /// Ctor
        /// </summary>
        /// <param name="underlyingObject">a COM object for the attachments collection</param>
        public OutlookAttachments(dynamic underlyingObject)
        {
            _underlyingObject = underlyingObject;
        }

        /// <summary>
        /// The underlying COM object
        /// </summary>
        private dynamic _underlyingObject = null;

        /// <summary>
        /// Adds an attachment to the list
        /// </summary>
        /// <param name="filename">a filename to attach as attachment</param>
        public void Add(string filename)
        {
            _underlyingObject.Add(filename);
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
