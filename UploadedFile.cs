using System;
using System.IO;
using System.Runtime.InteropServices;

namespace UploadManagerLib
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]   // VBScript‑friendly
    [Guid("4E1A7D9B-2BDE-4A94-9B3D-5FBF0A25C7C5")]
    public class UploadedFile
    {
        public string FieldName { get; internal set; }
        public string FileName { get; internal set; }
        public string ContentType { get; internal set; }
        public int Size => _buffer?.Length ?? 0;
        private byte[] _buffer;

        internal UploadedFile() { }

        internal void SetBuffer(byte[] buf) => _buffer = buf;

        // PHP‑like property names for convenience
        public string name => FileName;
        public string type => ContentType;
        public int size => Size;

        /// <summary>Save the uploaded file to disk.</summary>
        public void SaveAs(string path) => File.WriteAllBytes(path, _buffer);

        /// <summary>Return raw bytes (careful with memory!)</summary>
        public byte[] GetBytes() => _buffer;
    }
}
