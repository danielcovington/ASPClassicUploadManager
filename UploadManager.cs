using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

namespace UploadManagerLib
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [Guid("9224A1E3-085C-4B5E-93E5-0E1D3938B18C")]
    public class UploadManager
    {
        private readonly List<UploadedFile> _files = new List<UploadedFile>();

        /// <summary>
        /// Parses a multipart/form‑data request body obtained via Request.BinaryRead.
        /// Throws COMException with a helpful description when a problem occurs so
        /// Classic ASP can trap Err.Number/Err.Description.
        /// </summary>
        public void Parse(object body, string contentType)
        {
            try
            {
                _files.Clear();

                if (body == null)
                    throw new COMException("Request body is null.", unchecked((int)0x800A0005));

                if (string.IsNullOrWhiteSpace(contentType) ||
                    !contentType.StartsWith("multipart/form-data", StringComparison.OrdinalIgnoreCase))
                    throw new COMException("Request is not multipart/form-data.", unchecked((int)0x800A0005));

                byte[] bytes = (byte[])body;
                string boundary = GetBoundary(contentType);
                byte[] boundaryBytes = Encoding.ASCII.GetBytes(boundary);

                int pos = 0;
                while ((pos = IndexOf(bytes, boundaryBytes, pos)) != -1)
                {
                    int hdrStart = pos + boundaryBytes.Length + 2; // skip boundary + CRLF (\r\n)
                    if (hdrStart >= bytes.Length - 1 ||
                        (bytes[hdrStart] == '-' && bytes[hdrStart + 1] == '-'))
                        break; // reached final boundary

                    int hdrEnd = IndexOf(bytes, Encoding.ASCII.GetBytes("\r\n\r\n"), hdrStart);
                    if (hdrEnd == -1)
                        throw new COMException("Malformed multipart section: header terminator not found.", unchecked((int)0x800A0005));

                    string headerText = Encoding.ASCII.GetString(bytes, hdrStart, hdrEnd - hdrStart);

                    string fieldName = null;
                    string fileName = null;
                    string mime = "application/octet-stream";

                    foreach (string line in headerText.Split(new[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries))
                    {
                        if (line.StartsWith("Content-Disposition", StringComparison.OrdinalIgnoreCase))
                        {
                            foreach (string token in line.Split(';'))
                            {
                                string[] kv = token.Split('=');
                                if (kv.Length != 2) continue;
                                string key = kv[0].Trim();
                                string val = kv[1].Trim().Trim('"');

                                if (key.Equals("name", StringComparison.OrdinalIgnoreCase))
                                    fieldName = val;
                                else if (key.Equals("filename", StringComparison.OrdinalIgnoreCase))
                                    fileName = Path.GetFileName(val);
                            }
                        }
                        else if (line.StartsWith("Content-Type", StringComparison.OrdinalIgnoreCase))
                        {
                            int idx = line.IndexOf(':');
                            if (idx > -1) mime = line.Substring(idx + 1).Trim();
                        }
                    }

                    int dataStart = hdrEnd + 4; // position after \r\n\r\n
                    int boundaryMarkerPos = IndexOf(bytes, Encoding.ASCII.GetBytes("\r\n" + boundary), dataStart);
                    if (boundaryMarkerPos == -1)
                        throw new COMException("Terminating boundary not found (uploaded file may be too large or contain the boundary token).", unchecked((int)0x800A0005));

                    int dataEnd = boundaryMarkerPos - 2; // strip preceding \r\n
                    if (dataEnd < dataStart)
                        throw new COMException("Multipart parsing error: invalid data block indices.", unchecked((int)0x800A0005));

                    byte[] buffer = new byte[dataEnd - dataStart + 1];
                    Array.Copy(bytes, dataStart, buffer, 0, buffer.Length);

                    var f = new UploadedFile
                    {
                        FieldName = fieldName,
                        FileName = fileName,
                        ContentType = mime
                    };
                    f.SetBuffer(buffer);
                    _files.Add(f);

                    pos = dataEnd + 2; // continue scanning after CRLF
                }

                if (_files.Count == 0)
                    throw new COMException("No files were detected in the uploaded data.", unchecked((int)0x800A0005));
            }
            catch (COMException)
            {
                throw; // already formatted
            }
            catch (Exception ex)
            {
                throw new COMException(ex.Message, Marshal.GetHRForException(ex));
            }
        }

        // ---------------------------------------------------------------------
        // Helpers & utilities
        // ---------------------------------------------------------------------

        private static string GetBoundary(string contentType)
        {
            const string marker = "boundary=";
            int idx = contentType.IndexOf(marker, StringComparison.OrdinalIgnoreCase);
            if (idx < 0)
                throw new COMException("Boundary parameter missing in content type.", unchecked((int)0x800A0005));

            string boundary = contentType.Substring(idx + marker.Length).Trim();
            if (boundary.Length > 1 && boundary.StartsWith("\"") && boundary.EndsWith("\""))
                boundary = boundary.Trim('"');

            return "--" + boundary;
        }

        private static int IndexOf(byte[] haystack, byte[] needle, int start)
        {
            for (int i = start; i <= haystack.Length - needle.Length; i++)
            {
                bool match = true;
                for (int j = 0; j < needle.Length; j++)
                {
                    if (haystack[i + j] != needle[j]) { match = false; break; }
                }
                if (match) return i;
            }
            return -1;
        }

        // ---------------------------------------------------------------------
        // VBScript‑facing members
        // ---------------------------------------------------------------------

        public int FileCount
        {
            get
            {
                try { return _files.Count; }
                catch (Exception ex) { throw new COMException(ex.Message, Marshal.GetHRForException(ex)); }
            }
        }

        public UploadedFile Item(int index)
        {
            try { return _files[index]; }
            catch (Exception ex) { throw new COMException(ex.Message, Marshal.GetHRForException(ex)); }
        }

        public UploadedFile GetFileByField(string fieldName)
        {
            try { return _files.Find(f => string.Equals(f.FieldName, fieldName, StringComparison.OrdinalIgnoreCase)); }
            catch (Exception ex) { throw new COMException(ex.Message, Marshal.GetHRForException(ex)); }
        }

        [DispId(-4)]
        public IEnumerator GetEnumerator()
        {
            return _files.GetEnumerator();
        }
    }
}
