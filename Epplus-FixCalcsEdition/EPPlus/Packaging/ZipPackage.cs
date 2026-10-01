/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
 * See https://github.com/JanKallman/EPPlus for details.
 *
 * Copyright (C) 2011  Jan Källman
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.

 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  
 * See the GNU Lesser General Public License for more details.
 *
 * The GNU Lesser General Public License can be viewed at http://www.opensource.org/licenses/lgpl-license.php
 * If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
 *
 * All code and executables are provided "as is" with no warranty either express or implied. 
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 * Code change notes:
 * 
 * Author							Change						Date
 *******************************************************************************
 * Jan Källman		Added		25-Oct-2012
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.IO;
using System.IO.Compression;
using System.Xml;
using CompuMaster.Epplus4.Utils;
namespace CompuMaster.Epplus4.Packaging
{
    internal sealed class PackageMemoryStream : MemoryStream
    {
        internal long XmlSizeLimit { get; }

        internal PackageMemoryStream(long xmlSizeLimit)
        {
            XmlSizeLimit = xmlSizeLimit;
        }
    }

    /// <summary>
    /// Specifies whether the target is inside or outside the System.IO.Packaging.Package.
    /// </summary>
    public enum TargetMode
    {
        /// <summary>
        /// The relationship references a part that is inside the package.
        /// </summary>
        Internal = 0,
        /// <summary>
        /// The relationship references a resource that is external to the package.
        /// </summary>
        External = 1,
    }
    /// <summary>
    /// Represent an OOXML Zip package.
    /// </summary>
    public class ZipPackage : ZipPackageRelationshipBase
    {
        internal class ContentType
        {
            internal string Name;
            internal bool IsExtension;
            internal string Match;
            public ContentType(string name, bool isExtension, string match)
            {
                Name = name;
                IsExtension = isExtension;
                Match = match;
            }
        }
        Dictionary<string, ZipPackagePart> Parts = new Dictionary<string, ZipPackagePart>(StringComparer.OrdinalIgnoreCase);
        internal Dictionary<string, ContentType> _contentTypes = new Dictionary<string, ContentType>(StringComparer.OrdinalIgnoreCase);
        internal char _dirSeparator='/';
        internal ZipPackage()
        {
            AddNew();
        }

        private void AddNew()
        {
            _contentTypes.Add("xml", new ContentType(ExcelPackage.schemaXmlExtension, true, "xml"));
            _contentTypes.Add("rels", new ContentType(ExcelPackage.schemaRelsExtension, true, "rels"));
        }

        internal ZipPackage(Stream stream)
            : this(stream, ExcelPackageLoadLimits.Default)
        {
        }

        internal ZipPackage(Stream stream, ExcelPackageLoadLimits loadLimits)
        {
            if (loadLimits == null) throw new ArgumentNullException(nameof(loadLimits));
            bool hasContentTypeXml = false;
            if (stream == null || stream.Length == 0)
            {
                AddNew();
            }
            else
            {
                if (stream.Length > loadLimits.MaxInputBytes)
                    throw new InvalidDataException("The XLSX input exceeds the configured compressed input size limit.");
                ValidateCentralDirectoryEntryCount(stream, loadLimits.MaxZipEntries);
                var rels = new Dictionary<string, string>();
                stream.Seek(0, SeekOrigin.Begin);                
                ZipArchive zip = new ZipArchive(stream, ZipArchiveMode.Read, true);
                if (zip.Entries.Count == 0)
                {
                    zip.Dispose();
                    var repairedStream = RepairEmptyCentralDirectory(stream, loadLimits.MaxZipEntries);
                    if (repairedStream == null)
                    {
                        throw new InvalidDataException("The file is not an valid Package file. If the file is encrypted, please supply the password in the constructor.");
                    }
                    zip = new ZipArchive(repairedStream, ZipArchiveMode.Read, false);
                }
                using (zip)
                {
                    if (zip.Entries.Count == 0)
                    {
                        throw (new InvalidDataException("The file is not an valid Package file. If the file is encrypted, please supply the password in the constructor."));
                    }
                    if (zip.Entries.Count > loadLimits.MaxZipEntries)
                        throw new InvalidDataException("The XLSX package exceeds the configured ZIP entry count limit.");
                    long declaredTotal = 0;
                    foreach (ZipArchiveEntry entry in zip.Entries)
                    {
                        ValidateEntryMetadata(entry, loadLimits, ref declaredTotal);
                    }
                    if (zip.Entries[0].FullName.Contains("\\"))
                    {
                        _dirSeparator = '\\';
                    }
                    else
                    {
                        _dirSeparator = '/';
                    }
                    long extractedTotal = 0;
                    foreach (ZipArchiveEntry e in zip.Entries)
                    {
                        MemoryStream buffer = ReadEntry(e, loadLimits, ref extractedTotal);
                        if (e.Length == 0)
                        {
                            buffer.Dispose();
                            continue;
                        }
                        try
                        {
                            if (e.FullName.Equals("[content_types].xml", StringComparison.OrdinalIgnoreCase))
                            {
                                AddContentTypes(Encoding.UTF8.GetString(buffer.GetBuffer(), 0, (int)buffer.Length));
                                hasContentTypeXml = true;
                            }
                            else if (e.FullName.Equals($"_rels{_dirSeparator}.rels", StringComparison.OrdinalIgnoreCase))
                            {
                                ReadRelation(Encoding.UTF8.GetString(buffer.GetBuffer(), 0, (int)buffer.Length), "");
                            }
                            else
                            {
                                if (e.FullName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))
                                {
                                    rels.Add(GetUriKey(e.FullName), Encoding.UTF8.GetString(buffer.GetBuffer(), 0, (int)buffer.Length));
                                }
                                else
                                {
                                    var part = new ZipPackagePart(this, e);
                                    part.Stream = buffer;
                                    Parts.Add(GetUriKey(e.FullName), part);
                                    buffer = null;
                                }
                            }
                        }
                        finally { buffer?.Dispose(); }
                    }

                    foreach (var p in Parts)
                    {
                        string name = Path.GetFileName(p.Key);
                        string extension = Path.GetExtension(p.Key);
                        string relFile = string.Format("{0}_rels/{1}.rels", p.Key.Substring(0, p.Key.Length - name.Length), name);
                        if (rels.ContainsKey(relFile))
                        {
                            p.Value.ReadRelation(rels[relFile], p.Value.Uri.OriginalString);
                        }
                        if (_contentTypes.ContainsKey(p.Key))
                        {
                            p.Value.ContentType = _contentTypes[p.Key].Name;
                        }
                        else if (extension.Length > 1 && _contentTypes.ContainsKey(extension.Substring(1)))
                        {
                            p.Value.ContentType = _contentTypes[extension.Substring(1)].Name;
                        }
                    }
                    if (!hasContentTypeXml)
                    {
                        throw (new InvalidDataException("The file is not an valid Package file. If the file is encrypted, please supply the password in the constructor."));
                    }
                    if (!hasContentTypeXml)
                    {
                        throw (new InvalidDataException("The file is not an valid Package file. If the file is encrypted, please supply the password in the constructor."));
                    }
                }
            }
        }

        private static bool IsXmlEntry(string name)
        {
            return name.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) ||
                name.EndsWith(".rels", StringComparison.OrdinalIgnoreCase) ||
                name.EndsWith(".vml", StringComparison.OrdinalIgnoreCase);
        }

        private static void ValidateCentralDirectoryEntryCount(Stream stream, int maxEntries)
        {
            long end = stream.Length;
            long first = Math.Max(0, end - 65557);
            var tail = new byte[checked((int)(end - first))];
            stream.Seek(first, SeekOrigin.Begin);
            int read = 0;
            while (read < tail.Length)
            {
                int current = stream.Read(tail, read, tail.Length - read);
                if (current == 0) break;
                read += current;
            }
            for (int offset = read - 22; offset >= 0; offset--)
            {
                if (!HasZipSignature(tail, offset, 0x06054b50) ||
                    offset + 22 + BitConverter.ToUInt16(tail, offset + 20) != read)
                    continue;
                if (BitConverter.ToUInt16(tail, offset + 10) > maxEntries)
                    throw new InvalidDataException("The XLSX package exceeds the configured ZIP entry count limit.");
                break;
            }
        }

        private static void ValidateEntryMetadata(ZipArchiveEntry entry, ExcelPackageLoadLimits limits, ref long declaredTotal)
        {
            long maxEntryBytes = IsXmlEntry(entry.FullName)
                ? Math.Min(limits.MaxEntryBytes, limits.MaxXmlBytes)
                : limits.MaxEntryBytes;
            if (entry.Length > maxEntryBytes)
                throw new InvalidDataException("An XLSX ZIP entry exceeds the configured uncompressed entry size limit.");
            if (entry.Length > limits.MaxTotalUncompressedBytes - declaredTotal)
                throw new InvalidDataException("The XLSX package exceeds the configured total uncompressed size limit.");
            if (entry.Length > 0 && (entry.CompressedLength == 0 ||
                (double)entry.Length / entry.CompressedLength > limits.MaxCompressionRatio))
                throw new InvalidDataException("An XLSX ZIP entry exceeds the configured compression ratio limit.");
            declaredTotal += entry.Length;
        }

        private static MemoryStream ReadEntry(ZipArchiveEntry entry, ExcelPackageLoadLimits limits, ref long extractedTotal)
        {
            long maxEntryBytes = IsXmlEntry(entry.FullName)
                ? Math.Min(limits.MaxEntryBytes, limits.MaxXmlBytes)
                : limits.MaxEntryBytes;
            var result = new PackageMemoryStream(limits.MaxXmlBytes);
            try
            {
                long extracted = 0;
                var block = new byte[81920];
                using (var source = entry.Open())
                {
                    int bytesRead;
                    while ((bytesRead = source.Read(block, 0, block.Length)) != 0)
                    {
                        if (bytesRead > maxEntryBytes - extracted ||
                            bytesRead > limits.MaxTotalUncompressedBytes - extractedTotal - extracted ||
                            entry.CompressedLength == 0 ||
                            (double)(extracted + bytesRead) / entry.CompressedLength > limits.MaxCompressionRatio)
                            throw new InvalidDataException("An XLSX ZIP entry exceeds the configured resource limits.");
                        result.Write(block, 0, bytesRead);
                        extracted += bytesRead;
                    }
                }
                if (extracted != entry.Length)
                    throw new InvalidDataException("An XLSX ZIP entry has inconsistent uncompressed size metadata.");
                extractedTotal += extracted;
                result.Position = 0;
                return result;
            }
            catch
            {
                result.Dispose();
                throw;
            }
        }

        // Some older encrypted workbooks have valid central directory records but an
        // end-of-central-directory record whose entry count and offsets are all zero.
        // The former streaming ZIP reader ignored that record. Repair a private copy
        // only when the complete central directory chain can be validated.
        private static Stream RepairEmptyCentralDirectory(Stream stream, int maxEntries)
        {
            stream.Seek(0, SeekOrigin.Begin);
            byte[] data;
            using (var copy = new MemoryStream())
            {
                stream.CopyTo(copy);
                data = copy.ToArray();
            }
            int end = data.Length - 22;
            if (end < 46 || !HasZipSignature(data, 0, 0x04034b50) ||
                !HasZipSignature(data, end, 0x06054b50))
            {
                return null;
            }
            for (int offset = end + 4; offset < data.Length; offset++)
            {
                if (data[offset] != 0) return null;
            }

            int first = end;
            int count = 0;
            while (first >= 46 && count < ushort.MaxValue)
            {
                bool found = false;
                for (int offset = first - 46; offset >= 0; offset--)
                {
                    if (!HasZipSignature(data, offset, 0x02014b50)) continue;
                    int recordLength = 46 + BitConverter.ToUInt16(data, offset + 28) +
                        BitConverter.ToUInt16(data, offset + 30) + BitConverter.ToUInt16(data, offset + 32);
                    uint localOffset = BitConverter.ToUInt32(data, offset + 42);
                    if ((long)offset + recordLength != first || localOffset >= offset ||
                        !HasZipSignature(data, (int)localOffset, 0x04034b50)) continue;
                    first = offset;
                    count++;
                    if (count > maxEntries)
                        throw new InvalidDataException("The XLSX package exceeds the configured ZIP entry count limit.");
                    found = true;
                    break;
                }
                if (!found) break;
            }
            if (count == 0 || first == 0) return null;

            Buffer.BlockCopy(BitConverter.GetBytes((ushort)count), 0, data, end + 8, 2);
            Buffer.BlockCopy(BitConverter.GetBytes((ushort)count), 0, data, end + 10, 2);
            Buffer.BlockCopy(BitConverter.GetBytes((uint)(end - first)), 0, data, end + 12, 4);
            Buffer.BlockCopy(BitConverter.GetBytes((uint)first), 0, data, end + 16, 4);
            return new MemoryStream(data, false);
        }

        private static bool HasZipSignature(byte[] data, int offset, uint signature)
        {
            return offset >= 0 && offset <= data.Length - 4 &&
                BitConverter.ToUInt32(data, offset) == signature;
        }

        private void AddContentTypes(string xml)
        {
            var doc = new XmlDocument();
            XmlHelper.LoadXmlSafe(doc, xml, Encoding.UTF8);

            foreach (XmlElement c in doc.DocumentElement.ChildNodes)
            {
                ContentType ct;
                if (string.IsNullOrEmpty(c.GetAttribute("Extension")))
                {
                    ct = new ContentType(c.GetAttribute("ContentType"), false, c.GetAttribute("PartName"));
                }
                else
                {
                    ct = new ContentType(c.GetAttribute("ContentType"), true, c.GetAttribute("Extension"));
                }
                _contentTypes.Add(GetUriKey(ct.Match), ct);
            }
        }

#region Methods
        internal ZipPackagePart CreatePart(Uri partUri, string contentType)
        {
            return CreatePart(partUri, contentType, CompressionLevel.Default);
        }
        internal ZipPackagePart CreatePart(Uri partUri, string contentType, CompressionLevel compressionLevel)
        {
            if (PartExists(partUri))
            {
                throw (new InvalidOperationException("Part already exist"));
            }

            var part = new ZipPackagePart(this, partUri, contentType, compressionLevel);
            _contentTypes.Add(GetUriKey(part.Uri.OriginalString), new ContentType(contentType, false, part.Uri.OriginalString));
            Parts.Add(GetUriKey(part.Uri.OriginalString), part);
            return part;
        }
        internal ZipPackagePart GetPart(Uri partUri)
        {
            if (PartExists(partUri))
            {
                return Parts.Single(x => x.Key.Equals(GetUriKey(partUri.OriginalString),StringComparison.OrdinalIgnoreCase)).Value;
            }
            else
            {
                throw (new InvalidOperationException("Part does not exist."));
            }
        }

        internal string GetUriKey(string uri)
        {
            string ret = uri.Replace('\\', '/');
            if (ret[0] != '/')
            {
                ret = '/' + ret;
            }
            return ret;
        }
        internal bool PartExists(Uri partUri)
        {
            var uriKey = GetUriKey(partUri.OriginalString.ToLowerInvariant());
            return Parts.ContainsKey(uriKey);
            //return Parts.Keys.Any(x => x.Equals(uriKey, StringComparison.OrdinalIgnoreCase));
        }
#endregion

        internal void DeletePart(Uri Uri)
        {
            var delList=new List<object[]>(); 
            foreach (var p in Parts.Values)
            {
                foreach (var r in p.GetRelationships())
                {
                    if (UriHelper.ResolvePartUri(p.Uri, r.TargetUri).OriginalString.Equals(Uri.OriginalString, StringComparison.OrdinalIgnoreCase))
                    {                        
                        delList.Add(new object[]{r.Id, p});
                    }
                }
            }
            foreach (var o in delList)
            {
                ((ZipPackagePart)o[1]).DeleteRelationship(o[0].ToString());
            }
            var rels = GetPart(Uri).GetRelationships();
            while (rels.Count > 0)
            {
                rels.Remove(rels.First().Id);
            }
            rels=null;
            _contentTypes.Remove(GetUriKey(Uri.OriginalString));
            //remove all relations
            Parts.Remove(GetUriKey(Uri.OriginalString));
            
        }
        internal void Save(Stream stream)
        {
            var enc = Encoding.UTF8;
            using (ZipArchive archive = new ZipArchive(stream, ZipArchiveMode.Create, true))
            {
            /**** ContentType****/
            var entry = archive.CreateEntry("[Content_Types].xml", GetZipCompressionLevel(_compression));
            byte[] b = enc.GetBytes(GetContentTypeXml());
            using (var entryStream = entry.Open())
            {
                entryStream.Write(b, 0, b.Length);
            }
            /**** Top Rels ****/
            _rels.WriteZip(archive, $"_rels/.rels", _compression);
            ZipPackagePart ssPart=null;
            foreach(var part in Parts.Values)
            {
                if (part.ContentType != ExcelPackage.contentTypeSharedString)
                {
                    part.WriteZip(archive);
                }
                else
                {
                    ssPart = part;
                }
            }
            //Shared strings must be saved after all worksheets. The ss dictionary is populated when that workheets are saved (to get the best performance).
            if (ssPart != null)
            {
                ssPart.WriteZip(archive);
            }
            }
            
            //return ms;
        }

        internal static System.IO.Compression.CompressionLevel GetZipCompressionLevel(CompressionLevel level)
        {
            if (level == CompressionLevel.None)
            {
                return System.IO.Compression.CompressionLevel.NoCompression;
            }
            if ((int)level <= 3)
            {
                return System.IO.Compression.CompressionLevel.Fastest;
            }
            return System.IO.Compression.CompressionLevel.Optimal;
        }

        private string GetContentTypeXml()
        {
            StringBuilder xml = new StringBuilder("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">");
            foreach (ContentType ct in _contentTypes.Values)
            {
                if (ct.IsExtension)
                {
                    xml.AppendFormat("<Default ContentType=\"{0}\" Extension=\"{1}\"/>", ct.Name, ct.Match);
                }
                else
                {
                    xml.AppendFormat("<Override ContentType=\"{0}\" PartName=\"{1}\" />", ct.Name, GetUriKey(ct.Match));
                }
            }
            xml.Append("</Types>");
            return xml.ToString();
        }
        internal void Flush()
        {

        }
        internal void Close()
        {
            
        }
        CompressionLevel _compression = CompressionLevel.Default;
        public CompressionLevel Compression 
        { 
            get
            {
                return _compression;
            }
            set
            {
                foreach (var part in Parts.Values)
                {
                    if (part.CompressionLevel == _compression)
                    {
                        part.CompressionLevel = value;
                    }
                }
                _compression = value;
            }
        }
    }
}
