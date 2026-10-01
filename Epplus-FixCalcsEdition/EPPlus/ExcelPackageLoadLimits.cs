using System;

namespace CompuMaster.Epplus4
{
    /// <summary>
    /// Defines resource limits for loading XLSX packages.
    /// </summary>
#pragma warning disable CS3014 // The legacy EPPlus assembly has no assembly-level CLS declaration.
    [CLSCompliant(true)]
    public sealed class ExcelPackageLoadLimits
    {
        /// <summary>
        /// Gets the default resource limits for loading an XLSX package.
        /// </summary>
        public static ExcelPackageLoadLimits Default { get; } = new ExcelPackageLoadLimits();

        /// <summary>
        /// Initializes resource limits for loading XLSX packages.
        /// </summary>
        /// <param name="maxInputBytes">The maximum size of the compressed or encrypted input, in bytes.</param>
        /// <param name="maxZipEntries">The maximum number of ZIP entries, including empty entries.</param>
        /// <param name="maxEntryBytes">The maximum uncompressed size of one ZIP entry, in bytes.</param>
        /// <param name="maxTotalUncompressedBytes">The maximum sum of uncompressed ZIP entry sizes, in bytes.</param>
        /// <param name="maxXmlBytes">The maximum uncompressed size of one XML, relationships, or VML entry, in bytes.</param>
        /// <param name="maxCompressionRatio">The maximum ratio of uncompressed to compressed bytes for one ZIP entry.</param>
        /// <param name="maxPasswordHashIterations">The maximum password-hash iterations allowed for an encrypted package.</param>
        /// <param name="maxEncryptionMetadataBytes">The maximum size of encryption metadata, in bytes.</param>
        /// <exception cref="ArgumentOutOfRangeException">A resource limit is not positive.</exception>
        public ExcelPackageLoadLimits(
            long maxInputBytes = 128L * 1024 * 1024,
            int maxZipEntries = 20000,
            long maxEntryBytes = 512L * 1024 * 1024,
            long maxTotalUncompressedBytes = 1024L * 1024 * 1024,
            long maxXmlBytes = 512L * 1024 * 1024,
            double maxCompressionRatio = 200,
            int maxPasswordHashIterations = 1000000,
            long maxEncryptionMetadataBytes = 1024L * 1024)
        {
            if (maxInputBytes <= 0) throw new ArgumentOutOfRangeException(nameof(maxInputBytes));
            if (maxZipEntries <= 0) throw new ArgumentOutOfRangeException(nameof(maxZipEntries));
            if (maxEntryBytes <= 0) throw new ArgumentOutOfRangeException(nameof(maxEntryBytes));
            if (maxTotalUncompressedBytes <= 0) throw new ArgumentOutOfRangeException(nameof(maxTotalUncompressedBytes));
            if (maxXmlBytes <= 0) throw new ArgumentOutOfRangeException(nameof(maxXmlBytes));
            if (double.IsNaN(maxCompressionRatio) || double.IsInfinity(maxCompressionRatio) || maxCompressionRatio <= 0)
                throw new ArgumentOutOfRangeException(nameof(maxCompressionRatio));
            if (maxPasswordHashIterations <= 0) throw new ArgumentOutOfRangeException(nameof(maxPasswordHashIterations));
            if (maxEncryptionMetadataBytes <= 0) throw new ArgumentOutOfRangeException(nameof(maxEncryptionMetadataBytes));

            MaxInputBytes = maxInputBytes;
            MaxZipEntries = maxZipEntries;
            MaxEntryBytes = maxEntryBytes;
            MaxTotalUncompressedBytes = maxTotalUncompressedBytes;
            MaxXmlBytes = maxXmlBytes;
            MaxCompressionRatio = maxCompressionRatio;
            MaxPasswordHashIterations = maxPasswordHashIterations;
            MaxEncryptionMetadataBytes = maxEncryptionMetadataBytes;
        }

        /// <summary>
        /// Gets the maximum size of the compressed or encrypted input, in bytes.
        /// </summary>
        public long MaxInputBytes { get; }

        /// <summary>
        /// Gets the maximum number of ZIP entries, including empty entries.
        /// </summary>
        public int MaxZipEntries { get; }

        /// <summary>
        /// Gets the maximum uncompressed size of one ZIP entry, in bytes.
        /// </summary>
        public long MaxEntryBytes { get; }

        /// <summary>
        /// Gets the maximum sum of uncompressed ZIP entry sizes, in bytes.
        /// </summary>
        public long MaxTotalUncompressedBytes { get; }

        /// <summary>
        /// Gets the maximum uncompressed size of one XML, relationships, or VML entry, in bytes.
        /// </summary>
        public long MaxXmlBytes { get; }

        /// <summary>
        /// Gets the maximum ratio of uncompressed to compressed bytes for one ZIP entry.
        /// </summary>
        public double MaxCompressionRatio { get; }

        /// <summary>
        /// Gets the maximum password-hash iterations allowed for an encrypted package.
        /// </summary>
        public int MaxPasswordHashIterations { get; }

        /// <summary>
        /// Gets the maximum size of encryption metadata, in bytes.
        /// </summary>
        public long MaxEncryptionMetadataBytes { get; }
    }
#pragma warning restore CS3014
}
