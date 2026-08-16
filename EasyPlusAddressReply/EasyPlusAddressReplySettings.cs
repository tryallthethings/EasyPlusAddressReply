using System;
using System.IO;
using System.Xml;
using System.Xml.Serialization;

namespace EasyPlusAddressReply
{
    [Serializable]
    public sealed class EasyPlusAddressReplySettings
    {
        public bool Enabled { get; set; } = true;
        public bool UseDeliveryHeaderFallback { get; set; } = true;
        public bool WarnOnAmbiguousMatch { get; set; } = true;

        public EasyPlusAddressReplySettings Clone()
        {
            return new EasyPlusAddressReplySettings
            {
                Enabled = Enabled,
                UseDeliveryHeaderFallback = UseDeliveryHeaderFallback,
                WarnOnAmbiguousMatch = WarnOnAmbiguousMatch
            };
        }
    }

    internal sealed class SettingsStore
    {
        private const long MaxSettingsFileCharacters = 16 * 1024;

        private readonly string _directory;
        private readonly string _path;

        public SettingsStore()
        {
            _directory = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "EasyPlusAddressReply");
            _path = Path.Combine(_directory, "settings.xml");
        }

        public EasyPlusAddressReplySettings Load()
        {
            try
            {
                if (!File.Exists(_path))
                    return new EasyPlusAddressReplySettings();

                using (var stream = new FileStream(
                    _path,
                    FileMode.Open,
                    FileAccess.Read,
                    FileShare.Read))
                using (var reader = XmlReader.Create(stream, new XmlReaderSettings
                {
                    DtdProcessing = DtdProcessing.Prohibit,
                    XmlResolver = null,
                    MaxCharactersInDocument = MaxSettingsFileCharacters,
                    IgnoreComments = true,
                    IgnoreProcessingInstructions = true
                }))
                {
                    var settings = (EasyPlusAddressReplySettings)new XmlSerializer(
                        typeof(EasyPlusAddressReplySettings)).Deserialize(reader);

                    return settings ?? new EasyPlusAddressReplySettings();
                }
            }
            catch
            {
                // A corrupt, oversized, or inaccessible preference file must never break Outlook startup.
                return new EasyPlusAddressReplySettings();
            }
        }

        public void Save(EasyPlusAddressReplySettings settings)
        {
            if (settings == null)
                throw new ArgumentNullException(nameof(settings));

            Directory.CreateDirectory(_directory);
            string tempPath = _path + ".tmp";

            try
            {
                using (var stream = new FileStream(
                    tempPath,
                    FileMode.Create,
                    FileAccess.Write,
                    FileShare.None))
                {
                    new XmlSerializer(typeof(EasyPlusAddressReplySettings)).Serialize(stream, settings);
                    stream.Flush(true);
                }

                if (File.Exists(_path))
                {
                    try
                    {
                        File.Replace(tempPath, _path, null, true);
                    }
                    catch (Exception ex) when (ex is PlatformNotSupportedException || ex is IOException)
                    {
                        // File systems that cannot do an atomic replace still need the setting saved.
                        File.Delete(_path);
                        File.Move(tempPath, _path);
                    }
                }
                else
                {
                    File.Move(tempPath, _path);
                }

                tempPath = null;
            }
            finally
            {
                if (!string.IsNullOrEmpty(tempPath))
                {
                    try { File.Delete(tempPath); } catch { }
                }
            }
        }
    }
}
