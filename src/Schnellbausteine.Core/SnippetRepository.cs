using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace Schnellbausteine.Core
{
    public sealed class SnippetRepository
    {
        private readonly string root;
        private readonly string catalogPath;
        private readonly string snippetsPath;
        private readonly object gate = new object();

        public SnippetRepository(string rootDirectory)
        {
            if (string.IsNullOrWhiteSpace(rootDirectory)) throw new ArgumentException("Ein Speicherpfad ist erforderlich.", "rootDirectory");
            root = Path.GetFullPath(rootDirectory);
            catalogPath = Path.Combine(root, "catalog.json");
            snippetsPath = Path.Combine(root, "snippets");
        }

        public static string DefaultRoot
        {
            get { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Schnellbausteine", "Data"); }
        }

        public IList<Snippet> GetAll()
        {
            lock (gate)
            {
                return ReadCatalog().Snippets
                    .OrderBy(item => item.Name, StringComparer.CurrentCultureIgnoreCase)
                    .Select(Clone)
                    .ToList();
            }
        }

        public bool ContainsName(string name, string exceptId)
        {
            string normalized = NormalizeName(name);
            lock (gate)
            {
                return ReadCatalog().Snippets.Any(item =>
                    !string.Equals(item.Id, exceptId, StringComparison.Ordinal) &&
                    string.Equals(item.Name, normalized, StringComparison.CurrentCultureIgnoreCase));
            }
        }

        public Snippet Save(string name, XDocument fragment, bool overwrite)
        {
            if (fragment == null) throw new ArgumentNullException("fragment");
            string normalized = NormalizeName(name);

            lock (gate)
            {
                SnippetCatalog catalog = ReadCatalog();
                Snippet existing = catalog.Snippets.FirstOrDefault(candidate =>
                    string.Equals(candidate.Name, normalized, StringComparison.CurrentCultureIgnoreCase));

                if (existing != null && !overwrite)
                {
                    throw new SnippetNameConflictException(normalized);
                }

                DateTime now = DateTime.UtcNow;
                Snippet item = existing ?? new Snippet
                {
                    Id = Guid.NewGuid().ToString("N"),
                    FileName = Guid.NewGuid().ToString("N") + ".xml",
                    CreatedUtc = now
                };
                item.Name = normalized;
                item.UpdatedUtc = now;

                Directory.CreateDirectory(snippetsPath);
                AtomicFile.WriteAllText(Path.Combine(snippetsPath, item.FileName), fragment.ToString(SaveOptions.DisableFormatting));

                var items = catalog.Snippets.Where(candidate => !string.Equals(candidate.Id, item.Id, StringComparison.Ordinal)).ToList();
                items.Add(item);
                catalog.Snippets = items.ToArray();
                WriteCatalog(catalog);
                return Clone(item);
            }
        }

        public XDocument LoadFragment(string id)
        {
            lock (gate)
            {
                Snippet item = Find(ReadCatalog(), id);
                string path = Path.Combine(snippetsPath, item.FileName);
                if (!File.Exists(path)) throw new SnippetRepositoryException("Die Fragmentdatei fehlt: " + path);
                return XDocument.Load(path, LoadOptions.PreserveWhitespace);
            }
        }

        public Snippet Rename(string id, string newName)
        {
            string normalized = NormalizeName(newName);
            lock (gate)
            {
                SnippetCatalog catalog = ReadCatalog();
                Snippet item = Find(catalog, id);
                if (catalog.Snippets.Any(candidate =>
                    !string.Equals(candidate.Id, id, StringComparison.Ordinal) &&
                    string.Equals(candidate.Name, normalized, StringComparison.CurrentCultureIgnoreCase)))
                {
                    throw new SnippetNameConflictException(normalized);
                }

                item.Name = normalized;
                item.UpdatedUtc = DateTime.UtcNow;
                WriteCatalog(catalog);
                return Clone(item);
            }
        }

        public void Delete(string id)
        {
            lock (gate)
            {
                SnippetCatalog catalog = ReadCatalog();
                Snippet item = Find(catalog, id);
                catalog.Snippets = catalog.Snippets.Where(candidate => !string.Equals(candidate.Id, id, StringComparison.Ordinal)).ToArray();
                WriteCatalog(catalog);

                string path = Path.Combine(snippetsPath, item.FileName);
                if (File.Exists(path)) File.Delete(path);
            }
        }

        private SnippetCatalog ReadCatalog()
        {
            if (!File.Exists(catalogPath)) return new SnippetCatalog();
            try
            {
                using (var stream = File.OpenRead(catalogPath))
                {
                    var serializer = new DataContractJsonSerializer(typeof(SnippetCatalog));
                    var catalog = (SnippetCatalog)serializer.ReadObject(stream);
                    if (catalog == null || catalog.Version != 1 || catalog.Snippets == null)
                    {
                        throw new SnippetRepositoryException("Nicht unterstützte Katalogversion in " + catalogPath);
                    }

                    return catalog;
                }
            }
            catch (SnippetRepositoryException)
            {
                throw;
            }
            catch (Exception exception)
            {
                throw new SnippetRepositoryException("Der Schnellbaustein-Katalog kann nicht gelesen werden: " + catalogPath, exception);
            }
        }

        private void WriteCatalog(SnippetCatalog catalog)
        {
            Directory.CreateDirectory(root);
            var serializer = new DataContractJsonSerializer(typeof(SnippetCatalog));
            using (var stream = new MemoryStream())
            {
                serializer.WriteObject(stream, catalog);
                AtomicFile.WriteAllText(catalogPath, Encoding.UTF8.GetString(stream.ToArray()));
            }
        }

        private static Snippet Find(SnippetCatalog catalog, string id)
        {
            if (string.IsNullOrWhiteSpace(id)) throw new ArgumentException("Eine ID ist erforderlich.", "id");
            Snippet item = catalog.Snippets.FirstOrDefault(candidate => string.Equals(candidate.Id, id, StringComparison.Ordinal));
            if (item == null) throw new SnippetRepositoryException("Der Schnellbaustein wurde nicht gefunden: " + id);
            return item;
        }

        private static string NormalizeName(string name)
        {
            if (name == null) throw new ArgumentNullException("name");
            string value = name.Trim();
            if (value.Length == 0) throw new ArgumentException("Der Name darf nicht leer sein.", "name");
            if (value.Length > 100) throw new ArgumentException("Der Name darf höchstens 100 Zeichen enthalten.", "name");
            if (value.Any(char.IsControl)) throw new ArgumentException("Der Name enthält unzulässige Steuerzeichen.", "name");
            return value;
        }

        private static Snippet Clone(Snippet item)
        {
            return new Snippet
            {
                Id = item.Id,
                Name = item.Name,
                FileName = item.FileName,
                CreatedUtc = item.CreatedUtc,
                UpdatedUtc = item.UpdatedUtc
            };
        }
    }

    public sealed class SnippetNameConflictException : Exception
    {
        public SnippetNameConflictException(string name)
            : base("Ein Schnellbaustein mit dem Namen \"" + name + "\" existiert bereits.")
        {
        }
    }

    public sealed class SnippetRepositoryException : Exception
    {
        public SnippetRepositoryException(string message) : base(message) { }
        public SnippetRepositoryException(string message, Exception innerException) : base(message, innerException) { }
    }
}
