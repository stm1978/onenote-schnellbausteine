using System;
using System.Runtime.Serialization;

namespace Schnellbausteine.Core
{
    [DataContract]
    public sealed class Snippet
    {
        [DataMember(Order = 1)]
        public string Id { get; set; }

        [DataMember(Order = 2)]
        public string Name { get; set; }

        [DataMember(Order = 3)]
        public string FileName { get; set; }

        [DataMember(Order = 4)]
        public DateTime CreatedUtc { get; set; }

        [DataMember(Order = 5)]
        public DateTime UpdatedUtc { get; set; }
    }

    [DataContract]
    internal sealed class SnippetCatalog
    {
        public SnippetCatalog()
        {
            Version = 1;
            Snippets = new Snippet[0];
        }

        [DataMember(Order = 1)]
        public int Version { get; set; }

        [DataMember(Order = 2)]
        public Snippet[] Snippets { get; set; }
    }
}
