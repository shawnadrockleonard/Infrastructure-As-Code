using Microsoft.SharePoint.Client;
using System;

namespace InfrastructureAsCode.Core.Models.Minimal
{
    public class DefinitionFieldModel
    {
        public bool AutoIndexed { get; set; }
        public bool CanBeDeleted { get; set; }
        public string DefaultFormula { get; set; }
        public string DefaultValue { get; set; }
        public string Description { get; set; }
        public bool EnforceUniqueValues { get; set; }
        public FieldType FieldTypeKind { get; set; }
        public bool Filterable { get; set; }
        public string Group { get; set; }
        public bool Hidden { get; set; }
        public Guid Id { get; set; }
        public string InternalName { get; set; }
        public bool Indexed { get; set; }
        public string JSLink { get; set; }
        public bool NoCrawl { get; set; }
        public bool ReadOnlyField { get; set; }
        public bool Required { get; set; }
        public string Title { get; set; }

    }
}