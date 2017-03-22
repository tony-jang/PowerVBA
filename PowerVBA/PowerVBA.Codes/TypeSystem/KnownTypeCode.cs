using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes.TypeSystem
{
    public enum KnownTypeCode
    {
        None,
        Object,
        Boolean,
        Byte,
        Currency,
        DataObject,
        Date,
        Decimal,
        Double,
        Empty,
        Error,
        Char,
        Integer,
        Single,
        DateTime,
        Variant,
        String = 17,


        Enum,
        Type,
        Collection
    }

    public sealed class KnownTypeReference : ITypeReference
    {
        internal const int KnownTypeCodeCount = (int)KnownTypeCode.Collection + 1;

        public KnownTypeReference(KnownTypeCode knownTypeCode, string namespaceName, string name, int typeParameterCount = 0)
        {
            this.knownTypeCode = knownTypeCode;
            this.namespaceName = namespaceName;
            this.name = name;
            this.typeParameterCount = typeParameterCount;
        }

        static readonly KnownTypeReference[] knownTypeReferences = new KnownTypeReference[KnownTypeCodeCount] {
            null, // 없음
            new KnownTypeReference(KnownTypeCode.Object,        "", "Object"),
            new KnownTypeReference(KnownTypeCode.Boolean,       "", "Boolean"),
            new KnownTypeReference(KnownTypeCode.Byte,          "", "Char"),
            new KnownTypeReference(KnownTypeCode.Currency,      "", "Currency"),
            new KnownTypeReference(KnownTypeCode.DataObject,    "", "DataObject"),
            new KnownTypeReference(KnownTypeCode.Date,          "", "Date"),
            new KnownTypeReference(KnownTypeCode.Decimal,       "", "Decimal"),
            new KnownTypeReference(KnownTypeCode.Double,        "", "Double"),
            new KnownTypeReference(KnownTypeCode.Empty,         "", "Empty"),
            new KnownTypeReference(KnownTypeCode.Error,         "", "Error"),
            new KnownTypeReference(KnownTypeCode.Char,          "", "Char"),
            new KnownTypeReference(KnownTypeCode.Integer,       "", "Integer"),
            new KnownTypeReference(KnownTypeCode.Single,        "", "Single"),
            new KnownTypeReference(KnownTypeCode.DateTime,      "", "DateTime"),
            new KnownTypeReference(KnownTypeCode.Variant,       "", "Variant"),
            null,
            new KnownTypeReference(KnownTypeCode.String,        "", "String"),
            new KnownTypeReference(KnownTypeCode.Enum,          "", "Enum"),
            new KnownTypeReference(KnownTypeCode.Type,          "", "Type"),
            new KnownTypeReference(KnownTypeCode.Collection, "VBA", "Collection")
        };

        public static KnownTypeReference Get(KnownTypeCode typeCode)
        {
            return knownTypeReferences[(int)typeCode];
        }

        #region [  Get  ]

        public static readonly KnownTypeReference Object = Get(KnownTypeCode.Object);
        public static readonly KnownTypeReference Boolean = Get(KnownTypeCode.Boolean);
        public static readonly KnownTypeReference Byte = Get(KnownTypeCode.Byte);
        public static readonly KnownTypeReference Currency = Get(KnownTypeCode.Currency);
        public static readonly KnownTypeReference DataObject = Get(KnownTypeCode.DataObject);
        public static readonly KnownTypeReference Date = Get(KnownTypeCode.Date);
        public static readonly KnownTypeReference Decimal = Get(KnownTypeCode.Decimal);
        public static readonly KnownTypeReference Double = Get(KnownTypeCode.Double);
        public static readonly KnownTypeReference Empty = Get(KnownTypeCode.Empty);
        public static readonly KnownTypeReference Error = Get(KnownTypeCode.Error);
        public static readonly KnownTypeReference Char = Get(KnownTypeCode.Char);
        public static readonly KnownTypeReference Integer = Get(KnownTypeCode.Integer);
        public static readonly KnownTypeReference Single = Get(KnownTypeCode.Single);
        public static readonly KnownTypeReference DateTime = Get(KnownTypeCode.DateTime);
        public static readonly KnownTypeReference Variant = Get(KnownTypeCode.Variant);
        public static readonly KnownTypeReference String = Get(KnownTypeCode.String);
        public static readonly KnownTypeReference Enum = Get(KnownTypeCode.Enum);
        public static readonly KnownTypeReference Type = Get(KnownTypeCode.Type);
        public static readonly KnownTypeReference Collection = Get(KnownTypeCode.Collection);

        #endregion

        public IType Resolve(ITypeResolveContext context)
        {
            return context.Compilation.FindType(knownTypeCode);
        }
        readonly KnownTypeCode knownTypeCode;
        readonly string namespaceName;
        readonly string name;
        readonly int typeParameterCount;
        internal readonly KnownTypeCode baseType;
        public KnownTypeCode KnownTypeCode
        {
            get { return knownTypeCode; }
        }

        public string Namespace
        {
            get { return namespaceName; }
        }

        public string Name
        {
            get { return name; }
        }

        public int TypeParameterCount
        {
            get { return typeParameterCount; }
        }

        public override string ToString()
        {
            return GetVBANameByTypeCode(knownTypeCode) ?? ((string.IsNullOrEmpty(this.Namespace) ? "" : this.Namespace + ".") + this.Name);
        }

        public static string GetVBANameByTypeCode(KnownTypeCode knownTypeCode)
        {
            switch (knownTypeCode)
            {
                case KnownTypeCode.Object:
                    return "Object";
                case KnownTypeCode.Boolean:
                    return "Boolean";
                case KnownTypeCode.Byte:
                    return "Byte";
                case KnownTypeCode.DataObject:
                    return "DataObject";
                case KnownTypeCode.Date:
                    return "Date";
                case KnownTypeCode.Decimal:
                    return "Decimal";
                case KnownTypeCode.Double:
                    return "Double";
                case KnownTypeCode.Empty:
                    return "Empty";
                case KnownTypeCode.Error:
                    return "Error";
                case KnownTypeCode.Char:
                    return "Char";
                case KnownTypeCode.Integer:
                    return "Integer";
                case KnownTypeCode.Single:
                    return "Single";
                case KnownTypeCode.DateTime:
                    return "DateTime";
                case KnownTypeCode.Variant:
                    return "Variant";
                case KnownTypeCode.String:
                    return "String";
                default:
                    return null;
            }
        }


    }
}
