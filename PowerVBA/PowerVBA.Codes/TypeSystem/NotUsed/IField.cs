using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes.TypeSystem
{
    /// <summary>
    /// Represents a field or constant.
    /// </summary>
    public interface IField : IMember, IVariable
    {
        /// <summary>
        /// Gets the name of the field.
        /// </summary>
        new string Name { get; } // solve ambiguity between IMember.Name and IVariable.Name

        /// <summary>
        /// Gets the region where the field is declared.
        /// </summary>
        new DomRegion Region { get; } // solve ambiguity between IEntity.Region and IVariable.Region

        /// <summary>
        /// Gets whether this field is readonly.
        /// </summary>
        bool IsReadOnly { get; }

        /// <summary>
        /// Gets whether this field is volatile.
        /// </summary>
        bool IsVolatile { get; }

        /// <summary>
        /// Gets whether this field is a fixed size buffer (C#-like fixed).
        /// If this is true, then ConstantValue contains the size of the buffer.
        /// </summary>
        bool IsFixed { get; }

        new IMemberReference ToReference(); // solve ambiguity between IMember.ToReference() and IVariable.ToReference()
    }
}
