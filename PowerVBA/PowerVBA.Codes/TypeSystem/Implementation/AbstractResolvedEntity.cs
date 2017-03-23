using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes.TypeSystem.Implementation
{
    /// <summary>
    /// 해석되지 않은 독립체를 해석하는 <see cref="IEntity"/>를 구현합니다.
    /// </summary>
    public class AbstractResolvedEntity : IEntity
    {
        protected readonly IUnresolvedEntity unresolved;
        protected readonly ITypeResolveContext parentContext;
        

        public ITypeDefinition DeclaringTypeDefinition => parentContext.CurrentTypeDefinition;

        public IType DeclaringType => parentContext.CurrentTypeDefinition;

        public IAssembly ParentAssembly => parentContext.CurrentAssembly;

        public DomRegion Region => unresolved.Region;
        public DomRegion BodyRegion => unresolved.BodyRegion;

        public bool IsStatic => unresolved.IsStatic;
        public bool IsSynthetic => unresolved.IsSynthetic;
        public bool IsPrivate => unresolved.IsPrivate;
        public bool IsPublic => unresolved.IsPublic;
        public Accessibility Accessibility => unresolved.Accessibility;

        public string Name => unresolved.Name;
        public string FullName => unresolved.FullName;
        public string ReflectionName => unresolved.ReflectionName;


        public SymbolKind SymbolKind => unresolved.SymbolKind;
        public ICompilation Compilation => parentContext.Compilation;
        
        public string Namespace => unresolved.Namespace;

        

        public ISymbolReference ToReference()
        {
            throw new NotImplementedException();
        }
    }
}
