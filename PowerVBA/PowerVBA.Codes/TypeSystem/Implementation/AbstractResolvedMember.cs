//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;

//namespace PowerVBA.Codes.TypeSystem.Implementation
//{
//    /// <summary>
//    /// Implementation of <see cref="IMember"/> that resolves an unresolved member.
//    /// </summary>
//    public abstract class AbstractResolvedMember : AbstractResolvedEntity, IMember
//    {
//        protected new readonly IUnresolvedMember unresolved;
//        protected readonly ITypeResolveContext context;
//        volatile IType returnType;
//        IList<IMember> implementedInterfaceMembers;

//        protected AbstractResolvedMember(IUnresolvedMember unresolved, ITypeResolveContext parentContext)
//            : base(unresolved, parentContext)
//        {
//            this.unresolved = unresolved;
//            this.context = parentContext.WithCurrentMember(this);
//        }

//        IMember IMember.MemberDefinition
//        {
//            get { return this; }
//        }

//        public IType ReturnType => this.returnType ?? (this.returnType = unresolved.ReturnType.Resolve(context));

//        public IUnresolvedMember UnresolvedMember
//        {
//            get { return unresolved; }
//        }

//        public IList<IMember> ImplementedInterfaceMembers
//        {
//            get
//            {
//                IList<IMember> result = LazyInit.VolatileRead(ref this.implementedInterfaceMembers);
//                if (result != null)
//                {
//                    return result;
//                }
//                else
//                {
//                    return LazyInit.GetOrSet(ref implementedInterfaceMembers, FindImplementedInterfaceMembers());
//                }
//            }
//        }

//        public override DocumentationComment Documentation
//        {
//            get
//            {
//                IUnresolvedDocumentationProvider docProvider = unresolved.UnresolvedFile as IUnresolvedDocumentationProvider;
//                if (docProvider != null)
//                {
//                    var doc = docProvider.GetDocumentation(unresolved, this);
//                    if (doc != null)
//                        return doc;
//                }
//                return base.Documentation;
//            }
//        }
        
//        public TypeParameterSubstitution Substitution
//        {
//            get { return TypeParameterSubstitution.Identity; }
//        }

//        public abstract IMember Specialize(TypeParameterSubstitution substitution);

//        IMemberReference IMember.ToReference()
//        {
//            return (IMemberReference)ToReference();
//        }

//        public override ISymbolReference ToReference()
//        {
//            var declType = this.DeclaringType;
//            var declTypeRef = declType != null ? declType.ToTypeReference() : SpecialType.UnknownType;
//            if (IsExplicitInterfaceImplementation && ImplementedInterfaceMembers.Count == 1)
//            {
//                return new ExplicitInterfaceImplementationMemberReference(declTypeRef, ImplementedInterfaceMembers[0].ToReference());
//            }
//            else
//            {
//                return new DefaultMemberReference(this.SymbolKind, declTypeRef, this.Name);
//            }
//        }

//        public virtual IMemberReference ToMemberReference()
//        {
//            return (IMemberReference)ToReference();
//        }

//        internal IMethod GetAccessor(ref IMethod accessorField, IUnresolvedMethod unresolvedAccessor)
//        {
//            if (unresolvedAccessor == null)
//                return null;
//            IMethod result = LazyInit.VolatileRead(ref accessorField);
//            if (result != null)
//            {
//                return result;
//            }
//            else
//            {
//                return LazyInit.GetOrSet(ref accessorField, CreateResolvedAccessor(unresolvedAccessor));
//            }
//        }

//        protected virtual IMethod CreateResolvedAccessor(IUnresolvedMethod unresolvedAccessor)
//        {
//            return (IMethod)unresolvedAccessor.CreateResolved(context);
//        }
//    }
//}
