using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes.TypeSystem
{
    /// <summary>
    /// 타입 정의의 전체 이름을 포함합니다.
    /// 전체 타입 이름은 단일 어셈블리 내의 타입 정의를 고유하게 식별합니다.
    /// </summary>
    [Serializable]
    public struct FullTypeName : IEquatable<FullTypeName>
    {
        [Serializable]
        struct NestedTypeName
        {
            public readonly string Name;
            public readonly int AdditionalTypeParameterCount;

            public NestedTypeName(string name, int additionalTypeParameterCount)
            {
                if (name == null) throw new ArgumentException("이름은 null일 수 없습니다.");
                this.Name = name;
                this.AdditionalTypeParameterCount = additionalTypeParameterCount;

            }
        }


        //readonly TopLevelTypeName topLevelType;
        //readonly NestedTypeName[] nestedTypes;

        //FullTypeName(TopLevelTypeName topLevelTypeName, NestedTypeName[] nestedTypes)
        //{
        //    this.topLevelType = topLevelTypeName;
        //    this.nestedTypes = nestedTypes;
        //}
        /// <summary>
		/// Gets the name of the type.
		/// For nested types, this is the name of the innermost type.
		/// </summary>
		//public string Name
  //      {
  //          get
  //          {
  //              if (nestedTypes != null)
  //                  return nestedTypes[nestedTypes.Length - 1].Name;
  //              else
  //                  return topLevelType.Name;
  //          }
  //      }

        public bool Equals(FullTypeName other)
        {
            throw new NotImplementedException();
        }
    }
}
