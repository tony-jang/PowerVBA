using PowerVBA.Codes.Attributes;
using PowerVBA.Codes.Extension;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PowerVBA.Codes
{
    public class Locator
    {
        
        internal Locator(Stack<(CodeType, int)> init)
        {
            LocationList = init;
        }

        public Locator()
        {
            LocationList = new Stack<(CodeType, int)>();
        }

        public Stack<(CodeType, int)> LocationList { get; internal set; }

        public bool Insert(CodeType type, int line)
        {
            if (type.ContainsAttribute(typeof(CoverableAttribute)))
            {
                LocationList.Push((type, line));
                return true;
            }
            return false;
        }

        /// <summary>
        /// 로케이터가 가리키고 있는 맨 위의 아이템을 가져옵니다.
        /// 만약 Key와 로케이터의 맨 위의 아이템의 Key와 다를 경우 False를 반환합니다.
        /// </summary>
        /// <param name="Key">삭제할 CodeType입니다. 타입이 동일한지 확인합니다.</param>
        public bool Delete(CodeType type)
        {
            try
            {
                if (LocationList.Peek().Item1 == type)
                {
                    LocationList.Pop();
                    return true;
                }
            }
            catch (Exception)
            {
            }
            
            return false;
        }

        public bool Contains(CodeType type)
        {
            return LocationList
                .Where(i => i.Item1 == type)
                .Count() != 0;
        }

        public Locator Clone()
        {
            return new Locator(LocationList);
        }
    }
}
