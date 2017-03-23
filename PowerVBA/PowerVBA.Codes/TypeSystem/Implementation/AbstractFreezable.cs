using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes.TypeSystem.Implementation
{
    [Serializable]
    public abstract class AbstractFreezable : IFreezable
    {
        bool isFrozen;
        public bool IsFrozen { get => isFrozen; }

        public void Freeze()
        {
            if (!isFrozen)
            {
                FreezeInternal();
                isFrozen = true;
            }
        }

        /// <summary>
        /// 내부 고정
        /// </summary>
        protected virtual void FreezeInternal()
        {

        }
    }
}
