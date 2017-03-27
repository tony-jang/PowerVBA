using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes.TypeSystem
{
    public interface IFreezable
    {
        /// <summary>
        /// 이 인스턴스가 고정되어있는 경우 가져옵니다. 고정 된 인스턴스는 변경 가능하지 않으므로 스레드로부터 안전합니다.
        /// </summary>
        bool IsFrozen { get; }

        /// <summary>
        /// 이 인스턴스를 고정시킵니다.
        /// </summary>
        void Freeze();
    }
}
