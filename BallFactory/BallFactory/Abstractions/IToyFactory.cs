using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BallFactory.Abstractions
{
    public interface IToyFactory
    {
        Toy CreateNew();
    }
}
