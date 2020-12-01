using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HMGSeis
{
    class ZkPoints:IComparable<ZkPoints>
    {
        public double[] Coordinate { get; set; }
        public double x { get; set; }
        public double y { get; set; }
        public double z { get; set; }
        public int index { get; set; }
        public string Name { get; set; }
        //void public zkPoints(void);
        public  ZkPoints()
     {
            x = 0;y = 0;z = 0;
            index = 0;
            Name = "";
            Coordinate = new double[3];
        }
        public ZkPoints(int _index, string _Name, double _x,double _y,double _z)
        {
            index = _index;
            Name = _Name;
            x = _x;y = _y;z = _z;
            Coordinate = new double[3];
            Coordinate[0] = _x;
            Coordinate[1] = _y;
            Coordinate[2] = _z;
        }

        public int CompareTo(ZkPoints other)
        {
            return this.x.CompareTo(other.x);
        }
    }


    class ComparePoints_X : IComparer<ZkPoints>
    {
        public int Compare(ZkPoints x, ZkPoints y)
        {
            return x.x.CompareTo(y.x); 
        }
    }
}
