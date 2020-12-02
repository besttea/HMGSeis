using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HMGSeis
{
    class ZkPoints:IComparable<ZkPoints>
    {
        public double[] Coordinate
        { get;

          set; }
        public double x
        { get
            {
                return x;
            }
            set
            {
                x = value;
                this.Coordinate[0] = value;
            } }
        public double y {
            get
            {
                return y;
            }
            set
            {
                y = value;
                this.Coordinate[1] = value;
            }}
        public double z {
            get
            {
                return z;
            }
            set
            {
                z = value;
                this.Coordinate[2] = value;
            }
        }
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
