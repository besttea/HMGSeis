using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HMGSeis
{
   public class ZkPoints:IComparable<ZkPoints>
    {
        private double[] coordinate=new double[3] { 0,0,0};
        private double x = 0;
        private double y = 0;
        private double z = 0;
        private string coordSys = "GLOBAL";
        private int connectNumber = 0;

        private int index = 0;
        private string name = "";
       public double[] Coordinate {
            get
            { return coordinate; }
            set
            {
                coordinate[0] = x;
                coordinate[1] = y;
                coordinate[2] = z;
            }

        }
        public double X
        { get
            {
                return x;
            }
            set
            {
                x = value;
                this.coordinate[0] = value;
            } }
        public double Y {
            get
            {
                return y;
            }
            set
            {
                y = value;
                this.coordinate[1] = value;
            }}
        public double Z {
            get
            {
                return z;
            }
            set
            {
                z = value;
                this.coordinate[2] = value;
            }
        }
        public int Index
        {
            get { return index; }
            set { index = value; }
        }
        public string Name
        {
            get { return name; }
            set { name = value; }
                }

        public string CoordSys { get => coordSys; set => coordSys = value; }
        public int ConnectNumber { get => connectNumber; set => connectNumber = value; }

        //void public zkPoints(void);
        public  ZkPoints()
     {
            x = 0;y = 0;z = 0;
            index = 0;
            Name = "";
            coordinate[0] = 0;
            coordinate[1] = 0;
            coordinate[2] = 0;
        }
        public ZkPoints(int _index, string _Name, double _x,double _y,double _z)
        {
            index = _index;
            Name = _Name;
            x = _x;y = _y;z = _z;
            coordinate = new double[3];
            coordinate[0] = _x;
            coordinate[1] = _y;
            coordinate[2] = _z;
        }

        public int CompareTo(ZkPoints other)
        {
            return this.X.CompareTo(other.X);
        }
        public void Printme()
        {  
                Console.WriteLine(@"{0},{1},{2},{3},{4}", this.Index, this.Name, this.X, this.Y, this.Z);
        }
        public void Get_Points_ConnectorNum(int _connectNumber)
        {

            connectNumber = _connectNumber;

        }
    }


    class ComparePoints_X : IComparer<ZkPoints>
    {
        public int Compare(ZkPoints x, ZkPoints y)
        {
            return x.X.CompareTo(y.X); 
        }
    }
}
