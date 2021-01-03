using System.Collections.Generic;

namespace HMGSeis
{
    public class Solid 
    {

        private string [] elements=new string[]{"","","","", "", "", "", "" };
        private ZkPoints[] pointsMatrix=new ZkPoints[8];
        private int index =0;
        private double[] x = new double[8] { 0, 0, 0, 0, 0, 0, 0, 0 };
        private double[] y = new double[8] { 0, 0, 0, 0, 0, 0, 0, 0 };
        private double[] z = new double[8] { 0, 0, 0, 0, 0, 0, 0, 0 };
        private string name="";
        public string[] Elements
        {
            get => elements;
            set => elements = value;
        }
        public int Index { get => index; set => index = value; }
        public string Name { get => name; set => name = value; }
        internal ZkPoints[] PointsMatrix { get => pointsMatrix; set => pointsMatrix = value; }
        internal double[] X { get => x; set => x = value; }
        internal double[] Y { get => y; set => y = value; }
        internal double[] Z { get => z; set => z = value; }


        public Solid()
        {
        }
         public Solid(int _index,string _name, string[] _elements)
        {
            index=_index;
            name=_name;
            elements=_elements;
            
        }
        public Solid(int _index, string _name, double[] _x,double[]_y,double[]_z)
        {
            SetPointsArray(_index, _name, _x, _y, _z);

        }
        public Solid(int __index,string _name, ZkPoints[] _PointsMatrix)
        {
            SetPointsMatrix(__index,  _name, PointsMatrix);
        }
        public Solid(int __index, string _name, List<ZkPoints> _PointsMatrix)
        {
            SetPointsMatrix(__index, _name, PointsMatrix);
        }

        public void SetPointsMatrix(int _index,string _name, ZkPoints[] _PointsMatrix)
        {
            index = _index;
            name = _name;
            PointsMatrix = _PointsMatrix;
            for (int i = 0; i < 8; i++)
            {
                elements[i] = pointsMatrix[i].Name;
                x[i] = pointsMatrix[i].X;
                y[i] = pointsMatrix[i].Y;
                z[i] = pointsMatrix[i].Z;
                
            }


        }

        public void SetPointsMatrix(int _index, string _name,List <ZkPoints> _PointsMatrix)
        {
            index = _index;
            name = _name;
            
            for (int i = 0; i < 8; i++)
            {
                elements[i] = _PointsMatrix[i].Name;
                x[i] = _PointsMatrix[i].X;
                y[i] = _PointsMatrix[i].Y;
                z[i] = _PointsMatrix[i].Z;

            }


        }

        public void SetPointsArray(int _index, string _name, double[] _x, double[] _y, double[] _z)
        {
            index = _index;
            name = _name;
            x = _x;
            y = _y;
            z = _z;

        }

        public void GetPointsArray(ref double[] _x,ref double[] _y,ref double[] _z)
        {

            for (int i = 0; i < 8; i++)
            {

              _x[i] =x[i]; 
              _y[i] =y[i];
              _z[i] =z[i];

            }


        }


    }
}
