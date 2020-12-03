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
            index = _index;
            name = _name;
            x = _x;
            y = _y;
            z = _z;

        }

        public string[] Elements
         { get => elements; 
            set => elements = value; }
        public int Index { get => index; set => index = value; }
        public string Name { get => name; set => name = value; }
        internal ZkPoints[] PointsMatrix { get => pointsMatrix; set => pointsMatrix = value; }
        internal double[] X { get => x; set => x = value; }
        internal double[] Y { get => x; set => x = value; }
        internal double[] Z { get => x; set => x = value; }
    }
}
