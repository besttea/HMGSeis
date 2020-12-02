namespace HMGSeis
{
    public class Solid 
    {

        private int [] elements=new int[]{0,0,0,0,0,0,0,0};
        private ZkPoints[] pointsMatrix=new ZkPoints[8]();
        private int index =0;
        private string name="";

        public Solid()
        {
        }
         public Solid(int _index,string _name, int[] _elements,ZkPoints[] _pointMatrix)
        {
            index=_index;
            name=_name;
            elements=_elements;
            pointMatrix=_pointMatrix
        }       
        

        public int[] Elements
         { get => elements; 
            set => elements = value; }
        public int Index { get => index; set => index = value; }
        public string Name { get => name; set => name = value; }
        internal ZkPoints[] PointsMatrix1 { get => pointsMatrix; set => pointsMatrix = value; }
    }
}
