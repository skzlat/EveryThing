using System;
using System.Collections.Generic;
using Tekla.Structures.Geometry3d;
using Tekla.Structures.Solid;

namespace EveryThing
{
    public class ContactPlane : IComparable
    {
        int IComparable.CompareTo(object o)
        {
            var plane1 = new GeometricPlane(FirstFaceVertex[0], FirstFace.Normal);
            var dist1 = Distance.PointToPlane(SecondFaceVertex[0], plane1);
            var c = o as ContactPlane;
            var plane2 = new GeometricPlane(c.FirstFaceVertex[0], c.FirstFace.Normal);
            var dist2 = Distance.PointToPlane(c.SecondFaceVertex[0], plane2);
            if (dist1 < dist2) return -1;
            else if (dist1 > dist2) return 1;
            return 0;
        }
        public List<Point> FirstFaceVertex
        {
            get;
            set;
        }
        public List<Point> SecondFaceVertex
        {
            get;
            set;
        }
        public Face FirstFace
        {
            get;
            set;
        }
        public Face SecondFace
        {
            get;
            set;
        }
        public ContactPlane(IEnumerable<Point> firstFaceVertex, IEnumerable<Point> secondFaceVertex, Face firstFace, Face secondFace)
        {
            this.FirstFaceVertex = new List<Point>(firstFaceVertex);
            this.SecondFaceVertex = new List<Point>(secondFaceVertex);
            this.FirstFace = firstFace;
            this.SecondFace = secondFace;
        }
    }
}
