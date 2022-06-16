using System.Collections.Generic;
using Tekla.Structures.Geometry3d;
using Tekla.Structures.Solid;

namespace EveryThing
{
    /// <summary>
    /// Контакт грани и ребра
    /// </summary>
    public class ContactEdge
    {
        public List<Point> FaceVertices
        {
            get;
            private set;
        }
        public Face Face
        {
            get;
            private set;
        }

        public LineSegment Edge
        {
            get;
            private set;
        }

        public ContactEdge(List<Point> vertices, Face face, LineSegment edge)
        {
            this.FaceVertices = vertices;
            this.Face = face;
            this.Edge = edge;
        }

    }
}
