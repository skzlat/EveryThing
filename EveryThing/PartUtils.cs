using System;
using System.Collections.Generic;
using System.Linq;
using Tekla.Structures.Geometry3d;
using Tekla.Structures.Model;
using Tekla.Structures.Solid;

namespace EveryThing
{
    public class PartUtils
    {

        #region Контакт грани и ребра

        public static List<ContactEdge> GetContactEdges(Part firstPart, Part secondPart, double gapTolerance)
        {
            var faces1 = GetFaces(firstPart.GetSolid());
            var faces2 = GetFaces(secondPart.GetSolid());

            var edges1 = GetEdges(firstPart.GetSolid());
            var edges2 = GetEdges(secondPart.GetSolid());

            var list = (from f in faces1
                        from e in edges2
                        where IsContact(f, e, gapTolerance)
                        select new ContactEdge(GetFaceEdges(f), f, new LineSegment(e.StartPoint, e.EndPoint))).ToList();
            list.AddRange(from f in faces2
                          from e in edges1
                          where IsContact(f, e, gapTolerance)
                          select new ContactEdge(GetFaceEdges(f), f, new LineSegment(e.StartPoint, e.EndPoint)));
            return list;
        }

        private static bool IsContact(Face face, Edge edge, double gapTolerance)
        {
            var flag = false;
            var loopEnumerator = face.GetLoopEnumerator();
            loopEnumerator.MoveNext();
            var loop = loopEnumerator.Current as Loop;
            var vertexEnumerator = loop.GetVertexEnumerator();
            vertexEnumerator.MoveNext();
            var point = vertexEnumerator.Current as Point;

            var geomPlane = new GeometricPlane(point, face.Normal);
            var line = new Line(edge.StartPoint, edge.EndPoint);

            if (Parallel.LineToPlane(line, geomPlane) == false) return flag;
            if (Distance.PointToPlane(line.Origin, geomPlane) > gapTolerance) return flag;

            var polygon = GetFaceEdges(face);
            var intersectionPoints = GetIntersectionPointsBetweenLineAndPolygon(polygon, new Line(edge.StartPoint, edge.EndPoint));
            return intersectionPoints.Count != 0 || flag;
        }

        #region Получение коллекции точек пересечений между ОДНИМ полигоном и ОДНОЙ линией
        /// <summary>
        /// Получение коллекции точек пересечений между ОДНИМ полигоном и ОДНОЙ линией
        /// </summary>
        /// <param name="polygon">Полигон</param>
        /// <param name="line">Линия</param>
        /// <returns>Коллекция точек пересечений</returns>
        public static List<Point> GetIntersectionPointsBetweenLineAndPolygon(List<Point> polygon, Line line)
        {
            var edges = new List<LineSegment>();
            for (var i = 1; i < polygon.Count; i++)
            {
                edges.Add(new LineSegment(polygon[i - 1], polygon[i]));
            }
            edges.Add(new LineSegment(polygon[polygon.Count - 1], polygon[0]));
            //Line line = new Line(lineSegment.Point1, lineSegment.Point2);
            var intersectionPoints = new List<Point>();
            var intersectionPoints2 = new List<Point>();
            foreach (var s in edges)
            {
                cs_net_lib.Intersect.IntersectLineToLineSegment3D(line, s, ref intersectionPoints2);
                intersectionPoints.AddRange(intersectionPoints2);
            }
            return intersectionPoints;
        }

        #endregion

        private static List<Face> GetFaces(Solid solid)
        {
            var faces = new List<Face>();
            var faceEnumerator = solid.GetFaceEnumerator();
            while (faceEnumerator.MoveNext())
            {
                var current = faceEnumerator.Current as Face;
                if (current != null)
                    faces.Add(current);
            }
            return faces;
        }

        private static List<Edge> GetEdges(Solid solid)
        {
            var edges = new List<Edge>();
            var edgeEnumerator = solid.GetEdgeEnumerator();
            while (edgeEnumerator.MoveNext())
            {
                var current = edgeEnumerator.Current as Edge;
                if (current != null)
                    edges.Add(current);
            }
            return edges;
        }

        #endregion

        #region Контактные поверхности
        public static List<ContactPlane> GetContactPlanes(Part firstPart, Part secondPart, double gapTolerance)
        {
            var list = new List<ContactPlane>();
            var faceEnumerator = firstPart.GetSolid().GetFaceEnumerator();
            var faceEnumerator2 = secondPart.GetSolid().GetFaceEnumerator();
            while (faceEnumerator.MoveNext())
            {
                var current = faceEnumerator.Current;
                if (current == null) 
                    continue;
                var faceEdges = GetFaceEdges(current);
                if (faceEdges.Count > 0)
                {
                    var firstFacePlane = new GeometricPlane(faceEdges[0], current.Normal);
                    faceEnumerator2.Reset();
                    while (faceEnumerator2.MoveNext())
                    {
                        var current2 = faceEnumerator2.Current;
                        if (current2 != null)
                        {
                            var faceEdges2 = GetFaceEdges(current2);
                            if (faceEdges2.Count > 0)
                            {
                                var secondFacePlane = new GeometricPlane(faceEdges2[0], current2.Normal);
                                var flag = CheckIfAreOpposite(firstFacePlane, secondFacePlane, gapTolerance);
                                if (flag)
                                {
                                    var item = new ContactPlane(faceEdges, faceEdges2, current, current2);
                                    list.Add(item);
                                }
                            }
                        }
                    }
                }
            }
            return list;
        }

        public static List<Point> GetFaceEdges(Face face)
        {
            var list = new List<Point>();
            var loopEnumerator = face.GetLoopEnumerator();
            while (loopEnumerator.MoveNext())
            {
                if (loopEnumerator.Current is Loop current)
                {
                    var vertexEnumerator = current.GetVertexEnumerator();
                    while (vertexEnumerator.MoveNext())
                    {
                        var current2 = vertexEnumerator.Current as Point;
                        if (current2 != null)
                        {
                            list.Add(current2);
                        }
                    }
                }
            }
            return list;
        }

        private static bool CheckIfAreOpposite(GeometricPlane firstFacePlane, GeometricPlane secondFacePlane, double gapTolerance)
        {
            var result = false;
            if (Parallel.VectorToVector(firstFacePlane.Normal, secondFacePlane.Normal, 0.0017453292519943296))
            {
                var num = Distance.PointToPlane(secondFacePlane.Origin, firstFacePlane);
                var flag = Vector.Dot(firstFacePlane.Normal, secondFacePlane.Normal) + 1.0 < 0.0017453292519943296;
                if (flag && num - gapTolerance < 0.0001)
                {
                    result = true;
                }
            }
            return result;
        }
        #endregion

    }


    public class EqualityComparerPoint : EqualityComparer<Point>
    {
        public override bool Equals(Point x, Point y)
        {
            return Distance.PointToPoint(x, y) < 0.0001;
        }
        public override int GetHashCode(Point obj)
        {
            throw new NotImplementedException();
        }
    }

    public class EqualityComparerLineSegment : EqualityComparer<LineSegment>
    {
        public override bool Equals(LineSegment x, LineSegment y)
        {
            var result = false;
            if (Distance.PointToPoint(x.Point1, y.Point1) < 0.0001)
            {
                if (Distance.PointToPoint(x.Point2, y.Point2) < 0.0001)
                {
                    result = true;
                }
            }
            else if (Distance.PointToPoint(x.Point1, y.Point2) < 0.0001 && Distance.PointToPoint(x.Point2, y.Point1) < 0.0001)
            {
                result = true;
            }
            return result;
        }
        public override int GetHashCode(LineSegment obj)
        {
            throw new NotImplementedException();
        }
    }

    public class PointCompare : IComparer<Point>
    {
        public int Compare(Point x, Point y)
        {
            if (Math.Abs(x.X - y.X) < 0.0001)
            {
                if (Math.Abs(x.Y - y.Y) < 0.0001)
                {
                    return 0;
                }
                if (x.Y > y.Y)
                {
                    return 1;
                }
                if (x.Y < y.Y)
                {
                    return -1;
                }
            }
            else if (x.X > y.X)
            {
                return 1;
            }
            return -1;
        }
    }
}
