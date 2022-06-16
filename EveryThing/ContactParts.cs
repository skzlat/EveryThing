using System.Collections.Generic;
using Tekla.Structures.Model;

namespace EveryThing
{
    public static class ContactParts
    {
        // Получение соприкасающихся поверхностей двух деталей
        public static List<ContactPlane> GetContactPlanesOfParts(Part part1, Part part2, double gapTolerances)
        {
            var contactPlanes = PartUtils.GetContactPlanes(part1, part2, gapTolerances);
            return contactPlanes;
        }
    }
}
