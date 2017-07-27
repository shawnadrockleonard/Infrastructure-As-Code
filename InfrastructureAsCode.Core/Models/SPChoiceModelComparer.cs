using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Models
{
    /// <summary>
    /// Provides a model comparer to evaluate the choice model property
    /// </summary>
    public class SPChoiceModelComparer : IEqualityComparer<SPChoiceModel>
    {
        public bool Equals(SPChoiceModel x, SPChoiceModel y)
        {

            //Check whether the compared objects reference the same data.
            if (Object.ReferenceEquals(x, y)) return true;

            //Check whether any of the compared objects is null.
            if (Object.ReferenceEquals(x, null) || Object.ReferenceEquals(y, null))
                return false;

            //Check whether the products' properties are equal.
            return x.Choice == y.Choice && x.DefaultChoice == y.DefaultChoice;
        }

        // If Equals() returns true for a pair of objects 
        // then GetHashCode() must return the same value for these objects.

        public int GetHashCode(SPChoiceModel item)
        {
            //Check whether the object is null
            if (Object.ReferenceEquals(item, null)) return 0;

            //Get hash code for the Name field if it is not null.
            int hashName = item.Choice == null ? 0 : item.Choice.GetHashCode();

            //Get hash code for the Code field.
            int hashCode = item.DefaultChoice.GetHashCode();

            //Calculate the hash code for the product.
            return hashName ^ hashCode;
        }

    }
}