using System;

namespace Epplus.Mapper.Annotations
{
    [AttributeUsage(AttributeTargets.Property)]
    public class CellShiftAttribute : Attribute
    {
        public CellShiftAttribute(string address)
        {
            Address = address;
        }

        public string Address { get; }
    }
}
