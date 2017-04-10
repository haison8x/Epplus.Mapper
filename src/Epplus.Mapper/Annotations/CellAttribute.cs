using System;

namespace Epplus.Mapper.Annotations
{
    [AttributeUsage(AttributeTargets.Property)]
    public class CellAttribute : Attribute
    {
        public CellAttribute(string address)
        {
            Address = address;
        }

        public string Address { get; }
    }
}