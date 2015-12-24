using System.Collections.ObjectModel;

//no radar models map directly to this structure. 
namespace ImportPOC2
{
    public class ProductColorGroup
    {
        public string Code { get; set; }
        public string Description { get; set; }
        public string DisplayName { get; set; }
        public Collection<ColorGroup> CodeValueGroups { get; set; }
    }

    public class ColorGroup
    {
        //public string Code { get; set; }
        public string Description { get; set; }
        public string DisplayName { get; set; }
        //public bool IsDefault { get; set; }
        //public int DisplaySequence { get; set; }
        //public string MajorCodeValueGroupCode { get; set; }
        //public Hue ColorHue { get; set; }
        public Collection<SetCode> SetCodeValues { get; set; }
    }

    //public class Hue
    //{
    //    public string Code { get; set; }
    //    public string Description { get; set; }
    //    public string DisplayName { get; set; }
    //}

    public class SetCode
    {
        public long Id { get; set; }
        public string CodeValue { get; set; }
        //public int DisplaySequence { get; set; }
    }
}
