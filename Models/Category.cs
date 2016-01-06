
namespace ImportPOC2
{
    public class Category
    {
        public string Code { get; set; }
        public string Description { get; set; }
        public string Name { get; set; }
        public string DisplayName { get; set; }
        public string ParentCategoryCode { get; set; }
        public string ProductTypeCode { get; set; }
        public bool IsProductTypeSpecific { get; set; }
        public string IsAllowsAssign { get; set; }
        public string IsParent { get; set; }   
        public string IsPrimary { get; set; }
    }
}
