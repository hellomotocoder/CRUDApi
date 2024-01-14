namespace CRUDApi.Models
{
    public class SearchModel
    {
        public string FieldName { get; set; }
        public string FieldValue { get; set; }
        public int? PageNumber { get; set; } // Nullable int to handle null case if we dont ad
        public string SortBy { get; set; }
    }
}
