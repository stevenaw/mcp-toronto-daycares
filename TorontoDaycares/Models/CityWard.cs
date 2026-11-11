using CsvHelper;

namespace TorontoDaycares.Models
{
    [CsvRecord]
    public partial record CityWard
    {
        [CsvColumn("AREA_NAME")]
        public string Name { get; init; }

        [CsvColumn("AREA_LONG_CODE")]
        public int Number { get; init; }
    }
}
