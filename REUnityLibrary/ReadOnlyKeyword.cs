
namespace REUnityLibrary
{
    public class ReadOnlyKeyword
    {
        public Hyland.Unity.RecordType RecordType { get; set; }
        public string Name { get; set; }
        public string Value { get; set; }

        public ReadOnlyKeyword()
        {

        }

        public ReadOnlyKeyword(Hyland.Unity.RecordType recordType, string name, string value)
        {
            RecordType = recordType;
            Name = name;
            Value = value;
        }
    }

}
