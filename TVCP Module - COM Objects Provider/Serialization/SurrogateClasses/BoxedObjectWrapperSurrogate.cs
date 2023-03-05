namespace TVCP_Module___COM_Objects_Provider.Serialization.SurrogateClasses
{
    public class BoxedObjectWrapperSurrogate
    {
        public string SerializedObject { get; private set; }
        public string TypeName { get; private set; }

        public BoxedObjectWrapperSurrogate(string serializedObject, string typeName)
        {
            SerializedObject = serializedObject;
            TypeName = typeName;
        }
    }
}
