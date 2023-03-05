namespace TVCP_Module___COM_Objects_Provider.NamedPipeConnection
{
    public class BoxedObjectWrapper
    {
        public object BoxedObject { get; private set; }

        public BoxedObjectWrapper(object boxedObject)
        {
            BoxedObject = boxedObject;
        }
    }
}
