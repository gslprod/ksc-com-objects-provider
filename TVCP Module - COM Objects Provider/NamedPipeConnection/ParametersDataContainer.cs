using TVCP_Module___COM_Objects_Provider.Serialization.SurrogateClasses;

namespace TVCP_Module___COM_Objects_Provider.NamedPipeConnection
{
    public class ParametersDataContainer
    {
        public Dictionary<int, string>? StringParameters;
        public Dictionary<int, long>? NumParameters;
        public Dictionary<int, bool>? BoolParameters;
        public Dictionary<int, Guid>? KlAkObjectParameters;
        public Dictionary<int, BoxedObjectWrapper>? BoxedObjectWrappers;
    }
}
