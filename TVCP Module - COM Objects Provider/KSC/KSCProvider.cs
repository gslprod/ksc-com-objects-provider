using KLAKAUTLib;

namespace TVCP_Module___COM_Objects_Provider.KSC
{
    internal static class KSCProvider
    {
        private static readonly Dictionary<Guid, object> _registeredObjects = new();

        public static bool CheckForObjectsAviability()
        {
            try
            {
                _ = new KlAkProxy();

                return true;
            }
            catch
            {
                return false;
            }
        }

        public static void RegisterObject(object obj, Guid guid)
            => _registeredObjects.Add(guid, obj);

        public static object GetObjectByID(Guid id)
            => _registeredObjects[id];

        public static void UnregisterObject(Guid id)
            => _registeredObjects.Remove(id);

        public static bool ObjectWithIDExists(Guid id)
            => _registeredObjects.ContainsKey(id);

        public static bool ObjectExists(object obj)
            => _registeredObjects.ContainsValue(obj);

        public static Guid GetObjectID(object obj)
            => _registeredObjects.First((pair) => pair.Value == obj).Key;
    }
}
