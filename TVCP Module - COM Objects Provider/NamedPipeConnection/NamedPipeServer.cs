using KLAKAUTLib;
using System.IO.Pipes;
using System.Runtime.InteropServices;
using System.Text.Json;
using TVCP_Module___COM_Objects_Provider.KSC;
using TVCP_Module___COM_Objects_Provider.Serialization.SurrogateClasses.Converters;
using TVCP_Module___COM_Objects_Provider.Utilities;

namespace TVCP_Module___COM_Objects_Provider.NamedPipeConnection
{
    public static class NamedPipeServer
    {
        private const string OkAnswer = "OK";
        private const string CreateObjectCommand = "CreateObject";
        private const string ExecuteCommand = "Execute";
        private const string GetPropertyCommand = "GetProperty";
        private const string ReturnObjectIDCommand = "ReturnObjectID";
        private const string ReturnBoxedObjectCommand = "ReturnBoxedObject";
        private const string SetPropertyCommand = "SetProperty";
        private const string ValueIsRegisteredObjectCommand = "ValueIsRegisteredObject";
        private const string ValueIsBoxedObjectCommand = "ValueIsBoxedObject";
        private const string DestroyCommand = "Destroy";

        private static string _pipeName = "TVCPComObjectsProviderPipe";
        private static JsonSerializerOptions _serializerOptions = new()
        {
            Converters =
            {
                new BoxedObjectConverter()
            }
        };
        private static Type[] _registeredTypesToCreate = { typeof(KlAkProxyClass), typeof(KlAkUsersClass), typeof(KlAkHostsClass),
                                                   typeof(KlAkParamsClass), typeof(KlAkCollectionClass) };

        public static async void StartConnectionThread(int maxAmountOfThreads = NamedPipeServerStream.MaxAllowedServerInstances)
        {
            await Task.Run(() =>
            {
                while (true)
                {
                    using (NamedPipeServerStream pipeServer = new(_pipeName, PipeDirection.InOut, maxAmountOfThreads))
                    {
                        pipeServer.WaitForConnection();

                        try
                        {
                            var streamString = new StreamString(pipeServer);

                            var query = streamString.ReadString();
                            ProcessQuery(streamString, query);
                        }
                        catch (Exception ex)
                        {
                            Task.Run(() =>
                            {
                                MessageBox.Show($"{ex.GetType().Name} - {ex.Message}",
                                    "Ошибка обмена данными с TVCenter Project",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                            });
                        }
                    }
                }
            });
        }

        private static void ProcessQuery(StreamString streamString, string query)
        {
            var args = query.Split(" ");

            try
            {
                if (!KSCProvider.CheckForObjectsAviability())
                    throw new COMException("KSC objects is not supported on this PC.");

                if (args[0] == CreateObjectCommand)
                {
                    CreateAndRegisterObject(args[1], args[2]);
                    streamString.WriteString(OkAnswer);

                    return;
                }

                if (args[0] == ExecuteCommand)
                {
                    if (args[1] == ReturnObjectIDCommand)
                    {
                        var serializedValueStartIndex = args[0].Length + args[1].Length + args[2].Length + args[3].Length + 4;
                        var serializedParameters = query[serializedValueStartIndex..].NullStringToNull();

                        var returnValueID = ExecuteMethodOfObjectWithIDAndReturnID(args[2], args[3], serializedParameters).NullToString();
                        streamString.WriteString($"{OkAnswer} {returnValueID}");
                    }
                    else if (args[1] == ReturnBoxedObjectCommand)
                    {
                        var serializedValueStartIndex = args[0].Length + args[1].Length + args[2].Length + args[3].Length + 4;
                        var serializedParameters = query[serializedValueStartIndex..].NullStringToNull();

                        var returnValue = ExecuteMethodOfObjectWithIDAndReturnBoxedObject(args[2], args[3], serializedParameters).NullToString();
                        streamString.WriteString($"{OkAnswer} {returnValue}");
                    }
                    else
                    {
                        var serializedValueStartIndex = args[0].Length + args[1].Length + args[2].Length + 3;
                        var serializedParameters = query[serializedValueStartIndex..].NullStringToNull();

                        var returnValue = ExecuteMethodOfObjectWithID(args[1], args[2], serializedParameters).NullToString();
                        streamString.WriteString($"{OkAnswer} {returnValue}");
                    }

                    return;
                }

                if (args[0] == GetPropertyCommand)
                {
                    if (args[1] == ReturnObjectIDCommand)
                    {
                        var propertyValueID = GetObjectIDFromProperty(args[2], args[3]).NullToString();
                        streamString.WriteString($"{OkAnswer} {propertyValueID}");
                    }
                    else if (args[1] == ReturnBoxedObjectCommand)
                    {
                        var propertyValue = GetBoxedObjectFromProperty(args[2], args[3]).NullToString();
                        streamString.WriteString($"{OkAnswer} {propertyValue}");
                    }
                    else
                    {
                        var propertyValue = GetValueFromProperty(args[1], args[2]).NullToString();
                        streamString.WriteString($"{OkAnswer} {propertyValue}");
                    }

                    return;
                }

                if (args[0] == SetPropertyCommand)
                {
                    if (args[1] == ValueIsRegisteredObjectCommand)
                    {
                        SetRegisteredObjectToProperty(args[2], args[3], args[4]);
                    }
                    else if (args[1] == ValueIsBoxedObjectCommand)
                    {
                        var serializedValueStartIndex = args[0].Length + args[1].Length + args[2].Length + args[3].Length + 4;
                        var serializedValue = query[serializedValueStartIndex..].NullStringToNull();

                        SetBoxedObjectToProperty(args[2], args[3], serializedValue);
                    }
                    else
                    {
                        var serializedValueStartIndex = args[0].Length + args[1].Length + args[2].Length + 3;
                        var serializedValue = query[serializedValueStartIndex..].NullStringToNull();

                        SetValueToProperty(args[1], args[2], serializedValue);
                    }

                    streamString.WriteString(OkAnswer);

                    return;
                }

                if (args[0] == DestroyCommand)
                {
                    UnregisterObject(args[1]);
                    streamString.WriteString(OkAnswer);

                    return;
                }

                throw new ArgumentException($"Unknown query: {query}");
            }
            catch (Exception ex)
            {
                streamString.WriteString($"{ex.GetType().Name} - {ex.Message}");
                throw;
            }
        }

        private static void CreateAndRegisterObject(string typeName, string id)
        {
            var createdObject = Activator.CreateInstance(
                FindTypeToCreateByName(typeName)
                ?? throw new ArgumentException($"Type {typeName} not found."))
            ?? throw new ArgumentException($"Failed to create object of type {typeName}.");

            var objectID = Guid.Parse(id);
            if (KSCProvider.ObjectWithIDExists(objectID))
                throw new InvalidOperationException($"Object with id {objectID} already exists (object type: {createdObject.GetType().Name}).");

            KSCProvider.RegisterObject(createdObject, objectID);
        }

        private static string? ExecuteMethodOfObjectWithID(string id, string methodName, string? serializedParameters)
        {
            var objectID = Guid.Parse(id);
            if (!KSCProvider.ObjectWithIDExists(objectID))
                throw new KeyNotFoundException($"Object with id {objectID} not found.");

            var obj = KSCProvider.GetObjectByID(objectID);

            var parameters = DeserializeParameters(serializedParameters);

            var method = obj.GetType().GetMethod(methodName);
            if (method == null)
                throw new InvalidOperationException($"Method with name {methodName} of object {obj.GetType().Name} does not exist.");

            var returnValue = method.Invoke(obj, parameters);

            if (returnValue == null || method.ReturnType == typeof(void))
                return null;

            return SerializeObject(returnValue);
        }

        private static string? ExecuteMethodOfObjectWithIDAndReturnID(string id, string methodName, string? serializedParameters)
        {
            var objectID = Guid.Parse(id);
            if (!KSCProvider.ObjectWithIDExists(objectID))
                throw new KeyNotFoundException($"Object with id {objectID} not found.");

            var obj = KSCProvider.GetObjectByID(objectID);

            var parameters = DeserializeParameters(serializedParameters);

            var method = obj.GetType().GetMethod(methodName);
            if (method == null)
                throw new InvalidOperationException($"Method with name {methodName} of object {obj.GetType().Name} does not exist.");

            var returnValue = method.Invoke(obj, parameters);

            if (returnValue == null || method.ReturnType == typeof(void))
                return null;

            var propertyObjectValueID = ExistingOrNewID(obj);

            return propertyObjectValueID.ToString();
        }

        private static string? ExecuteMethodOfObjectWithIDAndReturnBoxedObject(string id, string methodName, string? serializedParameters)
        {
            var objectID = Guid.Parse(id);
            if (!KSCProvider.ObjectWithIDExists(objectID))
                throw new KeyNotFoundException($"Object with id {objectID} not found.");

            var obj = KSCProvider.GetObjectByID(objectID);

            var parameters = DeserializeParameters(serializedParameters);

            var method = obj.GetType().GetMethod(methodName);
            if (method == null)
                throw new InvalidOperationException($"Method with name {methodName} of object {obj.GetType().Name} does not exist.");

            var returnValue = method.Invoke(obj, parameters);

            if (returnValue == null || method.ReturnType == typeof(void))
                return null;

            return SerializeObject(new BoxedObjectWrapper(returnValue));
        }

        private static object[]? DeserializeParameters(string? serializedParameters)
        {
            if (serializedParameters == null)
                return null;

            var parametersContainer = DeserializeObject<ParametersDataContainer>(serializedParameters);

            if (parametersContainer == null)
                return null;

            var containsNums = parametersContainer.NumParameters != null && parametersContainer.NumParameters.Count > 0;
            var containsStrings = parametersContainer.StringParameters != null && parametersContainer.StringParameters.Count > 0;
            var containsBoolValues = parametersContainer.BoolParameters != null && parametersContainer.BoolParameters.Count > 0;
            var containsKlAkObjects = parametersContainer.KlAkObjectParameters != null && parametersContainer.KlAkObjectParameters.Count > 0;
            var containsBoxedObjects = parametersContainer.BoxedObjectWrappers != null && parametersContainer.BoxedObjectWrappers.Count > 0;

            int count = 0;
            if (containsNums)
                count += parametersContainer.NumParameters!.Count;
            if (containsStrings)
                count += parametersContainer.StringParameters!.Count;
            if (containsBoolValues)
                count += parametersContainer.BoolParameters!.Count;
            if (containsKlAkObjects)
                count += parametersContainer.KlAkObjectParameters!.Count;
            if (containsBoxedObjects)
                count += parametersContainer.BoxedObjectWrappers!.Count;

            object[] parameters = new object[count];

            if (containsStrings)
                foreach (var stringParameter in parametersContainer.StringParameters!)
                    parameters[stringParameter.Key] = stringParameter.Value;

            if (containsNums)
                foreach (var numParameter in parametersContainer.NumParameters!)
                    parameters[numParameter.Key] = numParameter.Value;

            if (containsBoolValues)
                foreach (var boolParameter in parametersContainer.BoolParameters!)
                    parameters[boolParameter.Key] = boolParameter.Value;

            if (containsKlAkObjects)
                foreach (var klAkObjectParameter in parametersContainer.KlAkObjectParameters!)
                {
                    var objectID = klAkObjectParameter.Value;
                    if (!KSCProvider.ObjectWithIDExists(objectID))
                        throw new KeyNotFoundException($"Object with id {objectID} not found.");

                    parameters[klAkObjectParameter.Key] = KSCProvider.GetObjectByID(objectID);
                }

            if (containsBoxedObjects)
                foreach (var boxedObjectWrapper in parametersContainer.BoxedObjectWrappers!)
                    parameters[boxedObjectWrapper.Key] = boxedObjectWrapper.Value.BoxedObject;

            return parameters;
        }

        private static string? GetValueFromProperty(string id, string propertyName)
        {
            var objectID = Guid.Parse(id);
            if (!KSCProvider.ObjectWithIDExists(objectID))
                throw new KeyNotFoundException($"Object with id {objectID} not found.");

            var obj = KSCProvider.GetObjectByID(objectID);

            var property = obj.GetType().GetProperty(propertyName);
            if (property == null)
                throw new InvalidOperationException($"Property with name {propertyName} of object {obj.GetType().Name} does not exist.");

            var propertyValue = property.GetValue(obj);
            if (propertyValue == null)
                return null;

            return SerializeObject(propertyValue);
        }

        private static string? GetBoxedObjectFromProperty(string id, string propertyName)
        {
            var objectID = Guid.Parse(id);
            if (!KSCProvider.ObjectWithIDExists(objectID))
                throw new KeyNotFoundException($"Object with id {objectID} not found.");

            var obj = KSCProvider.GetObjectByID(objectID);

            var property = obj.GetType().GetProperty(propertyName);
            if (property == null)
                throw new InvalidOperationException($"Property with name {propertyName} of object {obj.GetType().Name} does not exist.");

            var propertyValue = property.GetValue(obj);
            if (propertyValue == null)
                return null;

            return SerializeObject(new BoxedObjectWrapper(propertyValue));
        }

        private static string? GetObjectIDFromProperty(string id, string propertyName)
        {
            var objectID = Guid.Parse(id);
            if (!KSCProvider.ObjectWithIDExists(objectID))
                throw new KeyNotFoundException($"Object with id {objectID} not found.");

            var obj = KSCProvider.GetObjectByID(objectID);

            var property = obj.GetType().GetProperty(propertyName);
            if (property == null)
                throw new InvalidOperationException($"Property with name {propertyName} of object {obj.GetType().Name} does not exist.");

            var propertyValue = property.GetValue(obj);
            if (propertyValue == null)
                return null;

            var propertyObjectValueID = ExistingOrNewID(obj);

            return propertyObjectValueID.ToString();
        }

        private static void SetValueToProperty(string id, string propertyName, string? serializedValue)
        {
            var objectID = Guid.Parse(id);
            if (!KSCProvider.ObjectWithIDExists(objectID))
                throw new KeyNotFoundException($"Object with id {objectID} not found.");

            var obj = KSCProvider.GetObjectByID(objectID);

            var property = obj.GetType().GetProperty(propertyName);
            if (property == null)
                throw new InvalidOperationException($"Property with name {propertyName} of object {obj.GetType().Name} does not exist.");

            var value = serializedValue == null ? null : DeserializeObject(serializedValue, property.PropertyType);
            property.SetValue(obj, value);
        }

        private static void SetBoxedObjectToProperty(string id, string propertyName, string? serializedBoxedObject)
        {
            var objectID = Guid.Parse(id);
            if (!KSCProvider.ObjectWithIDExists(objectID))
                throw new KeyNotFoundException($"Object with id {objectID} not found.");

            var obj = KSCProvider.GetObjectByID(objectID);

            var property = obj.GetType().GetProperty(propertyName);
            if (property == null)
                throw new InvalidOperationException($"Property with name {propertyName} of object {obj.GetType().Name} does not exist.");

            var value = serializedBoxedObject == null ? null : DeserializeObject<BoxedObjectWrapper>(serializedBoxedObject);
            property.SetValue(obj, value?.BoxedObject);
        }

        private static void SetRegisteredObjectToProperty(string id, string propertyName, string idOfObjectToSet)
        {
            var objectID = Guid.Parse(id);
            if (!KSCProvider.ObjectWithIDExists(objectID))
                throw new KeyNotFoundException($"Object with id {objectID} not found.");
            var objectToSetID = Guid.Parse(idOfObjectToSet);
            if (!KSCProvider.ObjectWithIDExists(objectToSetID))
                throw new KeyNotFoundException($"Object with id {objectToSetID} not found.");

            var obj = KSCProvider.GetObjectByID(objectID);
            var objToSet = KSCProvider.GetObjectByID(objectToSetID);

            var property = obj.GetType().GetProperty(propertyName);
            if (property == null)
                throw new InvalidOperationException($"Property with name {propertyName} of object {obj.GetType().Name} does not exist.");

            property.SetValue(obj, objToSet);
        }

        private static void UnregisterObject(string id)
        {
            var objectID = Guid.Parse(id);
            if (!KSCProvider.ObjectWithIDExists(objectID))
                throw new KeyNotFoundException($"Object with id {objectID} not found.");

            KSCProvider.UnregisterObject(objectID);
        }

        private static string SerializeObject(object obj)
            => JsonSerializer.Serialize(obj, _serializerOptions);

        private static object? DeserializeObject(string serializedObject, Type objectType)
            => JsonSerializer.Deserialize(serializedObject, objectType, _serializerOptions);

        private static T? DeserializeObject<T>(string serializedObject)
            => JsonSerializer.Deserialize<T>(serializedObject, _serializerOptions);

        private static Guid ExistingOrNewID(object obj)
        {
            if (KSCProvider.ObjectExists(obj))
            {
                return KSCProvider.GetObjectID(obj);
            }
            else
            {
                var newID = Guid.NewGuid();
                KSCProvider.RegisterObject(obj, newID);

                return newID;
            }
        }

        private static Type? FindTypeToCreateByName(string name)
        {
            try
            {
                return Type.GetType(name, true, false);
            }
            catch
            {
                if (!name.EndsWith("Class"))
                    name += "Class";

                if (!Array.Exists(_registeredTypesToCreate, (type) => type.Name == name))
                    return null;

                return Array.Find(_registeredTypesToCreate, (type) => type.Name == name)!;
            }

        }
    }
}
