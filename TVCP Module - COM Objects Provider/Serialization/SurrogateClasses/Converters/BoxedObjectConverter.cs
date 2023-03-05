using System.Text.Json;
using System.Text.Json.Serialization;
using TVCP_Module___COM_Objects_Provider.NamedPipeConnection;

namespace TVCP_Module___COM_Objects_Provider.Serialization.SurrogateClasses.Converters
{
    public class BoxedObjectConverter : JsonConverter<BoxedObjectWrapper>
    {
        public override BoxedObjectWrapper? Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
        {
            var surrogate = JsonSerializer.Deserialize<BoxedObjectWrapperSurrogate>(ref reader, options);
            if (surrogate == null)
                return null;

            var boxedObj = JsonSerializer.Deserialize(surrogate.SerializedObject, Type.GetType(surrogate.TypeName, true, false)!, options);
            return new BoxedObjectWrapper(boxedObj!);
        }

        public override void Write(Utf8JsonWriter writer, BoxedObjectWrapper value, JsonSerializerOptions options)
        {
            var serializedObject = JsonSerializer.Serialize(value.BoxedObject, options);
            var surrogate = new BoxedObjectWrapperSurrogate(serializedObject, value.BoxedObject.GetType().FullName!);

            JsonSerializer.Serialize(writer, surrogate, options);
        }
    }
}
