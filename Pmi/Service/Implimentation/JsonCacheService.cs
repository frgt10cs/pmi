using Newtonsoft.Json;
using System.Collections.Generic;
using System.IO;
using Pmi.Service.Abstraction;
using DocumentFormat.OpenXml.Spreadsheet;
using System;

namespace Pmi.Service.Implimentation
{
    class JsonCacheService<T> : CacheService<T>
    {
        public JsonCacheService(string path) : base(path)
        {

        }

        public override T UploadCache()
        {
            T entity = JsonConvert.DeserializeObject<T>(File.ReadAllText(filePath));
            return entity;
        }

        public override void Cache(T entity)
        {
            string json = JsonConvert.SerializeObject(entity, new FontConvertor());
            File.WriteAllText(filePath, json);
        }
    }

    public class FontConvertor : JsonConverter<Font>
    {
        public override Font ReadJson(JsonReader reader, Type objectType, Font existingValue, bool hasExistingValue, JsonSerializer serializer)
        {
            Pmi.Builders.ExcelFontBuilder builder = new Builders.ExcelFontBuilder();
            builder.SetFontName(reader.ReadAsString());
            builder.SetFontSize((int)reader.ReadAsInt32());
            builder.SetColor(reader.ReadAsString());
            if (reader.ReadAsString() != null)
                builder.AddBold();
            if (reader.ReadAsString() != null)
                builder.AddItalic();
            return builder.GetFont();
        }

        public override void WriteJson(JsonWriter writer, Font value, JsonSerializer serializer)
        {
            writer.WriteStartObject();
            writer.WritePropertyName("FontName");
            writer.WriteValue(value.FontName.Val.Value);
            writer.WritePropertyName("FontSize");
            writer.WriteValue(value.FontSize.Val.Value);
            writer.WritePropertyName("Color");
            writer.WriteValue(value.Color?.Rgb.Value);
            writer.WritePropertyName("Bold");
            writer.WriteValue(value.Bold?.Val?.Value);
            writer.WritePropertyName("Italic");
            writer.WriteValue(value.Italic?.Val?.Value);
            writer.WriteEndObject();
        }
    }

    public class FontsConvertor : JsonConverter<List<Font>>
    {
        public override List<Font> ReadJson(JsonReader reader, Type objectType, List<Font> existingValue, bool hasExistingValue, JsonSerializer serializer)
        {
            throw new NotImplementedException();
        }

        public override void WriteJson(JsonWriter writer, List<Font> value, JsonSerializer serializer)
        {
            throw new NotImplementedException();
        }
    }
}
