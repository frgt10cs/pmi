using Newtonsoft.Json;
using System.Collections.Generic;
using System.IO;
using Pmi.Service.Abstraction;

namespace Pmi.Service.Implimentation
{       
    class JsonCacheService<T> : CacheService<T>
    {
        public JsonCacheService(string path):base(path)
        {

        }

        public override void  UploadCache()
        {
            elements = JsonConvert.DeserializeObject<List<T>>(File.ReadAllText(filePath));
        }                             
    
        public override void SaveChanges()
        {
            string json = JsonConvert.SerializeObject(elements);
            File.WriteAllText(filePath, json);
        }                
    }
}
