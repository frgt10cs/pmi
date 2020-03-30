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

        public override T UploadCache()
        {
            T entity = JsonConvert.DeserializeObject<T>(File.ReadAllText(filePath));
            return entity;
        }                             
    
        public override void Cache(T entity)
        {            
            string json = JsonConvert.SerializeObject(entity);
            File.WriteAllText(filePath, json);
        }                
    }
}
