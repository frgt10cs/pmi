using Newtonsoft.Json;
using Pmi.Service.Interface;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Security.Cryptography;

namespace Pmi.Service.Implimentation
{   
    class JsonCacheService<T> : ICacheService<T>
    {
        private string filePath;
        private List<T> elements;

        public JsonCacheService(string filePath)
        {
            if(File.Exists(filePath))
            {
                this.filePath = filePath;
            }
            else
            {
                
            }
        }

        public void UploadCache()
        {
            elements = JsonConvert.DeserializeObject<List<T>>(File.ReadAllText(filePath));
        }

        public void Add(T item)
        {
            elements.Add(item);
        }

        public void SetAll(List<T> items)
        {
            elements = items;
        }

        public List<T> GetAll()
        {                      
            return elements;
        }

        public void SaveChanges()
        {
            string json = JsonConvert.SerializeObject(elements);
            File.WriteAllText(filePath, json);
        }
        
        private string GetHash(string value)
        {
            using (var md5 = MD5.Create())
            {
                return Encoding.UTF8.GetString(md5.ComputeHash(Encoding.UTF8.GetBytes(value)));
            }
        }
    }
}
