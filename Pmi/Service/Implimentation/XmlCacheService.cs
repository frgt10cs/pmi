using Pmi.Service.Abstraction;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Xml.Serialization;

namespace Pmi.Service.Implimentation
{
    class XmlCacheService<T> : CacheService<T>
    {
        private readonly XmlSerializer formatter;

        public XmlCacheService(string filePath):base(filePath)
        {
             formatter = new XmlSerializer(typeof(T));            
        }

        public override void Cache(T entity)
        {            
            using (var fs = new FileStream(filePath, FileMode.OpenOrCreate))
            {
                formatter.Serialize(fs, entity);                
            }
        }

        public override T UploadCache()
        {
            using (var fs = new FileStream(filePath, FileMode.OpenOrCreate))
            {
                return (T)formatter.Deserialize(fs);
            }
        }
    }
}
